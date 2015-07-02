#include "StdAfx.h"
#include "AtlImage.h"

using namespace ATL;

CImage::CInitGDIPlus CImage::s_initGDIPlus;
CImage::CDCCache CImage::s_cache;

const DWORD CImage::createAlphaChannel = 0x01;

const DWORD CImage::excludeGIF   = 0x01;
const DWORD CImage::excludeBMP   = 0x02;
const DWORD CImage::excludeEMF   = 0x04;
const DWORD CImage::excludeWMF   = 0x08;
const DWORD CImage::excludeJPEG  = 0x10;
const DWORD CImage::excludePNG   = 0x20;
const DWORD CImage::excludeTIFF  = 0x40;
const DWORD CImage::excludeIcon  = 0x80;
const DWORD CImage::excludeOther = 0x80000000;
const DWORD CImage::excludeValid = 0x800000ff;
const DWORD CImage::excludeDefaultLoad = 0;
const DWORD CImage::excludeDefaultSave = CImage::excludeIcon|CImage::excludeEMF|CImage::excludeWMF;

CImageDC::CImageDC(const CImage& image) throw()
: m_image(image)
, m_hDC(image.GetDC())
{
	ATLASSERT(m_hDC != NULL);
}

CImageDC::~CImageDC() throw()
{
	m_image.ReleaseDC();
}

CImageDC::operator HDC() const throw()
{
	return(m_hDC);
}

CImage::CInitGDIPlus::CInitGDIPlus() throw()
: m_dwToken(0)
, m_nCImageObjects(0)
{
	__try
	{
		InitializeCriticalSection(&m_sect);
	}
	__except(STATUS_NO_MEMORY == GetExceptionCode())
	{
		ATLASSERT(FALSE);
	}
}

CImage::CInitGDIPlus::~CInitGDIPlus() throw()
{
	ReleaseGDIPlus();
	DeleteCriticalSection(&m_sect);
}

bool CImage::CInitGDIPlus::Init() throw()
{
	EnterCriticalSection(&m_sect);
	bool fRet = true;
	if (m_dwToken == 0)
	{
		Gdiplus::GdiplusStartupInput input;
		Gdiplus::GdiplusStartupOutput output;
		Gdiplus::Status status = Gdiplus::GdiplusStartup(&m_dwToken, &input, &output);
		if(status != Gdiplus::Ok)
			fRet = false;
	}
	LeaveCriticalSection(&m_sect);
	return fRet;
}

void CImage::CInitGDIPlus::ReleaseGDIPlus() throw()
{
	EnterCriticalSection(&m_sect);
	if(m_dwToken != 0)
	{
		Gdiplus::GdiplusShutdown(m_dwToken);
	}
	m_dwToken = 0;
	LeaveCriticalSection(&m_sect);
}

void CImage::CInitGDIPlus::IncreaseCImageCount() throw()
{
	EnterCriticalSection(&m_sect);
	m_nCImageObjects++;
	LeaveCriticalSection(&m_sect);
}

void CImage::CInitGDIPlus::DecreaseCImageCount() throw()
{
	EnterCriticalSection(&m_sect);
	if (--m_nCImageObjects == 0)
		ReleaseGDIPlus();
	LeaveCriticalSection(&m_sect);
}

CImage::CDCCache::CDCCache() throw()
{
	int iDC(0);
	
	for(iDC = 0; iDC < CIMAGE_DC_CACHE_SIZE; iDC++)
	{
		m_ahDCs[iDC] = NULL;
	}
}

CImage::CDCCache::~CDCCache() throw()
{
	int iDC(0);
	
	for(iDC = 0; iDC < CIMAGE_DC_CACHE_SIZE; iDC++)
	{
		if(m_ahDCs[iDC] != NULL)
		{
			::DeleteDC(m_ahDCs[iDC]);
		}
	}
}

HDC CImage::CDCCache::GetDC() throw()
{
	HDC hDC(NULL);
	
	for(int iDC = 0; iDC < CIMAGE_DC_CACHE_SIZE; iDC++)
	{
		hDC = (HDC)(InterlockedExchange(reinterpret_cast<PLONG>(&m_ahDCs[iDC]), NULL));
		if(hDC != NULL)
		{
			return(hDC);
		}
	}
	
	hDC = ::CreateCompatibleDC(NULL);
	
	return(hDC);
}

void CImage::CDCCache::ReleaseDC(HDC hDC) throw()
{
	for(int iDC = 0; iDC < CIMAGE_DC_CACHE_SIZE; iDC++)
	{
		HDC hOldDC;
		
		hOldDC = (HDC)(InterlockedExchange(reinterpret_cast<PLONG>(&m_ahDCs[iDC]), (LONG)hDC));
		if(hOldDC == NULL)
		{
			return;
		}
		else
		{
			hDC = hOldDC;
		}
	}
	if(hDC != NULL)
	{
		::DeleteDC(hDC);
	}
}

CImage::CImage() throw():
m_hBitmap(NULL),
m_pBits(NULL),
m_hDC(NULL),
m_nDCRefCount(0),
m_hOldBitmap(NULL),
m_nWidth(0),
m_nHeight(0),
m_nPitch(0),
m_nBPP(0),
m_iTransparentColor(-1),
m_bHasAlphaChannel(false),
m_bIsDIBSection(false)
{
	s_initGDIPlus.IncreaseCImageCount();
}

CImage::~CImage() throw()
{
	Destroy();
	s_initGDIPlus.DecreaseCImageCount();
}

CImage::operator HBITMAP() const throw()
{
	return(m_hBitmap);
}

#if WINVER >= 0x0500
BOOL CImage::AlphaBlend(HDC hDestDC, int xDest, int yDest, 
							   BYTE bSrcAlpha, BYTE bBlendOp) const throw()
{
	return(AlphaBlend(hDestDC, xDest, yDest, m_nWidth, m_nHeight, 0, 0, 
		m_nWidth, m_nHeight, bSrcAlpha, bBlendOp));
}

BOOL CImage::AlphaBlend(HDC hDestDC, const POINT& pointDest, 
							   BYTE bSrcAlpha, BYTE bBlendOp) const throw()
{
	return(AlphaBlend(hDestDC, pointDest.x, pointDest.y, m_nWidth, m_nHeight, 
		0, 0, m_nWidth, m_nHeight, bSrcAlpha, bBlendOp));
}

BOOL CImage::AlphaBlend(HDC hDestDC, int xDest, int yDest, 
							   int nDestWidth, int nDestHeight, int xSrc, int ySrc, int nSrcWidth, 
							   int nSrcHeight, BYTE bSrcAlpha, BYTE bBlendOp) const throw()
{
	BLENDFUNCTION blend;
	BOOL bResult;
	
	blend.SourceConstantAlpha = bSrcAlpha;
	blend.BlendOp = bBlendOp;
	blend.BlendFlags = 0;
	if(m_bHasAlphaChannel)
	{
		blend.AlphaFormat = 1;
	}
	else
	{
		blend.AlphaFormat = 0;
	}
	
	GetDC();
	
	bResult = ::AlphaBlend(hDestDC, xDest, yDest, nDestWidth, nDestHeight, m_hDC, 
		xSrc, ySrc, nSrcWidth, nSrcHeight, blend);
	
	ReleaseDC();
	
	return(bResult);
}

BOOL CImage::AlphaBlend(HDC hDestDC, const RECT& rectDest, 
							   const RECT& rectSrc, BYTE bSrcAlpha, BYTE bBlendOp) const throw()
{
	return(AlphaBlend(hDestDC, rectDest.left, rectDest.top, rectDest.right-
		rectDest.left, rectDest.bottom-rectDest.top, rectSrc.left, rectSrc.top, 
		rectSrc.right-rectSrc.left, rectSrc.bottom-rectSrc.top, bSrcAlpha, 
		bBlendOp));
}
#endif

void CImage::Attach(HBITMAP hBitmap, DIBOrientation eOrientation) throw()
{
	ATLASSERT(m_hBitmap == NULL);
	ATLASSERT(hBitmap != NULL);
	
	m_hBitmap = hBitmap;
	
	UpdateBitmapInfo(eOrientation);
}

BOOL CImage::BitBlt(HDC hDestDC, int xDest, int yDest, DWORD dwROP) const throw()
{
	return(BitBlt(hDestDC, xDest, yDest, m_nWidth, m_nHeight, 0, 0, dwROP));
}

BOOL CImage::BitBlt(HDC hDestDC, const POINT& pointDest, DWORD dwROP) const throw()
{
	return(BitBlt(hDestDC, pointDest.x, pointDest.y, m_nWidth, m_nHeight, 
		0, 0, dwROP));
}

BOOL CImage::BitBlt(HDC hDestDC, int xDest, int yDest, int nDestWidth, 
						   int nDestHeight, int xSrc, int ySrc, DWORD dwROP) const throw()
{
	BOOL bResult;
	
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(hDestDC != NULL);
	
	GetDC();
	
	bResult = ::BitBlt(hDestDC, xDest, yDest, nDestWidth, nDestHeight, 
		m_hDC, xSrc, ySrc, dwROP);
	
	ReleaseDC();
	
	return(bResult);
}

BOOL CImage::BitBlt(HDC hDestDC, const RECT& rectDest, 
						   const POINT& pointSrc, DWORD dwROP) const throw()
{
	return(BitBlt(hDestDC, rectDest.left, rectDest.top, rectDest.right-
		rectDest.left, rectDest.bottom-rectDest.top, pointSrc.x, pointSrc.y, 
		dwROP));
}

BOOL CImage::Create(int nWidth, int nHeight, int nBPP, DWORD dwFlags) throw()
{
	return(CreateEx(nWidth, nHeight, nBPP, BI_RGB, NULL, dwFlags));
}

BOOL CImage::CreateEx(int nWidth, int nHeight, int nBPP, DWORD eCompression, 
							 const DWORD* pdwBitfields, DWORD dwFlags) throw()
{
	LPBITMAPINFO pbmi = NULL;
	HBITMAP hBitmap = NULL;
	
	ATLASSERT((eCompression == BI_RGB) || (eCompression == BI_BITFIELDS));
	if(dwFlags&createAlphaChannel)
	{
		ATLASSERT((nBPP == 32) && (eCompression == BI_RGB));
	}
	
	pbmi = (LPBITMAPINFO)malloc(sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD)*256);
	if(pbmi == NULL)
		return FALSE;
	
	memset(&pbmi->bmiHeader, 0, sizeof(pbmi->bmiHeader));
	pbmi->bmiHeader.biSize = sizeof(pbmi->bmiHeader);
	pbmi->bmiHeader.biWidth = nWidth;
	pbmi->bmiHeader.biHeight = nHeight;
	pbmi->bmiHeader.biPlanes = 1;
	pbmi->bmiHeader.biBitCount = USHORT(nBPP);
	pbmi->bmiHeader.biCompression = eCompression;
	if(nBPP <= 8)
	{
		ATLASSERT(eCompression == BI_RGB);
#pragma warning(push)
#pragma warning(disable: 4068) //Disable unknown pragma warning that prefast pragma causes.
#pragma prefast(push)
#pragma prefast(disable: 203, "no buffer overrun here, buffer was alocated properly") 
		memset(pbmi->bmiColors, 0, 256 * sizeof(RGBQUAD));
#pragma prefast(pop)
#pragma warning(pop)
	}
	else 
	{
		if(eCompression == BI_BITFIELDS)
		{
			ATLASSERT(pdwBitfields != NULL);
			memcpy(pbmi->bmiColors, pdwBitfields, 3 * sizeof(DWORD));
		}
	}
	
	hBitmap = ::CreateDIBSection(NULL, pbmi, DIB_RGB_COLORS, &m_pBits, NULL, 0);
	if(hBitmap == NULL)
	{
		return(FALSE);
	}
	
	Attach(hBitmap, (nHeight < 0) ? DIBOR_TOPDOWN : DIBOR_BOTTOMUP);
	
	if(dwFlags&createAlphaChannel)
	{
		m_bHasAlphaChannel = true;
	}
	free(pbmi);
	return(TRUE);
}

void CImage::Destroy() throw()
{
	HBITMAP hBitmap;
	
	if(m_hBitmap != NULL)
	{
		hBitmap = Detach();
		::DeleteObject(hBitmap);
	}
}

HBITMAP CImage::Detach() throw()
{
	HBITMAP hBitmap;
	
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(m_hDC == NULL);
	
	hBitmap = m_hBitmap;
	m_hBitmap = NULL;
	m_pBits = NULL;
	m_nWidth = 0;
	m_nHeight = 0;
	m_nBPP = 0;
	m_nPitch = 0;
	m_iTransparentColor = -1;
	m_bHasAlphaChannel = false;
	m_bIsDIBSection = false;
	
	return(hBitmap);
}

BOOL CImage::Draw(HDC hDestDC, const RECT& rectDest) const throw()
{
	return(Draw(hDestDC, rectDest.left, rectDest.top, rectDest.right-
		rectDest.left, rectDest.bottom-rectDest.top, 0, 0, m_nWidth, 
		m_nHeight));
}

BOOL CImage::Draw(HDC hDestDC, int xDest, int yDest, int nDestWidth, int nDestHeight) const throw()
{
	return(Draw(hDestDC, xDest, yDest, nDestWidth, nDestHeight, 0, 0, m_nWidth, m_nHeight));
}

BOOL CImage::Draw(HDC hDestDC, const POINT& pointDest) const throw()
{
	return(Draw(hDestDC, pointDest.x, pointDest.y, m_nWidth, m_nHeight, 0, 0, m_nWidth, m_nHeight));
}

BOOL CImage::Draw(HDC hDestDC, int xDest, int yDest) const throw()
{
	return(Draw(hDestDC, xDest, yDest, m_nWidth, m_nHeight, 0, 0, m_nWidth, m_nHeight));
}

BOOL CImage::Draw(HDC hDestDC, const RECT& rectDest, const RECT& rectSrc) const throw()
{
	return(Draw(hDestDC, rectDest.left, rectDest.top, rectDest.right-
		rectDest.left, rectDest.bottom-rectDest.top, rectSrc.left, rectSrc.top, 
		rectSrc.right-rectSrc.left, rectSrc.bottom-rectSrc.top));
}

BOOL CImage::Draw(HDC hDestDC, int xDest, int yDest, int nDestWidth,
						 int nDestHeight, int xSrc, int ySrc, int nSrcWidth, int nSrcHeight) const throw()
{
	BOOL bResult;
	
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(hDestDC != NULL);
	ATLASSERT(nDestWidth > 0);
	ATLASSERT(nDestHeight > 0);
	ATLASSERT(nSrcWidth > 0);
	ATLASSERT(nSrcHeight > 0);
	
	GetDC();
	
#if WINVER >= 0x0500
	if((m_iTransparentColor != -1) && IsTransparencySupported())
	{
		bResult = ::TransparentBlt(hDestDC, xDest, yDest, nDestWidth, nDestHeight,
			m_hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, GetTransparentRGB());
	}
	else if(m_bHasAlphaChannel && IsTransparencySupported())
	{
		BLENDFUNCTION bf;
		
		bf.BlendOp = AC_SRC_OVER;
		bf.BlendFlags = 0;
		bf.SourceConstantAlpha = 0xff;
		bf.AlphaFormat = 1;
		bResult = ::AlphaBlend(hDestDC, xDest, yDest, nDestWidth, nDestHeight, 
			m_hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, bf);
	}
	else
#endif
	{
		bResult = ::StretchBlt(hDestDC, xDest, yDest, nDestWidth, nDestHeight, 
			m_hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, SRCCOPY);
	}
	
	ReleaseDC();
	
	return(bResult);
}

const void* CImage::GetBits() const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(IsDIBSection());
	
	return(m_pBits);
}

void* CImage::GetBits() throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(IsDIBSection());
	
	return(m_pBits);
}

int CImage::GetBPP() const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	
	return(m_nBPP);
}

void CImage::GetColorTable(UINT iFirstColor, UINT nColors, RGBQUAD* prgbColors) const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(m_pBits != NULL);
	ATLASSERT(IsIndexed());
	
	GetDC();
	
	::GetDIBColorTable(m_hDC, iFirstColor, nColors, prgbColors);
	
	ReleaseDC();
}

HDC CImage::GetDC() const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	
	m_nDCRefCount++;
	if(m_hDC == NULL)
	{
		m_hDC = s_cache.GetDC();
		m_hOldBitmap = HBITMAP(::SelectObject(m_hDC, m_hBitmap));
	}
	
	return(m_hDC);
}

bool CImage::ShouldExcludeFormat(REFGUID guidFileType, DWORD dwExclude) throw()
{
	static const GUID* apguidFormats[] =
	{
		&Gdiplus::ImageFormatGIF,
		&Gdiplus::ImageFormatBMP,
		&Gdiplus::ImageFormatEMF,
		&Gdiplus::ImageFormatWMF,
		&Gdiplus::ImageFormatJPEG,
		&Gdiplus::ImageFormatPNG,
		&Gdiplus::ImageFormatTIFF,
		&Gdiplus::ImageFormatIcon,
		NULL
	};

	ATLASSERT((dwExclude|excludeValid) == excludeValid);
	for(int iFormat = 0; apguidFormats[iFormat] != NULL; iFormat++)
	{
		if(guidFileType == *apguidFormats[iFormat])
		{
			return((dwExclude&(1<<iFormat)) != 0);
		}
	}
	
	return((dwExclude&excludeOther) != 0);
}

void CImage::BuildCodecFilterString(const Gdiplus::ImageCodecInfo* pCodecs, UINT nCodecs,
										   CSimpleString& strFilter, CSimpleArray< GUID >& aguidFileTypes, LPCTSTR pszAllFilesDescription, 
										   DWORD dwExclude, TCHAR chSeparator)
{
	USES_CONVERSION;
	if(pszAllFilesDescription != NULL)
	{
		aguidFileTypes.Add(const_cast<GUID&>(GUID_NULL));
	}
	
	CString strAllExtensions;
	CString strTempFilter;
	for(UINT iCodec = 0; iCodec < nCodecs; iCodec++)
	{
		const Gdiplus::ImageCodecInfo* pCodec = &pCodecs[iCodec];
		
		if(!ShouldExcludeFormat(pCodec->FormatID, dwExclude))
		{ 
			CString pwszFilenameExtension(W2CT(pCodec->FilenameExtension));
			strTempFilter += W2CT(pCodec->FormatDescription);
			strTempFilter += _T(" (");
			strTempFilter += pwszFilenameExtension;
			strTempFilter += _T(")");
			strTempFilter += chSeparator;
			strTempFilter += pwszFilenameExtension;
			strTempFilter += chSeparator;
			
			aguidFileTypes.Add(const_cast<GUID&>(pCodec->FormatID));
			
			if(!strAllExtensions.IsEmpty())
			{
				strAllExtensions += _T(";");
			}
			strAllExtensions += pwszFilenameExtension;
		}
	}
	
	if(pszAllFilesDescription != NULL)
	{
		strFilter += pszAllFilesDescription;
		strFilter += chSeparator;
		strFilter += strAllExtensions;
		strFilter += chSeparator;
	}
	strFilter += strTempFilter;
	
	strFilter += chSeparator;
	if(aguidFileTypes.GetSize() == 0)
	{
		strFilter += chSeparator;
	}
}

HRESULT CImage::GetImporterFilterString(CSimpleString& strImporters, CSimpleArray< GUID >& aguidFileTypes, 
												LPCTSTR pszAllFilesDescription/*= NULL */, 
												DWORD dwExclude/*= excludeDefaultLoad */, TCHAR chSeparator/*= '|' */)
{
	if(!InitGDIPlus())
	{
		return(E_FAIL);
	}
	
	UINT nCodecs;
	UINT nSize;
	Gdiplus::Status status;
	Gdiplus::ImageCodecInfo* pCodecs = NULL;
	
	status = Gdiplus::GetImageDecodersSize(&nCodecs, &nSize);
	pCodecs = static_cast< Gdiplus::ImageCodecInfo* >(malloc(nSize));
	
	if(pCodecs == NULL)
		return E_OUTOFMEMORY;
	
	status = Gdiplus::GetImageDecoders(nCodecs, nSize, pCodecs);
	BuildCodecFilterString(pCodecs, nCodecs, strImporters, aguidFileTypes, pszAllFilesDescription, dwExclude, chSeparator);
	
	free(pCodecs);
	
	return(S_OK);
}

HRESULT CImage::GetExporterFilterString(CSimpleString& strExporters, 
											   CSimpleArray< GUID >& aguidFileTypes, LPCTSTR pszAllFilesDescription/*= NULL */,
											   DWORD dwExclude/*= excludeDefaultSave */, TCHAR chSeparator/*= '|' */)
{
	if(!InitGDIPlus())
	{
		return(E_FAIL);
	}
	
	UINT nCodecs;
	UINT nSize;
	Gdiplus::Status status;
	Gdiplus::ImageCodecInfo* pCodecs;
	
	status = Gdiplus::GetImageDecodersSize(&nCodecs, &nSize);
	pCodecs = static_cast< Gdiplus::ImageCodecInfo* >(malloc(nSize));
	
	if(pCodecs == NULL)
		return E_OUTOFMEMORY;
	
	status = Gdiplus::GetImageDecoders(nCodecs, nSize, pCodecs);
	BuildCodecFilterString(pCodecs, nCodecs, strExporters, aguidFileTypes, pszAllFilesDescription, dwExclude, chSeparator);
	
	free(pCodecs);
	
	return(S_OK);
}

int CImage::GetHeight() const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	
	return(m_nHeight);
}

int CImage::GetMaxColorTableEntries() const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(IsDIBSection());
	
	if(IsIndexed())
	{
		return(1 << m_nBPP);
	}
	else
	{
		return(0);
	}
}

int CImage::GetPitch() const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(IsDIBSection());
	
	return(m_nPitch);
}

COLORREF CImage::GetPixel(int x, int y) const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT((x >= 0) && (x < m_nWidth));
	ATLASSERT((y >= 0) && (y < m_nHeight));
	
	GetDC();
	
	COLORREF clr = ::GetPixel(m_hDC, x, y);
	
	ReleaseDC();
	
	return(clr);
}

const void* CImage::GetPixelAddress(int x, int y) const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(IsDIBSection());
	ATLASSERT((x >= 0) && (x < m_nWidth));
	ATLASSERT((y >= 0) && (y < m_nHeight));
	
	return(LPBYTE(m_pBits)+(y*m_nPitch)+((x*m_nBPP)/8));
}

void* CImage::GetPixelAddress(int x, int y) throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(IsDIBSection());
	ATLASSERT((x >= 0) && (x < m_nWidth));
	ATLASSERT((y >= 0) && (y < m_nHeight));
	
	return(LPBYTE(m_pBits)+(y*m_nPitch)+((x*m_nBPP)/8));
}

LONG CImage::GetTransparentColor() const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT((m_nBPP == 4) || (m_nBPP == 8));
	
	return(m_iTransparentColor);
}

int CImage::GetWidth() const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	
	return(m_nWidth);
}

bool CImage::IsDIBSection() const throw()
{
	return(m_bIsDIBSection);
}

bool CImage::IsIndexed() const throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(IsDIBSection());
	
	return(m_nBPP <= 8);
}

bool CImage::IsNull() const throw()
{
	return(m_hBitmap == NULL);
}

HRESULT CImage::Load(IStream* pStream) throw()
{
	if(!InitGDIPlus())
	{
		return(E_FAIL);
	}
	
	Gdiplus::Bitmap bmSrc(pStream);
	if(bmSrc.GetLastStatus() != Gdiplus::Ok)
	{
		return(E_FAIL);
	}
	
	return(CreateFromGdiplusBitmap(bmSrc));
}

HRESULT CImage::Load(LPCTSTR pszFileName) throw()
{
	USES_CONVERSION; 
	if(!InitGDIPlus())
	{
		return(E_FAIL);
	}
	
	Gdiplus::Bitmap bmSrc(T2CW(pszFileName));
	if(bmSrc.GetLastStatus() != Gdiplus::Ok)
	{
		return(E_FAIL);
	}
	
	return(CreateFromGdiplusBitmap(bmSrc));
}

HRESULT CImage::CreateFromGdiplusBitmap(Gdiplus::Bitmap& bmSrc) throw()
{
	Gdiplus::PixelFormat eSrcPixelFormat = bmSrc.GetPixelFormat();
	UINT nBPP = 32;
	DWORD dwFlags = 0;
	Gdiplus::PixelFormat eDestPixelFormat = PixelFormat32bppRGB;
	if(eSrcPixelFormat&PixelFormatGDI)
	{
		nBPP = Gdiplus::GetPixelFormatSize(eSrcPixelFormat);
		eDestPixelFormat = eSrcPixelFormat;
	}
	if(Gdiplus::IsAlphaPixelFormat(eSrcPixelFormat))
	{
		nBPP = 32;
		dwFlags |= createAlphaChannel;
		eDestPixelFormat = PixelFormat32bppARGB;
	}
	
	BOOL bSuccess = Create(bmSrc.GetWidth(), bmSrc.GetHeight(), nBPP, dwFlags);
	if(!bSuccess)
	{
		return(E_FAIL);
	}
	Gdiplus::ColorPalette* pPalette = NULL;
	if(Gdiplus::IsIndexedPixelFormat(eSrcPixelFormat))
	{
		UINT nPaletteSize = bmSrc.GetPaletteSize();
		pPalette = static_cast< Gdiplus::ColorPalette* >(malloc(nPaletteSize));
		
		if(pPalette == NULL)
			return E_OUTOFMEMORY;
		
		bmSrc.GetPalette(pPalette, nPaletteSize);
		
		RGBQUAD argbPalette[256];
		ATLASSERT((pPalette->Count > 0) && (pPalette->Count <= 256));
		for(UINT iColor = 0; iColor < pPalette->Count; iColor++)
		{
			Gdiplus::ARGB color = pPalette->Entries[iColor];
			argbPalette[iColor].rgbRed = (BYTE)((color>>RED_SHIFT) & 0xff);
			argbPalette[iColor].rgbGreen = (BYTE)((color>>GREEN_SHIFT) & 0xff);
			argbPalette[iColor].rgbBlue = (BYTE)((color>>BLUE_SHIFT) & 0xff);
			argbPalette[iColor].rgbReserved = 0;
		}
		
		SetColorTable(0, pPalette->Count, argbPalette);
	}
	
	if(eDestPixelFormat == eSrcPixelFormat)
	{
		// The pixel formats are identical, so just memcpy the rows.
		Gdiplus::BitmapData data;
		Gdiplus::Rect rect(0, 0, GetWidth(), GetHeight());
		if(bmSrc.LockBits(&rect, Gdiplus::ImageLockModeRead, eSrcPixelFormat, &data)!=Gdiplus::Ok)
		{
			return E_OUTOFMEMORY;
		}
		
		size_t nBytesPerRow = _AtlAlignUp<size_t>(nBPP*GetWidth(), 8)/8;
		BYTE* pbDestRow = static_cast< BYTE* >(GetBits());
		BYTE* pbSrcRow = static_cast< BYTE* >(data.Scan0);
		for(int y = 0; y < GetHeight(); y++)
		{
			memcpy(pbDestRow, pbSrcRow, nBytesPerRow);
			pbDestRow += GetPitch();
			pbSrcRow += data.Stride;
		}
		
		bmSrc.UnlockBits(&data);
	}
	else
	{
		// Let GDI+ work its magic
		Gdiplus::Bitmap bmDest(GetWidth(), GetHeight(), GetPitch(), eDestPixelFormat, static_cast< BYTE* >(GetBits()));
		Gdiplus::Graphics gDest(&bmDest);
		
		gDest.DrawImage(&bmSrc, 0, 0);
	}
	
	free(pPalette);
	
	return(S_OK);
}

void CImage::LoadFromResource(HINSTANCE hInstance, LPCTSTR pszResourceName) throw()
{
	HBITMAP hBitmap;
	
	hBitmap = HBITMAP(::LoadImage(hInstance, pszResourceName, IMAGE_BITMAP, 0, 
		0, LR_CREATEDIBSECTION));
	
	Attach(hBitmap);
}

void CImage::LoadFromResource(HINSTANCE hInstance, UINT nIDResource) throw()
{
	LoadFromResource(hInstance, MAKEINTRESOURCE(nIDResource));
}

BOOL CImage::MaskBlt(HDC hDestDC, int xDest, int yDest, int nWidth, 
							int nHeight, int xSrc, int ySrc, HBITMAP hbmMask, int xMask, int yMask,
							DWORD dwROP) const throw()
{
	BOOL bResult;
	
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(hDestDC != NULL);
	
	GetDC();
	
	bResult = ::MaskBlt(hDestDC, xDest, yDest, nWidth, nHeight, m_hDC, xSrc, 
		ySrc, hbmMask, xMask, yMask, dwROP);
	
	ReleaseDC();
	
	return(bResult);
}

BOOL CImage::MaskBlt(HDC hDestDC, const RECT& rectDest, 
							const POINT& pointSrc, HBITMAP hbmMask, const POINT& pointMask, 
							DWORD dwROP) const throw()
{
	return(MaskBlt(hDestDC, rectDest.left, rectDest.top, rectDest.right-
		rectDest.left, rectDest.bottom-rectDest.top, pointSrc.x, pointSrc.y, 
		hbmMask, pointMask.x, pointMask.y, dwROP));
}

BOOL CImage::MaskBlt(HDC hDestDC, int xDest, int yDest, HBITMAP hbmMask, 
							DWORD dwROP) const throw()
{
	return(MaskBlt(hDestDC, xDest, yDest, m_nWidth, m_nHeight, 0, 0, hbmMask, 
		0, 0, dwROP));
}

BOOL CImage::MaskBlt(HDC hDestDC, const POINT& pointDest, HBITMAP hbmMask,
							DWORD dwROP) const throw()
{
	return(MaskBlt(hDestDC, pointDest.x, pointDest.y, m_nWidth, m_nHeight, 0, 
		0, hbmMask, 0, 0, dwROP));
}

BOOL CImage::PlgBlt(HDC hDestDC, const POINT* pPoints, int xSrc, 
						   int ySrc, int nSrcWidth, int nSrcHeight, HBITMAP hbmMask, int xMask, 
						   int yMask) const throw()
{
	BOOL bResult;
	
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(hDestDC != NULL);
	
	GetDC();
	
	bResult = ::PlgBlt(hDestDC, pPoints, m_hDC, xSrc, ySrc, nSrcWidth, 
		nSrcHeight, hbmMask, xMask, yMask);
	
	ReleaseDC();
	
	return(bResult);
}

BOOL CImage::PlgBlt(HDC hDestDC, const POINT* pPoints, 
						   const RECT& rectSrc, HBITMAP hbmMask, const POINT& pointMask) const throw()
{
	return(PlgBlt(hDestDC, pPoints, rectSrc.left, rectSrc.top, rectSrc.right-
		rectSrc.left, rectSrc.bottom-rectSrc.top, hbmMask, pointMask.x, 
		pointMask.y));
}

BOOL CImage::PlgBlt(HDC hDestDC, const POINT* pPoints, 
						   HBITMAP hbmMask) const throw()
{
	return(PlgBlt(hDestDC, pPoints, 0, 0, m_nWidth, m_nHeight, hbmMask, 0, 
		0));
}

void CImage::ReleaseDC() const throw()
{
	HBITMAP hBitmap;
	
	ATLASSERT(m_hDC != NULL);
	
	m_nDCRefCount--;
	if(m_nDCRefCount == 0)
	{
		hBitmap = HBITMAP(::SelectObject(m_hDC, m_hOldBitmap));
		ATLASSERT(hBitmap == m_hBitmap);
		s_cache.ReleaseDC(m_hDC);
		m_hDC = NULL;
	}
}

CLSID CImage::FindCodecForExtension(LPCTSTR pszExtension, const Gdiplus::ImageCodecInfo* pCodecs, UINT nCodecs)
{ 
	USES_CONVERSION; 
	WCHAR szExt[MAX_PATH] = { 0 }; 
	lstrcpynW(szExt, T2CW(pszExtension), MAX_PATH); 
	
	WCHAR szExtensions[MAX_PATH*4] = { 0 }; 
	
	for(UINT iCodec = 0; iCodec < nCodecs; iCodec++)
	{ 
		lstrcpynW(szExtensions, pCodecs[iCodec].FilenameExtension, MAX_PATH*4); 
		
		WCHAR * pEnd = szExtensions + lstrlenW(szExtensions); 
		WCHAR * pIter = szExtensions; 
		do
		{
			WCHAR * pToken = wcschr(pIter, L';'); 
			if (pToken) 
			{ 
				*pToken = '\0'; 
				pToken ++; 
			} 
			
			WCHAR * strExtension = ::PathFindExtensionW(pIter); 
			if (strExtension) 
			{ 
				if (lstrcmpiW(szExt, strExtension) == 0) 
				{ 
					return(pCodecs[iCodec].Clsid);
				} 
			} 
			
			if (pToken) { 
				pIter = pToken; 
			} else { 
				pIter += lstrlenW(pIter); 
			} 
		} while(pIter < pEnd);
	}
	
	return(CLSID_NULL);
}

CLSID CImage::FindCodecForFileType(REFGUID guidFileType, const Gdiplus::ImageCodecInfo* pCodecs, UINT nCodecs)
{
	for(UINT iCodec = 0; iCodec < nCodecs; iCodec++)
	{
		if(pCodecs[iCodec].FormatID == guidFileType)
		{
			return(pCodecs[iCodec].Clsid);
		}
	}
	
	return(CLSID_NULL);
}

HRESULT CImage::Save(IStream* pStream, REFGUID guidFileType) const throw()
{
	if(!InitGDIPlus())
	{
		return(E_FAIL);
	}
	
	UINT nEncoders;
	UINT nBytes;
	Gdiplus::Status status;
	
	status = Gdiplus::GetImageEncodersSize(&nEncoders, &nBytes);
	if(status != Gdiplus::Ok)
	{
		return(E_FAIL);
	}
	
	Gdiplus::ImageCodecInfo* pEncoders = static_cast< Gdiplus::ImageCodecInfo* >(malloc(nBytes));
	
	if(pEncoders == NULL)
		return E_OUTOFMEMORY;
	
	status = Gdiplus::GetImageEncoders(nEncoders, nBytes, pEncoders);
	if(status != Gdiplus::Ok)
	{
		return(E_FAIL);
	}
	
	CLSID clsidEncoder = FindCodecForFileType(guidFileType, pEncoders, nEncoders);
	if(clsidEncoder == CLSID_NULL)
	{
		return(E_FAIL);
	}
	
	if(m_bHasAlphaChannel)
	{
		ATLASSERT(m_nBPP == 32);
		Gdiplus::Bitmap bm(m_nWidth, m_nHeight, m_nPitch, PixelFormat32bppARGB, static_cast< BYTE* >(m_pBits));
		status = bm.Save(pStream, &clsidEncoder, NULL);
		if(status != Gdiplus::Ok)
		{
			return(E_FAIL);
		}
	}
	else
	{
		Gdiplus::Bitmap bm(m_hBitmap, NULL);
		status = bm.Save(pStream, &clsidEncoder, NULL);
		if(status != Gdiplus::Ok)
		{
			return(E_FAIL);
		}
	}
	free(pEncoders);
	
	return(S_OK);
}

HRESULT CImage::Save(LPCTSTR pszFileName, REFGUID guidFileType) const throw()
{
	if(!InitGDIPlus())
	{
		return(E_FAIL);
	}
	
	UINT nEncoders;
	UINT nBytes;
	Gdiplus::Status status;
	
	status = Gdiplus::GetImageEncodersSize(&nEncoders, &nBytes);
	if(status != Gdiplus::Ok)
	{
		return(E_FAIL);
	}
	
	USES_CONVERSION;
	Gdiplus::ImageCodecInfo* pEncoders = static_cast< Gdiplus::ImageCodecInfo* >(malloc(nBytes));
	
	if(pEncoders == NULL)
		return E_OUTOFMEMORY;
	
	status = Gdiplus::GetImageEncoders(nEncoders, nBytes, pEncoders);
	if(status != Gdiplus::Ok)
	{
		return(E_FAIL);
	}
	
	CLSID clsidEncoder = CLSID_NULL;
	if(guidFileType == GUID_NULL)
	{
		// Determine clsid from extension
		clsidEncoder = FindCodecForExtension(::PathFindExtension(pszFileName), pEncoders, nEncoders);
	}
	else
	{
		// Determine clsid from file type
		clsidEncoder = FindCodecForFileType(guidFileType, pEncoders, nEncoders);
	}
	if(clsidEncoder == CLSID_NULL)
	{
		return(E_FAIL);
	}
	
	LPCWSTR pwszFileName = T2CW(pszFileName);
#ifndef _UNICODE
	if(pwszFileName == NULL)
		return E_OUTOFMEMORY;
#endif
	if(m_bHasAlphaChannel)
	{
		ATLASSERT(m_nBPP == 32);
		Gdiplus::Bitmap bm(m_nWidth, m_nHeight, m_nPitch, PixelFormat32bppARGB, static_cast< BYTE* >(m_pBits));
		status = bm.Save(pwszFileName, &clsidEncoder, NULL);
		if(status != Gdiplus::Ok)
		{
			return(E_FAIL);
		}
	}
	else
	{
		Gdiplus::Bitmap bm(m_hBitmap, NULL);
		status = bm.Save(pwszFileName, &clsidEncoder, NULL);
		if(status != Gdiplus::Ok)
		{
			return(E_FAIL);
		}
	}
	
	free(pEncoders);
	
	return(S_OK);
}

void CImage::SetColorTable(UINT iFirstColor, UINT nColors, 
								  const RGBQUAD* prgbColors) throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(IsDIBSection());
	ATLASSERT(IsIndexed());
	
	GetDC();
	
	::SetDIBColorTable(m_hDC, iFirstColor, nColors, prgbColors);
	
	ReleaseDC();
}

void CImage::SetPixel(int x, int y, COLORREF color) throw()
{
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT((x >= 0) && (x < m_nWidth));
	ATLASSERT((y >= 0) && (y < m_nHeight));
	
	GetDC();
	
	::SetPixel(m_hDC, x, y, color);
	
	ReleaseDC();
}

void CImage::SetPixelIndexed(int x, int y, int iIndex) throw()
{
	SetPixel(x, y, PALETTEINDEX(iIndex));
}

void CImage::SetPixelRGB(int x, int y, BYTE r, BYTE g, BYTE b) throw()
{
	SetPixel(x, y, RGB(r, g, b));
}

LONG CImage::SetTransparentColor(LONG iTransparentColor) throw()
{
	LONG iOldTransparentColor;
	
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT((m_nBPP == 4) || (m_nBPP == 8));
	ATLASSERT(iTransparentColor < GetMaxColorTableEntries());
	ATLASSERT(iTransparentColor >= -1);
	
	iOldTransparentColor = m_iTransparentColor;
	m_iTransparentColor = iTransparentColor;
	
	return(iOldTransparentColor);
}

BOOL CImage::StretchBlt(HDC hDestDC, int xDest, int yDest, 
							   int nDestWidth, int nDestHeight, DWORD dwROP) const throw()
{
	return(StretchBlt(hDestDC, xDest, yDest, nDestWidth, nDestHeight, 0, 0, 
		m_nWidth, m_nHeight, dwROP));
}

BOOL CImage::StretchBlt(HDC hDestDC, const RECT& rectDest, 
							   DWORD dwROP) const throw()
{
	return(StretchBlt(hDestDC, rectDest.left, rectDest.top, rectDest.right-
		rectDest.left, rectDest.bottom-rectDest.top, 0, 0, m_nWidth, m_nHeight, 
		dwROP));
}

BOOL CImage::StretchBlt(HDC hDestDC, int xDest, int yDest, 
							   int nDestWidth, int nDestHeight, int xSrc, int ySrc, int nSrcWidth, 
							   int nSrcHeight, DWORD dwROP) const throw()
{
	BOOL bResult;
	
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(hDestDC != NULL);
	
	GetDC();
	
	bResult = ::StretchBlt(hDestDC, xDest, yDest, nDestWidth, nDestHeight, m_hDC,
		xSrc, ySrc, nSrcWidth, nSrcHeight, dwROP);
	
	ReleaseDC();
	
	return(bResult);
}

BOOL CImage::StretchBlt(HDC hDestDC, const RECT& rectDest, 
							   const RECT& rectSrc, DWORD dwROP) const throw()
{
	return(StretchBlt(hDestDC, rectDest.left, rectDest.top, rectDest.right-
		rectDest.left, rectDest.bottom-rectDest.top, rectSrc.left, rectSrc.top, 
		rectSrc.right-rectSrc.left, rectSrc.bottom-rectSrc.top, dwROP));
}

#if WINVER >= 0x0500
BOOL CImage::TransparentBlt(HDC hDestDC, int xDest, int yDest, 
								   int nDestWidth, int nDestHeight, UINT crTransparent) const throw()
{
	return(TransparentBlt(hDestDC, xDest, yDest, nDestWidth, nDestHeight, 0, 
		0, m_nWidth, m_nHeight, crTransparent));
}

BOOL CImage::TransparentBlt(HDC hDestDC, const RECT& rectDest, 
								   UINT crTransparent) const throw()
{
	return(TransparentBlt(hDestDC, rectDest.left, rectDest.top, 
		rectDest.right-rectDest.left, rectDest.bottom-rectDest.top, 
		crTransparent));
}

BOOL CImage::TransparentBlt(HDC hDestDC, int xDest, int yDest, 
								   int nDestWidth, int nDestHeight, int xSrc, int ySrc, int nSrcWidth, 
								   int nSrcHeight, UINT crTransparent) const throw()
{
	BOOL bResult;
	
	ATLASSERT(m_hBitmap != NULL);
	ATLASSERT(hDestDC != NULL);
	
	GetDC();
	
	if(crTransparent == CLR_INVALID)
	{
		crTransparent = GetTransparentRGB();
	}
	
	bResult = ::TransparentBlt(hDestDC, xDest, yDest, nDestWidth, nDestHeight,
		m_hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, crTransparent);
	
	ReleaseDC();
	
	return(bResult);
}

BOOL CImage::TransparentBlt(HDC hDestDC, const RECT& rectDest, 
								   const RECT& rectSrc, UINT crTransparent) const throw()
{
	return(TransparentBlt(hDestDC, rectDest.left, rectDest.top, 
		rectDest.right-rectDest.left, rectDest.bottom-rectDest.top, rectSrc.left, 
		rectSrc.top, rectSrc.right-rectSrc.left, rectSrc.bottom-rectSrc.top, 
		crTransparent));
}
#endif

BOOL CImage::IsTransparencySupported() throw()
{
	return(TRUE);
}

void CImage::UpdateBitmapInfo(DIBOrientation eOrientation)
{
	DIBSECTION dibsection;
	int nBytes;
	
	nBytes = ::GetObject(m_hBitmap, sizeof(DIBSECTION), &dibsection);
	if(nBytes == sizeof(DIBSECTION))
	{
		m_bIsDIBSection = true;
		m_nWidth = dibsection.dsBmih.biWidth;
		m_nHeight = abs(dibsection.dsBmih.biHeight);
		m_nBPP = dibsection.dsBmih.biBitCount;
		m_nPitch = ComputePitch(m_nWidth, m_nBPP);
		m_pBits = dibsection.dsBm.bmBits;
		if(eOrientation == DIBOR_DEFAULT)
		{
			eOrientation = (dibsection.dsBmih.biHeight > 0) ? DIBOR_BOTTOMUP : DIBOR_TOPDOWN;
		}
		if(eOrientation == DIBOR_BOTTOMUP)
		{
			m_pBits = LPBYTE(m_pBits)+((m_nHeight-1)*m_nPitch);
			m_nPitch = -m_nPitch;
		}
	}
	else
	{
		// Non - DIBSection
		ATLASSERT(nBytes == sizeof(BITMAP));
		m_bIsDIBSection = false;
		m_nWidth = dibsection.dsBm.bmWidth;
		m_nHeight = dibsection.dsBm.bmHeight;
		m_nBPP = dibsection.dsBm.bmBitsPixel;
		m_nPitch = 0;
		m_pBits = 0;
	}
	m_iTransparentColor = -1;
	m_bHasAlphaChannel = false;
}

void CImage::GenerateHalftonePalette(LPRGBQUAD prgbPalette)
{
	int r;
	int g;
	int b;
	int gray;
	LPRGBQUAD prgbEntry;
	
	prgbEntry = prgbPalette;
	for(r = 0; r < 6; r++)
	{
		for(g = 0; g < 6; g++)
		{
			for(b = 0; b < 6; b++)
			{
				prgbEntry->rgbBlue = BYTE(b*255/5);
				prgbEntry->rgbGreen = BYTE(g*255/5);
				prgbEntry->rgbRed = BYTE(r*255/5);
				prgbEntry->rgbReserved = 0;
				
				prgbEntry++;
			}
		}
	}
	
	for(gray = 0; gray < 20; gray++)
	{
		prgbEntry->rgbBlue = BYTE(gray*255/20);
		prgbEntry->rgbGreen = BYTE(gray*255/20);
		prgbEntry->rgbRed = BYTE(gray*255/20);
		prgbEntry->rgbReserved = 0;
		
		prgbEntry++;
	}
}

COLORREF CImage::GetTransparentRGB() const
{
	RGBQUAD rgb;
	
	ATLASSERT(m_hDC != NULL);
	ATLASSERT(m_iTransparentColor != -1);
	
	::GetDIBColorTable(m_hDC, m_iTransparentColor, 1, &rgb);
	
	return(RGB(rgb.rgbRed, rgb.rgbGreen, rgb.rgbBlue));
}

bool CImage::InitGDIPlus() throw()
{
	bool bSuccess = s_initGDIPlus.Init();
	return(bSuccess);
}
