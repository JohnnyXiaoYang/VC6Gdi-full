// Vc6Image.cpp: implementation of the CVc6Image class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Vc6Image.h"
#include <shlwapi.h>
#pragma comment(lib, "shlwapi.lib")

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

// Static member datas
using namespace Gdiplus;

CFormatMap CVc6Image::ms_formatMap;

static class GdiplusInfo
{
	ULONG_PTR m_gdiplusToken;
	GdiplusStartupInput m_gdiplusStartupInput;
public:
	GdiplusInfo()
	{
		CFormatMap& formatMap = CVc6Image::ms_formatMap;
		GdiplusStartup(&m_gdiplusToken, &m_gdiplusStartupInput, NULL);
		formatMap.SetAt(_T(".bmp"),  ImageFormatBMP);
		formatMap.SetAt(_T(".png"),  ImageFormatPNG);
		formatMap.SetAt(_T(".gif"),  ImageFormatGIF);
		formatMap.SetAt(_T(".jpg"),  ImageFormatJPEG);
		formatMap.SetAt(_T(".tif"),  ImageFormatTIFF);
		formatMap.SetAt(_T(".jpeg"), ImageFormatJPEG);
		formatMap.SetAt(_T(".tiff"), ImageFormatTIFF);
	}
	~GdiplusInfo()
	{
		GdiplusShutdown(m_gdiplusToken);
	}

} theGdiplus;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CVc6Image::CVc6Image()
: m_pBitmap(NULL)
{

}

CVc6Image::CVc6Image(LPCTSTR lpFileName)
: m_pBitmap(NULL)
{
	Load(lpFileName);
}

CVc6Image::CVc6Image(HMODULE hModule, LPCTSTR bmpResName)
: m_pBitmap(NULL)
{
	LoadBitmap(hModule, bmpResName);
}

CVc6Image::CVc6Image(HMODULE hModule, LPCTSTR lpResName, LPCTSTR lpResType)
: m_pBitmap(NULL)
{
	Load(hModule, lpResName, lpResType);
}

CVc6Image::~CVc6Image()
{
	__ReleaseBitmap();
}

BOOL CVc6Image::Attach(Gdiplus::Bitmap* pBitmap)
{
	ASSERT(m_pBitmap == NULL);
	m_pBitmap = pBitmap;
	return NULL != m_pBitmap;
}

Bitmap* CVc6Image::Detach()
{
	Bitmap* pRet = m_pBitmap;
	m_pBitmap = NULL;
	return pRet;
}

void CVc6Image::__ReleaseBitmap()
{
	if (m_pBitmap != NULL)
	{
		::delete m_pBitmap;
		m_pBitmap = NULL;
	}
}

BOOL CVc6Image::Load(LPCTSTR lpFileName)
{
	if (NULL != lpFileName)
	{
		__ReleaseBitmap();
		WCHAR szFileName[2048] = {0};
#ifdef _UNICODE
		lstrcpyn(szFileName, lpFileName, 2048);
#else
		mbstowcs(szFileName, lpFileName, 2048);
#endif
		m_pBitmap = ::new Bitmap(szFileName);
	}
	return m_pBitmap != NULL;
}

BOOL CVc6Image::LoadBitmap(HMODULE hModule, LPCTSTR bmpResName)
{
	BOOL bRet(FALSE);
	__ReleaseBitmap();
	LPCWSTR lpResName(NULL);
#ifdef _UNICODE
	lpResName = bmpResName;
#else
	lpResName = MAKEINTRESOURCEW(bmpResName);
#endif
	m_pBitmap = ::new Bitmap(hModule, lpResName);
	if (m_pBitmap != NULL)
	{
		bRet = TRUE;
	}
	return bRet;
}

BOOL CVc6Image::Load(HMODULE hModule, LPCTSTR lpResName, LPCTSTR lpResType)
{
	BOOL bRet(FALSE);
	HRSRC hResource = FindResource(hModule, lpResName, lpResType);
	if (hResource != NULL)
	{
		__ReleaseBitmap();
		DWORD imageSize = SizeofResource(hModule, hResource);
		if (imageSize > 0)
		{
			LPVOID pResourceData = LockResource(LoadResource(hModule, hResource));
			if (NULL != pResourceData)
			{
				HGLOBAL hBuffer  = GlobalAlloc(GMEM_MOVEABLE, imageSize);
				if (NULL != hBuffer)
				{
					LPVOID pBuffer = GlobalLock(hBuffer);
					if (NULL != pBuffer)
					{
						CopyMemory(pBuffer, pResourceData, imageSize);
						IStream* pStream = NULL;
						if (CreateStreamOnHGlobal(hBuffer, TRUE, &pStream) == S_OK)
						{
							m_pBitmap = ::new Bitmap(pStream);
							if (NULL != m_pBitmap)
							{
								bRet = TRUE;
							}
						}
						GlobalUnlock(hBuffer);
					}
				}
			}		
		}	
	}
	return bRet;
}

int CVc6Image::__GetEncoderClsid(REFGUID imageFormat, CLSID* pClsid, BOOL bGetEncoders /*= TRUE*/)
{
	UINT  num = 0;          // number of image encoders
	UINT  size = 0;         // size of the image encoder array in bytes
	
	ImageCodecInfo* pImageCodecInfo = NULL;
	
	if (bGetEncoders) GetImageEncodersSize(&num, &size);
	else GetImageDecodersSize(&num, &size);
	
	if (size == 0) return -1;  // Failure
	
	pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
	if (pImageCodecInfo == NULL) return -1;  // Failure
	
	if (bGetEncoders) GetImageEncoders(num, size, pImageCodecInfo);
	else GetImageDecoders(num, size, pImageCodecInfo);
	
	for(UINT j = 0; j < num; ++j)
	{
		if (pImageCodecInfo[j].FormatID == imageFormat)
		{
			*pClsid = pImageCodecInfo[j].Clsid;
			free(pImageCodecInfo);
			return j;  // Success
		}    
	}
	free(pImageCodecInfo);
	return -1;  // Failure
}

BOOL CVc6Image::Save(LPCTSTR lpFileName, REFGUID guidFileType /*= GUID_NULL*/)
{
	ASSERT(NULL != m_pBitmap);
	BOOL bRet(FALSE);
	CLSID clsid;
	if (GUID_NULL != guidFileType)
	{
		__GetEncoderClsid(guidFileType, &clsid);
	}
	else
	{
		CString strExtension = PathFindExtension(lpFileName);
		strExtension.MakeLower();
		GUID guid = ms_formatMap[strExtension];
		__GetEncoderClsid(guid, &clsid);
	}
	WCHAR szFileName[2048] = {0};
#ifdef _UNICODE
	lstrcpyn(szFileName, lpFileName, 2048);
#else
	mbstowcs(szFileName, lpFileName, 2048);
#endif
	if (Ok == m_pBitmap->Save(szFileName, &clsid))
	{
		bRet = TRUE;
	}
	return bRet;
}

UINT CVc6Image::GetWidth( ) const
{
	ASSERT(NULL != m_pBitmap);
	return m_pBitmap->GetWidth();
}

UINT CVc6Image::GetHeight() const
{
	ASSERT(NULL != m_pBitmap);
	return m_pBitmap->GetHeight();
}

Bitmap* CVc6Image::CopyImage(int x, int y, int width, int height, PixelFormat format) const
{
	ASSERT(NULL != m_pBitmap);
	return m_pBitmap->Clone(x, y, width, height, format);
}

BOOL CVc6Image::Draw(HDC hdc, int x, int y, int srcx, int srcy, int srcWidth, int srcHeight)
{
	ASSERT(NULL != m_pBitmap);
	BOOL bRet(FALSE);
	Graphics graphics(hdc);
	if (Ok == graphics.DrawImage(m_pBitmap, x, y, srcx, srcy, srcWidth, srcHeight, UnitPixel))
	{
		bRet = TRUE;
	}
	return bRet;
}

BOOL CVc6Image::Draw(HDC hdc, int x, int y, int width, int height, int srcx, int srcy, int srcWidth, int srcHeight)
{
	ASSERT(NULL != m_pBitmap);
	BOOL bRet(FALSE);
	Graphics graphics(hdc);
	Rect rect(x, y, width, height);
	if (Ok == graphics.DrawImage(m_pBitmap, rect, srcx, srcy, srcWidth, srcHeight, UnitPixel))
	{
		bRet = TRUE;
	}
	return bRet;
}

BOOL CVc6Image::Draw(HDC hdc, int x, int y, int width, int height)
{
	ASSERT(NULL != m_pBitmap);
	BOOL bRet(FALSE);
	Graphics graphics(hdc);
	if (Ok == graphics.DrawImage(m_pBitmap, x, y, width, height))
	{
		bRet = TRUE;
	}
	return bRet;
}
