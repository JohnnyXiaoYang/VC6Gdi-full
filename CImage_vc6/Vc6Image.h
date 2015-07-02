// Vc6Image.h: interface for the CVc6Image class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_VC6IMAGE_H__CDB8CCF9_C57D_41D4_B0A4_9E7A0CA1E084__INCLUDED_)
#define AFX_VC6IMAGE_H__CDB8CCF9_C57D_41D4_B0A4_9E7A0CA1E084__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <Gdiplus.h>
#include <afxtempl.h>
#pragma comment(lib, "Gdiplus.lib")

typedef CMap<CString, LPCTSTR, GUID, REFGUID> CFormatMap;

class CVc6Image
{
	friend class GdiplusInfo;
public:
	CVc6Image();
	CVc6Image(LPCTSTR lpFileName);
	CVc6Image(HMODULE hModule, LPCTSTR bmpResName);
	CVc6Image(HMODULE hModule, LPCTSTR lpResName, LPCTSTR lpResType);
	virtual ~CVc6Image();

	Gdiplus::Bitmap* Detach();
	BOOL Attach(Gdiplus::Bitmap* pBitmap);

public:
	//guidFileType: ImageFormatBMP;ImageFormatGIF;ImageFormatPNG;ImageFormatTIFF;ImageFormatJPEG
	BOOL Load(LPCTSTR lpFileName);
	BOOL LoadBitmap(HMODULE hModule, LPCTSTR bmpResName);
	BOOL Load(HMODULE hModule, LPCTSTR lpResName, LPCTSTR lpResType);
	BOOL Save(LPCTSTR lpFileName, REFGUID guidFileType = GUID_NULL);

	// About Images
public:
	UINT GetWidth() const;
	UINT GetHeight() const;

	// About Operations
public:
	Gdiplus::Bitmap* CopyImage(int x, int y, int width, int height, Gdiplus::PixelFormat format) const; // Must be  manually delete.

	// About Show in Windows
public:
	BOOL Draw(HDC hdc, int x, int y, int width, int height);
	BOOL Draw(HDC hdc, int x, int y, int srcx, int srcy, int srcWidth, int srcHeight);
	BOOL Draw(HDC hdc, int x, int y, int width, int height, int srcx, int srcy, int srcWidth, int srcHeight);

private: 
	void __ReleaseBitmap();
	static int __GetEncoderClsid(REFGUID imageFormat, CLSID* pClsid, BOOL bGetEncoders = TRUE);

private:
	Gdiplus::Bitmap *m_pBitmap;
	static CFormatMap ms_formatMap;
};

#endif // !defined(AFX_VC6IMAGE_H__CDB8CCF9_C57D_41D4_B0A4_9E7A0CA1E084__INCLUDED_)
