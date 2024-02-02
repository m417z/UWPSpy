#include "stdafx.h"

#include "flash_area.h"

#include "resource.h"

// Based on code from WinSpy++.
// https://github.com/m417z/winspy
// www.catch22.net
// Copyright (C) 2012 James Brown

#define WC_FLASHWINDOW L"FlashAreaClass"
#define ALPHA_LEVEL 128

// Creates a stream object initialized with the data from an executable
// resource.
IStream* CreateStreamOnResource(HMODULE hModule, PCWSTR lpName, PCWSTR lpType) {
    IStream* ipStream = NULL;

    // find the resource
    HRSRC hRes = FindResource(hModule, lpName, lpType);
    DWORD dwResourceSize;
    HGLOBAL hglbImage;
    PVOID pvSourceResourceData;
    HGLOBAL hgblResourceData;
    PVOID pvResourceData;

    if (hRes == NULL)
        goto Return;

    // load the resource
    dwResourceSize = SizeofResource(hModule, hRes);
    hglbImage = LoadResource(hModule, hRes);

    if (hglbImage == NULL)
        goto Return;

    // lock the resource, getting a pointer to its data
    pvSourceResourceData = LockResource(hglbImage);

    if (pvSourceResourceData == NULL)
        goto Return;

    // allocate memory to hold the resource data
    hgblResourceData = GlobalAlloc(GMEM_MOVEABLE, dwResourceSize);

    if (hgblResourceData == NULL)
        goto Return;

    // get a pointer to the allocated memory
    pvResourceData = GlobalLock(hgblResourceData);

    if (pvResourceData == NULL)
        goto FreeData;

    // copy the data from the resource to the new memory block
    CopyMemory(pvResourceData, pvSourceResourceData, dwResourceSize);

    GlobalUnlock(hgblResourceData);

    // create a stream on the HGLOBAL containing the data
    // Specify that the HGLOBAL will be automatically free'd on the last
    // Release()
    if (SUCCEEDED(CreateStreamOnHGlobal(hgblResourceData, TRUE, &ipStream)))
        goto Return;

FreeData:

    // couldn't create stream; free the memory
    GlobalFree(hgblResourceData);

Return:

    // no need to unlock or free the resource
    return ipStream;
}

// Now that we have an IStream pointer to the data of the image,
// we can use WIC to load that image. An important step in this
// process is to use WICConvertBitmapSource to ensure that the image
// is in a 32bpp format suitable for direct conversion into a DIB.
// This method assumes that the input image is in the PNG format;
// for a splash screen, this is an excellent choice because it allows
// an alpha channel as well as lossless compression of the source image.
// (To make the splash screen image as small as possible,
// I highly recommend the PNGOUT compression utility.)

// Loads a PNG image from the specified stream (using Windows Imaging
// Component).
IWICBitmapSource* LoadBitmapFromStream(IStream* ipImageStream) {
    // initialize return value
    IWICBitmapSource* ipBitmap = NULL;

    // load WIC's PNG decoder
    IWICBitmapDecoder* ipDecoder = NULL;
    IID i = IID_IWICBitmapDecoder;

    if (FAILED(CoCreateInstance(CLSID_WICPngDecoder1, NULL,
                                CLSCTX_INPROC_SERVER,
                                i,  //__uuidof(ipDecoder)
                                (void**)&ipDecoder))) {
        return ipBitmap;
    }

    // load the PNG
    if (FAILED(ipDecoder->Initialize(ipImageStream,
                                     WICDecodeMetadataCacheOnLoad))) {
        ipDecoder->Release();
        return ipBitmap;
    }

    // check for the presence of the first frame in the bitmap
    UINT nFrameCount = 0;

    if (FAILED(ipDecoder->GetFrameCount(&nFrameCount)) || nFrameCount != 1) {
        ipDecoder->Release();
        return ipBitmap;
    }

    // load the first frame (i.e., the image)
    IWICBitmapFrameDecode* ipFrame = NULL;

    if (FAILED(ipDecoder->GetFrame(0, &ipFrame))) {
        ipDecoder->Release();
        return ipBitmap;
    }

    // convert the image to 32bpp BGRA format with pre-multiplied alpha
    //   (it may not be stored in that format natively in the PNG resource,
    //   but we need this format to create the DIB to use on-screen)
    WICConvertBitmapSource(GUID_WICPixelFormat32bppPBGRA, ipFrame, &ipBitmap);
    ipFrame->Release();
    ipDecoder->Release();
    return ipBitmap;
}

// Creates a 32-bit DIB from the specified WIC bitmap.
HBITMAP CreateHBITMAP(IWICBitmapSource* ipBitmap, PVOID* bits) {
    // initialize return value
    HBITMAP hbmp = NULL;

    // get image attributes and check for valid image
    UINT width = 0;
    UINT height = 0;

    if (FAILED(ipBitmap->GetSize(&width, &height)) || width == 0 ||
        height == 0) {
        return hbmp;
    }

    // prepare structure giving bitmap information (negative height indicates a
    // top-down DIB)
    BITMAPINFO bminfo;

    ZeroMemory(&bminfo, sizeof(bminfo));

    bminfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bminfo.bmiHeader.biWidth = width;
    bminfo.bmiHeader.biHeight = -((LONG)height);
    bminfo.bmiHeader.biPlanes = 1;
    bminfo.bmiHeader.biBitCount = 32;
    bminfo.bmiHeader.biCompression = BI_RGB;

    // create a DIB section that can hold the image
    void* pvImageBits = NULL;

    HDC hdcScreen = GetDC(NULL);
    hbmp = CreateDIBSection(hdcScreen, &bminfo, DIB_RGB_COLORS, &pvImageBits,
                            NULL, 0);
    ReleaseDC(NULL, hdcScreen);

    if (bits)
        *bits = pvImageBits;

    if (hbmp == NULL) {
        return hbmp;
    }

    // extract the image into the HBITMAP
    const UINT cbStride = width * 4;
    const UINT cbImage = cbStride * height;

    if (FAILED(ipBitmap->CopyPixels(NULL, cbStride, cbImage,
                                    static_cast<BYTE*>(pvImageBits)))) {
        // couldn't extract image; delete HBITMAP
        DeleteObject(hbmp);

        hbmp = NULL;
    }

    return hbmp;
}

// Loads the PNG bitmap into a 32bit HBITMAP.
HBITMAP LoadPNGImage(HMODULE hModule, UINT id, OUT VOID** bits) {
    HBITMAP hbmpSplash = NULL;

    // load the PNG image data into a stream
    IStream* ipImageStream =
        CreateStreamOnResource(hModule, MAKEINTRESOURCE(id), L"PNG");

    if (ipImageStream == NULL) {
        return hbmpSplash;
    }

    // load the bitmap with WIC
    IWICBitmapSource* ipBitmap = LoadBitmapFromStream(ipImageStream);

    if (ipBitmap == NULL) {
        ipImageStream->Release();
        return hbmpSplash;
    }

    // create a HBITMAP containing the image
    hbmpSplash = CreateHBITMAP(ipBitmap, bits);

    ipBitmap->Release();
    ipImageStream->Release();
    return hbmpSplash;
}

int GetRectWidth(RECT* rect) {
    return rect->right - rect->left;
}

int GetRectHeight(RECT* rect) {
    return rect->bottom - rect->top;
}

void UpdateLayeredWindowContent(HWND hwnd, RECT rc, HBITMAP hbmp, BYTE alpha) {
    POINT ptZero = {0, 0};
    POINT pt;
    SIZE size;
    BLENDFUNCTION blend = {AC_SRC_OVER, 0, 0, AC_SRC_ALPHA};
    blend.SourceConstantAlpha = alpha;

    pt.x = rc.left;
    pt.y = rc.top;

    size.cx = GetRectWidth(&rc);
    size.cy = GetRectHeight(&rc);

    HDC hdcSrc = GetDC(0);
    HDC hdcMem = CreateCompatibleDC(hdcSrc);

    HANDLE hbmpOld = SelectObject(hdcMem, hbmp);

    UpdateLayeredWindow(hwnd, hdcSrc, &pt, &size, hdcMem, &ptZero, RGB(0, 0, 0),
                        &blend, ULW_ALPHA);

    SelectObject(hdcMem, hbmpOld);
    DeleteDC(hdcMem);

    ReleaseDC(0, hdcSrc);
}

HBITMAP ExpandNineGridImage(SIZE outputSize, HBITMAP hbmSrc, RECT edges) {
    HDC hdcScreen, hdcDst, hdcSrc;
    HBITMAP hbmDst;
    void* pBits;
    HANDLE hOldSrc, hOldDst;
    BITMAP bmSrc;

    // Create a 32bpp DIB of the desired size, this is the output bitmap.
    BITMAPINFOHEADER bih = {sizeof(bih)};

    bih.biWidth = outputSize.cx;
    bih.biHeight = outputSize.cy;
    bih.biPlanes = 1;
    bih.biBitCount = 32;
    bih.biCompression = BI_RGB;
    bih.biSizeImage = 0;

    hdcScreen = GetDC(0);
    hbmDst = CreateDIBSection(hdcScreen, (BITMAPINFO*)&bih, DIB_RGB_COLORS,
                              &pBits, 0, 0);

    // Determine size of the source image.
    GetObject(hbmSrc, sizeof(bmSrc), &bmSrc);

    // Prep DCs
    hdcSrc = CreateCompatibleDC(hdcScreen);
    hOldSrc = SelectObject(hdcSrc, hbmSrc);

    hdcDst = CreateCompatibleDC(hdcScreen);
    hOldDst = SelectObject(hdcDst, hbmDst);

    // Sizes of the nine-grid edges
    int cxEdgeL = edges.left;
    int cxEdgeR = edges.right;
    int cyEdgeT = edges.top;
    int cyEdgeB = edges.bottom;

    // Precompute sizes and coordinates of the interior boxes
    // (that is, the source and dest rects with the edges subtracted out).
    int cxDstInner = outputSize.cx - (cxEdgeL + cxEdgeR);
    int cyDstInner = outputSize.cy - (cyEdgeT + cyEdgeB);
    int cxSrcInner = bmSrc.bmWidth - (cxEdgeL + cxEdgeR);
    int cySrcInner = bmSrc.bmHeight - (cyEdgeT + cyEdgeB);

    int xDst1 = cxEdgeL;
    int xDst2 = outputSize.cx - cxEdgeR;
    int yDst1 = cyEdgeT;
    int yDst2 = outputSize.cy - cyEdgeB;

    int xSrc1 = cxEdgeL;
    int xSrc2 = bmSrc.bmWidth - cxEdgeR;
    int ySrc1 = cyEdgeT;
    int ySrc2 = bmSrc.bmHeight - cyEdgeB;

    // Upper-left corner
    BitBlt(hdcDst, 0, 0, cxEdgeL, cyEdgeT, hdcSrc, 0, 0, SRCCOPY);

    // Upper-right corner
    BitBlt(hdcDst, xDst2, 0, cxEdgeR, cyEdgeT, hdcSrc, xSrc2, 0, SRCCOPY);

    // Lower-left corner
    BitBlt(hdcDst, 0, yDst2, cxEdgeL, cyEdgeB, hdcSrc, 0, ySrc2, SRCCOPY);

    // Lower-right corner
    BitBlt(hdcDst, xDst2, yDst2, cxEdgeR, cyEdgeB, hdcSrc, xSrc2, ySrc2,
           SRCCOPY);

    // Left side
    StretchBlt(hdcDst, 0, yDst1, cxEdgeL, cyDstInner, hdcSrc, 0, ySrc1, cxEdgeL,
               cySrcInner, SRCCOPY);

    // Right side
    StretchBlt(hdcDst, xDst2, yDst1, cxEdgeR, cyDstInner, hdcSrc, xSrc2, ySrc1,
               cxEdgeR, cySrcInner, SRCCOPY);

    // Top side
    StretchBlt(hdcDst, xDst1, 0, cxDstInner, cyEdgeT, hdcSrc, xSrc1, 0,
               cxSrcInner, cyEdgeT, SRCCOPY);

    // Bottom side
    StretchBlt(hdcDst, xDst1, yDst2, cxDstInner, cyEdgeB, hdcSrc, xSrc1, ySrc2,
               cxSrcInner, cyEdgeB, SRCCOPY);

    // Middle
    StretchBlt(hdcDst, xDst1, yDst1, cxDstInner, cyDstInner, hdcSrc, xSrc1,
               ySrc1, cxSrcInner, cySrcInner, SRCCOPY);

    SelectObject(hdcSrc, hOldSrc);
    SelectObject(hdcDst, hOldDst);

    DeleteDC(hdcSrc);
    DeleteDC(hdcDst);

    ReleaseDC(0, hdcScreen);

    return hbmDst;
}

HBITMAP MakeFlashWindowBitmap(HMODULE hModule, RECT rc) {
    static HBITMAP hbmBox;

    if (hbmBox == 0) {
        hbmBox = LoadPNGImage(hModule, IDB_SELBOX, NULL);
    }

    RECT edges = {2, 2, 2, 2};
    SIZE size;

    size.cx = GetRectWidth(&rc);
    size.cy = GetRectHeight(&rc);

    return ExpandNineGridImage(size, hbmBox, edges);
}

HWND FlashArea(HWND hParentWnd, HMODULE hModule, const RECT& rc, bool show) {
    static BOOL fInitializedWindowClass = FALSE;

    HWND hwnd;

    if (!fInitializedWindowClass) {
        WNDCLASSEX wc = {sizeof(wc)};

        wc.style = 0;
        wc.lpszClassName = WC_FLASHWINDOW;
        wc.lpfnWndProc = DefWindowProc;

        RegisterClassEx(&wc);

        fInitializedWindowClass = TRUE;
    }

    hwnd = CreateWindowEx(WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_LAYERED,
                          WC_FLASHWINDOW, 0, WS_POPUP, rc.left, rc.top,
                          rc.right - rc.left, rc.bottom - rc.top, hParentWnd, 0,
                          0, NULL);

    if (hwnd) {
        HBITMAP hbmp = MakeFlashWindowBitmap(hModule, rc);
        UpdateLayeredWindowContent(hwnd, rc, hbmp, ALPHA_LEVEL);
        DeleteObject(hbmp);

        UINT flags = SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE;
        if (show) {
            flags |= SWP_SHOWWINDOW;
        }

        SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags);
    }

    return hwnd;
}
