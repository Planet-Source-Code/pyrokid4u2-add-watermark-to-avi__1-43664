Attribute VB_Name = "Module3"
Option Explicit
Public Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long 'returns fourcc
Public Declare Function VideoForWindowsVersion Lib "msvfw32.dll" () As Long
Public Declare Function AVIFileOpen Lib "avifil32.dll" (ByRef ppfile As Long, ByVal szFile As String, ByVal uMode As Long, ByVal pclsidHandler As Long) As Long  'HRESULT
Public Declare Function AVIFileInfo Lib "avifil32.dll" (ByVal pfile As Long, pfi As AVI_FILE_INFO, ByVal lSize As Long) As Long 'HRESULT
Public Declare Function AVIFileCreateStream Lib "avifil32.dll" Alias "AVIFileCreateStreamA" (ByVal pfile As Long, ByRef ppavi As Long, ByRef psi As AVI_STREAM_INFO) As Long
Public Declare Function AVISaveOptions Lib "avifil32.dll" (ByVal hwnd As Long, ByVal uiFlags As Long, ByVal nStreams As Long, ByRef ppavi As Long, ByRef ppOptions As Long) As Long
Public Declare Function AVISave Lib "avifil32.dll" Alias "AVISaveVA" (ByVal szFile As String, ByVal pclsidHandler As Long, ByVal lpfnCallback As Long, ByVal nStreams As Long, ByRef ppaviStream As Long, ByRef ppCompOptions As Long) As Long
Public Declare Function AVISaveOptionsFree Lib "avifil32.dll" (ByVal nStreams As Long, ByRef ppOptions As Long) As Long
Public Declare Function AVIMakeCompressedStream Lib "avifil32.dll" (ByRef ppsCompressed As Long, ByVal psSource As Long, ByRef lpOptions As AVI_COMPRESS_OPTIONS, ByVal pclsidHandler As Long) As Long
Public Declare Function AVIStreamSetFormat Lib "avifil32.dll" (ByVal pavi As Long, ByVal lPos As Long, ByRef lpFormat As Any, ByVal cbFormat As Long) As Long
Public Declare Function AVIStreamWrite Lib "avifil32.dll" (ByVal pavi As Long, ByVal lStart As Long, ByVal lSamples As Long, ByVal lpBuffer As Long, ByVal cbBuffer As Long, ByVal dwFlags As Long, ByRef plSampWritten As Long, ByRef plBytesWritten As Long) As Long
Public Declare Function AVIStreamReadFormat Lib "avifil32.dll" (ByVal pAVIStream As Long, ByVal lPos As Long, ByVal lpFormatBuf As Long, ByRef sizeBuf As Long) As Long
Public Declare Function AVIStreamRead Lib "avifil32.dll" (ByVal pAVIStream As Long, ByVal lStart As Long, ByVal lSamples As Long, ByVal lpBuffer As Long, ByVal cbBuffer As Long, ByRef pBytesWritten As Long, ByRef pSamplesWritten As Long) As Long
Public Declare Function AVIStreamGetFrameOpen Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef bih As Any) As Long
Public Declare Function AVIStreamGetFrame Lib "avifil32.dll" (ByVal pGetFrameObj As Long, ByVal lPos As Long) As Long
Public Declare Function AVIStreamGetFrameClose Lib "avifil32.dll" (ByVal pGetFrameObj As Long) As Long
Public Declare Function AVIFileGetStream Lib "avifil32.dll" (ByVal pfile As Long, ByRef ppaviStream As Long, ByVal fccType As Long, ByVal lParam As Long) As Long
Public Declare Function AVIMakeFileFromStreams Lib "avifil32.dll" (ByRef ppfile As Long, ByVal nStreams As Long, ByVal pAVIStreamArray As Long) As Long
Public Declare Function AVIStreamInfo Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef psi As AVI_STREAM_INFO, ByVal lSize As Long) As Long
Public Declare Function AVIStreamStart Lib "avifil32.dll" (ByVal pavi As Long) As Long
Public Declare Function AVIStreamLength Lib "avifil32.dll" (ByVal pavi As Long) As Long
Public Declare Function AVIStreamRelease Lib "avifil32.dll" (ByVal pavi As Long) As Long 'ULONG
Public Declare Function AVIStreamClose Lib "avifil32.dll" Alias "AVIStreamRelease" (ByVal pavi As Long) As Long 'ULONG
Public Declare Function AVIFileRelease Lib "avifil32.dll" (ByVal pfile As Long) As Long
Public Declare Function AVIFileClose Lib "avifil32.dll" Alias "AVIFileRelease" (ByVal pfile As Long) As Long
Public Declare Function AVIMakeStreamFromClipboard Lib "avifil32.dll" (ByVal cfFormat As Long, ByVal hGlobal As Long, ByRef ppstream As Long) As Long
Public Declare Function AVIPutFileOnClipboard Lib "avifil32.dll" (ByVal pAVIFile As Long) As Long
Public Declare Function AVIGetFromClipboard Lib "avifil32.dll" (ByRef ppAVIFile As Long) As Long
Public Declare Function AVIClearClipboard Lib "avifil32.dll" () As Long

Public Declare Sub AVIFileInit Lib "avifil32.dll" ()
Public Declare Sub AVIFileExit Lib "avifil32.dll" ()

Public Const BMP_MAGIC_COOKIE As Integer = 19778

Public Type BITMAPFILEHEADER '14 bytes
  bfType As Integer
  bfSize As Long
  bfReserved1 As Integer
  bfReserved2 As Integer
  bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER '40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type BITMAPINFOHEADER_MJPEG '68 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
   biExtDataOffset As Long
   JPEGSize As Long
   JPEGProcess As Long
   JPEGColorSpaceID As Long
   JPEGBitsPerSample As Long
   JPEGHSubSampling As Long
   JPEGVSubSampling As Long
End Type
        
Public Type AVI_RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type AVI_STREAM_INFO
    fccType As Long
    fccHandler As Long
    dwFlags As Long
    dwCaps As Long
    wPriority As Integer
    wLanguage As Integer
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
    dwInitialFrames As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As AVI_RECT
    dwEditCount  As Long
    dwFormatChangeCount As Long
    szName As String * 64
End Type

Public Type AVI_FILE_INFO
    dwMaxBytesPerSecond As Long
    dwFlags As Long
    dwCaps As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwLength As Long
    dwEditCount As Long
    szFileType As String * 64
End Type

Public Type AVI_COMPRESS_OPTIONS
    fccType As Long
    fccHandler As Long
    dwKeyFrameEvery As Long
    dwQuality As Long
    dwBytesPerSecond As Long
    dwFlags As Long
    lpFormat As Long
    cbFormat As Long
    lpParms As Long
    cbParms As Long
    dwInterleaveEvery As Long
End Type
Global Const AVIERR_OK As Long = 0&

Public Const SEVERITY_ERROR    As Long = &H80000000
Public Const FACILITY_ITF      As Long = &H40000
Public Const AVIERR_BASE       As Long = &H4000

Global Const AVIERR_BADFLAGS    As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 105) '-2147205015
Global Const AVIERR_BADPARAM    As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 106) '-2147205014
Global Const AVIERR_BADSIZE     As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 107) '-2147205013
Global Const AVIERR_USERABORT   As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 198) '-2147204922

Global Const AVIFILEINFO_HASINDEX         As Long = &H10
Global Const AVIFILEINFO_MUSTUSEINDEX     As Long = &H20
Global Const AVIFILEINFO_ISINTERLEAVED    As Long = &H100
Global Const AVIFILEINFO_WASCAPTUREFILE   As Long = &H10000
Global Const AVIFILEINFO_COPYRIGHTED      As Long = &H20000

Global Const AVIFILECAPS_CANREAD          As Long = &H1
Global Const AVIFILECAPS_CANWRITE         As Long = &H2
Global Const AVIFILECAPS_ALLKEYFRAMES     As Long = &H10
Global Const AVIFILECAPS_NOCOMPRESSION    As Long = &H20

Global Const AVICOMPRESSF_INTERLEAVE     As Long = &H1           '// interleave
Global Const AVICOMPRESSF_DATARATE       As Long = &H2           '// use a data rate
Global Const AVICOMPRESSF_KEYFRAMES      As Long = &H4           '// use keyframes
Global Const AVICOMPRESSF_VALID          As Long = &H8           '// has valid data?

Global Const OF_READ  As Long = &H0
Global Const OF_WRITE As Long = &H1
Global Const OF_SHARE_DENY_WRITE As Long = &H20
Global Const OF_CREATE As Long = &H1000

Global Const AVIIF_KEYFRAME  As Long = &H10

Global Const DIB_RGB_COLORS  As Long = 0    '/* color table in RGBs */
Global Const DIB_PAL_COLORS  As Long = 1    '/* color table in palette indices */

Global Const BI_RGB          As Long = 0
Global Const BI_RLE8         As Long = 1
Global Const BI_RLE4         As Long = 2
Global Const BI_BITFIELDS    As Long = 3

Global Const streamtypeVIDEO       As Long = 1935960438 'equivalent to: mmioStringToFOURCC("vids", 0&)
Global Const streamtypeAUDIO       As Long = 1935963489 'equivalent to: mmioStringToFOURCC("auds", 0&)
Global Const streamtypeMIDI        As Long = 1935960429 'equivalent to: mmioStringToFOURCC("mids", 0&)
Global Const streamtypeTEXT        As Long = 1937012852 'equivalent to: mmioStringToFOURCC("txts", 0&)

Global Const AVIGETFRAMEF_BESTDISPLAYFMT  As Long = 1

Global Const ICMF_CHOOSE_KEYFRAME           As Long = &H1     '// show KeyFrame Every box
Global Const ICMF_CHOOSE_DATARATE           As Long = &H2     '// show DataRate box
Global Const ICMF_CHOOSE_PREVIEW            As Long = &H4     '// allow expanded preview dialog
Global Const ICMF_CHOOSE_ALLCOMPRESSORS     As Long = &H8     '// don't only show those that
                                                              '// can handle the input format
Public Declare Function SetRect Lib "user32.dll" (ByRef lprc As AVI_RECT, ByVal xLeft As Long, ByVal yTop As Long, ByVal xRight As Long, ByVal yBottom As Long) As Long
Public Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
Public Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal dwLen As Long)

Public Const HEAP_ZERO_MEMORY As Long = &H8

Global gfAbort As Boolean
Private Declare Function AlphaBlending Lib "msimg32.dll" Alias "AlphaBlend" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BF As Long) As Long
Private Declare Function DrawTransparent Lib "msimg32.dll" Alias "TransparentBlt" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

Public Function AlphaBlend(ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long
  Dim lngBlend As Long
  lngBlend = Val("&h" & Hex(AlphaSource) & "00" & "00")
  AlphaBlending destHDC, XDest, YDest, destWidth, destHeight, srcHDC, xSrc, ySrc, srcWidth, srcHeight, lngBlend
End Function

Public Function TransparentBlt(ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal TransparentColor As Long) As Long
  DrawTransparent destHDC, XDest, YDest, destWidth, destHeight, srcHDC, xSrc, ySrc, srcWidth, srcHeight, TransparentColor
End Function
