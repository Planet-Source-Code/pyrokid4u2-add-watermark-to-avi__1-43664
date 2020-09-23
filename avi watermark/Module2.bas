Attribute VB_Name = "Module1"
 Option Explicit
 Private Declare Function AlphaBlend _
  Lib "msimg32" ( _
  ByVal hDestDC As Long, _
  ByVal x As Long, ByVal y As Long, _
  ByVal nWidth As Long, _
  ByVal nHeight As Long, _
  ByVal hSrcDC As Long, _
  ByVal xSrc As Long, _
  ByVal ySrc As Long, _
  ByVal widthSrc As Long, _
  ByVal heightSrc As Long, _
  ByVal dreamAKA As Long) _
  As Boolean 'only Windows 98 or Latter
 Dim Num As Byte, nN%, nBlend&


' /**************************************************************************
' *
' *  'UTILITY FUNCTIONS FOR WORKING WITH AVI FILES
' *
' ***************************************************************************/
Public Sub DebugPrintAVIStreamInfo(asi As AVI_STREAM_INFO)
Debug.Print ""
Debug.Print "**** AVI_STREAM_INFO (START) ****"
 With asi
    Debug.Print "fccType = " & .fccType
    Debug.Print "fccHandler = " & .fccHandler
    Debug.Print "dwFlags = " & .dwFlags
    Debug.Print "dwCaps = " & .dwCaps
    Debug.Print "wPriority = " & .wPriority
    Debug.Print "wLanguage = " & .wLanguage
    Debug.Print "dwScale = " & .dwScale
    Debug.Print "dwRate = " & .dwRate
    Debug.Print "dwStart = " & .dwStart
    Debug.Print "dwLength = " & .dwLength
    Debug.Print "dwInitialFrames = " & .dwInitialFrames
    Debug.Print "dwSuggestedBufferSize = " & .dwSuggestedBufferSize
    Debug.Print "dwQuality = " & .dwQuality
    Debug.Print "dwSampleSize = " & .dwSampleSize
    Debug.Print "rcFrame.left = " & .rcFrame.Left
    Debug.Print "rcFrame.top = " & .rcFrame.Top
    Debug.Print "rcFrame.right = " & .rcFrame.Right
    Debug.Print "rcFrame.bottom = " & .rcFrame.Bottom
    Debug.Print "dwEditCount = " & .dwEditCount
    Debug.Print "dwFormatChangeCount = " & .dwFormatChangeCount
    Debug.Print "szName = " & .szName
 End With
 Debug.Print "**** AVI_STREAM_INFO (END) ****"
 Debug.Print ""
End Sub

Public Sub DebugPrintAVIFileInfo(afi As AVI_FILE_INFO)
Debug.Print "**** AVI_FILE_INFO (START) ****"
 With afi
    Debug.Print "dwMaxBytesPerSecond = " & .dwMaxBytesPerSecond
    Debug.Print "dwFlags = " & .dwFlags
    Debug.Print "dwCaps = " & .dwCaps
    Debug.Print "dwStreams = " & .dwStreams
    Debug.Print "dwSuggestedBufferSize = " & .dwSuggestedBufferSize
    Debug.Print "dwWidth = " & .dwWidth
    Debug.Print "dwHeight = " & .dwHeight
    Debug.Print "dwScale = " & .dwScale
    Debug.Print "dwRate = " & .dwRate
    Debug.Print "dwLength = " & .dwLength
    Debug.Print "dwEditCount = " & .dwEditCount
    Debug.Print "szFileType = " & .szFileType
 End With
 Debug.Print "**** AVI_FILE_INFO (END) ****"
 Debug.Print ""
End Sub

