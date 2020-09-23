Attribute VB_Name = "Module4"
'Private types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'variable just in the module
Dim glo_from As Long
Dim glo_to As Long
Dim glo_AliasName As String
Dim glo_hWnd As Long
Private Function AVIFillInfoStruct(ByVal filename As String, ByRef pfi As AVI_FILE_INFO) As Boolean
    Dim hr As Long 'HRESULT (0 = success)
    Dim pfile As Long 'PAVIFILE pointer
    
    Call AVIFileInit     'opens AVIFile library

    'open the AVI file
    hr = AVIFileOpen(pfile, filename, 0&, 0&) 'OF_SHARE_DENY_WRITE
    If hr <> 0 Then
        MsgBox "Unable to open file:" & vbCrLf & filename, vbCritical, App.Title
        Exit Function
    End If
    
    'file the info struct
    hr = AVIFileInfo(pfile, pfi, 108)
    If hr <> 0 Then
        MsgBox "Unable to read AVI file info from:" & vbCrLf & filename, vbCritical, App.Title
        Call AVIFileRelease(pfile) 'closes the file
        Call AVIFileExit        'releases AVIFile library
        Exit Function
    End If
    
    'close the file
    Call AVIFileRelease(pfile) 'closes the file
    Call AVIFileExit        'releases AVIFile library
    
    AVIFillInfoStruct = True 'indicate success
End Function
Public Function AVIFileFrameRate(ByVal filename As String) As Long
    Dim pfi As AVI_FILE_INFO
    
    If AVIFillInfoStruct(filename, pfi) Then
      'return fps
        On Error Resume Next 'ignore div by zero errors
        AVIFileFrameRate = pfi.dwRate / pfi.dwScale
        
    End If

End Function

