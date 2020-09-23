Attribute VB_Name = "Module2"
'This function is awesome.
'I'm sorry to the author of this code I downloaded it
'a while ago and I forgot who it was.

Public Sub Raised3D(obj As Object)
    ' Gives the effect of a raised line around
    ' the form or picturebox
    ' Hold the original scale mode
    Dim nScaleMode As Integer
    ' Used for user defined scale only
    Dim sngScaleTop As Single
    Dim sngScaleLeft As Single
    Dim sngScaleWidth As Single
    Dim sngScaleHeight As Single
    If (TypeOf obj Is PictureBox) Or (TypeOf obj Is Form) Then
        nScaleMode = obj.ScaleMode
        If nScaleMode = 0 Then
            sngScaleTop = obj.ScaleTop
            sngScaleLeft = obj.ScaleLeft
            sngScaleWidth = obj.ScaleWidth
            sngScaleHeight = obj.ScaleHeight
        End If
        obj.ScaleMode = 3
        obj.Line (1, 1)-(obj.ScaleWidth - 1, 1), vb3DHighlight
        obj.Line (1, 2)-(obj.ScaleWidth, 2), vb3DShadow
        obj.Line (1, 2)-(1, obj.ScaleHeight), vb3DHighlight
        obj.Line (2, 2)-(2, obj.ScaleHeight), vb3DShadow
        obj.Line (1, obj.ScaleHeight - 2)-(obj.ScaleWidth, obj.ScaleHeight - 2), vb3DHighlight
        obj.Line (1, obj.ScaleHeight - 1)-(obj.ScaleWidth, obj.ScaleHeight - 1), vb3DShadow
        obj.Line (obj.ScaleWidth - 2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 2, 1), vb3DHighlight
        obj.Line (obj.ScaleWidth - 1, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, 1), vb3DShadow
        obj.ScaleMode = nScaleMode
        If nScaleMode = 0 Then
            obj.ScaleTop = sngScaleTop
            obj.ScaleWidth = sngScaleWidth
            obj.ScaleLeft = sngScaleLeft
            obj.ScaleHeight = sngScaleHeight
        End If
    End If
End Sub
