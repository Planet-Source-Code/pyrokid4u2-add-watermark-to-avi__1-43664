VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AVI Watermark"
   ClientHeight    =   3690
   ClientLeft      =   2115
   ClientTop       =   570
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5205
   Begin VB.Frame Frame2 
      Caption         =   "Preview"
      Height          =   3975
      Left            =   120
      TabIndex        =   32
      Top             =   3720
      Width           =   4935
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   3615
         Left            =   120
         ScaleHeight     =   237
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   309
         TabIndex        =   33
         Top             =   240
         Width           =   4695
         Begin VB.Timer Timer6 
            Interval        =   10
            Left            =   2040
            Top             =   480
         End
         Begin VB.PictureBox Picture9 
            Height          =   495
            Left            =   3720
            ScaleHeight     =   435
            ScaleWidth      =   315
            TabIndex        =   36
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Timer Timer5 
            Interval        =   10
            Left            =   2520
            Top             =   360
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Command12"
            Height          =   375
            Left            =   720
            TabIndex        =   35
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Timer Timer4 
            Interval        =   2000
            Left            =   2880
            Top             =   600
         End
      End
   End
   Begin VB.Frame wholeframe 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      TabIndex        =   14
      Top             =   -120
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "Load AVI File"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Append Watermark"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save AVI File"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   29
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   4935
         Begin VB.CommandButton Command13 
            Caption         =   "About"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3720
            TabIndex        =   39
            Top             =   1920
            Width           =   1095
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   175
            Left            =   120
            Max             =   255
            TabIndex        =   38
            Top             =   1440
            Value           =   125
            Width           =   1095
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Help"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3720
            TabIndex        =   19
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Preview"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   34
            Top             =   1200
            Width           =   1095
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Height          =   1935
            Left            =   1320
            Picture         =   "Form1.frx":0000
            ScaleHeight     =   125
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   149
            TabIndex        =   21
            Top             =   240
            Width           =   2295
            Begin VB.PictureBox Picture2 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   420
               Left            =   360
               Picture         =   "Form1.frx":DB02
               ScaleHeight     =   28
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   99
               TabIndex        =   24
               Top             =   1440
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   960
               TabIndex        =   22
               Text            =   "Text4"
               Top             =   2160
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   1080
               TabIndex        =   23
               Text            =   "Text3"
               Top             =   1680
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Sample Movie "
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1455
               Left            =   240
               TabIndex        =   40
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2295
               Left            =   0
               TabIndex        =   26
               Top             =   -360
               Width           =   2295
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Watermark -->"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   570
               TabIndex        =   25
               Top             =   1580
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Image picture10 
               Height          =   345
               Left            =   1320
               Picture         =   "Form1.frx":FC14
               Top             =   480
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Image Image1 
               Height          =   345
               Left            =   960
               Picture         =   "Form1.frx":1032A
               Top             =   720
               Visible         =   0   'False
               Width           =   375
            End
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Open Watermark"
            Height          =   495
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form1.frx":10A40
            Left            =   2280
            List            =   "Form1.frx":10A4D
            TabIndex        =   18
            Text            =   "Select One"
            Top             =   2190
            Width           =   1215
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            Height          =   1935
            Left            =   1320
            ScaleHeight     =   1875
            ScaleWidth      =   2235
            TabIndex        =   27
            Top             =   240
            Width           =   2295
            Begin VB.Image Image2 
               Height          =   345
               Left            =   0
               Picture         =   "Form1.frx":10A76
               Top             =   0
               Visible         =   0   'False
               Width           =   375
            End
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Transparency:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            TabIndex        =   37
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            FillStyle       =   7  'Diagonal Cross
            Height          =   855
            Left            =   120
            Top             =   240
            Width           =   1095
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            FillStyle       =   7  'Diagonal Cross
            Height          =   855
            Left            =   3720
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Allignment:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1320
            TabIndex        =   28
            Top             =   2235
            Width           =   1575
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2040
         Top             =   1440
      End
      Begin VB.CommandButton Command6 
         Caption         =   "About"
         Height          =   280
         Left            =   120
         TabIndex        =   15
         Top             =   3435
         Width           =   4935
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   165
         Left            =   120
         TabIndex        =   16
         Top             =   3195
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   291
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   360
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   735
      Left            =   4920
      TabIndex        =   13
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      Height          =   405
      Left            =   4800
      Picture         =   "Form1.frx":1118C
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   12
      Top             =   8040
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   480
      Left            =   4320
      Picture         =   "Form1.frx":118A2
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   7
      Top             =   9720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   735
      Left            =   5400
      TabIndex        =   11
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   4800
      Picture         =   "Form1.frx":192C4
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   6
      Top             =   9480
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Stretch"
      Enabled         =   0   'False
      Height          =   230
      Left            =   0
      TabIndex        =   5
      Top             =   9840
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "Height"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Text            =   "Width"
      Top             =   9210
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   9555
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Left            =   3960
      Top             =   8160
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2880
      Top             =   8160
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Watermark Details:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Transparent"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Transparency:"
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image picture5 
      Height          =   2055
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   8160
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim buttondown As Boolean, aviframerate As Integer
Dim xd As Integer, yd As Integer, combobefore
Dim FinalY As Integer, FinalX As Integer, FinalHeight As Integer, FinalWidth As Integer, append As Boolean
Dim FinalStretch As Boolean, setWH As Boolean
Dim transparencysetting As Integer
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Dim Frameone As Integer, Framedone As Integer
Private Declare Function GdiAlphaBlend Lib "gdi32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
'Big thanks to Ray Mercer for some of the code you see below.

    Dim Res As Long         'result code
    Dim ofd As CommonDialog     'OpenFileDialog class
    Dim szFile As String    'filename
    Dim pAVIFile As Long    'pointer to AVI file interface (PAVIFILE handle)
    Dim pAVIStream As Long  'pointer to AVI stream interface (PAVISTREAM handle)
    Dim numFrames As Long   'number of frames in video stream
    Dim firstframe As Long  'position of the first video frame
    Dim fileInfo As AVI_FILE_INFO       'file info struct
    Dim streamInfo As AVI_STREAM_INFO   'stream info struct
    Dim dib As cdib
    Dim pGetFrameObj As Long    'pointer to GetFrame interface
    Dim pDIB As Long            'pointer to packed DIB in memory
    Dim bih As BITMAPINFOHEADER 'infoheader to pass to GetFrame functions
    Dim i As Long

Private Sub Check1_Click()
If Check1.Value = 1 Then
Picture2.AutoSize = True
FinalStretch = True
Else
Picture2.AutoSize = False
FinalStretch = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Me.Height = 8280
Else
    Me.Height = 4170
End If
End Sub

Private Sub Combo1_Change()
Command12_Click
End Sub

Private Sub Combo1_GotFocus()
Timer4 = False
End Sub

Private Sub Combo1_LostFocus()
Timer4 = True
End Sub

Private Sub Combo1_Scroll()
Command12_Click
End Sub

'Private Sub Combo1_Scroll()
'Command12_Click
'End Sub

Private Sub Command1_Click()
If append = True Then GoTo appendmid
Command2.Enabled = False
CommonDialog1.filename = ""
'res

'Get the name of an AVI file to work with
Set ofd = CommonDialog1
With ofd
    .Filter = "AVI Files|*.avi"
    .DialogTitle = "Open AVI File"
    .ShowOpen
    If .filename = "" Then
    MsgBox "You must select a valid filename!", vbCritical, "Error"
    Exit Sub
    End If
End With
aviframerate = AVIFileFrameRate(CommonDialog1.filename)
'Dim pAVIFile As Long 'pointer to AVI File (PAVIFILE handle)

Res = AVIFileOpen(pAVIFile, ofd.filename, OF_SHARE_DENY_WRITE, 0&)
If Res <> AVIERR_OK Then GoTo ErrorOut
Me.Caption = "AVI Watermark (AVI Loaded!)"

Res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
If Res <> AVIERR_OK Then GoTo ErrorOut

'get the starting position of the stream (some streams may not start simultaneously)
firstframe = AVIStreamStart(pAVIStream)
If firstframe = -1 Then GoTo ErrorOut 'this function returns -1 on error

'get the length of video stream in frames
numFrames = AVIStreamLength(pAVIStream)
If numFrames = -1 Then GoTo ErrorOut ' this function returns -1 on error

Me.Caption = "AVI Watermark (AVI Details: Frames:" & numFrames & ")"

MsgBox "AVI File Details:" & vbCrLf & _
"PAVISTREAM handle is " & pAVIStream & vbCrLf & _
"Video stream length - " & numFrames & vbCrLf & _
"Stream starts on frame #" & firstframe, vbInformation, "Information"

'get file info struct (UDT)
Res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))
If Res <> AVIERR_OK Then GoTo ErrorOut

'print file info to Debug Window
Call DebugPrintAVIFileInfo(fileInfo)

'get stream info struct (UDT)
Res = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
If Res <> AVIERR_OK Then GoTo ErrorOut

'print stream info to Debug Window
Call DebugPrintAVIStreamInfo(streamInfo)
Command2.Enabled = True
Exit Sub
appendmid:
append = False
'MsgBox "Appending"
    'set bih attributes which we want GetFrame functions to return
    With bih
        .biBitCount = 24
        .biClrImportant = 0
        .biClrUsed = 0
        .biCompression = BI_RGB
        .biHeight = streamInfo.rcFrame.Bottom - streamInfo.rcFrame.Top
        .biPlanes = 1
        .biSize = 40
        .biWidth = streamInfo.rcFrame.Right - streamInfo.rcFrame.Left
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biSizeImage = (((.biWidth * 3) + 3) And &HFFFC) * .biHeight 'calculate total size of RGBQUAD scanlines (DWORD aligned)
    End With
    'GoTo now
'goagain:
'    bih.biBitCount = 8 'Small, yes, but at least it works :)
'now:
    'init AVISTreamGetFrame* functions and create GETFRAME object
    'pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, ByVal AVIGETFRAMEF_BESTDISPLAYFMT) 'tell AVIStream API what format we expect and input stream
    pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, bih) 'force function to return 24bit DIBS
    If pGetFrameObj = 0 Then 'That didn't work. Let's try something else
    pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, ByVal AVIGETFRAMEF_BESTDISPLAYFMT)
        If pGetFrameObj = 0 Then 'Well, if it's gonna be stuborn with us, choose another avi :(
            MsgBox "No suitable decompressor found for this video stream!", vbInformation, App.Title
            GoTo ErrorOut
        End If
    End If
    
    'create a DIB class to load the frames into
    Set dib = New cdib
        If Text3 = "true" Then Text3 = "false": Exit Sub
        pDIB = AVIStreamGetFrame(pGetFrameObj, i)  'returns "packed DIB"
        If dib.CreateFromPackedDIBPointer(pDIB) Then
On Error Resume Next
Dim NoRepeat As Boolean
            MkDir App.Path & "\temporary"
            Kill App.Path & "\temporary\" & firstframe & ".jpg"
            Call dib.WriteToFile(App.Path & "\temporary\" & firstframe & ".jpg") 'This'll probably take up a lot of space, so we'll delete 'em all at the end :)
Else
    MsgBox "There was an error writing the AVI to the screen.", vbCritical, "Error"
    Exit Sub
End If


            Picture6.Picture = LoadPicture(App.Path & "\temporary\" & i & ".jpg")
            'picture5.Picture = LoadPicture(App.Path & "\temporary\" & i & ".bmp")
            'If setWH <> True Then
            'Picture6.Width = Picture6.ScaleWidth 'picture5.Width
            'Picture6.Height = Picture6.ScaleHeight ' picture5.Height
            'setWH = True
            'End If
            'Picture6.Cls
            'Picture6.Refresh
            'Picture6.PaintPicture picture2.Picture, picture2.Left, picture2.Top, picture2.Width, picture2.Height
                 'Picture3.Visible = True
  Dim w As Long, h As Long
  w = Picture1.ScaleWidth
  h = Picture1.ScaleHeight
  'Picture6.Circle (10, 0), 6, vbBlack
  'vbBlack
  'Picture1.ScaleMode = 3
  'Dim asdf As Integer, fdsa As Integer
  'asdf = Picture1.ScaleWidth
  'fdsa = Picture1.ScaleHeight
  'Picture1.ScaleMode = 1
  fdsa = Picture6.ScaleHeight - Picture2.ScaleHeight - 5
  'MsgBox asdf
  If Combo1.ListIndex = 0 Then
    asdf = Picture6.ScaleWidth / 2 - (Picture2.ScaleWidth / 2)
  ElseIf Combo1.ListIndex = 1 Then
    asdf = Picture6.ScaleWidth - Picture2.ScaleWidth - 2
  ElseIf Combo1.ListIndex = 2 Then
    asdf = 2
  Else
    asdf = Picture6.ScaleWidth / 2 - (Picture2.ScaleWidth / 2)
  End If
    'Call AlphaBlend(Picture1.hdc, asdf, fdsa, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, 125)
  'MsgBox fdsa
  Call AlphaBlend(Picture6.hdc, asdf, fdsa, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, transparencysetting)
  Refresh
  '(Picture2.Left * Picture6.Width) / asdf
  '(Picture2.Top * Picture6.Height) / fdsa
  'Int((Picture2.ScaleWidth * Picture6.Width) / Picture1.Width), Int((Picture2.ScaleHeight * Picture6.Height) / Picture1.Height)
' 4 '   ' 5 '
'---' = '---'
' 2 '   ' x '
                 DoEvents
                 'Picture3.Visible = False
            'icture6.   (((Picture2.Top - 360) * Picture6.Height) / Picture1.Height)
            'BitBlt picture2, X, Y, W, H, TDC, 0, 0, SRCCOPY

            'DirectLoad App.Path & "\temporary\wtrmrk.bmp", Picture6.hdc, picture2.Picture, picture2.Width, picture2.Height, picture2.Left, picture2.Top
            'picture6.
            If Check2.Value = 0 Then GoTo noconfirm
                If MsgBox("Is the watermark the way you want it?" & vbCrLf & "Press yes to continue or no to stop.", vbQuestion + vbYesNo, "Confirm") = vbYes Then GoTo noconfirm
            Exit Sub
noconfirm:
    Set dib = New cdib
    For i = firstframe To (numFrames - 1) + firstframe
        If Text3 = "true" Then Text3 = "false": Exit Sub
        pDIB = AVIStreamGetFrame(pGetFrameObj, i)  'returns "packed DIB"
        If dib.CreateFromPackedDIBPointer(pDIB) Then
On Error Resume Next
            MkDir App.Path & "\temporary"
writefile:
On Error GoTo killfile
            Call dib.WriteToFile(App.Path & "\temporary\" & i & ".jpg") 'This'll probably take up a lot of space, so we'll delete 'em all at the end :)


            Picture6.Picture = LoadPicture(App.Path & "\temporary\" & i & ".jpg")
            'picture5.Picture = LoadPicture(App.Path & "\temporary\" & i & ".bmp")
            'If setWH <> True Then
            'Picture6.Width = Picture6.ScaleWidth 'picture5.Width
            'Picture6.Height = Picture6.ScaleHeight ' picture5.Height
            'setWH = True
            'End If
            'Picture6.Cls
            'Picture6.Refresh
            'Picture6.PaintPicture picture2.Picture, picture2.Left, picture2.Top, picture2.Width, picture2.Height
                 'Picture3.Visible = True

  'Picture6.Circle (10, 0), 6, vbBlack
  'vbBlack
  'Picture1.ScaleMode = 3
  'Dim asdf As Integer, fdsa As Integer
  'asdf = Picture1.ScaleWidth
  'fdsa = Picture1.ScaleHeight
  'Picture1.ScaleMode = 1
  fdsa = Picture6.ScaleHeight - Picture2.ScaleHeight - 5
  'MsgBox asdf
  If Combo1.ListIndex = 0 Then
    asdf = Picture6.ScaleWidth / 2 - (Picture2.ScaleWidth / 2)
  ElseIf Combo1.ListIndex = 1 Then
    asdf = Picture6.ScaleWidth - Picture2.ScaleWidth - 2
  ElseIf Combo1.ListIndex = 2 Then
    asdf = 2
  Else
    asdf = Picture6.ScaleWidth / 2 - (Picture2.ScaleWidth / 2)
  End If
    'Call AlphaBlend(Picture1.hdc, asdf, fdsa, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, 125)
  'MsgBox fdsa
  Call AlphaBlend(Picture6.hdc, asdf, fdsa, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, 125)
  Refresh
  '(Picture2.Left * Picture6.Width) / asdf
  '(Picture2.Top * Picture6.Height) / fdsa
  'Int((Picture2.ScaleWidth * Picture6.Width) / Picture1.Width), Int((Picture2.ScaleHeight * Picture6.Height) / Picture1.Height)
' 4 '   ' 5 '
'---' = '---'
' 2 '   ' x '
                 DoEvents


            DoEvents
            Kill App.Path & "\temporary\" & i & ".jpg"
            DoEvents
            SavePicture Picture6.Image, App.Path & "\temporary\" & i & ".jpg"
            NoRepeat = False
            GoTo nextfile
killfile:
            If NoRepeat = True Then
                NoRepeat = False
                GoTo ErrorOut
            End If
            NoRepeat = True
            Kill App.Path & "\temp\" & i & ".jpg"
            GoTo writefile
nextfile:
            'txtStatus = "Bitmap " & i + 1 & " of " & numFrames & " written to app folder"
            'txtStatus.Refresh
            Me.Caption = "AVI Watermark (Frame " & i + 1 & " of " & numFrames & ")"
        ProgressBar1.Value = Int((i / numFrames) * 100)
        Else
            GoTo ErrorOut
        End If

        
        
    Next
    Frameone = firstframe
    Framedone = numFrames + firstframe
    ProgressBar1.Value = 100
    Me.Caption = "AVI Watermark"
    Set dib = Nothing
    Command3.Enabled = True

    'And now for the error handling :)
Exit Sub
ErrorOut:
    If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile) '// closes the file
    End If

    If (rc <> AVIERR_OK) Then 'if there was an error then show feedback to user
        MsgBox "There was an error working with the file:" _
                & vbCrLf & szFile, vbInformation, App.Title
    End If


End Sub

Private Sub Command12_Click()
'Picture9.Picture = Picture1.Picture
'Picture1.Picture = ""
'Picture1.Picture = Picture1.Picture
Picture1.Cls
'picture1.
  fdsa = Picture1.ScaleHeight - Picture2.ScaleHeight - 5
  'MsgBox asdf
  If Combo1.ListIndex = 0 Then
    asdf = Picture1.ScaleWidth / 2 - (Picture2.ScaleWidth / 2)
  ElseIf Combo1.ListIndex = 1 Then
    asdf = Picture1.ScaleWidth - Picture2.ScaleWidth - 2
  ElseIf Combo1.ListIndex = 2 Then
    asdf = 2
  Else
    asdf = Picture1.ScaleWidth / 2 - (Picture2.ScaleWidth / 2)
  End If
    Call AlphaBlend(Picture1.hdc, asdf, fdsa, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, transparencysetting)

  'Call AlphaBlend(Picture6.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, 125)
  Refresh
'DoEvents
End Sub

Private Sub Command13_Click()
frmAbout.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
Command3.Enabled = False

    MkDir App.Path & "\temporary"
    Kill App.Path & "\temporary\wtrmrk.bmp"
    SavePicture Picture2.Picture, App.Path & "\temporary\wtrmrk.bmp"
    append = True
    'Timer1 = True
    'Text4.Text = "true"
    Call Command1_Click
End Sub

Private Sub Command3_Click()
With CommonDialog1
.filename = ""
.DialogTitle = "Save AVI File"
.Filter = "AVI Files|*.avi"
.ShowSave
If .filename <> "" Then
WriteAvi .filename, Frameone, Framedone
End If
End With
End Sub

Private Sub Command4_Click()
On Error GoTo errorhandle
With CommonDialog1
.DialogTitle = "Open Watermark"
.Filter = ""
.ShowOpen

If .filename <> "" Then
Picture2.Picture = LoadPicture(.filename)
Else
MsgBox "You must choose an image to load.", vbCritical, "Error"
End If
.filename = ""
End With
Exit Sub
errorhandle:
If Err.Number = 481 Then
MsgBox "Please select a valid image.", vbCritical, "Error"
End If
Dim asdf As Integer, fdsa As Integer

  fdsa = Picture1.ScaleHeight - Picture2.ScaleHeight - 5
  'MsgBox asdf
  If Combo1.ListIndex = 0 Then
    asdf = Picture1.ScaleWidth / 2 - (Picture1.ScaleWidth / 2)
  ElseIf Combo1.ListIndex = 1 Then
    asdf = Picture2.ScaleWidth - Picture1.ScaleWidth - 2
  Else
    asdf = 2
  End If
Picture2.Left = asdf
Picture2.Top = fdsa
End Sub

Private Sub Command5_Click()
Text1_set
FinalHeight = Val(Text1)
Text2_set
FinalWidth = Val(Text2)

End Sub

Private Sub Command6_Click()
Load frmAbout
frmAbout.Show
End Sub





Private Sub Command9_Click()
On Error GoTo nofile
Shell "notepad " & App.Path & "\readme.txt", vbMaximizedFocus
Exit Sub
nofile:
MsgBox "The help file does not exist. If you downloaded this program, try downloading it again.", vbCritical, "Error"
End Sub

Private Sub Form_Load()

transparencysetting = 125
Command12_Click
Check1_Click
Call AVIFileInit
Me.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call AVIFileExit
End Sub

Private Sub HScroll1_Change()
transparencysetting = HScroll1.Value
transparencysetting = 255 - transparencysetting
Command12_Click
End Sub

Private Sub HScroll1_Scroll()
transparencysetting = HScroll1.Value
transparencysetting = 255 - transparencysetting
Command12_Click

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
buttondown = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button <> 1 Then Exit Sub
If buttondown <> True Then Exit Sub
Picture2.Top = y - 365 '- picture2.Height '- 20 '+ yd + (yd / 2)
Picture2.Left = x - 5 '+ (Picture1.Left) '- 5 '- 20 '- xd
'Picture1.Refresh
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'If Button <> 1 Then Exit Sub
If buttondown <> True Then Exit Sub
Picture2.Top = y - 365 '- picture2.Height ' - 20 '+ yd + (yd / 2)
Picture2.Left = x - 5 '- 20 '- xd
'Picture1.Refresh
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 1 Then
'If buttondown = False Then
'    buttondown = True
'    Label6.Visible = False
'Else
'    buttondown = False
'End If
'End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If buttondown <> True Then Exit Sub
If x > xd Then
Picture2.Left = Picture2.Left + (x - xd) + x + x '(X - 20) ' 15 ' - 20
End If

If y > yd Then
Picture2.Top = Picture2.Top + (y - yd) + y + y ' (Y - 20) '- 20
End If

    xd = x '+ (X - 20) - 20
    yd = y '+ (y - 20) - 20

End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'buttondown = False
End Sub

Private Sub Text1_set()
On Error Resume Next
Picture2.Height = Text1
Picture2.ScaleHeight = Text1

End Sub

Private Sub Text2_set()
On Error Resume Next
Picture2.Width = Text2
Picture2.ScaleWidth = Text2
End Sub

Private Sub Picture6_Resize()
If Check2.Value = 0 Then Exit Sub
Frame2.Width = Picture6.Width + 240
Frame2.Height = Picture6.Height + 360
Picture6.Left = (Frame2.Width - Picture6.Width) / 2
Picture6.Top = ((Frame2.Height + 50) - Picture6.Height) / 2
Form1.Height = Frame2.Height + Frame2.Top + 550
If Form1.Width > Frame2.Width Then GoTo noform
Form1.Width = Frame2.Width + 360
wholeframe.Left = (Form1.Width - wholeframe.Width) / 2
noform:
Frame2.Left = (Form1.ScaleWidth - Frame2.Width) / 2
End Sub

Private Sub Timer1_Timer()
append = False
Timer1 = False
End Sub

Private Function WriteAvi(aviname As String, firstframe As Integer, lastFrame As Integer)
 
    Dim InitDir As String
    Dim szOutputAVIFile As String
    Dim Res As Long
    Dim pfile As Long 'ptr PAVIFILE
    Dim bmp As cdib
    Dim ps As Long 'ptr PAVISTREAM
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim BI As BITMAPINFOHEADER
    Dim opts As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim i As Long
    
    szOutputAVIFile = aviname
        
'    Open the file for writing
    Res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (Res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set bmp = New cdib
    If bmp.CreateFromFile(App.Path & "\temporary\" & firstframe & ".jpg") <> True Then
        MsgBox "Could not load frames into AVI!", vbExclamation, "Error"
        GoTo error
    End If

'   Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&                             '// default AVI handler
        .dwScale = 1
        .dwRate = Val(aviframerate)                        '// fps
        .dwSuggestedBufferSize = bmp.SizeImage       '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)       '// rectangle for stream
    End With
    
    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

'   And create the stream
    Res = AVIFileCreateStream(pfile, ps, strhdr)
    If (Res <> AVIERR_OK) Then GoTo error

    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
    Res = AVISaveOptions(Me.hwnd, ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, 1, ps, pOpts)                     'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If Res <> 1 Then 'In C TRUE = 1
        Call AVISaveOptionsFree(1, pOpts)
        GoTo error
    End If
    
    'make compressed stream
    Res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If Res <> AVIERR_OK Then GoTo error
    
    'set format of stream according to the bitmap
    With BI
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With
    
    'set the format of the compressed stream
    Res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (Res <> AVIERR_OK) Then GoTo error
'Set the progress bar values
ProgressBar1.Value = 0
        ProgressBar1.Min = firstframe
        ProgressBar1.Max = lastFrame
'   Now write out each video frame
    For i = firstframe To lastFrame - 1
        bmp.CreateFromFile (App.Path & "\temporary\" & i & ".jpg") 'load the bitmap (ignore errors)
        Res = AVIStreamWrite(psCompressed, i, 1, bmp.PointerToBits, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
        If Res <> AVIERR_OK Then GoTo error
        'Show user feedback
        On Error Resume Next 'Slight bug it doesnt wat to show the last frame
        If Check2.Value = 1 Then Picture6.Picture = LoadPicture(App.Path & "\temporary\" & i & ".jpg"): Picture2.Refresh: DoEvents
        Me.Caption = "AVI Watermark (Frame " & i & "/" & lastFrame - firstframe & " saved)"
        On Error GoTo error 'Set error handling back to normal
        'Set the progress bar
        ProgressBar1.Value = i
        Kill App.Path & "\temporary\" & i & ".jpg"
    Next
    ProgressBar1.Value = lastFrame
    
    Me.Caption = "AVI Watermark"
ShellExecLaunchFile aviname, "", ""
error:
'   Now close the file
    Set file = Nothing
    Set bmp = Nothing
    
    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pfile <> 0) Then Call AVIFileClose(pfile)

    Call AVIFileExit

    If (Res <> AVIERR_OK) Then
        MsgBox "There was an error writing the file.", vbInformation, App.Title
    End If

End Function

Private Sub Timer4_Timer()
Command12_Click
End Sub

Private Sub Timer5_Timer()
Dim asdf As Integer, fdsa As Integer

  fdsa = Picture1.ScaleHeight - Picture2.ScaleHeight - 5
  'MsgBox asdf
  If Combo1.ListIndex = 0 Then
    asdf = Picture1.ScaleWidth / 2 - (Picture2.ScaleWidth / 2)
  ElseIf Combo1.ListIndex = 1 Then
    asdf = Picture1.ScaleWidth - Picture2.ScaleWidth - 2
  ElseIf Combo1.ListIndex = 2 Then
    asdf = 2
  Else
    asdf = Picture1.ScaleWidth / 2 - (Picture2.ScaleWidth / 2)
  End If
Picture2.Left = asdf
Picture2.Top = fdsa
End Sub

Private Sub Timer6_Timer()
If Combo1.ListIndex <> combobefore Then
Command12_Click
End If
combobefore = Combo1.ListIndex
End Sub

