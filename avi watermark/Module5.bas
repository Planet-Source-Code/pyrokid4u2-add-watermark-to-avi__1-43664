Attribute VB_Name = "Module5"
Option Explicit

      Public Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String, _
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

      Public Declare Function GetDesktopWindow Lib "user32" () As Long

    Const SW_HIDE = 0                '  Hides the window and activates another window.
    Const SW_MAXIMIZE = 3            '  Maximizes the specified window.
    Const SW_MINIMIZE = 6            '  Minimizes the specified window and activates the next top-level window in the z-order.
    Const SW_RESTORE = 9             '  Activates and displays the window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    Const SW_SHOW = 5                '  Activates the window and displays it in its current size and position.
    Const SW_SHOWDEFAULT = 10        '  Sets the show state based on the SW_ flag specified in theSTARTUPINFO structure passed to theCreateProcess function by the program that started the application. An application should callShowWindow with this flag to set the initial show state of its main window.
    Const SW_SHOWMAXIMIZED = 3       '  Activates the window and displays it as a maximized window.
    Const SW_SHOWMINIMIZED = 2       '  Activates the window and displays it as a minimized window.
    Const SW_SHOWMINNOACTIVE = 7     '  Displays the window as a minimized window. The active window remains active.
    Const SW_SHOWNA = 8              '  Displays the window in its current state. The active window remains active.
    Const SW_SHOWNOACTIVATE = 4      '  Displays a window in its most recent size and position. The active window remains active.
    Const SW_SHOWNORMAL = 1          '  Activates and displays a window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
    Const SE_ERR_FNF = 2                 '  File not found
    Const SE_ERR_PNF = 3                 '  Path not found
    Const SE_ERR_ACCESSDENIED = 5        '  Access denied
    Const SE_ERR_OOM = 8                 '  Out of memory
    Const SE_ERR_DLLNOTFOUND = 32        '  DLL not found
    Const SE_ERR_SHARE = 26              '  A sharing violation occurred
    Const SE_ERR_ASSOCINCOMPLETE = 27    '  Incomplete or invalid file association
    Const SE_ERR_DDETIMEOUT = 28         '  DDE Time out
    Const SE_ERR_DDEFAIL = 29            '  DDE transaction failed
    Const SE_ERR_DDEBUSY = 30            '  DDE busy
    Const SE_ERR_NOASSOC = 31            '  No association for file extension
    Const ERROR_BAD_FORMAT = 11&         '  Invalid EXE file or error in EXE image
    Const ERROR_FILE_NOT_FOUND = 2&      '  The specified file was not found.
    Const ERROR_PATH_NOT_FOUND = 3&      '  The specified path was not found.
    Const ERROR_BAD_EXE_FORMAT = 193&    '  The .exe file is invalid (non-Win32® .exe or error in .exe image).


Public Function ShellExecLaunchFile(ByVal strPathFile As String, ByVal strOpenInPath As String, ByVal strArguments As String) As Long

    Dim Scr_hDC As Long
    
    'Get the Desktop handle
    Scr_hDC = GetDesktopWindow()
    
    'Launch File
    ShellExecLaunchFile = ShellExecute(Scr_hDC, "Open", strPathFile, "", strOpenInPath, SW_SHOWNORMAL)

End Function

