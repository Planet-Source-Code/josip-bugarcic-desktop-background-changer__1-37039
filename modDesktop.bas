Attribute VB_Name = "modDesktop"
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long


Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40

Public Const SPI_SETDESKWALLPAPER = 20
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2


Enum BrowseForFolderFlags
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_EDITBOX = &H10
    BIF_RETURNFSANCESTORS = &H8
End Enum

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long


Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Public Const SW_RESTORE = 9
Public Const SW_MINIMIZE = 6
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    sTip As String * 64
End Type


Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Sub setWallpaper(ByVal strBitmapImage As String)
Dim lngSuccess As Long
lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, strBitmapImage, SPIF_UPDATEINIFILE)
With SPIF_UPDATEINIFILE = &H1
End With
End Sub

Public Function BrowseForFolder(hwnd As Long, Optional Title As String, Optional Flags As BrowseForFolderFlags) As String

    Dim iNull As Integer
    Dim IDList As Long
    Dim Result As Long
    Dim Path As String
    Dim bi As BrowseInfo
     
    If Flags = 0 Then Flags = BIF_RETURNONLYFSDIRS
     
    With bi
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(Title, "")
        .ulFlags = Flags
    End With

    IDList = SHBrowseForFolder(bi)
     
    If IDList Then
        Path = String$(300, 0)
        Result = SHGetPathFromIDList(IDList, Path)
        iNull = InStr(Path, vbNullChar)
        If iNull Then Path = Left$(Path, iNull - 1)
    End If

    BrowseForFolder = Path
End Function

