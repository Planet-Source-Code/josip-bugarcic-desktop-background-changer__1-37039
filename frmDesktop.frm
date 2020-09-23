VERSION 5.00
Begin VB.Form frmDesktop 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7800
   Icon            =   "frmDesktop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6600
      Top             =   4920
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuSetTimer 
         Caption         =   "Set Timer"
      End
      Begin VB.Menu mnuChooseFolder 
         Caption         =   "Choose Folder"
      End
   End
End
Attribute VB_Name = "frmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCount As Integer
Dim maxCount As Integer

Dim colBMP As Collection

Dim MyPath As String

Dim Tic As NOTIFYICONDATA

Private Sub Form_Load()

Dim intT As Integer
Dim strP As String

intT = CInt(GetSetting("DesktopChange", "StartUp", "timer", "3000"))
strP = GetSetting("DesktopChange", "StartUp", "folder", "c:\")

SetSysTray NIM_ADD

MyPath = strP
Timer1.Interval = intT

BrowseFolder (MyPath)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Dim msg As Long
        Dim sFilter As String
        msg = X / Screen.TwipsPerPixelX
        Select Case msg
            Case WM_RBUTTONUP
                PopupMenu mnuMain
            Case WM_LBUTTONDBLCLK
                
        End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "DesktopChange", "StartUp", "folder", MyPath
SaveSetting "DesktopChange", "StartUp", "timer", Timer1.Interval


Shell_NotifyIcon NIM_DELETE, Tic
End Sub

Private Sub mnuChooseFolder_Click()
Dim strValue As String

strValue = BrowseForFolder(Me.hwnd, "Choose a folder which contains BMP files:", BIF_DONTGOBELOWDOMAIN)

If strValue <> "" Then
    MyPath = strValue & "\"
End If

BrowseFolder (MyPath)
End Sub

Private Sub mnuClose_Click()
Unload Me

Set frmDesktop = Nothing

End Sub

Private Sub mnuSetTimer_Click()
frmSetTimer.Show
End Sub

Private Sub Timer1_Timer()

If intCount > colBMP.Count Then intCount = 1
setWallpaper colBMP(intCount)
intCount = intCount + 1
End Sub

Private Sub BrowseFolder(mPath As String)


Dim MyName As String
Dim i As Integer

MyName = Dir(mPath)

i = 0
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
      
        If (GetAttr(mPath & MyName) And vbDirectory) <> vbDirectory Then
            If UCase(Right(MyName, 4)) = ".BMP" Then
                If i = 0 Then
                    If IsObject(colBMP) Then Set colBMP = Nothing
                    Set colBMP = New Collection
                End If
                colBMP.Add (mPath & MyName)
                i = i + 1
            End If
        End If
    End If
   MyName = Dir
Loop


intCount = 1

Select Case i
    Case 0
        MsgBox "You don't have any .BMP file in selected folder"
        Exit Sub
    Case 1
        setWallpaper colBMP(1)
    Case Else
        Timer1.Enabled = True
        setWallpaper colBMP(intCount)
        intCount = intCount + 1
End Select

End Sub

Sub SetSysTray(ByVal nNIM As Integer)


Dim rc As Long

Tic.cbSize = Len(Tic)
Tic.hwnd = Me.hwnd
Tic.uID = vbNull
Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE


Tic.hIcon = Me.Icon
Tic.sTip = "Desktop Background Changer - www.bugarcic.net"

rc = Shell_NotifyIcon(nNIM, Tic)


End Sub

