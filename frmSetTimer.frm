VERSION 5.00
Begin VB.Form frmSetTimer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Timer"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDown 
      Height          =   195
      Left            =   1920
      Picture         =   "frmSetTimer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   930
      Width           =   255
   End
   Begin VB.CommandButton cmdUp 
      Height          =   195
      Left            =   1920
      Picture         =   "frmSetTimer.frx":0096
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Specify number of seconds (no more then 30000)"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmSetTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strTime As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDown_Click()
If CDbl(txtTime.Text) > 1 Then
txtTime.Text = CDbl(txtTime.Text) - 1
End If

End Sub

Private Sub cmdOK_Click()
strTime = txtTime.Text
frmDesktop.Timer1.Interval = CDbl(strTime) * 1000
Unload Me

End Sub

Private Sub cmdUp_Click()
If CDbl(txtTime.Text) < 30000 Then
txtTime.Text = CDbl(txtTime.Text) + 1
End If

End Sub

Private Sub Form_Load()
strTime = frmDesktop.Timer1.Interval / 1000
Me.txtTime.Text = strTime

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmSetTimer = Nothing
End Sub

Private Sub txtTime_GotFocus()
txtTime.SelStart = 0
txtTime.SelLength = Len(txtTime.Text)

End Sub

Private Sub txtTime_Validate(Cancel As Boolean)
If Not IsNumeric(txtTime.Text) Then
    txtTime.Text = strTime
    Cancel = True
Else
    If CDbl(txtTime.Text) > 30000 Then
        txtTime.Text = strTime
        Cancel = True
    End If
    
    If CDbl(txtTime.Text) <= 0 Then
        txtTime.Text = strTime
        Cancel = True
    End If

End If

End Sub
