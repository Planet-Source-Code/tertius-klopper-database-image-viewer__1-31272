VERSION 5.00
Begin VB.Form frmChangePass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4770
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2438
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1118
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtConfirmPassword 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtNewPassword 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtOldPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm New Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "New Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Old Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ChgUserName As String
Dim OldPassword As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Public Property Let UserName(ByVal vdata As String)
ChgUserName = vdata
End Property

Public Property Let OldPass(ByVal vdata As String)
OldPassword = vdata
End Property

Private Sub cmdOk_Click()
Dim ChgPassRs As ADODB.Recordset

Set ChgPassRs = New ADODB.Recordset
ChgPassRs.Open "Select * from login where username ='" & ChgUserName & "'", DB, adOpenStatic, adLockOptimistic
If ChgPassRs.RecordCount <> 0 Then
   ChgPassRs!Password = txtNewPassword.Text
   ChgPassRs.Update
   ChgPassRs.Close
ElseIf ChgPassRs.RecordCount = 0 Then
 MsgBox "Unable to Change Password"
 ChgPassRs.Close
End If
Unload Me
frmUserSetup.FromListUpdate
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'catch both "Enter" keys on keyboard
    If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeySeparator) Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()
Me.Height = 2025
Me.Width = 4860
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 700)
End Sub

Private Sub txtConfirmPassword_Change()
If txtNewPassword.Text = txtConfirmPassword.Text And txtNewPassword.Text <> "" Then
  cmdOk.Enabled = True
ElseIf txtNewPassword.Text <> txtConfirmPassword.Text Then
  cmdOk.Enabled = False
End If
End Sub

Private Sub txtOldPassword_LostFocus()
 If OldPassword <> txtOldPassword.Text And txtOldPassword.Text <> "" Then
  txtNewPassword.Enabled = False
  txtConfirmPassword.Enabled = False
  txtOldPassword.Text = ""
  txtOldPassword.SetFocus
 ElseIf OldPassword = txtOldPassword.Text Then
  txtNewPassword.Enabled = True
  txtConfirmPassword.Enabled = True
  txtNewPassword.SetFocus
 End If
End Sub
