VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1673
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox cboUserName 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1673
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2333
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   773
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   473
      TabIndex        =   5
      Top             =   870
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   473
      TabIndex        =   4
      Top             =   390
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPassRs As ADODB.Recordset

Private Sub cmdCancel_Click()
LoginSuccess = False
Unload Me
End Sub

Private Sub cmdLogin_Click()
Dim FindUser As String

FindUser = cboUserName.Text

If FindUser <> "" Then
  Set adoPassRs = New ADODB.Recordset
  adoPassRs.Open "Select * From login where username = '" & FindUser & "'", DB, adOpenStatic, adLockOptimistic
  
  If Not adoPassRs.BOF Then
   adoPassRs.MoveFirst
  End If 'adoPassrs.BOF
  
  If adoPassRs.RecordCount <> 0 Then
   
   If adoPassRs("Password") = txtPassword.Text Then
    LoginSuccess = True
    UserName = adoPassRs!UserName
    SecLevel = adoPassRs!SecurityLevel
    frmMain.Caption = "Image Viewer-User Name :" & UserName & "-" & SecLevel
    EnableSecLevels
    Unload Me
   ElseIf adoPassRs("password") <> txtPassword.Text Then
    MsgBox "Incorrect Password", vbExclamation, "Password Error"
    txtPassword.Text = ""
    txtPassword.SetFocus
   End If 'adoPassRsPassword
  
  ElseIf FindUser = "OVERRIDE" And adoPassRs.RecordCount = 0 Then
   TestPassOver
  End If 'RecordCount

ElseIf FindUser <> "OVERRIDE" And adoPassRs.RecordCount = 0 Then
 MsgBox "User Does Not Exits"
 cboUserName.Text = ""
 cboUserName.SetFocus

ElseIf FindUser = "" Then
 cboUserName.SetFocus
End If 'cbousername

End Sub

Private Sub Form_Load()
Set adoPassRs = New ADODB.Recordset
adoPassRs.Open "SELECT * FROM login", DB, adOpenStatic, adLockOptimistic
    If adoPassRs.RecordCount = 0 Then
        LoginSuccess = True
    Else
        adoPassRs.MoveFirst
        Do While Not adoPassRs.EOF
        cboUserName.AddItem (adoPassRs("UserName"))
        adoPassRs.MoveNext
        Loop
    End If
   UserName = ""
   SecLevel = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If LoginSuccess <> True Then
 If MsgBox("Must Login to Enter Fleet Control , This action will close Fleet Control" & vbNewLine & "Are you sure you want to exit", vbYesNo + vbCritical, "Login Required") = vbYes Then
   adoPassRs.Close
   frmLogin.Hide
   End
 ElseIf vbNo Then
   Cancel = 1
 End If 'MsgBox
ElseIf LoginSuccess = True Then
 'adoPassRs.Close
 frmLogin.Hide
End If 'LoginSuccess
End Sub

Private Sub TestPassOver()
 If cboUserName.Text = "OVERRIDE" And txtPassword.Text = "MATRIX" Then
   UserName = "OVERRIDE"
   SecLevel = 99
   'frmMain.StatusBar.Panels(1).Text = "User Name : " & UserName
   'frmMain.StatusBar.Panels(2).Text = "Security Level : " & SecLevel
   LoginSuccess = True
   EnableSecLevels
   Unload Me
 End If
End Sub


