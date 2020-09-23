VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Setup"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   8610
   Begin VB.CommandButton cmdNewUser 
      Caption         =   "New User"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdGetPasswords 
      Caption         =   "Show Password"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change Password"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdDeleteUser 
      Caption         =   "Delete User"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdateUser 
      Caption         =   "Update User"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstUsers 
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12632256
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Users"
         Object.Width           =   5646
      EndProperty
   End
   Begin TabDlg.SSTab tabUserInfo 
      Height          =   2775
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "User Security Info"
      TabPicture(0)   =   "frmUserSetup.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtUserName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtPassword"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboSecLvl"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.ComboBox cboSecLvl 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtUserName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Security Level :"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Password :"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "User Name :"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmUserSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LUsersRs As ADODB.Recordset
Dim UserRs As ADODB.Recordset
Dim ChkSecLvlRs As ADODB.Recordset

Private Sub cboSecLvl_LostFocus()
Dim checkrs As ADODB.Recordset
Set checkrs = New ADODB.Recordset
checkrs.Open "select * from login where securitylevel = 99", DB, adOpenStatic, adLockOptimistic
If UserName = "NO USER" And cboSecLvl.Text <> "99" And txtUserName <> "" And checkrs.RecordCount = 0 Then
 MsgBox "Must make at least on user level 99", vbCritical, "Warning"
 cboSecLvl.Text = ""
 cboSecLvl.SetFocus
ElseIf cboSecLvl.Text <> "" And cboSecLvl.Text >= "10" And cboSecLvl.Text <= "99" Then
 cmdNewUser.Enabled = True
 If cmdNewUser.Caption = "Save User" Then
  cmdNewUser.SetFocus
 End If
'ElseIf cboSecLvl.Text = "" Then
' cboSecLvl.SetFocus
End If
End Sub

Private Sub cmdChangePassword_Click()
If SecLevel >= SecSldLvl(5) Then
 frmChangePass.UserName = txtUserName.Text
 frmChangePass.OldPass = txtPassword.Text
 frmChangePass.Show
 
ElseIf SecLevel < SecSldLvl(5) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If
End Sub

Private Sub cmdDeleteUser_Click()
If SecLevel >= SecSldLvl(4) Then
Set ChkSecLvlRs = New ADODB.Recordset
ChkSecLvlRs.Open "Select Count(SecurityLevel) as NumAdmin from login where securitylevel = 99", DB, adOpenStatic, adLockOptimistic

If cboSecLvl.Text <> 99 Then
 If Not UserRs.EOF Or UserRs.BOF Then
  If MsgBox("Are You Sure ?", vbQuestion + vbYesNo) = vbYes Then
   If InputBox("Please Enter Password", "Delete User") = UserRs!Password Then
    UserRs.Delete
   End If
  End If
 End If
 LinkInput
 DoList
 
ElseIf cboSecLvl.Text = 99 Then
  If UserRs.RecordCount = 1 Then
   If Not UserRs.EOF Or UserRs.BOF Then
    If MsgBox("Are You Sure ?", vbQuestion + vbYesNo) = vbYes Then
     If InputBox("Please Enter Password", "Delete User") = UserRs!Password Then
      UserRs.Delete
     End If
    End If
   End If
   LinkInput
   DoList
  ElseIf ChkSecLvlRs!NumAdmin >= 2 Then
   If Not UserRs.EOF Or UserRs.BOF Then
    If MsgBox("Are You Sure ?", vbQuestion + vbYesNo) = vbYes Then
     If InputBox("Please Enter Password", "Delete User") = UserRs!Password Then
      UserRs.Delete
     End If
    End If
   End If
   LinkInput
   DoList
   
  ElseIf ChkSecLvlRs!NumAdmin < 2 Then
  MsgBox "One Person must have Level 99 Access"
  End If 'Num Admin
End If 'CboSecLvl

ElseIf SecLevel < SecSldLvl(4) Then
  MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If

End Sub

Private Sub cmdGetPasswords_Click()
If SecLevel >= SecSldLvl(6) Then
 If cmdGetPasswords.Caption = "Show Password" Then
  cmdGetPasswords.Caption = "Hide Password"
  txtPassword.PasswordChar = ""
 ElseIf cmdGetPasswords.Caption = "Hide Password" Then
  cmdGetPasswords.Caption = "Show Password"
  txtPassword.PasswordChar = "*"
 End If
ElseIf SecLevel < SecSldLvl(6) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If
End Sub

Private Sub cmdNewUser_Click()
If SecLevel >= SecSldLvl(2) Then

 If cmdNewUser.Caption = "New User" Then
   cmdNewUser.Caption = "Save User"
   cboSecLvl.Text = 10
   LinkInput
   UserRs.AddNew
   txtUserName.Enabled = True
   txtPassword.Enabled = True
   cmdNewUser.Enabled = False
   cboSecLvl.Text = ""
   txtUserName.SetFocus
 ElseIf cmdNewUser.Caption = "Save User" Then
   cmdNewUser.Caption = "New User"
   txtUserName.Enabled = False
   txtPassword.Enabled = False
      UserRs.Update
      UserRs.Close
      LinkInput
      DoList
      Exit Sub
 End If 'cmdNewUser

ElseIf SecLevel < SecSldLvl(2) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If
End Sub

Private Sub cmdUpdateUser_Click()
Dim ChkAdminRs As ADODB.Recordset
If SecLevel >= SecSldLvl(3) Then
  Set ChkSecLvlRs = New ADODB.Recordset
  Set ChkAdminRs = New ADODB.Recordset
  ChkAdminRs.Open "Select * from login where Username ='" & UserRs!UserName & "'", DB, adOpenStatic, adLockOptimistic
  ChkSecLvlRs.Open "Select Count(SecLevel) as NumAdmin from login where seclevel = 99", DB, adOpenStatic, adLockOptimistic
  If ChkSecLvlRs!NumAdmin >= 2 Then
    UserRs.Update
    UserRs.Close
    LinkInput
    DoList
  ElseIf ChkSecLvlRs!NumAdmin = 1 And ChkAdminRs!SecLevel <> 99 Then
    UserRs.Update
    UserRs.Close
    LinkInput
    DoList
  ElseIf ChkSecLvlRs!NumAdmin < 2 And cboSecLvl.Text <> 99 Then
   MsgBox "One Person must have Level 99 Access"
   LinkInput
   DoList
  End If
ElseIf SecLevel < SecSldLvl(3) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If

End Sub

Private Sub Form_Activate()
cmdNewUser.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'catch both "Enter" keys on keyboard
    If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeySeparator) Then
        SendKeys "{tab}"
    End If
End Sub
    

Private Sub Form_Load()
Dim r As Integer
Me.Height = 4365
Me.Width = 8700
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 380)
For r = 10 To 99
cboSecLvl.AddItem (r)
Next
LinkInput
DoList

End Sub
Private Sub LinkInput()
Set UserRs = New ADODB.Recordset
UserRs.Open "Select * from login", DB, adOpenStatic, adLockOptimistic
Set txtUserName.DataSource = UserRs
txtUserName.DataField = "UserName"
Set txtPassword.DataSource = UserRs
txtPassword.DataField = "Password"
Set cboSecLvl.DataSource = UserRs
cboSecLvl.DataField = "SecurityLevel"

End Sub
Private Sub lstUsers_Click()
  FromListUpdate
End Sub
Public Sub FromListUpdate()
    On Error GoTo ExiThis
    If Not UserRs.BOF Then UserRs.MoveFirst
    If Not lstUsers.SelectedItem.Text = Empty Then
        UserRs.Find "UserName='" & Trim(lstUsers.SelectedItem.Text) & "'"
    End If
ExiThis:
End Sub

Private Sub txtPassword_LostFocus()

If txtPassword.Text <> "" And txtUserName.Text <> "" Then
 If txtPassword.Text <> InputBox("Please Reenter your password", "Reenter Password") Then
  txtPassword.Text = ""
  txtPassword.SetFocus
 End If
ElseIf txtPassword.Text = "" And txtUserName.Text <> "" Then
 txtPassword.SetFocus
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
     KeyAscii = ValidateInput(KeyAscii, Text_Input, True)
End Sub

Private Sub DoList()


Set LUsersRs = New ADODB.Recordset
LUsersRs.Open "Select username from login order by Username", DB, adOpenStatic, adLockOptimistic
lstUsers.ListItems.Clear
If Not LUsersRs.BOF Then LUsersRs.MoveFirst
Do While Not LUsersRs.EOF
    lstUsers.ListItems.Add , , LUsersRs("UserName")
    LUsersRs.MoveNext
Loop
If LUsersRs.RecordCount = 0 Then
   cmdDeleteUser.Enabled = False
   cmdChangePassword.Enabled = False
   cmdGetPasswords.Enabled = False
   cmdUpdateUser.Enabled = False
   frmMain.mnuLogOut.Enabled = False
   UserName = "NO USER"
ElseIf LUsersRs.RecordCount <> 0 Then
   cmdDeleteUser.Enabled = True
   cmdChangePassword.Enabled = True
   cmdGetPasswords.Enabled = True
   cmdUpdateUser.Enabled = True
   frmMain.mnuLogOut.Enabled = True
End If
lstUsers.Refresh
End Sub

Private Sub txtUserName_LostFocus()
Dim UserXrs As New ADODB.Recordset

Set UserXrs = New ADODB.Recordset
  UserXrs.Open "SELECT * FROM login where username = '" & txtUserName.Text & "'", DB, adOpenStatic, adLockOptimistic
   If UserXrs.RecordCount <> 0 Then
     MsgBox "User Name already Exits, Try another User Name", vbCritical
     txtUserName.Text = ""
     txtPassword.Text = ""
     txtUserName.SetFocus
   End If 'RecodCount

End Sub
