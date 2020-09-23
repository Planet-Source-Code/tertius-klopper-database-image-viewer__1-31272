VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmImageTypes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Types"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6300
   Begin MSComctlLib.ProgressBar prgLoad 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdImgType 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4683
      _Version        =   393216
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Image type"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdNewImgType 
      Caption         =   "Insert Image Type"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "frmImageTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim ImgTypeDelRs As ADODB.Recordset
Dim ImgType As String
On Error GoTo DelError
If SecLevel >= SecSldLvl(9) Then
Set ImgTypeDelRs = New ADODB.Recordset
grdImgType.Row = grdImgType.Row
grdImgType.Col = 1: ImgType = grdImgType.Text

ImgTypeDelRs.Open "Select * from imageType where imagetype = '" & ImgType & "'", DB, adOpenStatic, adLockOptimistic
If ImgTypeDelRs.RecordCount <> 0 Then
   ImgTypeDelRs.Delete
End If
ClearGrid
CreateGrid
GetImgTypeData
ElseIf SecLevel < SecSldLvl(9) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If
DelError:
If err.Number = -2147217887 Then
 MsgBox "Record Cannot be Deleted because it has Images Related to it, Remove Them First", vbCritical, "Error Deleting"

End If
End Sub

Private Sub cmdNewImgType_Click()
Dim ImgTypeRs As ADODB.Recordset
Dim r As Integer
If SecLevel >= SecSldLvl(8) Then

If cmdNewImgType.Caption = "Insert Image Type" Then
cmdNewImgType.Caption = "Save Image Type"
ClearGrid
CreateGrid
cmdDelete.Enabled = False
grdImgType.SetFocus
grdImgType.Col = 1
grdImgType.Rows = grdImgType.Rows
grdImgType.Row = 1
grdImgType_EnterCell
ElseIf cmdNewImgType.Caption = "Save Image Type" Then
cmdNewImgType.Caption = "Insert Image Type"
cmdDelete.Enabled = True

Set ImgTypeRs = New ADODB.Recordset
ImgTypeRs.Open "Select * from ImageType", DB, adOpenStatic, adLockOptimistic
For r = 1 To grdImgType.Rows - 2
grdImgType.Row = r
 With ImgTypeRs
  .AddNew
  grdImgType.Col = 1: !ImageType = grdImgType.Text
  grdImgType.Col = 2: !seclvl = Val(grdImgType.Text)
  .Update
 End With
Next
 ClearGrid
 CreateGrid
 GetImgTypeData
End If
ElseIf SecLevel < SecSldLvl(8) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If
End Sub

Private Sub GetImgTypeData()
Dim r As Integer
Dim GetImgTypeRs As ADODB.Recordset
prgLoad.Value = 0
Set GetImgTypeRs = New ADODB.Recordset
GetImgTypeRs.Open "Select * from ImageType", DB, adOpenStatic, adLockOptimistic
If GetImgTypeRs.RecordCount <> 0 Then
  prgLoad.Max = GetImgTypeRs.RecordCount
   For r = 1 To GetImgTypeRs.RecordCount
     grdImgType.Row = r
    With GetImgTypeRs
     grdImgType.Rows = grdImgType.Rows + 1
     prgLoad.Value = r
     grdImgType.Col = 1: grdImgType.Text = !ImageType
     grdImgType.Col = 2: grdImgType.Text = !seclvl
     .MoveNext
    End With
   Next
ElseIf GetImgTypeRs.RecordCount = 0 Then
  prgLoad.Max = 1
End If
If grdImgType.Rows > 2 Then
  grdImgType.Rows = grdImgType.Rows - 1
End If
End Sub
Private Sub ClearGrid()
grdImgType.Clear
txtEdit.Visible = False
End Sub


Private Sub Form_Load()
Me.Height = 3885
Me.Width = 6390
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 380)
ClearGrid
CreateGrid
GetImgTypeData
End Sub
Private Sub CreateGrid()
grdImgType.Cols = 3
grdImgType.Rows = 2
grdImgType.Row = 0

grdImgType.Col = 1: grdImgType.Text = "Image Type"
grdImgType.Col = 2: grdImgType.Text = "Access level"
grdImgType.ColWidth(0) = 300

grdImgType.ColWidth(2) = Len("Access Level") * 95 + 250
grdImgType.ColWidth(1) = grdImgType.Width - grdImgType.ColWidth(2) - 600
End Sub

Private Sub grdImgType_EnterCell()
  Select Case grdImgType.Col
    Case 1, 2
     If grdImgType.Row = grdImgType.Rows - 1 Then
      With txtEdit
       .Move grdImgType.CellLeft + grdImgType.Left, _
        grdImgType.CellTop + grdImgType.Top, grdImgType.CellWidth - 25, _
        grdImgType.CellHeight - 25
       .Text = grdImgType.Text
       If Len(.Text) > 0 Then
        .SelStart = 0
        .SelLength = Len(.Text)
       End If
        .Visible = True
        .ZOrder 0
        .SetFocus
       End With
     End If
  End Select
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ImgTypeRs As ADODB.Recordset
    Select Case KeyCode
      Case vbKeyEscape
        With txtEdit
          .Text = Empty
          .Visible = False
        End With
          grdImgType.SetFocus
      Case vbKeyReturn
        With txtEdit
         If .Text = Empty And grdImgType.Col = 1 Then
          MsgBox "Must Supply a Image Type"
          .SetFocus
          Exit Sub
         ElseIf .Text = Empty And grdImgType.Col = 2 Then
          MsgBox "Must Supply a Access Level"
          .SetFocus
          Exit Sub
         ElseIf Not .Text = Empty Then
           Select Case grdImgType.Col
            Case 1
             Set ImgTypeRs = New ADODB.Recordset
             ImgTypeRs.Open "Select * from imagetype where imagetype = '" & .Text & "'", DB, adOpenStatic, adLockOptimistic
              If ImgTypeRs.RecordCount <> 0 Then
               MsgBox "Image type Already Exists"
               .SetFocus
               Exit Sub
              End If
            Case 2
              If Val(.Text) < 10 Or Val(.Text) > 99 Then
                MsgBox "Access Level Must be between 10 and 99"
                .SetFocus
                Exit Sub
              End If
           
           End Select
           grdImgType.Text = .Text
         End If
          .Visible = False
          .Text = Empty
        
        End With
        Select Case grdImgType.Col
         Case 1
          grdImgType.Col = 2
          grdImgType_EnterCell
         Case 2
          If grdImgType.Col = 2 And grdImgType.Text <> "" Then
            grdImgType.Rows = grdImgType.Rows + 1
            grdImgType.Row = grdImgType.Row + 1
            grdImgType.Col = 1
            grdImgType_EnterCell
         End If
        End Select



    End Select
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If grdImgType.Col = 1 Then
     KeyAscii = ValidateInput(KeyAscii, Text_Input, True)
ElseIf grdImgType.Col = 2 Then
     KeyAscii = ValidateInput(KeyAscii, Numeric_Input)
End If

End Sub

