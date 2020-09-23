VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSecLvl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Levels"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8070
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8705
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Security Settings"
      TabPicture(0)   =   "frmSecLvl.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Image Settings"
      TabPicture(1)   =   "frmSecLvl.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Database Functions"
      TabPicture(2)   =   "frmSecLvl.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame6 
         Caption         =   "Imports and Export"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   40
         Top             =   600
         Width           =   7575
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   14
            Left            =   2280
            TabIndex        =   43
            Top             =   315
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   15
            Left            =   2280
            TabIndex        =   44
            Top             =   795
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin VB.Label Label19 
            Caption         =   "Export Images into Directory :"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label18 
            Caption         =   "Import Image into Database :"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Functions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   35
         Top             =   2160
         Width           =   7575
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   16
            Left            =   2280
            TabIndex        =   36
            Top             =   360
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   17
            Left            =   2280
            TabIndex        =   45
            Top             =   840
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   18
            Left            =   2280
            TabIndex        =   46
            Top             =   1320
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin VB.Label Label17 
            Caption         =   "Compact Database :"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1365
            Width           =   1935
         End
         Begin VB.Label Label16 
            Caption         =   "Restore Database :"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   885
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "Backup Database :"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   405
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Database Image Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74880
         TabIndex        =   25
         Top             =   2280
         Width           =   7575
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   10
            Left            =   2280
            TabIndex        =   31
            Top             =   360
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   11
            Left            =   2280
            TabIndex        =   32
            Top             =   840
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   12
            Left            =   2280
            TabIndex        =   33
            Top             =   1320
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   13
            Left            =   2280
            TabIndex        =   34
            Top             =   1800
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin VB.Label Label10 
            Caption         =   "Modify Image in Database :"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   878
            Width           =   2055
         End
         Begin VB.Label Label14 
            Caption         =   "Export Image from Database :"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1845
            Width           =   2055
         End
         Begin VB.Label Label13 
            Caption         =   "Delete Image from Database :"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1358
            Width           =   2175
         End
         Begin VB.Label Label12 
            Caption         =   "Add Image to Database"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   405
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Image Types"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   7575
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   7
            Left            =   2280
            TabIndex        =   20
            Top             =   322
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   8
            Left            =   2280
            TabIndex        =   21
            Top             =   762
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   9
            Left            =   2280
            TabIndex        =   30
            Top             =   1200
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin VB.Label Label9 
            Caption         =   "Add Image Type :"
            Height          =   255
            Left            =   165
            TabIndex        =   24
            Top             =   800
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "View Image Types :"
            Height          =   255
            Left            =   165
            TabIndex        =   23
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Delete Image Type :"
            Height          =   255
            Left            =   165
            TabIndex        =   22
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "User Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   7575
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   2
            Left            =   2280
            TabIndex        =   9
            Top             =   322
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   3
            Left            =   2280
            TabIndex        =   10
            Top             =   772
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   4
            Left            =   2280
            TabIndex        =   11
            Top             =   1222
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   5
            Left            =   2280
            TabIndex        =   12
            Top             =   1672
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   6
            Left            =   2280
            TabIndex        =   13
            Top             =   2122
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin VB.Label Label7 
            Caption         =   "Unhide of User Password :"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Changing of User Password :"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1710
            Width           =   2175
         End
         Begin VB.Label Label5 
            Caption         =   "Deleting of Users :"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1260
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Updating of Users :"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   810
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Adding of New User :"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Security Levels"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   7575
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   0
            Left            =   2280
            TabIndex        =   6
            Top             =   322
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldSecLevel 
            Height          =   330
            Index           =   1
            Left            =   2280
            TabIndex        =   7
            Top             =   802
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   582
            _Version        =   393216
            Min             =   10
            Max             =   99
            SelStart        =   10
            Value           =   10
         End
         Begin VB.Label Label2 
            Caption         =   "View Security Levels :"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Change Security Levels :"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   1815
         End
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOKSave 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2895
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "frmSecLvl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOKSave_Click()
If SecLevel >= SecSldLvl(1) Then
If MsgBox("This will save new Security Levels, Continue", vbExclamation + vbYesNo) = vbYes Then
  SaveSecLvl
  Unload Me
End If
ElseIf SecLevel < SecSldLvl(1) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
 LoadSecLvl
End If
End Sub

Private Sub Form_Load()
Me.Height = 5985
Me.Width = 8160
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 360)
LoadSecLvl
End Sub
Private Sub LoadSecLvl()
Dim r As Integer
RetrieveSecLevels
For r = 0 To 18
sldSecLevel(r).Value = SecSldLvl(r)
Next
End Sub

Private Sub SaveSecLvl()
Dim SecLvlRs As ADODB.Recordset
Set SecLvlRs = New ADODB.Recordset
SecLvlRs.Open "select * from seclevels", DB, adOpenStatic, adLockOptimistic

If SecLvlRs.RecordCount <> 0 Then
 SecLvlRs!viewseclvl = sldSecLevel(0).Value
 SecLvlRs!chgseclvl = sldSecLevel(1).Value
 SecLvlRs!addnewuser = sldSecLevel(2).Value
 SecLvlRs!upduser = sldSecLevel(3).Value
 SecLvlRs!delUser = sldSecLevel(4).Value
 SecLvlRs!chgpass = sldSecLevel(5).Value
 SecLvlRs!unhidePass = sldSecLevel(6).Value
 SecLvlRs!ViewImgType = sldSecLevel(7).Value
 SecLvlRs!AddImgType = sldSecLevel(8).Value
 SecLvlRs!DeleteImgType = sldSecLevel(9).Value
 SecLvlRs!AddImgtoDBA = sldSecLevel(10).Value
 SecLvlRs!ModifyImgInDba = sldSecLevel(11).Value
 SecLvlRs!DeleteImgFromDba = sldSecLevel(12).Value
 SecLvlRs!ExportImgFromDba = sldSecLevel(13).Value
 SecLvlRs!ImportAllImg = sldSecLevel(14).Value
 SecLvlRs!ExportAllImg = sldSecLevel(15).Value
 SecLvlRs!BackupDBA = sldSecLevel(16).Value
 SecLvlRs!RestoreDBA = sldSecLevel(17).Value
 SecLvlRs!CompactDBA = sldSecLevel(18).Value

 SecLvlRs.Update
 SecLvlRs.Close
ElseIf SecLvlRs.RecordCount = 0 Then
 SecLvlRs.AddNew
 SecLvlRs!viewseclvl = sldSecLevel(0).Value
 SecLvlRs!chgseclvl = sldSecLevel(1).Value
 SecLvlRs!addnewuser = sldSecLevel(2).Value
 SecLvlRs!upduser = sldSecLevel(3).Value
 SecLvlRs!delUser = sldSecLevel(4).Value
 SecLvlRs!chgpass = sldSecLevel(5).Value
 SecLvlRs!unhidePass = sldSecLevel(6).Value
 SecLvlRs!ViewImgType = sldSecLevel(7).Value
 SecLvlRs!AddImgType = sldSecLevel(8).Value
 SecLvlRs!DeleteImgType = sldSecLevel(9).Value
 SecLvlRs!AddImgtoDBA = sldSecLevel(10).Value
 SecLvlRs!ModifyImgInDba = sldSecLevel(11).Value
 SecLvlRs!DeleteImgFromDba = sldSecLevel(12).Value
 SecLvlRs!ExportImgFromDba = sldSecLevel(13).Value
 SecLvlRs!ImportAllImg = sldSecLevel(14).Value
 SecLvlRs!ExportAllImg = sldSecLevel(15).Value
 SecLvlRs!BackupDBA = sldSecLevel(16).Value
 SecLvlRs!RestoreDBA = sldSecLevel(17).Value
 SecLvlRs!CompactDBA = sldSecLevel(18).Value
 SecLvlRs.Update
 SecLvlRs.Close
End If
End Sub

