VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImageType 
   Caption         =   "Image Type Setup"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3113
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   1673
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   233
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin MSComctlLib.Slider sldSecLevel 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   10
      Min             =   10
      Max             =   99
      SelStart        =   10
      Value           =   10
   End
   Begin VB.ComboBox cboImageType 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Security Level Needed to Access Image Type"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Image Type :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmImageType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

