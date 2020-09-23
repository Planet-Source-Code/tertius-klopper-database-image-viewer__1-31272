VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Image Viewer"
   ClientHeight    =   5790
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9915
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   120
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogOut 
         Caption         =   "&Log Out"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Begin VB.Menu mnuImage 
            Caption         =   "&Image"
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSlideShow 
         Caption         =   "&SlideShow"
      End
      Begin VB.Menu mnuThumbNails 
         Caption         =   "&Thumbnails"
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBSlideshow 
         Caption         =   "&DB SlideShow"
      End
      Begin VB.Menu mnuDBThumNails 
         Caption         =   "D&B Thumnails"
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportAllImg 
         Caption         =   "&Import Images"
      End
      Begin VB.Menu mnuExportAllImg 
         Caption         =   "&Export Images"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuImageType 
         Caption         =   "&Image Types"
      End
      Begin VB.Menu mnuUserSetup 
         Caption         =   "&User Setup"
      End
      Begin VB.Menu mnuSeclvlSetup 
         Caption         =   "&Securirty Level Setup"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuBackupDBA 
         Caption         =   "&Backup Database"
      End
      Begin VB.Menu mnuRestoreDBA 
         Caption         =   "&Restore Database"
      End
      Begin VB.Menu mnusep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompactDBA 
         Caption         =   "&Compact Database"
      End
      Begin VB.Menu mnusep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgSettings 
         Caption         =   "Porgram Settings"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "&PopUpMenus"
      Visible         =   0   'False
      Begin VB.Menu mnuDba 
         Caption         =   "NormalThumb"
         Begin VB.Menu mnuAddToDba 
            Caption         =   "&Add To Database"
         End
         Begin VB.Menu mnuSep5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDelImage 
            Caption         =   "&Delete Image"
         End
      End
      Begin VB.Menu mnuDbaThumb 
         Caption         =   "DBAThumb"
         Begin VB.Menu mnuChgImgType 
            Caption         =   "&Change Image Type"
         End
         Begin VB.Menu mnuDelfromDba 
            Caption         =   "&Delete Image from Database"
         End
         Begin VB.Menu mnusep6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExportImg 
            Caption         =   "Export Image to File"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPath As String

Private Sub MDIForm_Load()
ConfigExt
End Sub

Private Sub mnuAddToDba_Click()
frmAddToDBA.AddtoDba FileName
End Sub

Private Sub mnuBackupDBA_Click()
frmBackupDba.Show
End Sub

Private Sub mnuChgImgType_Click()
'Change Image Types and Sub Types
frmModifyImage.LookUpImage frmDBThumbnails.ImgID
End Sub

Private Sub mnuCompactDBA_Click()
frmCompactDba.Show
End Sub

Private Sub mnuDBSlideshow_Click()
frmSelectImageTypes.DisplayOpt = "Slideshow"
frmSelectImageTypes.Show

End Sub

Private Sub mnuDBThumNails_Click()
frmMain.mnuDBThumNails.Enabled = False
frmSelectImageTypes.DisplayOpt = "Thumbnails"
frmSelectImageTypes.Show
End Sub

Private Sub mnuDelfromDba_Click()
'Delete Image from Database
frmDBThumbnails.DelImgFromDBA
End Sub

Private Sub mnuDelImage_Click()
  If MsgBox("Are You Sure", vbQuestion + vbYesNo, "Delete Image") = vbYes Then
    Kill FileName
    Unload frmThumbs
    ShowThumbs sPath & "\"
  End If
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuExportAllImg_Click()
If SecLevel >= SecSldLvl(14) Then
frmExportAllImg.Show
ElseIf SecLevel < SecSldLvl(14) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If

End Sub

Private Sub mnuExportImg_Click()
'Export Image To a File
frmDBThumbnails.ExportImage
End Sub

Private Sub mnuImage_Click()
With comDialog
 .Filter = "All Graphic Files|*.bmp;*.jpg;*.gif;*.emf;*.wmf|Windows Bitmaps|*.bmp|" & _
 "JPEG Filter|*.jpg|GIF Filter|*.gif|Windows Metafile|*.wmf|" & _
 "Enhanced Metafile|*.emf"
 .ShowOpen
End With
If comDialog.FileName <> "" Then
 OpenImage comDialog.FileName
End If
End Sub
Public Sub OpenImage(FileName As String)
  Dim NewImage As frmImages
  Set NewImage = New frmImages
     With NewImage
       .Caption = FileName
       .WindowState = vbMaximized
       frmWidth = .Width
       frmHeight = .Height
       .LoadImage FileName
       .CenterImage
       .Show
      If .Tag = "error" Then Unload NewImage
     End With
End Sub

Private Sub mnuImageType_Click()
frmImageTypes.Show
End Sub

Private Sub mnuImportAllImg_Click()
If SecLevel >= SecSldLvl(14) Then

frmImportAllImg.Show
ElseIf SecLevel < SecSldLvl(14) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If
End Sub

Private Sub mnuLogOut_Click()
RetrieveSecLevels
frmLogin.Show vbModal
End Sub

Private Sub mnuProgSettings_Click()
frmAppOptions.Show
End Sub

Private Sub mnuRestoreDBA_Click()
frmRestoreDba.Show
End Sub

Private Sub mnuSeclvlSetup_Click()
frmSecLvl.Show
End Sub

Private Sub mnuSlideShow_Click()
Dim sPath As String
   sPath = fBrowseForFolder(Me.hwnd, "Select SlideShow Folder")
   If sPath <> "" Then StartSlideShow sPath & "\"
End Sub
Private Sub StartSlideShow(sPath As String)
Dim NumOfPics As Integer
NumOfPics = 0
frmLoading.Show
With frmSlideShow
   NumOfPics = .ScanFolder(sPath)
   If NumOfPics > 1 Then
    .NextPic
    .CenterImage
    .WindowState = vbMaximized
    .Show
    .StartSlideShow True
   End If
End With
Unload frmLoading
End Sub
Private Sub ShowThumbs(sPath As String)
Dim NumOfPics As Integer
frmLoading.Show
With frmThumbs
   NumOfPics = .ScanFolder(sPath)
  .Caption = "" ' "Thumbnails : Total Images : " & NumOfPics
  .DisplayThumbs
  .Show
  frmMain.mnuThumbNails.Enabled = False
End With
Unload frmLoading
End Sub

Private Sub mnuThumbNails_Click()
sPath = fBrowseForFolder(Me.hwnd, "Select folder to display as thumbnails")
If sPath <> "" Then ShowThumbs sPath & "\"
End Sub

Private Sub mnuUserSetup_Click()
frmUserSetup.Show
End Sub
