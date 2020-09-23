Attribute VB_Name = "GlobalDecl"
'Programmer : Tertius Klopper
'Date Last Modified : 23/01/2002
'Program Name : Image Viewer
'Copyright © 2002 Tertius Klopper
'
'Note : This source code is copyrighted and may not be redistrubuted
'any other site that planet-source-code.com. May use this code
'in any of you private applications. But must ask permission
'of programmer before using it in any commersial applications.
'
'Please : Send me any changes that you may make to this
'source code if it add for functions or improve the performance
'of the application. Also please report any bug's that you may find
'in the source code. Please give a detailed description of the
'bug and the conditions that caused that bug. Thank You
'

'Sample Code found on PSC used in this application is
'
'Program Name - InputValidation
'Programmer   - Matt Trigwell
'Comments     - None

'Sample Code found on PSC that help clarify somthing

'Program Name - Cyber Crypt
'Programmer   - Mark Withers
'Comments     - Help to better understand ZLIB

'Progarm Description
'This is a Image Viewer that enable you to view images from a file or
'images that was placed in an Access database. Images that is placed in the
'access database can be proteced by a security level. Allowing only certain
'users access to that image.

'Functions of File based image
'Open single Image and enlarge image on the screen
'Create Thumbnails of entire directory sorting out none picture formats
'View each Thumbnail by itself as single image and enable you to enlarge it
'Create a slideshow from a entire directory or from the thumbnails you
'are viewing.
'Functions of Database images
'Allow you to place images in database based on dirrent image type
'Image types has three levels
'  First Level  - Image Main Type         -  Transportation
'  Second Level - Image Sub Type          -  Vehicles
'  Third Level  - Image Sub Minimum Type  -  Cars
'Import all images from a directory and place them directly in image types
'that you want.
'Export all images saved in a specific image type to a directory
'Allow you to create Thumbnails of images in database
'Allow you to create a SlideShow from images in the database

'Database is also password protected and not able to open with Access
'Enable you to Create Users and assign a security level to that user
'Create main Image types (First Level) and assign a security level to that
'image type, only user above that security level may then access that
'image type

'Multiple level security system placed in the application allows you to
'restrict access to all major function of the application involved with
'the database.

'Note on Pictures in Database : Due to the fact or the nature of picture used
'when they are saved in the database ( use a OLE Field ) there size normaly increase
'by two. That is why I am compressing the picture.
'If for some reasone I am placing the picture wrong in the database
'plz inform me. Or if you have a better and faster compreesion method plz send
'me a copy
'
'Thank You

'If you like this application please vote for it on planet source code

Option Explicit
Global Const dbPass = "jklmo-7125"
Global Const DEFSOURCE = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="
Global Const DBName = ";Jet OLEDB:Database Password=jklmo-7125;"
Public DB As ADODB.Connection
Public dbPath As String
Public NoDba As Boolean
Public UserName As String
Public SecLevel As Integer
Dim NoUsers As Boolean
Public LoginSuccess As Boolean
Public SecSldLvl(0 To 20) As Integer
Public FileName As String
Public frmHeight As Integer
Public frmWidth As Integer
Public OnlyFile As String
Public CompLevel As Integer

Private Type Ext
   Descr As String
   strData As String
End Type

'Public Const BIF_BROWSEFORCOMPUTER = &H1000
'Public Const BIF_BROWSEFORPRINTER = &H2000
'Public Const BIF_BROWSEINCLUDEFILES = &H4000
'Public Const BIF_DONTGOBELOWDOMAIN = &H2
'Public Const BIF_RETURNFSANCESTORS = &H8
'Public Const BIF_RETURNONLYFSDIRS = &H1
'Public Const BIF_STATUSTEXT = &H4

'Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public SuppExt(1 To 7) As Ext

Public Sub ConfigExt()
      SuppExt(1).Descr = "All Images"
      SuppExt(1).strData = ".bmp;*.emf;*.gif;*.ico;*.jpg;*.wmf"
      SuppExt(2).Descr = "BMP Image"
      SuppExt(2).strData = ".bmp"
      SuppExt(3).Descr = "Enhanced Metafile"
      SuppExt(3).strData = ".emf"
      SuppExt(4).Descr = "GIF Image"
      SuppExt(4).strData = ".gif"
      SuppExt(5).Descr = "Icon"
      SuppExt(5).strData = ".ico"
      SuppExt(6).Descr = "JPG Image"
      SuppExt(6).strData = ".jpg"
      SuppExt(7).Descr = "Windows Metafile"
      SuppExt(7).strData = ".wmf"
End Sub

Public Function VerifyFile(FileName As String) As Boolean
Dim fs
 Set fs = CreateObject("Scripting.FileSystemObject")
   If fs.fileexists(FileName) Then
     VerifyFile = True
   Else
     VerifyFile = False
   End If
End Function

Public Function CountImages(sPath As String) As Integer
  On Error Resume Next
  Dim fs, f, fc, f1
  Dim rndpos As Long
  Dim ChkExt As Integer
  Dim Found As Integer
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.GetFolder(sPath)
  Set fc = f.Files
  frmLoading.Caption = "Checking for Supported Images, please wait..."
  For Each f1 In fc
     DoEvents
     For ChkExt = 1 To 7
        If InStr(1, f1.Name, SuppExt(ChkExt).strData, vbTextCompare) Then
          OnlyFile = f1.Name
          Found = Found + 1
        End If
     Next ChkExt
  Next
  CountImages = Found
End Function

Public Sub OpenDB()
On Error GoTo err
Set DB = New ADODB.Connection

LoadDBPath
 If dbPath = "" Then
   With frmMain.comDialog
    .Filter = "Image Database|IMAGEDBA.MDB"
    .ShowOpen
   End With
    dbPath = frmMain.comDialog.FileName
   If frmMain.comDialog.CancelError = True Then
    End
   End If
 End If
If dbPath <> "" Then
  SaveDBPath
  DB.Open DEFSOURCE & dbPath & DBName
End If
err:
 Select Case err.Number
  Case -2147467259
   NoDba = True
    MsgBox "Database " & dbPath & DBName & "could not be found Restore Database", vbYesNo
    
    'Else'
     End
   'End If
  End Select
End Sub
Public Sub CheckUsers()

Dim CheckUser As ADODB.Recordset
Set CheckUser = New ADODB.Recordset
CheckUser.Open "SELECT * FROM LOGIN", DB, adOpenStatic, adLockOptimistic
If CheckUser.RecordCount = 0 Then
 NoUsers = True
 UserName = "NO USER"
 SecLevel = 99
frmMain.Caption = "Image Viewer-User Name :" & UserName & "-" & SecLevel
ElseIf CheckUser.RecordCount <> 0 Then
 NoUsers = False
End If 'CheckUser.RecordCount
If UserName = "NO USER" Then
 frmMain.mnuLogOut.Enabled = False
End If
If NoUsers <> True Then
  frmLogin.Show vbModal
End If 'NoUsers
End Sub
Sub Main()
OpenDB
RetrieveSecLevels
frmMain.Show
CheckUsers
End Sub
Public Sub ConnectDenv()
'To Ensure That the path to the database change if needed
 DEnv.conImage.Open DEFSOURCE & App.Path & DBName
End Sub

Public Function TestNull(iField As ADODB.Field) As String
If IsNull(iField) Then
  TestNull = ""
Else
  TestNull = iField
End If
End Function

Public Function ChgNull(iField As String) As String
'Stop Database from Getting Null Values
If iField = "" Then
  ChgNull = ""
Else
 ChgNull = iField
End If
End Function


Private Sub SaveDBPath()
CreateKey ("HKEY_CURRENT_USER\Software\ImageViewer")
SetStringValue "HKEY_CURRENT_USER\Software\ImageViewer", "DBLocation", dbPath
End Sub

Public Sub LoadDBPath()
dbPath = GetStringValue("HKEY_CURRENT_USER\Software\ImageViewer", "DBLocation")
CompLevel = Val(GetStringValue("HKEY_CURRENT_USER\Software\ImageViewer", "CompLevel"))
If FileEx(dbPath) <> True Then
 dbPath = ""
End If
End Sub


