Attribute VB_Name = "SecFunc"
Option Explicit

Public Sub RetrieveSecLevels()
Dim SecLvlRs As ADODB.Recordset
Set SecLvlRs = New ADODB.Recordset
SecLvlRs.Open "Select * from SecLevels", DB, adOpenStatic, adLockOptimistic

 If SecLvlRs.RecordCount <> 0 Then
    SecSldLvl(0) = SecLvlRs!viewseclvl
    SecSldLvl(1) = SecLvlRs!chgseclvl
    SecSldLvl(2) = SecLvlRs!addnewuser
    SecSldLvl(3) = SecLvlRs!upduser
    SecSldLvl(4) = SecLvlRs!delUser
    SecSldLvl(5) = SecLvlRs!chgpass
    SecSldLvl(6) = SecLvlRs!unhidePass
    SecSldLvl(7) = SecLvlRs!ViewImgType
    SecSldLvl(8) = SecLvlRs!AddImgType
    SecSldLvl(9) = SecLvlRs!DeleteImgType
    SecSldLvl(10) = SecLvlRs!AddImgtoDBA
    SecSldLvl(11) = SecLvlRs!ModifyImgInDba
    SecSldLvl(12) = SecLvlRs!DeleteImgFromDba
    SecSldLvl(13) = SecLvlRs!ExportImgFromDba
    SecSldLvl(14) = SecLvlRs!ImportAllImg
    SecSldLvl(15) = SecLvlRs!ExportAllImg
    SecSldLvl(16) = SecLvlRs!BackupDBA
    SecSldLvl(17) = SecLvlRs!RestoreDBA
    SecSldLvl(18) = SecLvlRs!CompactDBA
     End If
 
End Sub

Public Sub EnableSecLevels()
If UserName <> "NO USER" And UserName <> "OVERRIDE" Then

 If SecLevel >= SecSldLvl(0) Then
   frmMain.mnuSeclvlSetup.Enabled = True
 ElseIf SecLevel < SecSldLvl(0) Then
   frmMain.mnuSeclvlSetup.Enabled = False
 End If
 
 If SecLevel >= SecSldLvl(7) Then
   frmMain.mnuImageType.Enabled = True
 ElseIf SecLevel < SecSldLvl(7) Then
   frmMain.mnuImageType.Enabled = False
 End If
 If SecLevel >= SecSldLvl(10) Then
   frmMain.mnuAddToDba.Enabled = True
 ElseIf SecLevel < SecSldLvl(10) Then
   frmMain.mnuAddToDba.Enabled = False
 End If
 
End If 'Username
End Sub

