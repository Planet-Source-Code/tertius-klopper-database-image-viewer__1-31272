Attribute VB_Name = "FileFunc"
Option Explicit
Dim m_strPicData As String
Dim m_lngPicLen As Long

Public Function ReadFromFile(iFileName As String, Optional iComp As Boolean) As String
        Open iFileName For Binary Access Read As #1
            m_strPicData = Space$(LOF(1))
            Get #1, , m_strPicData
        Close #1
        If iComp = True Then
         'm_strPicData = HuffmanEncode(m_strPicData, True)
         CompStr m_strPicData, CompLevel
         ReadFromFile = m_strPicData
        ElseIf iComp = False Then
         ReadFromFile = m_strPicData
       End If
End Function


Public Function ReadFromDba(iField As ADODB.Field, Optional iFileTo As String, Optional iUnComp As Boolean) As String
Dim TempFile As String
Dim Pos As Integer
Dim strSize As Long


    m_lngPicLen = iField.ActualSize
    If m_lngPicLen > 0 Then
        m_strPicData = iField.GetChunk(m_lngPicLen)
        If iUnComp = True Then
         'm_strPicData = HuffmanDecode(m_strPicData)
         Pos = InStr(m_strPicData, "-")
         strSize = Val(Left(m_strPicData, Pos - 1))
         m_strPicData = Right(m_strPicData, Len(m_strPicData) - Pos)
         DeCompStr m_strPicData, strSize
        End If
        If iFileTo <> "" Then
         TempFile = UCase(iFileTo)
        ElseIf iFileTo = "" Then
         TempFile = UCase(App.Path & "\tmp.bmp")
        End If
        Open TempFile For Binary As #1
            Put #1, , m_strPicData
        Close #1
        ReadFromDba = TempFile
    Else
        ReadFromDba = ""
    End If

End Function

