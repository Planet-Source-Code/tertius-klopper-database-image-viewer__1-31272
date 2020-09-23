Attribute VB_Name = "HuffmanCoding"
'################################################################
' Huffman Coding Compression / Decompression Algorithm
' Created 1 August 2000 by James Vincent Carnicelli
'
' NOTES
'
' The Huffman algorithm, named after its inventor, was created
' around about 1952.  It's the method used by most commercial
' compression utilities, like PKZIP, and even by the JPEG image
' file format.  It's generally thought to offer an average of
' 50% compression, given a typical mix of text and binary data.
' For long strings that contain lots of repeating characters or
' only a handful of different characters, the compression ratio
' could get as high as 80%.  While efficient, this algorithm is
' not guaranteed to result in a compressed string that is
' smaller than the original source.
'
' This is a less-than-optimal implementation of this compression
' algorithm.  It's very simple to use in your programs (even if
' it is difficult to understand how it works).  You need only
' call:
'
'         Compressed = HuffmanEncode(SourceText, [Force])
'
' passing in the text you want compressed.  If the compressed
' version is actually larger than the original source, this
' algorithm spits out a special string that contains a four-
' byte header and the original source string, so the resulting
' string will always be at most four bytes larger than the
' source string.  If you pass in True for Force, the result will
' always be huffman-encoded, bypassing this neat optimization.
' Be aware that the output is binary data, so it might not work
' nicely with some things like text boxes, certain Windows
' API calls, and some SQL and database field formats.
'
' To decode a string encoded with HuffmanEncode, simply call
' the following:
'
'         UncompressedText = HuffmanDecode(Compressed)
'
' One cool application of this algorithm is encryption.  Because
' Huffman coding relies on variable-bit-length character
' representations, it's next to impossible to decrypt a string
' compressed with this algorithm without recognizing the
' lookup tables in the header as the key to decrypting it.  You
' could even strip out this lookup table and keep it as a
' private key to be shared only with those you want.  Without
' the lookup table, even someone equiped with this very code
' would not likely be able to decrypt the string.
'
' One last thing.  While I've tested this algorithm with plain
' text strings and even some binary files, I don't know how
' much data you can cram into the compression engine before it
' breaks.  With luck, it's something like 2GB.  In that case,
' though, this would be pretty slow.  Also, I have not proven
' beyond a doubt that this won't choke on some data, so I would
' encourage you to do so to your satisfaction before putting
' this into full production.  Be sure to let me know if you find
' anything interesting.
'################################################################



'NOTES BY Tertius Klopper
'
'NOT USED IN THIS APPLICATION - THIS IS TO SLOW A WAY TO COMPRESS
'THE LONG BINARY DATA OF AN OLE OBJECT
'
'THIS WHOULD WORK BETTER IF YOU WOULD LIKE TO ENCRYPT THE OLE OBJECT
'BUT WILL SLOW DOWN THE APPLICATION ABOUT 70 % ON SLOWER COMPUTERS
Option Explicit

Private Enum HuffmanTreeNodeParts
    htnWeight = 1
    htnIsLeaf = 2
    htnAsciiCode = 3
    htnBitCode = 4
    htnLeftSubtree = 5
    htnRightSubtree = 6
End Enum


'Compress the text.
Public Function HuffmanEncode(Text As String, Optional Force As Boolean) As String
    Dim TextLen As Long, Char As Byte, i As Long, j As Long
    Dim CodeCounts(255) As Long, BitStrings(255), BitString
    Dim HuffmanTrees As Collection
    Dim HTRootNode As Collection, HTNode As Collection
    Dim NextByte As Byte, BitPos As Integer, Temp As String
    
    'Initialize for processing.
    TextLen = Len(Text)
    Set HuffmanTrees = New Collection
    
    'Is there anything to encode?
    If TextLen = 0 Then
        HuffmanEncode = "HE0" & vbCr  'Version 0 = Plain text
        Exit Function  'No point in continuing
    End If
    
    HuffmanEncode = "HE2" & vbCr  'Version 1
    
    'Count how many times each ASCII code is encountered in text.
    For i = 1 To TextLen
        Char = Asc(Mid(Text, i, 1))
        CodeCounts(Char) = CodeCounts(Char) + 1
    Next
    
    'Initialize the forest of Huffman trees; one for each ASCII
    'character used.
    For i = 0 To UBound(CodeCounts)
        If CodeCounts(i) > 0 Then
            Set HTNode = NewNode
            S HTNode, htnAsciiCode, Chr(i)
            S HTNode, htnWeight, CDbl(CodeCounts(i) / TextLen)
            S HTNode, htnIsLeaf, True
            
            'Now place it in its reverse-ordered position.
            For j = 1 To HuffmanTrees.Count + 1
                If j > HuffmanTrees.Count Then
                    HuffmanTrees.Add HTNode
                    Exit For
                End If
                If HTNode(htnWeight) >= HuffmanTrees(j)(htnWeight) Then
                    HuffmanTrees.Add HTNode, , j
                    Exit For
                End If
            Next
        End If
    Next
    
    'Now assemble all these single-level Huffman trees into
    'one single tree, where all the leaves have the ASCII codes
    'associated with them.
    If HuffmanTrees.Count = 1 Then
        Set HTNode = NewNode
        S HTNode, htnLeftSubtree, HuffmanTrees(1)
        S HTNode, htnWeight, 1
        HuffmanTrees.Remove (1)
        HuffmanTrees.Add HTNode
    End If
    While HuffmanTrees.Count > 1
        Set HTNode = NewNode
        S HTNode, htnRightSubtree, HuffmanTrees(HuffmanTrees.Count)
        HuffmanTrees.Remove HuffmanTrees.Count
        S HTNode, htnLeftSubtree, HuffmanTrees(HuffmanTrees.Count)
        HuffmanTrees.Remove HuffmanTrees.Count
        S HTNode, htnWeight, HTNode(htnLeftSubtree)(htnWeight) + HTNode(htnRightSubtree)(htnWeight)
        
        'Place this new tree it in its reverse-ordered position.
        For j = 1 To HuffmanTrees.Count + 1
            If j > HuffmanTrees.Count Then
                HuffmanTrees.Add HTNode
                Exit For
            End If
            If HTNode(htnWeight) >= HuffmanTrees(j)(htnWeight) Then
                HuffmanTrees.Add HTNode, , j
                Exit For
            End If
        Next
    Wend
    Set HTRootNode = HuffmanTrees(1)
    AttachBitCodes BitStrings, HTRootNode, Array()
    For i = 0 To UBound(BitStrings)
        If Not IsEmpty(BitStrings(i)) Then
            Set HTNode = BitStrings(i)
            Temp = Temp & HTNode(htnAsciiCode) _
              & BitsToString(HTNode(htnBitCode))
        End If
    Next
    HuffmanEncode = HuffmanEncode & Len(Temp) & vbCr & Temp
    
    'The next part of the header is a checksum value, which
    'we'll use later to verify our decompression.
    Char = 0
    For i = 1 To TextLen
        Char = Char Xor Asc(Mid(Text, i, 1))
    Next
    HuffmanEncode = HuffmanEncode & Chr(Char)
    
    'The final part of the header identifies how many bytes
    'the original text strings contains.  We will probably
    'have a few unused bits in the last byte that we need to
    'account for.  Additionally, this serves as a final check
    'for corruption.
    HuffmanEncode = HuffmanEncode & TextLen & vbCr
    
    'Now we can encode the data by exchanging each ASCII byte for
    'its appropriate bit string.
    BitPos = -1
    Char = 0
    Temp = ""
    For i = 1 To TextLen
        BitString = BitStrings(Asc(Mid(Text, i, 1)))(htnBitCode)
        'Add each bit to the end of the output stream's 1-byte buffer.
        For j = 0 To UBound(BitString)
            BitPos = BitPos + 1
            If BitString(j) = 1 Then
                Char = Char + 2 ^ BitPos
            End If
            'If the bit buffer is full, dump it to the output stream.
            If BitPos >= 7 Then
                Temp = Temp & Chr(Char)
                'If the temporary output buffer is full, dump it
                'to the final output stream.
                If Len(Temp) > 1024 Then
                    HuffmanEncode = HuffmanEncode & Temp
                    Temp = ""
                End If
                BitPos = -1
                Char = 0
            End If
        Next
    Next
    If BitPos > -1 Then
        Temp = Temp & Chr(Char)
    End If
    If Len(Temp) > 0 Then
        HuffmanEncode = HuffmanEncode & Temp
    End If
    
    'If it takes up more space compressed because the source is
    'small and the header is big, we'll leave it uncompressed
    'and prepend it with a 4 byte header.
    If Len(HuffmanEncode) > TextLen And Not Force Then
        HuffmanEncode = "HE0" & vbCr & Text
    End If
End Function


'Decompress the string back into its original text.
Public Function HuffmanDecode(ByVal Text As String) As String
    Dim Pos As Long, Temp As String, Char As Byte, Bits
    Dim i As Long, j As Long, CharsFound As Long, BitPos As Integer
    Dim CheckSum As Byte, SourceLen As Long, TextLen As Long
    Dim HTRootNode As Collection, HTNode As Collection
    
    'If this was left uncompressed, this will be easy.
    If Left(Text, 4) = "HE0" & vbCr Then
        HuffmanDecode = Mid(Text, 5)
        Exit Function
    End If
    
    'If this is any version other than 2, we'll bow out.
    If Left(Text, 4) <> "HE2" & vbCr Then
        err.Raise vbObjectError, "HuffmanDecode()", _
          "The data either was not compressed with HE2 or is corrupt"
    End If
    Text = Mid(Text, 5)
    
    'Extract the ASCII character bit-code table's byte length.
    Pos = InStr(1, Text, vbCr)
    If Pos = 0 Then
        err.Raise vbObjectError, "HuffmanDecode()", _
          "The data either was not compressed with HE2 or is corrupt"
    End If
    On Error Resume Next
    TextLen = Left(Text, Pos - 1)
    If err.Number <> 0 Then
        On Error GoTo 0
        err.Raise vbObjectError, "HuffmanDecode()", _
          "The header is corrupt"
    End If
    On Error GoTo 0
    Text = Mid(Text, Pos + 1)
    Temp = Left(Text, TextLen)
    Text = Mid(Text, TextLen + 1)
    'Now extract the ASCII character bit-code table.
    Set HTRootNode = NewNode
    Pos = 1
    While Pos <= Len(Temp)
        Char = Asc(Mid(Temp, Pos, 1))
        Pos = Pos + 1
        Bits = StringToBits(Pos, Temp)
        Set HTNode = HTRootNode
        For j = 0 To UBound(Bits)
            If Bits(j) = 1 Then
                If HTNode(htnLeftSubtree) Is Nothing Then
                    S HTNode, htnLeftSubtree, NewNode
                End If
                Set HTNode = HTNode(htnLeftSubtree)
            Else
                If HTNode(htnRightSubtree) Is Nothing Then
                    S HTNode, htnRightSubtree, NewNode
                End If
                Set HTNode = HTNode(htnRightSubtree)
            End If
        Next
        S HTNode, htnIsLeaf, True
        S HTNode, htnAsciiCode, Chr(Char)
        S HTNode, htnBitCode, Bits
    Wend
    
    'Extract the checksum.
    CheckSum = Asc(Left(Text, 1))
    Text = Mid(Text, 2)
    
    'Extract the length of the original string.
    Pos = InStr(1, Text, vbCr)
    If Pos = 0 Then
        err.Raise vbObjectError, "HuffmanDecode()", _
          "The header is corrupt"
    End If
    On Error Resume Next
    SourceLen = Left(Text, Pos - 1)
    If err.Number <> 0 Then
        On Error GoTo 0
        err.Raise vbObjectError, "HuffmanDecode()", _
          "The header is corrupt"
    End If
    On Error GoTo 0
    Text = Mid(Text, Pos + 1)
    TextLen = Len(Text)
    
    'Now that we've processed the header, let's decode the actual data.
    i = 1
    BitPos = -1
    Set HTNode = HTRootNode
    Temp = ""
    While CharsFound < SourceLen
        If BitPos = -1 Then
            If i > TextLen Then
                err.Raise vbObjectError, "HuffmanDecode()", _
                  "Expecting more bytes in data stream"
            End If
            Char = Asc(Mid(Text, i, 1))
            i = i + 1
        End If
        BitPos = BitPos + 1
        
        If (Char And 2 ^ BitPos) > 0 Then
            Set HTNode = HTNode(htnLeftSubtree)
        Else
            Set HTNode = HTNode(htnRightSubtree)
        End If
        If HTNode Is Nothing Then
            'Uh oh.  We've followed the tree to a Huffman tree to a dead
            'end, which won't happen unless the data is corrupt.
            err.Raise vbObjectError, "HuffmanDecode()", _
              "The header (lookup table) is corrupt"
        End If
        
        If HTNode(htnIsLeaf) Then
            Temp = Temp & HTNode(htnAsciiCode)
            If Len(Temp) > 1024 Then
                HuffmanDecode = HuffmanDecode & Temp
                Temp = ""
            End If
            CharsFound = CharsFound + 1
            Set HTNode = HTRootNode
        End If
        
        If BitPos >= 7 Then BitPos = -1
    Wend
    If Len(Temp) > 0 Then
        HuffmanDecode = HuffmanDecode & Temp
    End If
    If i <= TextLen Then
        err.Raise vbObjectError, "HuffmanDecode()", _
          "Found extra bytes at end of data stream"
    End If
    
    'Verify data to check for corruption.
    If Len(HuffmanDecode) <> SourceLen Then
        err.Raise vbObjectError, "HuffmanDecode()", _
          "Data corrupt because check sums do not match"
    End If
    Char = 0
    For i = 1 To SourceLen
        Char = Char Xor Asc(Mid(HuffmanDecode, i, 1))
    Next
    If Char <> CheckSum Then
        err.Raise vbObjectError, "HuffmanDecode()", _
          "Data corrupt because check sums do not match"
    End If
End Function



'----------------------------------------------------------------
' Everything below here is only for supporting the two main
' routines above.
'----------------------------------------------------------------


'Follows the tree, now built, to its end leaf nodes, where the
'character codes are, in order to tell those character codes
'what their bit string representations are.
Private Sub AttachBitCodes(BitStrings, HTNode As Collection, ByVal Bits)
    If HTNode Is Nothing Then Exit Sub
    If HTNode(htnIsLeaf) Then
        S HTNode, htnBitCode, Bits
        Set BitStrings(Asc(HTNode(htnAsciiCode))) = HTNode
    Else
        ReDim Preserve Bits(UBound(Bits) + 1)
        Bits(UBound(Bits)) = 1
        AttachBitCodes BitStrings, HTNode(htnLeftSubtree), Bits
        Bits(UBound(Bits)) = 0
        AttachBitCodes BitStrings, HTNode(htnRightSubtree), Bits
    End If
End Sub

'Turns a string of '0' and '1' characters into a string of bytes
'containing the bits, preceeded by 1 byte indicating the
'number of bits represented.
Private Function BitsToString(Bits) As String
    Dim Char As Byte, i As Long
    BitsToString = Chr(UBound(Bits) + 1)  'Number of bits
    For i = 0 To UBound(Bits)
        If i Mod 8 = 0 Then
            If i > 0 Then BitsToString = BitsToString & Chr(Char)
            Char = 0
        End If
        If Bits(i) = 1 Then  'Bit value = 1
            'Mask the bit into its proper position in the byte
            Char = Char + 2 ^ (i Mod 8)
        End If
    Next
    BitsToString = BitsToString & Chr(Char)
End Function

'The opposite of BitsToString() function.
Private Function StringToBits(StartPos As Long, Bytes As String)
    Dim Char As Byte, i As Long, BitCount As Long, Bits
    Bits = Array()
    BitCount = Asc(Mid(Bytes, StartPos, 1))
    StartPos = StartPos + 1
    For i = 0 To BitCount - 1
        If i Mod 8 = 0 Then
            Char = Asc(Mid(Bytes, StartPos, 1))
            StartPos = StartPos + 1
        End If
        ReDim Preserve Bits(UBound(Bits) + 1)
        If (Char And 2 ^ (i Mod 8)) > 0 Then   'Bit value = 1
            Bits(UBound(Bits)) = 1
        Else  'Bit value = 0
            Bits(UBound(Bits)) = 0
        End If
    Next
    StringToBits = Bits
End Function

'Remove the specified item and put the specified value in its place.
Private Sub S(Col As Collection, Index As HuffmanTreeNodeParts, Value)
    Col.Remove Index
    If Index > Col.Count Then
        Col.Add Value
    Else
        Col.Add Value, , Index
    End If
End Sub

'Creates a new Huffman tree node with the default values set.
Private Function NewNode() As Collection
    Dim Node As New Collection
    Node.Add 0  'htnWeight
    Node.Add False  'htnIsLeaf
    Node.Add Chr(0)  'htnAsciiCode
    Node.Add ""  'htnBitCode
    Node.Add Nothing  'htnLeftSubtree
    Node.Add Nothing  'htnRightSubtree
    Set NewNode = Node
End Function


