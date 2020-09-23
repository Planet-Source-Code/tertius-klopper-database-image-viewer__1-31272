Attribute VB_Name = "ZLibComp"
'If you find a way to compress the pictures so that they are smaller
'and that the compression is done faster then please send me a copy of it

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal level As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Const m_def_CompressedSize = 0
Const m_def_OriginalSize = 0

Dim m_CompressedSize As Long
Dim m_OriginalSize As Long

Enum CZErrors
    Z_OK = 0
    Z_STREAM_END = 1
    Z_NEED_DICT = 2
    Z_ERRNO = -1
    Z_STREAM_ERROR = -2
    Z_DATA_ERROR = -3
    Z_MEM_ERROR = -4
    Z_BUF_ERROR = -5
    Z_VERSION_ERROR = -6
End Enum

Enum CompressionLevels
    Z_NO_COMPRESSION = 0
    Z_BEST_SPEED = 1
    'note that levels 2-8 exist, too
    Z_BEST_COMPRESSION = 9
    Z_DEFAULT_COMPRESSION = -1
End Enum

Public Function CompStr(TheString As String, CompressionLevel As Integer) As Long
Dim orgSize As Long
Dim ret As Long

'Allocate string space for the buffers


Dim CmpSize As Long
Dim TBuff As String
orgSize = Len(TheString)
TBuff = String(orgSize + (orgSize * 0.01) + 12, 0)
CmpSize = Len(TBuff)
'Compress string (temporary string buffer) data
ret = compress2(ByVal TBuff, CmpSize, ByVal TheString, Len(TheString), CompressionLevel)
'Crop the string and set it to the actual string.
TheString = Left$(TBuff, CmpSize)
'Set compressed size of string.
'CompressedSize = CmpSize
'Cleanup
TBuff = ""
'Return error code (if any)
TheString = orgSize & "-" & TheString
CompStr = ret
End Function

Public Function DeCompStr(TheString As String, OrigSize As Long) As Long
Dim result As Long
'Allocate string space
Dim CmpSize As Long
Dim TBuff As String
'originalsize = OrigSize
TBuff = String(OrigSize + (OrigSize * 0.01) + 12, 0)
CmpSize = Len(TBuff)

'Decompress
result = uncompress(ByVal TBuff, CmpSize, ByVal TheString, Len(TheString))
'Make string the size of the uncompressed string
TheString = Left$(TBuff, CmpSize)
'CompressedSize = CmpSize
'Return error code (if any)
DeCompStr = result
End Function
