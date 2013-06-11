Attribute VB_Name = "md5"
Option Explicit

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647

Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21

Private mlngState(4) As Long
Private mlngByteCount As Long
Private mbytBuffer(63) As Byte

Private Property Get RegisterA() As String
    RegisterA = mlngState(1)
End Property

Private Property Get RegisterB() As String
    RegisterB = mlngState(2)
End Property

Private Property Get RegisterC() As String
    RegisterC = mlngState(3)
End Property

Private Property Get RegisterD() As String
    RegisterD = mlngState(4)
End Property

Public Function DigestFileToHexStr(ByVal Filename As String) As String
Dim intFile          As Integer
Dim lngBufferSize    As Long
On Error GoTo ErrHandler
   lngBufferSize = UBound(mbytBuffer) + 1
   intFile = FreeFile
   Open Filename For Binary Access Read As #intFile
   MD5Init
   Do While Not EOF(intFile)
      Get #intFile, , mbytBuffer
      If Loc(intFile) < LOF(intFile) Then
         mlngByteCount = mlngByteCount + lngBufferSize
         MD5Transform mbytBuffer
      End If
   Loop
   mlngByteCount = mlngByteCount + (LOF(intFile) Mod lngBufferSize)
   Close #intFile
   MD5Final
   DigestFileToHexStr = JoinStateValues
   Exit Function
ErrHandler:
   If Not (intFile = 0) Then
      Close #intFile
   End If
   Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function DigestStrToHexStr(SourceString As String) As String
    MD5Init
    MD5Update Len(SourceString), StringToArray(SourceString)
    MD5Final
    DigestStrToHexStr = JoinStateValues
End Function


Private Function StringToArray(InString As String) As Byte()
Dim lngIndex      As Long
Dim bytBuffer()   As Byte
   ReDim bytBuffer(Len(InString))
   For lngIndex = 0 To Len(InString) - 1
      bytBuffer(lngIndex) = Asc(Mid(InString, lngIndex + 1, 1))
   Next lngIndex
   StringToArray = bytBuffer
End Function

Private Function JoinStateValues() As String
    JoinStateValues = LongToHexString(mlngState(1)) & LongToHexString(mlngState(2)) & LongToHexString(mlngState(3)) & LongToHexString(mlngState(4))
End Function

Private Function LongToHexString(ByVal plngNum As Long) As String
Dim a           As Byte
Dim b           As Byte
Dim c           As Byte
Dim d           As Byte
Dim strTemp     As String
    a = plngNum And &HFF&
    If a < 16 Then
        strTemp = "0" & Hex(a)
    Else
        strTemp = Hex(a)
    End If
     b = (plngNum And &HFF00&) \ 256
     If b < 16 Then
         strTemp = strTemp & "0" & Hex(b)
     Else
         strTemp = strTemp & Hex(b)
     End If
     c = (plngNum And &HFF0000) \ 65536
     If c < 16 Then
         strTemp = strTemp & "0" & Hex(c)
     Else
         strTemp = strTemp & Hex(c)
     End If
     If plngNum < 0 Then
         d = ((plngNum And &H7F000000) \ 16777216) Or &H80&
     Else
         d = (plngNum And &HFF000000) \ 16777216
     End If
     
     If d < 16 Then
         strTemp = strTemp & "0" & Hex(d)
     Else
         strTemp = strTemp & Hex(d)
     End If
     LongToHexString = strTemp
End Function

Public Sub MD5Init()
    mlngByteCount = 0
    mlngState(1) = UnsignedToLong(1732584193#)
    mlngState(2) = UnsignedToLong(4023233417#)
    mlngState(3) = UnsignedToLong(2562383102#)
    mlngState(4) = UnsignedToLong(271733878#)
End Sub

Public Sub MD5Final()
    Dim dblBits As Double
    
    Dim padding(72) As Byte
    Dim lngBytesBuffered As Long
    
    padding(0) = &H80
    
    dblBits = mlngByteCount * 8
    
    ' Pad out
    lngBytesBuffered = mlngByteCount Mod 64
    If lngBytesBuffered <= 56 Then
        MD5Update 56 - lngBytesBuffered, padding
    Else
        MD5Update 120 - mlngByteCount, padding
    End If
    
    
    padding(0) = UnsignedToLong(dblBits) And &HFF&
    padding(1) = UnsignedToLong(dblBits) \ 256 And &HFF&
    padding(2) = UnsignedToLong(dblBits) \ 65536 And &HFF&
    padding(3) = UnsignedToLong(dblBits) \ 16777216 And &HFF&
    padding(4) = 0
    padding(5) = 0
    padding(6) = 0
    padding(7) = 0
    
    MD5Update 8, padding
End Sub

'
' Break up input stream into 64 byte chunks
'
Public Sub MD5Update(InputLen As Long, InputBuffer() As Byte)
    Dim II As Integer
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim lngBufferedBytes As Long
    Dim lngBufferRemaining As Long
    Dim lngRem As Long
    
    lngBufferedBytes = mlngByteCount Mod 64
    lngBufferRemaining = 64 - lngBufferedBytes
    mlngByteCount = mlngByteCount + InputLen
    ' Use up old buffer results first
    If InputLen >= lngBufferRemaining Then
        For II = 0 To lngBufferRemaining - 1
            mbytBuffer(lngBufferedBytes + II) = InputBuffer(II)
        Next II
        MD5Transform mbytBuffer
        
        lngRem = (InputLen) Mod 64
        ' The transfer is a multiple of 64 lets do some transformations
        For I = lngBufferRemaining To InputLen - II - lngRem Step 64
            For J = 0 To 63
                mbytBuffer(J) = InputBuffer(I + J)
            Next J
            MD5Transform mbytBuffer
        Next I
        lngBufferedBytes = 0
    Else
      I = 0
    End If
    
    ' Buffer any remaining input
    For K = 0 To InputLen - I - 1
        mbytBuffer(lngBufferedBytes + K) = InputBuffer(I + K)
    Next K
    
End Sub

'
' MD5 Transform
'
Private Sub MD5Transform(Buffer() As Byte)
    Dim x(16) As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    
    a = mlngState(1)
    b = mlngState(2)
    c = mlngState(3)
    d = mlngState(4)
    
    Decode 64, x, Buffer

    ' Round 1
    FF a, b, c, d, x(0), S11, -680876936
    FF d, a, b, c, x(1), S12, -389564586
    FF c, d, a, b, x(2), S13, 606105819
    FF b, c, d, a, x(3), S14, -1044525330
    FF a, b, c, d, x(4), S11, -176418897
    FF d, a, b, c, x(5), S12, 1200080426
    FF c, d, a, b, x(6), S13, -1473231341
    FF b, c, d, a, x(7), S14, -45705983
    FF a, b, c, d, x(8), S11, 1770035416
    FF d, a, b, c, x(9), S12, -1958414417
    FF c, d, a, b, x(10), S13, -42063
    FF b, c, d, a, x(11), S14, -1990404162
    FF a, b, c, d, x(12), S11, 1804603682
    FF d, a, b, c, x(13), S12, -40341101
    FF c, d, a, b, x(14), S13, -1502002290
    FF b, c, d, a, x(15), S14, 1236535329
    
    ' Round 2
    GG a, b, c, d, x(1), S21, -165796510
    GG d, a, b, c, x(6), S22, -1069501632
    GG c, d, a, b, x(11), S23, 643717713
    GG b, c, d, a, x(0), S24, -373897302
    GG a, b, c, d, x(5), S21, -701558691
    GG d, a, b, c, x(10), S22, 38016083
    GG c, d, a, b, x(15), S23, -660478335
    GG b, c, d, a, x(4), S24, -405537848
    GG a, b, c, d, x(9), S21, 568446438
    GG d, a, b, c, x(14), S22, -1019803690
    GG c, d, a, b, x(3), S23, -187363961
    GG b, c, d, a, x(8), S24, 1163531501
    GG a, b, c, d, x(13), S21, -1444681467
    GG d, a, b, c, x(2), S22, -51403784
    GG c, d, a, b, x(7), S23, 1735328473
    GG b, c, d, a, x(12), S24, -1926607734
    
    ' Round 3
    HH a, b, c, d, x(5), S31, -378558
    HH d, a, b, c, x(8), S32, -2022574463
    HH c, d, a, b, x(11), S33, 1839030562
    HH b, c, d, a, x(14), S34, -35309556
    HH a, b, c, d, x(1), S31, -1530992060
    HH d, a, b, c, x(4), S32, 1272893353
    HH c, d, a, b, x(7), S33, -155497632
    HH b, c, d, a, x(10), S34, -1094730640
    HH a, b, c, d, x(13), S31, 681279174
    HH d, a, b, c, x(0), S32, -358537222
    HH c, d, a, b, x(3), S33, -722521979
    HH b, c, d, a, x(6), S34, 76029189
    HH a, b, c, d, x(9), S31, -640364487
    HH d, a, b, c, x(12), S32, -421815835
    HH c, d, a, b, x(15), S33, 530742520
    HH b, c, d, a, x(2), S34, -995338651
    
    ' Round 4
    II a, b, c, d, x(0), S41, -198630844
    II d, a, b, c, x(7), S42, 1126891415
    II c, d, a, b, x(14), S43, -1416354905
    II b, c, d, a, x(5), S44, -57434055
    II a, b, c, d, x(12), S41, 1700485571
    II d, a, b, c, x(3), S42, -1894986606
    II c, d, a, b, x(10), S43, -1051523
    II b, c, d, a, x(1), S44, -2054922799
    II a, b, c, d, x(8), S41, 1873313359
    II d, a, b, c, x(15), S42, -30611744
    II c, d, a, b, x(6), S43, -1560198380
    II b, c, d, a, x(13), S44, 1309151649
    II a, b, c, d, x(4), S41, -145523070
    II d, a, b, c, x(11), S42, -1120210379
    II c, d, a, b, x(2), S43, 718787259
    II b, c, d, a, x(9), S44, -343485551
    
    
    mlngState(1) = LongOverflowAdd(mlngState(1), a)
    mlngState(2) = LongOverflowAdd(mlngState(2), b)
    mlngState(3) = LongOverflowAdd(mlngState(3), c)
    mlngState(4) = LongOverflowAdd(mlngState(4), d)

'  /* Zeroize sensitive information.
'*/
'  MD5_memset ((POINTER)x, 0, sizeof (x));
    
End Sub

Private Sub Decode(Length As Integer, OutputBuffer() As Long, InputBuffer() As Byte)
    Dim intDblIndex As Integer
    Dim intByteIndex As Integer
    Dim dblSum As Double
    
    intDblIndex = 0
    For intByteIndex = 0 To Length - 1 Step 4
        dblSum = InputBuffer(intByteIndex) + InputBuffer(intByteIndex + 1) * 256# + InputBuffer(intByteIndex + 2) * 65536# + InputBuffer(intByteIndex + 3) * 16777216#
        OutputBuffer(intDblIndex) = UnsignedToLong(dblSum)
        intDblIndex = intDblIndex + 1
    Next intByteIndex
End Sub

'
' FF, GG, HH, and II transformations for rounds 1, 2, 3, and 4.
' Rotation is separate from addition to prevent recomputation.
'
Private Function FF(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, ac As Long) As Long
    a = LongOverflowAdd4(a, (b And c) Or (Not (b) And d), x, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function GG(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, ac As Long) As Long
    a = LongOverflowAdd4(a, (b And d) Or (c And Not (d)), x, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function HH(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, ac As Long) As Long
    a = LongOverflowAdd4(a, b Xor c Xor d, x, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function II(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, ac As Long) As Long
    a = LongOverflowAdd4(a, c Xor (b Or Not (d)), x, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function
'
' Rotate a long to the right
'
Function LongLeftRotate(Value As Long, bits As Long) As Long
    Dim lngSign As Long
    Dim lngI As Long
    bits = bits Mod 32
    If bits = 0 Then LongLeftRotate = Value: Exit Function
    For lngI = 1 To bits
        lngSign = Value And &HC0000000
        Value = (Value And &H3FFFFFFF) * 2
        Value = Value Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
    Next
    LongLeftRotate = Value
End Function

'
' Function to add two unsigned numbers together as in C.
' Overflows are ignored!
'
Private Function LongOverflowAdd(Val1 As Long, Val2 As Long) As Long
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    LongOverflowAdd = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

'
' Function to add two unsigned numbers together as in C.
' Overflows are ignored!
'
Private Function LongOverflowAdd4(Val1 As Long, Val2 As Long, val3 As Long, val4 As Long) As Long
Dim lngHighWord   As Long
Dim lngLowWord    As Long
Dim lngOverflow   As Long
    lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&) + (val3 And &HFFFF&) + (val4 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + ((val3 And &HFFFF0000) \ 65536) + ((val4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    LongOverflowAdd4 = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

Private Function UnsignedToLong(ByVal Value As Double) As Long
   If Value < 0 Or Value >= OFFSET_4 Then
      Error 6
   Else
      If Value <= MAXINT_4 Then
         UnsignedToLong = Value
      Else
         UnsignedToLong = Value - OFFSET_4
      End If
   End If
End Function

Private Function LongToUnsigned(ByVal Value As Long) As Double
   If Value < 0 Then
      LongToUnsigned = Value + OFFSET_4
   Else
      LongToUnsigned = Value
   End If
End Function
