VERSION 1.0 CLASS

BEGIN

    MultiUse = -1

END

Attribute VB_Name = "CRYPTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This uses LF instead of CRLF, so the lines above must include an empty line after the keyword.
'
Option Explicit

Private Const BITS_TO_A_BYTE As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE
Private m_lOnBits(30)
Private m_l2Power(30)

Private hashMode As String
Private hash As String ' save as hex string

Private Sub Class_Initialize()
    m_lOnBits(0) = 1           ' 00000000000000000000000000000001
    m_lOnBits(1) = 3           ' 00000000000000000000000000000011
    m_lOnBits(2) = 7           ' 00000000000000000000000000000111
    m_lOnBits(3) = 15          ' 00000000000000000000000000001111
    m_lOnBits(4) = 31          ' 00000000000000000000000000011111
    m_lOnBits(5) = 63          ' 00000000000000000000000000111111
    m_lOnBits(6) = 127         ' 00000000000000000000000001111111
    m_lOnBits(7) = 255         ' 00000000000000000000000011111111
    m_lOnBits(8) = 511         ' 00000000000000000000000111111111
    m_lOnBits(9) = 1023        ' 00000000000000000000001111111111
    m_lOnBits(10) = 2047       ' 00000000000000000000011111111111
    m_lOnBits(11) = 4095       ' 00000000000000000000111111111111
    m_lOnBits(12) = 8191       ' 00000000000000000001111111111111
    m_lOnBits(13) = 16383      ' 00000000000000000011111111111111
    m_lOnBits(14) = 32767      ' 00000000000000000111111111111111
    m_lOnBits(15) = 65535      ' 00000000000000001111111111111111
    m_lOnBits(16) = 131071     ' 00000000000000011111111111111111
    m_lOnBits(17) = 262143     ' 00000000000000111111111111111111
    m_lOnBits(18) = 524287     ' 00000000000001111111111111111111
    m_lOnBits(19) = 1048575    ' 00000000000011111111111111111111
    m_lOnBits(20) = 2097151    ' 00000000000111111111111111111111
    m_lOnBits(21) = 4194303    ' 00000000001111111111111111111111
    m_lOnBits(22) = 8388607    ' 00000000011111111111111111111111
    m_lOnBits(23) = 16777215   ' 00000000111111111111111111111111
    m_lOnBits(24) = 33554431   ' 00000001111111111111111111111111
    m_lOnBits(25) = 67108863   ' 00000011111111111111111111111111
    m_lOnBits(26) = 134217727  ' 00000111111111111111111111111111
    m_lOnBits(27) = 268435455  ' 00001111111111111111111111111111
    m_lOnBits(28) = 536870911  ' 00011111111111111111111111111111
    m_lOnBits(29) = 1073741823 ' 00111111111111111111111111111111
    m_lOnBits(30) = 2147483647 ' 01111111111111111111111111111111

    m_l2Power(0) = 1           ' 00000000000000000000000000000001
    m_l2Power(1) = 2           ' 00000000000000000000000000000010
    m_l2Power(2) = 4           ' 00000000000000000000000000000100
    m_l2Power(3) = 8           ' 00000000000000000000000000001000
    m_l2Power(4) = 16          ' 00000000000000000000000000010000
    m_l2Power(5) = 32          ' 00000000000000000000000000100000
    m_l2Power(6) = 64          ' 00000000000000000000000001000000
    m_l2Power(7) = 128         ' 00000000000000000000000010000000
    m_l2Power(8) = 256         ' 00000000000000000000000100000000
    m_l2Power(9) = 512         ' 00000000000000000000001000000000
    m_l2Power(10) = 1024       ' 00000000000000000000010000000000
    m_l2Power(11) = 2048       ' 00000000000000000000100000000000
    m_l2Power(12) = 4096       ' 00000000000000000001000000000000
    m_l2Power(13) = 8192       ' 00000000000000000010000000000000
    m_l2Power(14) = 16384      ' 00000000000000000100000000000000
    m_l2Power(15) = 32768      ' 00000000000000001000000000000000
    m_l2Power(16) = 65536      ' 00000000000000010000000000000000
    m_l2Power(17) = 131072     ' 00000000000000100000000000000000
    m_l2Power(18) = 262144     ' 00000000000001000000000000000000
    m_l2Power(19) = 524288     ' 00000000000010000000000000000000
    m_l2Power(20) = 1048576    ' 00000000000100000000000000000000
    m_l2Power(21) = 2097152    ' 00000000001000000000000000000000
    m_l2Power(22) = 4194304    ' 00000000010000000000000000000000
    m_l2Power(23) = 8388608    ' 00000000100000000000000000000000
    m_l2Power(24) = 16777216   ' 00000001000000000000000000000000
    m_l2Power(25) = 33554432   ' 00000010000000000000000000000000
    m_l2Power(26) = 67108864   ' 00000100000000000000000000000000
    m_l2Power(27) = 134217728  ' 00001000000000000000000000000000
    m_l2Power(28) = 268435456  ' 00010000000000000000000000000000
    m_l2Power(29) = 536870912  ' 00100000000000000000000000000000
    m_l2Power(30) = 1073741824 ' 01000000000000000000000000000000
End Sub

Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
    Dim lX4 As Long
    Dim lY4 As Long
    Dim lX8 As Long
    Dim lY8 As Long
    Dim lResult As Long

    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000

    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If

    AddUnsigned = lResult
End Function

Private Function ConvertToWordArray(sMessage As String) As Long()
    Dim lMessageLength As Long
    Dim lNumberOfWords As Long
    Dim lWordArray() As Long
    Dim lBytePosition As Long
    Dim lByteCount As Long
    Dim lWordCount As Long

    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448

    lMessageLength = Len(sMessage)

    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)

    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)

    ConvertToWordArray = lWordArray
End Function

Private Function ConvertToWordArray2(sMessage As String) As Long()
    Dim lMessageLength  As Long
    Dim lNumberOfWords  As Long
    Dim lWordArray()    As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long
    Dim lByte           As Long

    Const MODULUS_BITS As Long = 512
    Const CONGRUENT_BITS As Long = 448

    lMessageLength = Len(sMessage)

    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)

    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        lByte = AscB(Mid(sMessage, lByteCount + 1, 1))
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)

    ConvertToWordArray2 = lWordArray
End Function

Private Function todo(s)
  Debug.Print "TODO: " & s
End Function

Public Sub createHash(Optional ByVal name As String = "sha256")
  hashMode = name
End Sub

Private Function calculateMD5(words() AS Long) AS Long()
    Dim rVals As Variant: rVals = Array( _ 
        7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, _ 
        5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, _ 
        4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, _ 
        6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21)
    Dim kVals As Variant: kVals = Array( _
        &HD76AA478, &HE8C7B756, &H242070DB, &HC1BDCEEE, &HF57C0FAF, &H4787C62A, &HA8304613, &HFD469501, _
        &H698098D8, &H8B44F7AF, &HFFFF5BB1, &H895CD7BE, &H6B901122, &HFD987193, &HA679438E, &H49B40821, _ 
        &HF61E2562, &HC040B340, &H265E5A51, &HE9B6C7AA, &HD62F105D, &H2441453, &HD8A1E681, &HE7D3FBC8, _ 
        &H21E1CDE6, &HC33707D6, &HF4D50D87, &H455A14ED, &HA9E3E905, &HFCEFA3F8, &H676F02D9, &H8D2A4C8A, _ 
        &HFFFA3942, &H8771F681, &H6D9D6122, &HFDE5380C, &HA4BEEA44, &H4BDECFA9, &HF6BB4B60, &HBEBFBC70, _ 
        &H289B7EC6, &HEAA127FA, &HD4EF3085, &H4881D05, &HD9D4D039, &HE6DB99E5, &H1FA27CF8, &HC4AC5665, _ 
        &HF4292244, &H432AFF97, &HAB9423A7, &HFC93A039, &H655B59C3, &H8F0CCC92, &HFFEFF47D, &H85845DD1, _
        &H6FA87E4F, &HFE2CE6E0, &HA3014314, &H4E0811A1, &HF7537E82, &HBD3AF235, &H2AD7D2BB, &HEB86D391)
    Dim result(3) As Long
    result(0) = &H67452301
    result(1) = &HEFCDAB89
    result(2) = &H98BADCFE
    result(3) = &H10325476
    Dim k, i, f, g, v
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    For k = 0 To UBound(words) Step 16
        a = result(0)
        b = result(1)
        c = result(2)
        d = result(3)
        For i = 0 To 63
            If i < 16 Then
                f = (b And c) Or ((Not b) And d)
                g = i
            ElseIf i < 32 Then
                f = (d And b) Or ((Not d) And c)
                g = (5 * i + 1) Mod 16
            ElseIf i < 48 Then
                f = b Xor c Xor d
                g = (3 * i + 5) Mod 16
            Else
                f = c Xor (b Or (Not d))
                g = (7 * i) Mod 16
            End If
            v = AddUnsigned(a, AddUnsigned(AddUnsigned(f, words(k + g)), kVals(i)))
            a = d
            d = c
            c = b
            b = AddUnsigned(RotateLeft(v, rVals(i)), b)
        Next i
        result(0) = AddUnsigned(a, result(0))
        result(1) = AddUnsigned(b, result(1))
        result(2) = AddUnsigned(c, result(2))
        result(3) = AddUnsigned(d, result(3))
    Next
    calculateMD5 = result
End Function

Private Function rotr(ByVal x As Long, ByVal n As Integer) As Long
    rotr = RShift(x, (n And m_lOnBits(4))) Or LShift(x, (32 - (n And m_lOnBits(4))))
End Function

Private Function rs(ByVal x As Long, ByVal n As Integer) As Long
    rs = RShift(x, CInt(n And m_lOnBits(4)))
End Function

Private Function ch(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    ch = (x And y) Xor ((Not x) And z)
End Function

Private Function maj(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    maj = (x And y) Xor (x And z) Xor (y And z)
End Function

Private Function sigma(ByVal x As Long) As Long
    sigma = rotr(x, 2) Xor rotr(x, 13) Xor rotr(x, 22)
End Function

Private Function sigmaPrime(ByVal x As Long) As Long
    sigmaPrime = rotr(x, 6) Xor rotr(x, 11) Xor rotr(x, 25)
End Function

Private Function littleSigma(ByVal x As Long) As Long
    littleSigma = rotr(x, 7) Xor rotr(x, 18) Xor rs(x, 3)
End Function

Private Function littleSigmaPrime(ByVal x As Long) As Long
    littleSigmaPrime = rotr(x, 17) Xor rotr(x, 19) Xor rs(x, 10)
End Function

Private Function calculateSHA1(words() AS Long) AS Long()
    todo "calculateSHA1"
End Function

Private Function calculateSHA256(words() AS Long) AS Long()
    Dim kVals As Variant: kVals = Array( _
        &H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
        &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
        &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
        &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
        &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
        &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
        &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
        &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2)
    Dim result(7) As Long
    result(0) = &H6A09E667
    result(1) = &HBB67AE85
    result(2) = &H3C6EF372
    result(3) = &HA54FF53A
    result(4) = &H510E527F
    result(5) = &H9B05688C
    result(6) = &H1F83D9AB
    result(7) = &H5BE0CD19
    Dim W(63)   As Long
    Dim k, i, j, t1, t2
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    For k = 0 To UBound(words) Step 16
        a = result(0)
        b = result(1)
        c = result(2)
        d = result(3)
        e = result(4)
        f = result(5)
        g = result(6)
        h = result(7)
        For i = 0 To 63
            If i < 16 Then
                W(i) = words(k + i)
            Else
                W(i) = AddUnsigned(AddUnsigned(AddUnsigned(littleSigmaPrime(W(i - 2)), W(i - 7)), littleSigma(W(i - 15))), W(i - 16))
            End If
            t1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(h, sigmaPrime(e)), ch(e, f, g)), kVals(i)), W(i))
            t2 = AddUnsigned(sigma(a), maj(a, b, c))
            h = g
            g = f
            f = e
            e = AddUnsigned(d, t1)
            d = c
            c = b
            b = a
            a = AddUnsigned(t1, t2)
        Next
        result(0) = AddUnsigned(a, result(0))
        result(1) = AddUnsigned(b, result(1))
        result(2) = AddUnsigned(c, result(2))
        result(3) = AddUnsigned(d, result(3))
        result(4) = AddUnsigned(e, result(4))
        result(5) = AddUnsigned(f, result(5))
        result(6) = AddUnsigned(g, result(6))
        result(7) = AddUnsigned(h, result(7))
    Next
    calculateSHA256 = result
End Function
Private Function calculateSHA384(words() AS Long) AS Long()
    todo "calculateSHA384"
End Function

Private Function calculateSHA512(words() AS Long) AS Long()
    todo "calculateSHA512"
End Function

Private Function rLongToHex(lValue As Long) As String
    Dim lByte As Long
    Dim lCount As Long

    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        rLongToHex = rLongToHex & Right("0" & Hex(lByte), 2)
    Next
End Function

Private Function LongToHex(lValue As Long) As String
    LongToHex = Right("00000000" & Hex(lValue), 8)
End Function

Private Function rLongsToHex(arr() As Long) As String
    Dim i As Long, hexStr As String
    For i = 0 To UBound(arr)
        hexStr = hexStr & rLongToHex(arr(i))
    Next i
    rLongsToHex = LCase(hexStr)
End Function

Private Function longsToHex(arr() As Long) As String
    Dim i As Long, hexStr As String
    For i = 0 To UBound(arr)
        hexStr = hexStr & LongToHex(arr(i))
    Next i
    longsToHex = LCase(hexStr)
End Function

Public Sub update(s As String)
    Select Case hashMode
        Case "md5"
            hash = rLongsToHex(calculateMD5(ConvertToWordArray(s)))
        Case "sha1"
            hash = longsToHex(calculateSHA1(ConvertToWordArray2(s)))
        Case "sha256"
            hash = longsToHex(calculateSHA256(ConvertToWordArray2(s)))
        Case "sha384"
            hash = longsToHex(calculateSHA384(ConvertToWordArray2(s)))
        Case "sha512"
            hash = longsToHex(calculateSHA512(ConvertToWordArray2(s)))
    End Select
End Sub

Private Function bytesToHex(bytes() As Byte) As String
    Dim i As Long, s As String
    For i = LBound(bytes) To UBound(bytes)
        s = s & LCase(Right("0" & Hex(bytes(i)), 2))
    Next i
    bytesToHex = s
End Function

Private Function bytesToBase64(bytes() As Byte) As String
    Const BASE64_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim byte1 As Byte, byte2 As Byte, byte3 As Byte
    Dim s As String
    Dim i As Long: i = LBound(bytes)
    Do While i <= UBound(bytes)
        byte1 = bytes(i)
        If i + 1 <= UBound(bytes) Then byte2 = bytes(i + 1) Else byte2 = 0
        If i + 2 <= UBound(bytes) Then byte3 = bytes(i + 2) Else byte3 = 0

        ' Base64 encoding: take 3 bytes = 24 bits, split into 4 x 6 bits
        s = s & Mid(BASE64_CHARS, ((byte1 And &HFC) \ 4) + 1, 1)
        s = s & Mid(BASE64_CHARS, (((byte1 And &H3) * 16 + (byte2 And &HF0) \ 16) Mod 64) + 1, 1)

        If i + 1 <= UBound(bytes) Then
            s = s & Mid(BASE64_CHARS, (((byte2 And &HF) * 4 + (byte3 And &HC0) \ 64) Mod 64) + 1, 1)
        Else
            s = s & "="
        End If

        If i + 2 <= UBound(bytes) Then
            s = s & Mid(BASE64_CHARS, (byte3 And &H3F) + 1, 1)
        Else
            s = s & "="
        End If

        i = i + 3
    Loop

    bytesToBase64 = s
End Function

Private Function hexToBytes(hexString As String) As Byte()
    Dim i As Long
    Dim bytes() As Byte

    ' Sanitize input
    hexString = Replace(hexString, " ", "")
    If Len(hexString) Mod 2 <> 0 Then
        hexString = "0" & hexString
    End If

    ' Convert hex to byte array
    ReDim bytes(0 To Len(hexString) \ 2 - 1)

    For i = LBound(bytes) To UBound(bytes)
        bytes(i) = CByte("&H" & Mid(hexString, i * 2 + 1, 2))
    Next i

    hexToBytes = bytes
End Function

Private Function hexToBase64(hexString As String) As String
    hexToBase64 = bytesToBase64(hexToBytes(hexString))
End Function

Public Function digest(Optional ByVal mode As String = "hex") As String
    Select Case mode
        Case "hex"
            digest = hash
        Case "base64"
            digest = hexToBase64(hash)
        Case Else
            todo mode
    End Select
End Function
