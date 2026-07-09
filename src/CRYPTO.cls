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

' Accumulated message bytes (streaming buffer). mBuf holds mCap allocated bytes,
' of which the first mLen are the actual message; update() appends, digest()
' hashes mBuf(0 .. mLen - 1).
Private mBuf() As Byte
Private mLen As Long
Private mCap As Long

' HMAC state: when isHmac is set, digest() keys the hash with hmacKeyBytes.
Private isHmac As Boolean
Private hmacKeyBytes() As Byte

' A 64-bit value, split into two 32-bit halves (VBA's Long is only 32-bit).
' Used by the SHA-384/512 implementation, which operates on 64-bit words.
Private Type T64
    hi As Long
    lo As Long
End Type

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

' Converts a string to an array of 32-bit words (little-endian, padding and bit length included)
' as required for MD5 hashing.
'
' Example:
'   Input string: "hello world"
'
'   After processing:
'   - The first 11 bytes come from ASCII
'   - Append 0x80 (10000000)
'   - Zero padding until length is congruent to 448 mod 512 (56 bytes)
'   - The message length in bits (88 = 0x58) is appended as a 64-bit little-endian integer
'
'   Example output (array of 16 Longs, 32 bits each):
'     Index | Hex bytes (VBA) | Packed hex string (little-endian) | ASCII (packed chars)
'     [ 0]  | &H6C6C6568      | 6c6c6568                          | "hell"
'     [ 1]  | &H6F77206F      | 6f77206f                          | "o wo"
'     [ 2]  | &H80646C72      | 80646c72                          | "rld" & vbLf & Chr(&H80)
'     [ 3]  | &H00000000      | 00000000
'     [ 4]  | &H00000000      | 00000000
'     [ 5]  | &H00000000      | 00000000
'     [ 6]  | &H00000000      | 00000000
'     [ 7]  | &H00000000      | 00000000
'     [ 8]  | &H00000000      | 00000000
'     [ 9]  | &H00000000      | 00000000
'     [10]  | &H00000000      | 00000000
'     [11]  | &H00000000      | 00000000
'     [12]  | &H00000000      | 00000000
'     [13]  | &H00000000      | 00000000
'     [14]  | &H00000058      | 00000058 ' bit length (low dword): 88 bits
'     [15]  | &H00000000      | 00000000 ' bit length (high dword)
'
'   This array can now be used with MD5 core processing.
Private Function bytesToWordArrayMD5(bytes() As Byte, ByVal msgLen As Long) As Long()
    Dim lNumberOfWords As Long
    Dim lWordArray() As Long
    Dim lBytePosition As Long
    Dim lByteCount As Long
    Dim lWordCount As Long

    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448

    lNumberOfWords = (((msgLen + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)

    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= msgLen
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(bytes(lByteCount), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(msgLen, 3)
    lWordArray(lNumberOfWords - 1) = RShift(msgLen, 29)

    bytesToWordArrayMD5 = lWordArray
End Function

' Converts a string to an array of 32-bit words (big-endian, padding and bit length included)
' as required for SHA-1, SHA-256, etc.
'
' Example:
'   Input string: "hello world"
'
'   After processing:
'   - The first bytes are from ASCII
'   - Append 0x80 (10000000)
'   - Zero padding until length is congruent to 448 mod 512 (56 bytes)
'   - The message length in bits (88 = 0x58) is appended as a 64-bit big-endian integer
'
'   Output (array of 16 Longs, 32 bits each for a 512-bit block):
'     [ 0] ... [13] = input bits, padding, and zeroes as needed
'     [14] = high dword of message bit length (for short ASCII, this is 0)
'     [15] = low dword of message bit length (e.g. 0x00000058 for "hello world")
'
'   This array is used for SHA core processing.
Private Function bytesToWordArraySHA(bytes() As Byte, ByVal msgLen As Long) As Long()
    Dim lNumberOfWords  As Long
    Dim lWordArray()    As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long

    Const MODULUS_BITS As Long = 512
    Const CONGRUENT_BITS As Long = 448

    lNumberOfWords = (((msgLen + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)

    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= msgLen
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(bytes(lByteCount), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 1) = LShift(msgLen, 3)
    lWordArray(lNumberOfWords - 2) = RShift(msgLen, 29)

    bytesToWordArraySHA = lWordArray
End Function

' Converts msgLen bytes to an array of 32-bit words (big-endian) padded to
' 1024-bit blocks with a 128-bit message-length field, for SHA-384/SHA-512.
'
' The layout matches bytesToWordArraySHA but uses 1024-bit blocks (32 Longs
' each) and reserves the final 16 bytes for the length. Each 64-bit SHA word is
' therefore the pair (lWordArray(2j), lWordArray(2j + 1)). msgLen*8 always fits
' in the low 64 bits, so the upper two length Longs stay zero.
Private Function bytesToWordArraySHA512(bytes() As Byte, ByVal msgLen As Long) As Long()
    Dim lNumberOfWords  As Long
    Dim lWordArray()    As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long

    Const MODULUS_BITS As Long = 1024
    Const CONGRUENT_BITS As Long = 896

    lNumberOfWords = (((msgLen + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)

    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= msgLen
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(bytes(lByteCount), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 1) = LShift(msgLen, 3)
    lWordArray(lNumberOfWords - 2) = RShift(msgLen, 29)

    bytesToWordArraySHA512 = lWordArray
End Function

Private Function todo(s)
  Debug.Print "TODO: " & s
End Function

Public Sub createHash(Optional ByVal name As String = "sha256")
    Select Case LCase(name)
        Case "md5", "sha1", "sha224", "sha256", "sha384", "sha512", "sha512-256", "sha512-224"
            hashMode = LCase(name)
        Case Else
            Err.Raise vbObjectError + 1, "CRYPTO.createHash", "unknown hash algorithm: " & name
    End Select
    mLen = 0
    mCap = 0
    Erase mBuf
    isHmac = False
    Erase hmacKeyBytes
End Sub

' Starts an HMAC keyed by `key` (a String, encoded as UTF-8, or a Byte array)
' using hash algorithm `name`. Then feed the message with update() and finalize
' with digest()/digestBytes(), exactly as for a plain hash.
Public Sub createHmac(Optional ByVal name As String = "sha256", Optional key As Variant = "")
    createHash name
    hmacKeyBytes = toBytes(key)
    isHmac = True
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

Private Function calculateSHA1(words() As Long) As Long()
    Dim H(4) As Long
    H(0) = &H67452301
    H(1) = &HEFCDAB89
    H(2) = &H98BADCFE
    H(3) = &H10325476
    H(4) = &HC3D2E1F0

    Dim W(79) As Long
    Dim i As Long, t As Long
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, temp As Long
    Dim k As Long, f As Long

    For t = 0 To UBound(words) Step 16
        ' Prepare the message schedule
        For i = 0 To 15
            W(i) = words(t + i)
        Next i
        For i = 16 To 79
            W(i) = RotateLeft(W(i - 3) Xor W(i - 8) Xor W(i - 14) Xor W(i - 16), 1)
        Next i

        a = H(0)
        b = H(1)
        c = H(2)
        d = H(3)
        e = H(4)

        For i = 0 To 79
            If i >= 0 And i <= 19 Then
                f = (b And c) Or ((Not b) And d)
                k = &H5A827999
            ElseIf i >= 20 And i <= 39 Then
                f = b Xor c Xor d
                k = &H6ED9EBA1
            ElseIf i >= 40 And i <= 59 Then
                f = (b And c) Or (b And d) Or (c And d)
                k = &H8F1BBCDC
            Else
                f = b Xor c Xor d
                k = &HCA62C1D6
            End If

            temp = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(RotateLeft(a, 5), f), e), k), W(i))
            e = d
            d = c
            c = RotateLeft(b, 30)
            b = a
            a = temp
        Next i

        H(0) = AddUnsigned(H(0), a)
        H(1) = AddUnsigned(H(1), b)
        H(2) = AddUnsigned(H(2), c)
        H(3) = AddUnsigned(H(3), d)
        H(4) = AddUnsigned(H(4), e)
    Next t

    calculateSHA1 = H
End Function

' SHA-256 / SHA-224 compression core. iv is an 8-element array of the initial
' hash values (SHA-256 and SHA-224 differ only in iv and in the output length,
' which the caller truncates). Returns the full 8-word (256-bit) state.
Private Function calculateSHA256Core(words() As Long, iv) As Long()
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
    Dim m As Long
    For m = 0 To 7
        result(m) = CLng(iv(m))
    Next m
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
    calculateSHA256Core = result
End Function

' --- 64-bit helpers -----------------------------------------------------------
' VBA's Long is 32-bit, but SHA-384/512 operate on 64-bit words. Each 64-bit
' value is carried in a T64 (hi/lo Long halves). RShift/LShift below are 32-bit
' zero-fill (logical) shifts provided earlier in this module. To stay within
' classic VBA, function results are always assigned to a variable before being
' passed on (UDT arguments are ByRef, so passing a call result directly is not
' portable).

' Builds a 64-bit value from its high and low 32-bit halves.
Private Function mk64(ByVal hi As Long, ByVal lo As Long) As T64
    mk64.hi = hi
    mk64.lo = lo
End Function

' Interprets a 32-bit Long as an unsigned value in the range [0, 2^32).
Private Function unsignedVal(ByVal x As Long) As Double
    If x < 0 Then
        unsignedVal = CDbl(x) + 4294967296#
    Else
        unsignedVal = CDbl(x)
    End If
End Function

' Converts an unsigned value in [0, 2^32) back to a signed 32-bit Long.
Private Function doubleToLong(ByVal d As Double) As Long
    If d >= 2147483648# Then
        doubleToLong = CLng(d - 4294967296#)
    Else
        doubleToLong = CLng(d)
    End If
End Function

' 64-bit modular addition (mod 2^64), carrying from the low half to the high half.
Private Function Add64(a As T64, b As T64) As T64
    Dim loSum As Double
    Dim carry As Long
    loSum = unsignedVal(a.lo) + unsignedVal(b.lo)
    If loSum >= 4294967296# Then
        carry = 1
        loSum = loSum - 4294967296#
    Else
        carry = 0
    End If
    Add64.lo = doubleToLong(loSum)
    Add64.hi = AddUnsigned(AddUnsigned(a.hi, b.hi), carry)
End Function

Private Function Xor64(a As T64, b As T64) As T64
    Xor64.hi = a.hi Xor b.hi
    Xor64.lo = a.lo Xor b.lo
End Function

Private Function And64(a As T64, b As T64) As T64
    And64.hi = a.hi And b.hi
    And64.lo = a.lo And b.lo
End Function

Private Function Not64(a As T64) As T64
    Not64.hi = Not a.hi
    Not64.lo = Not a.lo
End Function

' Rotates a 64-bit value right by n bits (0 <= n < 64).
Private Function Rotr64(a As T64, ByVal n As Long) As T64
    Dim hi As Long, lo As Long, m As Long
    hi = a.hi
    lo = a.lo
    n = n Mod 64
    If n = 0 Then
        Rotr64.hi = hi
        Rotr64.lo = lo
    ElseIf n < 32 Then
        Rotr64.hi = RShift(hi, n) Or LShift(lo, 32 - n)
        Rotr64.lo = RShift(lo, n) Or LShift(hi, 32 - n)
    ElseIf n = 32 Then
        Rotr64.hi = lo
        Rotr64.lo = hi
    Else
        m = n - 32
        Rotr64.hi = RShift(lo, m) Or LShift(hi, 32 - m)
        Rotr64.lo = RShift(hi, m) Or LShift(lo, 32 - m)
    End If
End Function

' Logical (zero-fill) right shift of a 64-bit value by n bits (0 <= n < 64).
Private Function Shr64(a As T64, ByVal n As Long) As T64
    Dim hi As Long, lo As Long
    hi = a.hi
    lo = a.lo
    If n = 0 Then
        Shr64.hi = hi
        Shr64.lo = lo
    ElseIf n < 32 Then
        Shr64.hi = RShift(hi, n)
        Shr64.lo = RShift(lo, n) Or LShift(hi, 32 - n)
    ElseIf n = 32 Then
        Shr64.hi = 0
        Shr64.lo = hi
    Else
        Shr64.hi = 0
        Shr64.lo = RShift(hi, n - 32)
    End If
End Function

Private Function bigSigma0_512(x As T64) As T64
    Dim r1 As T64, r2 As T64, r3 As T64, t As T64
    r1 = Rotr64(x, 28)
    r2 = Rotr64(x, 34)
    r3 = Rotr64(x, 39)
    t = Xor64(r1, r2)
    bigSigma0_512 = Xor64(t, r3)
End Function

Private Function bigSigma1_512(x As T64) As T64
    Dim r1 As T64, r2 As T64, r3 As T64, t As T64
    r1 = Rotr64(x, 14)
    r2 = Rotr64(x, 18)
    r3 = Rotr64(x, 41)
    t = Xor64(r1, r2)
    bigSigma1_512 = Xor64(t, r3)
End Function

Private Function littleSigma0_512(x As T64) As T64
    Dim r1 As T64, r2 As T64, r3 As T64, t As T64
    r1 = Rotr64(x, 1)
    r2 = Rotr64(x, 8)
    r3 = Shr64(x, 7)
    t = Xor64(r1, r2)
    littleSigma0_512 = Xor64(t, r3)
End Function

Private Function littleSigma1_512(x As T64) As T64
    Dim r1 As T64, r2 As T64, r3 As T64, t As T64
    r1 = Rotr64(x, 19)
    r2 = Rotr64(x, 61)
    r3 = Shr64(x, 6)
    t = Xor64(r1, r2)
    littleSigma1_512 = Xor64(t, r3)
End Function

Private Function ch64(x As T64, y As T64, z As T64) As T64
    Dim nx As T64, t1 As T64, t2 As T64
    nx = Not64(x)
    t1 = And64(x, y)
    t2 = And64(nx, z)
    ch64 = Xor64(t1, t2)
End Function

Private Function maj64(x As T64, y As T64, z As T64) As T64
    Dim t1 As T64, t2 As T64, t3 As T64, t As T64
    t1 = And64(x, y)
    t2 = And64(x, z)
    t3 = And64(y, z)
    t = Xor64(t1, t2)
    maj64 = Xor64(t, t3)
End Function

' SHA-512 family compression core. Hhi/Hlo are the eight 64-bit initial hash
' values (high/low halves). Returns the full 8-word (512-bit) state as 16 Longs
' (hi, lo, hi, lo, ...); shorter variants (384, 512/256, 512/224) are produced by
' the caller truncating the byte output.
Private Function calculateSHA512Core(words() As Long, Hhi, Hlo) As Long()
    Dim Khi As Variant: Khi = Array( _
        &H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
        &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
        &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
        &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
        &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
        &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
        &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
        &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2, _
        &HCA273ECE, &HD186B8C7, &HEADA7DD6, &HF57D4F7F, &H6F067AA, &HA637DC5, &H113F9804, &H1B710B35, _
        &H28DB77F5, &H32CAAB7B, &H3C9EBE0A, &H431D67C4, &H4CC5D4BE, &H597F299C, &H5FCB6FAB, &H6C44198C)
    Dim Klo As Variant: Klo = Array( _
        &HD728AE22, &H23EF65CD, &HEC4D3B2F, &H8189DBBC, &HF348B538, &HB605D019, &HAF194F9B, &HDA6D8118, _
        &HA3030242, &H45706FBE, &H4EE4B28C, &HD5FFB4E2, &HF27B896F, &H3B1696B1, &H25C71235, &HCF692694, _
        &H9EF14AD2, &H384F25E3, &H8B8CD5B5, &H77AC9C65, &H592B0275, &H6EA6E483, &HBD41FBD4, &H831153B5, _
        &HEE66DFAB, &H2DB43210, &H98FB213F, &HBEEF0EE4, &H3DA88FC2, &H930AA725, &HE003826F, &HA0E6E70, _
        &H46D22FFC, &H5C26C926, &H5AC42AED, &H9D95B3DF, &H8BAF63DE, &H3C77B2A8, &H47EDAEE6, &H1482353B, _
        &H4CF10364, &HBC423001, &HD0F89791, &H654BE30, &HD6EF5218, &H5565A910, &H5771202A, &H32BBD1B8, _
        &HB8D2D0C8, &H5141AB53, &HDF8EEB99, &HE19B48A8, &HC5C95A63, &HE3418ACB, &H7763E373, &HD6B2B8A3, _
        &H5DEFB2FC, &H43172F60, &HA1F0AB72, &H1A6439EC, &H23631E28, &HDE82BDE9, &HB2C67915, &HE372532B, _
        &HEA26619C, &H21C0C207, &HCDE0EB1E, &HEE6ED178, &H72176FBA, &HA2C898A6, &HBEF90DAE, &H131C471B, _
        &H23047D84, &H40C72493, &H15C9BEBC, &H9C100D4C, &HCB3E42B6, &HFC657E2A, &H3AD6FAEC, &H4A475817)

    ' hs = hash state. (Named "hs" not "H": VBA identifiers are case-insensitive,
    ' so an array "H" would collide with the working variable "h" below.)
    Dim hs(7) As T64
    Dim i As Long, k As Long
    For i = 0 To 7
        hs(i) = mk64(CLng(Hhi(i)), CLng(Hlo(i)))
    Next i

    Dim W(79) As T64
    Dim a As T64, b As T64, c As T64, d As T64, e As T64, f As T64, g As T64, h As T64
    Dim t1 As T64, t2 As T64
    Dim s0 As T64, s1 As T64, tmp As T64
    Dim bs1e As T64, chv As T64, kv As T64, bs0a As T64, mjv As T64
    For k = 0 To UBound(words) Step 32
        For i = 0 To 15
            W(i) = mk64(words(k + 2 * i), words(k + 2 * i + 1))
        Next i
        For i = 16 To 79
            s1 = littleSigma1_512(W(i - 2))
            s0 = littleSigma0_512(W(i - 15))
            tmp = Add64(s1, W(i - 7))
            tmp = Add64(tmp, s0)
            W(i) = Add64(tmp, W(i - 16))
        Next i

        a = hs(0): b = hs(1): c = hs(2): d = hs(3)
        e = hs(4): f = hs(5): g = hs(6): h = hs(7)

        For i = 0 To 79
            bs1e = bigSigma1_512(e)
            chv = ch64(e, f, g)
            kv = mk64(CLng(Khi(i)), CLng(Klo(i)))
            t1 = Add64(h, bs1e)
            t1 = Add64(t1, chv)
            t1 = Add64(t1, kv)
            t1 = Add64(t1, W(i))
            bs0a = bigSigma0_512(a)
            mjv = maj64(a, b, c)
            t2 = Add64(bs0a, mjv)
            h = g: g = f: f = e
            e = Add64(d, t1)
            d = c: c = b: b = a
            a = Add64(t1, t2)
        Next i

        hs(0) = Add64(hs(0), a)
        hs(1) = Add64(hs(1), b)
        hs(2) = Add64(hs(2), c)
        hs(3) = Add64(hs(3), d)
        hs(4) = Add64(hs(4), e)
        hs(5) = Add64(hs(5), f)
        hs(6) = Add64(hs(6), g)
        hs(7) = Add64(hs(7), h)
    Next k

    Dim result() As Long
    ReDim result(15)
    For i = 0 To 7
        result(2 * i) = hs(i).hi
        result(2 * i + 1) = hs(i).lo
    Next i
    calculateSHA512Core = result
End Function

' Converts an array of Longs to its big-endian byte sequence (for SHA output).
Private Function longsToBytesBE(arr() As Long) As Byte()
    Dim bytes() As Byte
    Dim i As Long, v As Long, base As Long
    ReDim bytes((UBound(arr) - LBound(arr) + 1) * 4 - 1)
    For i = 0 To UBound(arr) - LBound(arr)
        v = arr(LBound(arr) + i)
        base = i * 4
        bytes(base) = CByte(RShift(v, 24) And &HFF&)
        bytes(base + 1) = CByte(RShift(v, 16) And &HFF&)
        bytes(base + 2) = CByte(RShift(v, 8) And &HFF&)
        bytes(base + 3) = CByte(v And &HFF&)
    Next i
    longsToBytesBE = bytes
End Function

' Converts an array of Longs to its little-endian byte sequence (for MD5 output).
Private Function longsToBytesLE(arr() As Long) As Byte()
    Dim bytes() As Byte
    Dim i As Long, v As Long, base As Long
    ReDim bytes((UBound(arr) - LBound(arr) + 1) * 4 - 1)
    For i = 0 To UBound(arr) - LBound(arr)
        v = arr(LBound(arr) + i)
        base = i * 4
        bytes(base) = CByte(v And &HFF&)
        bytes(base + 1) = CByte(RShift(v, 8) And &HFF&)
        bytes(base + 2) = CByte(RShift(v, 16) And &HFF&)
        bytes(base + 3) = CByte(RShift(v, 24) And &HFF&)
    Next i
    longsToBytesLE = bytes
End Function

' Returns the first n bytes of b (used to truncate to a shorter digest length).
Private Function truncateBytes(b() As Byte, ByVal n As Long) As Byte()
    Dim r() As Byte, i As Long
    ReDim r(n - 1)
    For i = 0 To n - 1
        r(i) = b(i)
    Next i
    truncateBytes = r
End Function

' Encodes a VBA (UTF-16) string as UTF-8 bytes, matching how Node/openssl hash
' text. Handles the full BMP plus surrogate pairs (code points > U+FFFF).
' Returns an unallocated (zero-length) array for an empty string.
Private Function utf8Bytes(s As String) As Byte()
    Dim outArr() As Byte
    Dim cap As Long, cnt As Long, i As Long, L As Long
    Dim cu As Long, cu2 As Long, cp As Long
    L = Len(s)
    If L = 0 Then Exit Function
    cap = L + 8
    ReDim outArr(cap - 1)
    i = 1
    Do While i <= L
        cu = AscW(Mid(s, i, 1)) And &HFFFF&
        If cu >= &HD800& And cu <= &HDBFF& And i < L Then
            cu2 = AscW(Mid(s, i + 1, 1)) And &HFFFF&
            If cu2 >= &HDC00& And cu2 <= &HDFFF& Then
                cp = &H10000 + ((cu - &HD800&) * &H400&) + (cu2 - &HDC00&)
                i = i + 1
            Else
                cp = cu
            End If
        Else
            cp = cu
        End If
        i = i + 1

        If cnt + 4 > cap Then
            cap = cap * 2
            ReDim Preserve outArr(cap - 1)
        End If

        If cp < &H80 Then
            outArr(cnt) = CByte(cp): cnt = cnt + 1
        ElseIf cp < &H800 Then
            outArr(cnt) = CByte(&HC0 Or RShift(cp, 6)): cnt = cnt + 1
            outArr(cnt) = CByte(&H80 Or (cp And &H3F)): cnt = cnt + 1
        ElseIf cp < &H10000 Then
            outArr(cnt) = CByte(&HE0 Or RShift(cp, 12)): cnt = cnt + 1
            outArr(cnt) = CByte(&H80 Or (RShift(cp, 6) And &H3F)): cnt = cnt + 1
            outArr(cnt) = CByte(&H80 Or (cp And &H3F)): cnt = cnt + 1
        Else
            outArr(cnt) = CByte(&HF0 Or RShift(cp, 18)): cnt = cnt + 1
            outArr(cnt) = CByte(&H80 Or (RShift(cp, 12) And &H3F)): cnt = cnt + 1
            outArr(cnt) = CByte(&H80 Or (RShift(cp, 6) And &H3F)): cnt = cnt + 1
            outArr(cnt) = CByte(&H80 Or (cp And &H3F)): cnt = cnt + 1
        End If
    Loop
    ReDim Preserve outArr(cnt - 1)
    utf8Bytes = outArr
End Function

' Number of elements in a Byte array, or 0 if it is unallocated.
Private Function byteLen(b() As Byte) As Long
    On Error Resume Next
    byteLen = UBound(b) - LBound(b) + 1
    If Err.Number <> 0 Then byteLen = 0
    On Error GoTo 0
End Function

' Grows mBuf so it can hold at least `need` bytes (doubling to keep appends cheap).
Private Sub ensureCapacity(ByVal need As Long)
    If need <= mCap Then Exit Sub
    Dim newCap As Long
    newCap = mCap
    If newCap < 64 Then newCap = 64
    Do While newCap < need
        newCap = newCap * 2
    Loop
    ReDim Preserve mBuf(newCap - 1)
    mCap = newCap
End Sub

' Appends src to the message buffer.
Private Sub appendBytes(src() As Byte)
    Dim n As Long, i As Long
    n = byteLen(src)
    If n = 0 Then Exit Sub
    ensureCapacity mLen + n
    For i = 0 To n - 1
        mBuf(mLen + i) = src(LBound(src) + i)
    Next i
    mLen = mLen + n
End Sub

' Hashes the first n bytes of `bytes` with the current hashMode and returns the
' raw digest bytes.
Private Function rawHash(bytes() As Byte, ByVal n As Long) As Byte()
    Dim iv As Variant, Hhi As Variant, Hlo As Variant
    Select Case hashMode
        Case "md5"
            rawHash = longsToBytesLE(calculateMD5(bytesToWordArrayMD5(bytes, n)))
        Case "sha1"
            rawHash = longsToBytesBE(calculateSHA1(bytesToWordArraySHA(bytes, n)))
        Case "sha224"
            iv = Array(&HC1059ED8, &H367CD507, &H3070DD17, &HF70E5939, &HFFC00B31, &H68581511, &H64F98FA7, &HBEFA4FA4)
            rawHash = truncateBytes(longsToBytesBE(calculateSHA256Core(bytesToWordArraySHA(bytes, n), iv)), 28)
        Case "sha256"
            iv = Array(&H6A09E667, &HBB67AE85, &H3C6EF372, &HA54FF53A, &H510E527F, &H9B05688C, &H1F83D9AB, &H5BE0CD19)
            rawHash = longsToBytesBE(calculateSHA256Core(bytesToWordArraySHA(bytes, n), iv))
        Case "sha384"
            Hhi = Array(&HCBBB9D5D, &H629A292A, &H9159015A, &H152FECD8, &H67332667, &H8EB44A87, &HDB0C2E0D, &H47B5481D)
            Hlo = Array(&HC1059ED8, &H367CD507, &H3070DD17, &HF70E5939, &HFFC00B31, &H68581511, &H64F98FA7, &HBEFA4FA4)
            rawHash = truncateBytes(longsToBytesBE(calculateSHA512Core(bytesToWordArraySHA512(bytes, n), Hhi, Hlo)), 48)
        Case "sha512"
            Hhi = Array(&H6A09E667, &HBB67AE85, &H3C6EF372, &HA54FF53A, &H510E527F, &H9B05688C, &H1F83D9AB, &H5BE0CD19)
            Hlo = Array(&HF3BCC908, &H84CAA73B, &HFE94F82B, &H5F1D36F1, &HADE682D1, &H2B3E6C1F, &HFB41BD6B, &H137E2179)
            rawHash = longsToBytesBE(calculateSHA512Core(bytesToWordArraySHA512(bytes, n), Hhi, Hlo))
        Case "sha512-256"
            Hhi = Array(&H22312194, &H9F555FA3, &H2393B86B, &H96387719, &H96283EE2, &HBE5E1E25, &H2B0199FC, &HEB72DDC)
            Hlo = Array(&HFC2BF72C, &HC84C64C2, &H6F53B151, &H5940EABD, &HA88EFFE3, &H53863992, &H2C85B8AA, &H81C52CA2)
            rawHash = truncateBytes(longsToBytesBE(calculateSHA512Core(bytesToWordArraySHA512(bytes, n), Hhi, Hlo)), 32)
        Case "sha512-224"
            Hhi = Array(&H8C3D37C8, &H73E19966, &H1DFAB7AE, &H679DD514, &HF6D2B69, &H77E36F73, &H3F9D85A8, &H1112E6AD)
            Hlo = Array(&H19544DA2, &H89DCD4D6, &H32FF9C82, &H582F9FCF, &H7BD44DA8, &H4C48942, &H6A1D36C8, &H91D692A1)
            rawHash = truncateBytes(longsToBytesBE(calculateSHA512Core(bytesToWordArraySHA512(bytes, n), Hhi, Hlo)), 28)
        Case Else
            Err.Raise vbObjectError + 1, "CRYPTO", "unknown hash algorithm: " & hashMode
    End Select
End Function

' Input block size (in bytes) of the current hash, needed for HMAC padding.
Private Function blockSize() As Long
    Select Case hashMode
        Case "sha384", "sha512", "sha512-256", "sha512-224"
            blockSize = 128
        Case Else
            blockSize = 64
    End Select
End Function

' HMAC(K, msg) = H((K' XOR opad) || H((K' XOR ipad) || msg)) for the current
' hashMode, where K' is the key hashed down if longer than the block, then
' zero-padded to the block size. Pure over its arguments (used by both the
' instance HMAC flow and PBKDF2).
Private Function hmacRaw(key() As Byte, ByVal keyLen As Long, msg() As Byte, ByVal msgLen As Long) As Byte()
    Dim bs As Long: bs = blockSize()
    Dim k() As Byte, kSrc() As Byte, kSrcLen As Long, i As Long
    ReDim k(bs - 1)
    If keyLen > bs Then
        kSrc = rawHash(key, keyLen)
        kSrcLen = byteLen(kSrc)
    Else
        kSrc = key
        kSrcLen = keyLen
    End If
    For i = 0 To bs - 1
        If i < kSrcLen Then k(i) = kSrc(i) Else k(i) = 0
    Next i

    ' inner = H((K' XOR 0x36) || msg)
    Dim inner() As Byte
    ReDim inner(bs + msgLen - 1)
    For i = 0 To bs - 1
        inner(i) = k(i) Xor &H36
    Next i
    For i = 0 To msgLen - 1
        inner(bs + i) = msg(i)
    Next i
    Dim innerHash() As Byte, iLen As Long
    innerHash = rawHash(inner, bs + msgLen)
    iLen = byteLen(innerHash)

    ' outer = H((K' XOR 0x5C) || inner)
    Dim outer() As Byte
    ReDim outer(bs + iLen - 1)
    For i = 0 To bs - 1
        outer(i) = k(i) Xor &H5C
    Next i
    For i = 0 To iLen - 1
        outer(bs + i) = innerHash(i)
    Next i
    hmacRaw = rawHash(outer, bs + iLen)
End Function

Private Function computeHmac() As Byte()
    computeHmac = hmacRaw(hmacKeyBytes, byteLen(hmacKeyBytes), mBuf, mLen)
End Function

' Output length in bytes of the current hash algorithm.
Private Function digestLen() As Long
    Select Case hashMode
        Case "md5": digestLen = 16
        Case "sha1": digestLen = 20
        Case "sha224", "sha512-224": digestLen = 28
        Case "sha256", "sha512-256": digestLen = 32
        Case "sha384": digestLen = 48
        Case "sha512": digestLen = 64
        Case Else
            Err.Raise vbObjectError + 1, "CRYPTO", "unknown hash algorithm: " & hashMode
    End Select
End Function

' Coerces a String (UTF-8) or Byte-array argument to a Byte array.
Private Function toBytes(v) As Byte()
    Select Case VarType(v)
        Case vbString
            toBytes = utf8Bytes(CStr(v))
        Case vbArray + vbByte
            toBytes = v
        Case Else
            Err.Raise 5, "CRYPTO", "expected a String or a Byte array"
    End Select
End Function

' PBKDF2 (RFC 2898) using HMAC-<name> as the PRF. Derives keyLen bytes from
' password and salt with `iterations` rounds. password/salt may be String
' (UTF-8) or Byte arrays. NOTE: this configures the object's algorithm, so it
' resets any in-progress update() state.
Public Function pbkdf2Bytes(password, salt, ByVal iterations As Long, ByVal keyLen As Long, Optional ByVal name As String = "sha256") As Byte()
    If iterations < 1 Then Err.Raise 5, "CRYPTO.pbkdf2", "iterations must be >= 1"
    If keyLen < 1 Then Err.Raise 5, "CRYPTO.pbkdf2", "keyLen must be >= 1"
    createHash name

    Dim pw() As Byte, sl() As Byte, pwLen As Long, slLen As Long
    pw = toBytes(password): pwLen = byteLen(pw)
    sl = toBytes(salt): slLen = byteLen(sl)

    Dim hLen As Long: hLen = digestLen()
    Dim dk() As Byte: ReDim dk(keyLen - 1)
    Dim dkPos As Long: dkPos = 0

    ' msg = salt || INT_32_BE(blockIndex)
    Dim msg() As Byte: ReDim msg(slLen + 4 - 1)
    Dim i As Long, j As Long, c As Long, blockIndex As Long
    For i = 0 To slLen - 1
        msg(i) = sl(i)
    Next i

    Dim u() As Byte, t() As Byte
    blockIndex = 1
    Do While dkPos < keyLen
        msg(slLen) = CByte(RShift(blockIndex, 24) And &HFF&)
        msg(slLen + 1) = CByte(RShift(blockIndex, 16) And &HFF&)
        msg(slLen + 2) = CByte(RShift(blockIndex, 8) And &HFF&)
        msg(slLen + 3) = CByte(blockIndex And &HFF&)

        u = hmacRaw(pw, pwLen, msg, slLen + 4)   ' U_1
        t = u                                    ' T = U_1
        For c = 2 To iterations
            u = hmacRaw(pw, pwLen, u, hLen)       ' U_c = PRF(pw, U_{c-1})
            For j = 0 To hLen - 1
                t(j) = t(j) Xor u(j)
            Next j
        Next c

        For j = 0 To hLen - 1
            If dkPos < keyLen Then
                dk(dkPos) = t(j)
                dkPos = dkPos + 1
            End If
        Next j
        blockIndex = blockIndex + 1
    Loop

    pbkdf2Bytes = dk
End Function

' PBKDF2 with the derived key returned as a "hex", "base64" or "base64url" string.
Public Function pbkdf2(password, salt, ByVal iterations As Long, ByVal keyLen As Long, Optional ByVal name As String = "sha256", Optional ByVal outMode As String = "hex") As String
    Dim dk() As Byte
    dk = pbkdf2Bytes(password, salt, iterations, keyLen, name)
    pbkdf2 = encodeBytes(dk, outMode)
End Function

' Finalizes the current message as either a plain hash or an HMAC.
Private Function finalBytes() As Byte()
    If isHmac Then
        finalBytes = computeHmac()
    Else
        finalBytes = rawHash(mBuf, mLen)
    End If
End Function

' Appends data to the message. data may be a String (encoded as UTF-8) or a
' Byte array. Call repeatedly to hash data in pieces, then finalize with digest.
Public Sub update(data)
    Dim b() As Byte
    b = toBytes(data)
    appendBytes b
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

' Encodes bytes as a "hex", "base64" or "base64url" string.
Private Function encodeBytes(raw() As Byte, ByVal mode As String) As String
    Select Case LCase(mode)
        Case "hex"
            encodeBytes = bytesToHex(raw)
        Case "base64"
            encodeBytes = bytesToBase64(raw)
        Case "base64url"
            ' URL-safe alphabet, no padding (RFC 4648 §5), as used by JWT etc.
            encodeBytes = Replace(Replace(Replace(bytesToBase64(raw), "+", "-"), "/", "_"), "=", "")
        Case Else
            Err.Raise 5, "CRYPTO", "unknown output mode: " & mode
    End Select
End Function

' Returns the digest as a "hex", "base64" or "base64url" string. May be called
' more than once; it does not consume the buffered message.
Public Function digest(Optional ByVal mode As String = "hex") As String
    Dim raw() As Byte
    raw = finalBytes()
    digest = encodeBytes(raw, mode)
End Function

' Returns the digest as a raw Byte array.
Public Function digestBytes() As Byte()
    digestBytes = finalBytes()
End Function

' Compares the digest against an expected value in constant time, to avoid
' leaking (via early-exit timing) how many leading characters matched -- the
' standard way to verify a MAC. mode is the encoding of `expected` ("hex" by
' default; "base64"/"base64url" also accepted). Hex compares case-insensitively.
Public Function verify(ByVal expected As String, Optional ByVal mode As String = "hex") As Boolean
    Dim actual As String
    actual = digest(mode)
    If LCase(mode) = "hex" Then
        actual = LCase(actual)
        expected = LCase(expected)
    End If
    verify = constantTimeEquals(actual, expected)
End Function

' Timing-safe string equality: always scans the full length and never exits
' early, so the running time does not depend on where the first mismatch is.
Private Function constantTimeEquals(ByVal a As String, ByVal b As String) As Boolean
    Dim diff As Long, n As Long, i As Long, la As Long, lb As Long
    la = Len(a)
    lb = Len(b)
    n = la
    If lb > n Then n = lb
    diff = la Xor lb
    For i = 1 To n
        Dim ca As Long, cb As Long
        If i <= la Then ca = Asc(Mid(a, i, 1)) Else ca = 0
        If i <= lb Then cb = Asc(Mid(b, i, 1)) Else cb = 0
        diff = diff Or (ca Xor cb)
    Next i
    constantTimeEquals = (diff = 0)
End Function
