Attribute VB_Name = "CRYPTO_test"
'
' Test suite for CRYPTO.cls.
'
' How to run:
'   1. Import CRYPTO.cls and this file into the VBA project (see README).
'   2. In the Immediate window (Ctrl+G) type:  testAll
'   3. Read the summary line. Any FAIL prints the expected/actual digests.
'
Option Explicit

Private passCount As Long
Private failCount As Long

Private Sub assertEqual(ByVal label As String, ByVal expected As String, ByVal actual As String)
    If expected = actual Then
        passCount = passCount + 1
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL  " & label
        Debug.Print "        expected: " & expected
        Debug.Print "        actual:   " & actual
    End If
End Sub

' Hash a whole string (UTF-8) in one shot.
Private Sub test(ByVal mode As String, ByVal s As String, ByVal format As String, ByVal expected As String)
    Dim c As CRYPTO: Set c = New CRYPTO
    c.createHash mode
    c.update s
    assertEqual mode & " / " & format & " / len=" & Len(s), expected, c.digest(format)
End Sub

Private Sub testAll()
    passCount = 0
    failCount = 0

    ' --- empty string ---------------------------------------------------------
    test "md5", "", "hex", "d41d8cd98f00b204e9800998ecf8427e"
    test "sha1", "", "hex", "da39a3ee5e6b4b0d3255bfef95601890afd80709"
    test "sha224", "", "hex", "d14a028c2a3a2bc9476102bb288234c415a2b01f828ea62ac5b3e42f"
    test "sha256", "", "hex", "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    test "sha384", "", "hex", "38b060a751ac96384cd9327eb1b1e36a21fdb71114be07434c0cc7bf63f6e1da274edebfe76f65fbd51ad2f14898b95b"
    test "sha512", "", "hex", "cf83e1357eefb8bdf1542850d66d8007d620e4050b5715dc83f4a921d36ce9ce47d0d13c5d85f2b0ff8318d2877eec2f63b931bd47417a81a538327af927da3e"
    test "sha512-256", "", "hex", "c672b8d1ef56ed28ab87c3622c5114069bdd3ad7b8f9737498d0c01ecef0967a"
    test "sha512-224", "", "hex", "6ed0dd02806fa89e25de060c19d3ac86cabb87d6a0ddd05c333b84f4"

    ' --- "abc" ----------------------------------------------------------------
    test "md5", "abc", "hex", "900150983cd24fb0d6963f7d28e17f72"
    test "sha1", "abc", "hex", "a9993e364706816aba3e25717850c26c9cd0d89d"
    test "sha224", "abc", "hex", "23097d223405d8228642a477bda255b32aadbce4bda0b3f7e36c9da7"
    test "sha256", "abc", "hex", "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    test "sha384", "abc", "hex", "cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed8086072ba1e7cc2358baeca134c825a7"
    test "sha512", "abc", "hex", "ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f"
    test "sha512-256", "abc", "hex", "53048e2681941ef99b2e29b76b4c7dabe4c2d0c634fc6d46e0e2f13107e7af23"
    test "sha512-224", "abc", "hex", "4634270f707b6a54daae7530460842e20e37ed265ceee9a43e8924aa"

    ' --- "hello world" --------------------------------------------------------
    test "md5", "hello world", "hex", "5eb63bbbe01eeed093cb22bb8f5acdc3"
    test "sha1", "hello world", "hex", "2aae6c35c94fcfb415dbe95f408b9ce91ee846ed"
    test "sha224", "hello world", "hex", "2f05477fc24bb4faefd86517156dafdecec45b8ad3cf2522a563582b"
    test "sha256", "hello world", "hex", "b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9"
    test "sha384", "hello world", "hex", "fdbd8e75a67f29f701a4e040385e2e23986303ea10239211af907fcbb83578b3e417cb71ce646efd0819dd8c088de1bd"
    test "sha512", "hello world", "hex", "309ecc489c12d6eb4cc40f50c902f2b4d0ed77ee511a7c7a9bcd3ca86d4cd86f989dd35bc5ff499670da34255b45b0cfd830e81f605dcf7dc5542e93ae9cd76f"
    test "sha512-256", "hello world", "hex", "0ac561fac838104e3f2e4ad107b4bee3e938bf15f2b15f009ccccd61a913f017"
    test "sha512-224", "hello world", "hex", "22e0d52336f64a998085078b05a6e37b26f8120f43bf4db4c43a64ee"

    ' --- base64 output --------------------------------------------------------
    test "md5", "hello world", "base64", "XrY7u+Ae7tCTyyK7j1rNww=="
    test "sha1", "hello world", "base64", "Kq5sNclPz7QV2+lfQIuc6R7oRu0="
    test "sha224", "hello world", "base64", "LwVHf8JLtPrv2GUXFW2v3s7EW4rTzyUipWNYKw=="
    test "sha256", "hello world", "base64", "uU0nuZNNPgilLlLX2n2r+sSE7+N6U4DukIj3rOLvzek="
    test "sha512-256", "hello world", "base64", "CsVh+sg4EE4/LkrRB7S+4+k4vxXysV8AnMzNYakT8Bc="

    ' --- padding boundaries: SHA-256 family straddles 55/56/63/64 bytes -------
    test "sha256", String(55, "a"), "hex", "9f4390f8d30c2dd92ec9f095b65e2b9ae9b0a925a5258e241c9f1e910f734318"
    test "sha256", String(56, "a"), "hex", "b35439a4ac6f0948b6d6f9e3c6af0f5f590ce20f1bde7090ef7970686ec6738a"
    test "sha256", String(63, "a"), "hex", "7d3e74a05d7db15bce4ad9ec0658ea98e3f06eeecf16b4c6fff2da457ddc2f34"
    test "sha256", String(64, "a"), "hex", "ffe054fe7ae0cb6dc65c3af9b61d5209f439851db43d0ba5997337df154668eb"
    test "sha256", String(65, "a"), "hex", "635361c48bb9eab14198e76ea8ab7f1a41685d6ad62aa9146d301d4f17eb0ae0"
    test "sha256", String(200, "a"), "hex", "c2a908d98f5df987ade41b5fce213067efbcc21ef2240212a41e54b5e7c28ae5"

    ' --- padding boundaries: SHA-512 family straddles 111/112/127/128 bytes ---
    test "sha512", String(111, "a"), "hex", "fa9121c7b32b9e01733d034cfc78cbf67f926c7ed83e82200ef86818196921760b4beff48404df811b953828274461673c68d04e297b0eb7b2b4d60fc6b566a2"
    test "sha512", String(112, "a"), "hex", "c01d080efd492776a1c43bd23dd99d0a2e626d481e16782e75d54c2503b5dc32bd05f0f1ba33e568b88fd2d970929b719ecbb152f58f130a407c8830604b70ca"
    test "sha512", String(127, "a"), "hex", "828613968b501dc00a97e08c73b118aa8876c26b8aac93df128502ab360f91bab50a51e088769a5c1eff4782ace147dce3642554199876374291f5d921629502"
    test "sha512", String(128, "a"), "hex", "b73d1929aa615934e61a871596b3f3b33359f42b8175602e89f7e06e5f658a243667807ed300314b95cacdd579f3e33abdfbe351909519a846d465c59582f321"
    test "sha512", String(200, "a"), "hex", "4b11459c33f52a22ee8236782714c150a3b2c60994e9acee17fe68947a3e6789f31e7668394592da7bef827cddca88c4e6f86e4df7ed1ae6cba71f3e98faee9f"

    testUtf8
    testBytesInput
    testStreaming
    testHmac
    testErrors

    Debug.Print "-----------------------------------------"
    Debug.Print passCount & " passed, " & failCount & " failed"
End Sub

' Non-ASCII input must be hashed as UTF-8 to match openssl/Node.
' String = "a" + U+00E9 (e-acute) + U+4F60 (你) + U+1F30D (earth emoji),
' built with ChrW so the .bas file stays plain ASCII on disk.
Private Sub testUtf8()
    Dim s As String
    s = "a" & ChrW(&HE9) & ChrW(&H4F60) & ChrW(&HD83C) & ChrW(&HDF0D)
    test "md5", s, "hex", "43124c7cb8dcda6bb21d599fb21547a9"
    test "sha256", s, "hex", "73d8bad36f3b67ca0ecfb3159b4ed65aed03c7c89ddf73ef6ebcdf4f6556dc0f"
    test "sha512", s, "hex", "6db8e7967b418fcfd163e7cb0708d0db96d15ddff9e2492ac02f1fc0cbead9ed45e55150c0cb6158faee44ff022d0c1267cb664896486038b45fd15ac85ed7b6"
End Sub

' update() also accepts a raw Byte array; here bytes 0x00..0x09.
Private Sub testBytesInput()
    Dim c As CRYPTO: Set c = New CRYPTO
    Dim b(9) As Byte, i As Long
    For i = 0 To 9
        b(i) = CByte(i)
    Next i
    c.createHash "sha256"
    c.update b
    assertEqual "bytes(0..9) sha256", "1f825aa2f0020ef7cf91dfa30da4668d791c5d4824fc8e41354b89ec05795ab3", c.digest("hex")
End Sub

' Multiple update() calls accumulate; the result equals hashing the whole thing.
Private Sub testStreaming()
    Dim c As CRYPTO: Set c = New CRYPTO
    c.createHash "sha256"
    c.update "hello"
    c.update " "
    c.update "world"
    assertEqual "streaming sha256", "b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9", c.digest("hex")

    ' digest() is repeatable and does not consume the buffer.
    assertEqual "digest is repeatable", "b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9", c.digest("hex")
End Sub

' HMAC one-shot helper (String key + String message).
Private Sub testHmacStr(ByVal mode As String, ByVal key As String, ByVal msg As String, ByVal expected As String)
    Dim c As CRYPTO: Set c = New CRYPTO
    c.createHmac mode, key
    c.update msg
    assertEqual "hmac-" & mode & " / keylen=" & Len(key), expected, c.digest("hex")
End Sub

' HMAC known-answer vectors (RFC 2202 / 4231 style), all algorithms, plus keys
' that are empty, exactly the block size, and longer than the block (hashed down).
Private Sub testHmac()
    testHmacStr "sha256", "key", "The quick brown fox jumps over the lazy dog", "f7bc83f430538424b13298e6aa6fb143ef4d59a14946175997479dbc2d1a3cd8"
    testHmacStr "sha1", "key", "The quick brown fox jumps over the lazy dog", "de7c9b85b8b78aa6bc8a7a36f70a90701c9db4d9"
    testHmacStr "md5", "key", "The quick brown fox jumps over the lazy dog", "80070713463e7749b90c2dc24911e275"
    testHmacStr "sha256", "", "", "b613679a0814d9ec772f95d778c35fc5ff1697c493715653c6c712144292c5ad"
    testHmacStr "sha256", "secret", "", "f9e66e179b6747ae54108f82f8ade8b3c25d76fd30afde6c395822c530196169"
    testHmacStr "sha224", "key", "hello world", "818a1ce4c3a9ebc0506529a83234d9c77926b614892dbd62658444f4"
    testHmacStr "sha512", "key", "hello world", "ea0625a5ff1cd1653a327f8a4ae2f478fc51405c73ddac3a8a05a7a810310a6a14d7c8b4d284013493a6016ecadc772cfd98ed6cbe745949c5e6119fafb63b54"
    testHmacStr "sha384", "key", "hello world", "b7e365fa38bb22d6553614a63095564a0411866e65aac7b835d02d0b24245f4dc48696c9d970ac20f24105be7dc60133"
    testHmacStr "sha512-256", "key", "hello world", "0ba53ab0afef0dec4ff94412bfffa096ba69c40fb03319308b8fc23f53e3661d"
    testHmacStr "sha512-224", "key", "hello world", "a8e5289077249fdfe86fa129bafd135b3ff0c3cfb1ad587887f5e247"

    ' key longer than the block is hashed down first (64-byte and 128-byte blocks)
    testHmacStr "sha256", String(100, "a"), "msg", "1b110355f805afa1c9cbb6cf7065062139d2fb7b9eb28c7ae7581ea99cff6b8e"
    testHmacStr "sha512", String(200, "a"), "msg", "a31fbe02f4a9d748da17b18a8ba6dbce437bc2c1570eb220175203c2e30f3436f127b88c973b7308ab1295dab2a116c9b5d8e22fef562d58f87912afa8be9668"
    ' key exactly one block long
    testHmacStr "sha256", String(64, "k"), "exactly-block-key", "7c746358b2c7e9e4eff6c562edb13a9f65c24e51afe61607aac5da8d8f4a4dce"

    ' HMAC also supports a Byte-array key and streaming update()
    Dim c As CRYPTO: Set c = New CRYPTO
    Dim kb(9) As Byte, i As Long
    For i = 0 To 9
        kb(i) = CByte(i)
    Next i
    c.createHmac "sha256", kb
    c.update "data"
    assertEqual "hmac byte-key sha256", "cd8b852f159404c44469ac7b8dca5c830d96ef0bc1a8f917fece641ff696a840", c.digest("hex")

    Set c = New CRYPTO
    c.createHmac "sha256", "key"
    c.update "The quick brown fox "
    c.update "jumps over the lazy dog"
    assertEqual "hmac streaming sha256", "f7bc83f430538424b13298e6aa6fb143ef4d59a14946175997479dbc2d1a3cd8", c.digest("hex")
End Sub

' Unknown algorithm / output mode must raise instead of failing silently.
Private Sub testErrors()
    Dim c As CRYPTO, raised As Boolean

    raised = False
    On Error Resume Next
    Set c = New CRYPTO
    c.createHash "sha999"
    raised = (Err.Number <> 0)
    On Error GoTo 0
    assertEqual "unknown algorithm raises", "True", CStr(raised)

    raised = False
    Set c = New CRYPTO
    c.createHash "sha256"
    c.update "x"
    On Error Resume Next
    Dim d As String
    d = c.digest("no-such-format")
    raised = (Err.Number <> 0)
    On Error GoTo 0
    assertEqual "unknown output mode raises", "True", CStr(raised)
End Sub
