Private Sub test(mode, s, format, expected)
    Dim crypto As New CRYPTO
    crypto.createHash (mode)
    Debug.Print "crypto.createHash(""" & mode & """)"

    crypto.update (s)
    Debug.Print "crypto.update(""" & s & """)"

    Dim actual: actual = crypto.digest(format)
    Debug.Print "crypto.digest(""" & format & """)"
    Debug.Print "Expected: " & expected
    Debug.Print "Actual:   " & actual

    Debug.Assert expected = actual

    Set crypto = Nothing
End Sub
Private Sub testAll()
    Dim s: s = "hello world"
    test "md5", s, "hex", "5eb63bbbe01eeed093cb22bb8f5acdc3"
    test "md5", s, "base64", "XrY7u+Ae7tCTyyK7j1rNww=="
    test "sha1", s, "hex", "2aae6c35c94fcfb415dbe95f408b9ce91ee846ed"
    test "sha1", s, "base64", "Kq5sNclPz7QV2+lfQIuc6R7oRu0="
    test "sha256", s, "hex", "b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9"
    test "sha256", s, "base64", "uU0nuZNNPgilLlLX2n2r+sSE7+N6U4DukIj3rOLvzek="
    ' test "sha384", s, "hex", "fdbd8e75a67f29f701a4e040385e2e23986303ea10239211af907fcbb83578b3e417cb71ce646efd0819dd8c088de1bd"
    ' test "sha384", s, "base64", "/b2OdaZ/KfcBpOBAOF4uI5hjA+oQI5IRr5B/y7g1eLPkF8txzmRu/QgZ3YwIjeG9"
    ' test "sha512", s, "hex", "309ecc489c12d6eb4cc40f50c902f2b4d0ed77ee511a7c7a9bcd3ca86d4cd86f989dd35bc5ff499670da34255b45b0cfd830e81f605dcf7dc5542e93ae9cd76f"
    ' test "sha512", s, "base64", "MJ7MSJwS1utMxA9QyQLytNDtd+5RGnx6m808qG1M2G+YndNbxf9JlnDaNCVbRbDP2DDoH2Bdz33FVC6TrpzXbw=="
End Sub
