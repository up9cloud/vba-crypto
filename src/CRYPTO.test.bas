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
    test "sha256", s, "hex", "b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9"
    test "sha256", s, "base64", "uU0nuZNNPgilLlLX2n2r+sSE7+N6U4DukIj3rOLvzek="
End Sub
