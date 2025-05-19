# vba-crypto

crypto library for VBA.

## Usage

- Download `./src/CRYPTO.cls`
- Open your excel file, go to the macro editor `Microsoft Visual Basic for Applications` (ALT + F11)
- Go to: File > Import File, to choose the `CRYPTO.cls` file from your download folder.

```vb
Function hash(s)
  Dim o As New CRYPTO
  o.createHash("sha256") ' md5, sha256
  o.update(s)
  hash = o.digest("hex") ' hex, base64
  Set o = Nothing
End Function

Debug.Print hash("hello world")
' b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9'
```

The above works the same way as the following does with .NET objects:

```vb
Private Function stringToUTFBytes(aString)
    Dim UTF8: Set UTF8 = CreateObject("System.Text.UTF8Encoding")
    stringToUTFBytes = UTF8.GetBytes_4(aString)
    Set UTF8 = Nothing
End Function
Private Function sha256hashBytes(aBytes)
    Dim SHA256: Set SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    SHA256.Initialize
    sha256hashBytes = SHA256.ComputeHash_2((aBytes))
    Set SHA256 = Nothing
End Function
Private Function bytesToHex(aBytes)
    Dim hexStr, x
    For x = 1 To LenB(aBytes)
        hexStr = Hex(AscB(MidB((aBytes), x, 1)))
        If Len(hexStr) = 1 Then
          hexStr = "0" & hexStr
        End If
        bytesToHex = bytesToHex & hexStr
    Next
End Function
Function hash(s)
  hash = bytesToHex(sha256hashBytes(stringToUTFBytes(s)))
End Function

Debug.Print hash("hello world")
' b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9
```

## Credits

- chrisfeas: https://www.mrexcel.com/board/threads/sha256-hash-function.468698/page-2#post-5110793
- Phil Fresle: http://web.archive.org/web/20070808024050if_/http://www.frez.co.uk/SHA.zip
- GustavBrock: https://github.com/GustavBrock/VBA.Cryptography
- Nodejs crypto: https://nodejs.org/api/crypto.html
- sha***: https://github.com/mozilla/pdf.js/blob/master/src/core/calculate_sha_other.js
- md5: https://stackoverflow.com/questions/125785/password-hash-function-for-excel-vba/125844#125844
- ChatGPT: https://chatgpt.com/
