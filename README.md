# vba-crypto

crypto library for VBA.

## Usage

- Download `./src/CRYPTO.cls`
- Open your excel file, go to the macro editor `Microsoft Visual Basic for Applications` (ALT + F11)
- Go to: File > Import File, to choose the `CRYPTO.cls` file from your download folder.

```vb
Function hash(s)
  Dim o As New CRYPTO
  o.createHash "sha256"   ' md5, sha1, sha224, sha256, sha384, sha512, sha512-256, sha512-224
  o.update s
  hash = o.digest("hex")  ' hex, base64
  Set o = Nothing
End Function

Debug.Print hash("hello world")
' b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9
```

### HMAC

Use `createHmac(name, key)` instead of `createHash`; everything else is the same. The key may be a `String` (UTF-8) or a `Byte` array.

```vb
Dim o As New CRYPTO
o.createHmac "sha256", "my-secret-key"
o.update "message"
Debug.Print o.digest("hex")
' 1423df78108155d1bef5dac2bb5d8eae873eb35b417c0184d6a33416b669136e
```

### PBKDF2

`pbkdf2(password, salt, iterations, keyLen, [name], [outMode])` derives a key from a password (RFC 2898, using HMAC-`name` as the PRF). `password`/`salt` may be `String` (UTF-8) or `Byte` arrays.

```vb
Dim o As New CRYPTO
Debug.Print o.pbkdf2("password", "salt", 100000, 32, "sha256")   ' 64 hex chars
```

`pbkdf2Bytes(...)` returns the derived key as a raw `Byte()`. (Both configure the object's algorithm, so they reset any in-progress `update()`.)

### API

- `createHash(name)` — pick the algorithm (see the list above). Also resets the object so it can be reused.
- `createHmac(name, key)` — start an HMAC with the given algorithm and key (String or Byte array).
- `pbkdf2(password, salt, iterations, keyLen, [name], [outMode])` / `pbkdf2Bytes(...)` — PBKDF2 key derivation, as a string or raw `Byte()`.
- `update(data)` — feed data. `data` may be a `String` (encoded as **UTF-8**, so non-ASCII matches openssl/Node) or a `Byte` array. Call it multiple times to hash data in pieces:

  ```vb
  Dim o As New CRYPTO
  o.createHash "sha256"
  o.update "hello "
  o.update "world"       ' same result as hashing "hello world" at once
  Debug.Print o.digest("hex")
  ```

- `digest(mode)` — finalize as a `"hex"`, `"base64"` or `"base64url"` string. It does not consume the buffer, so it can be called more than once.
- `digestBytes()` — finalize as a raw `Byte()` array.
- `verify(expected, [mode])` — timing-safe comparison of the digest against `expected` (avoids leaking, via early-exit timing, how many leading characters matched — the correct way to check a MAC). `mode` is the encoding of `expected` (`"hex"` default, `"base64"`/`"base64url"` accepted); hex compares case-insensitively.

  ```vb
  Dim o As New CRYPTO
  o.createHmac "sha256", key
  o.update message
  If o.verify(expectedMac) Then ' ... trusted
  ```

Unknown algorithm or output-mode names raise a run-time error rather than returning an empty string.

### Testing

`src/CRYPTO.test.bas` is a self-contained test suite (known-answer vectors, padding-boundary lengths, UTF-8, byte input, streaming, and error cases).

1. Import both `CRYPTO.cls` and `CRYPTO.test.bas` into the VBA project.
2. Open the Immediate window with `Ctrl+G`, type `testAll`, and press Enter.
3. The last line prints e.g. `83 passed, 0 failed`; any failure prints the expected and actual digests above the summary.

### Limitations

The message length is tracked in bytes as a 32-bit value, so inputs larger than ~256 MB are not supported (well beyond typical spreadsheet use).

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
- Claude: https://claude.ai
