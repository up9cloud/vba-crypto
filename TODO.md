# TODO / Roadmap

## Done

- **PBKDF2** — `pbkdf2()` / `pbkdf2Bytes()`, using HMAC-<name> as the PRF.

## Out of scope (by decision)

These are intentionally **not** implemented in pure VBA. When they are genuinely
needed, calling a .NET/C# object via COM is the more appropriate solution than
reimplementing them here.

- **File hashing** (e.g. `updateFile(path)`) — if you really need to hash a
  file, call a .NET object directly; that is the correct tool for the job.
- **True incremental / low-memory streaming** — `update()` currently buffers the
  whole message in memory (O(n)). Processing arbitrarily large inputs block by
  block belongs in a .NET object rather than this library.
- **SHA-3 / Keccak** — uncommon in practice; not needed.
