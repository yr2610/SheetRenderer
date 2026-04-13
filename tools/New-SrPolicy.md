# `New-SrPolicy.ps1` Usage

This script converts a plain-text allow list into `sr_policy.dat`.

## Input format

Prepare a UTF-8 text file such as `allowed-users.txt`.

One user per line:

```txt
# user[,expireDate]
DOMAIN\UserA
DOMAIN\UserB,2026-12-31
MACHINE\UserC
```

- Blank lines are ignored.
- Lines starting with `#` are ignored.
- `expireDate` is optional.
- Date format must be `yyyy-MM-dd`.

## Basic usage

Run from the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\New-SrPolicy.ps1 -InputPath .\allowed-users.txt
```

This writes `sr_policy.dat` to:

```text
.\SheetRenderer\sr_policy.dat
```

## Custom output path

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\New-SrPolicy.ps1 -InputPath .\allowed-users.txt -OutputPath C:\path\to\sr_policy.dat
```

## Notes

- `sr_policy.dat` is gzip-compressed and base64-encoded text.
- It is meant to provide lightweight obscurity, not strong security.
- The add-in reads `sr_policy.dat` from its own output directory.
