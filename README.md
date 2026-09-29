# Initial Setup Needed
## Commands to run to clear variables

### Powershell

```
$env:REQUESTS_CA_BUNDLE=""
$env:SSL_CERT_FILE=""
$env:CURL_CA_BUNDLE=""
```

### Batch file in Windows

```batch file in windows
@set REQUESTS_CA_BUNDLE=
@set SSL_CERT_FILE=
@set CURL_CA_BUNDLE=
```

### If encounter an error like below it is usually related to one of these three variable above.

```ps
PS C:\Users\WaldoDa1\OneDrive - Republic Services\test> pip install pyinstaller
WARNING: Cache entry deserialization failed, entry ignored
WARNING: There was an error checking the latest version of pip.
Defaulting to user installation because normal site-packages is not writeable
WARNING: Cache entry deserialization failed, entry ignored
ERROR: Could not install packages due to an OSError: Could not find a suitable TLS CA certificate bundle, invalid path:
```
You can type echo %CURL_CA_BUNDLE% (any of the above environment variables) to ensure nothing is set to cause this error.  If the output comes back as empty or %CURL_CA_BUNDLE%, you should be fine.  Any other value, run the set commands above.


## Compilation: This python script was compiled (modules are case sensitive) using the command line below.

```bash
python -m PyInstaller --onefile rally_export.py
```

