# Aliniant AES Ribbon Host

Classic Outlook **COM add-in** that hosts Ribbon XML on **Home (Mail)** so AES can use:

- Large icons for **AES ON/OFF**, **Short Scan**, **Full Scan**, **Deep Scan**, **GURI**, **Aura**
- Stacked **Diagnostics / Settings / View Logs**
- **Live** ON/OFF label + green tick / red cross icons

Actions call your existing VBA AES CommandBar buttons (`Execute`), so scan/settings/diagnostics logic stays in VBA.
If the button is missing, the host falls back to `Application.Run` for the macro.
**GURI** additionally launches `<install root>\guri\gui.py --raise` directly when VBA has no GURI button yet.
**Aura** launches `guri\gui.py --raise --aura` (opens the Aura / data-broker removal tab).

## Install

1. Close Outlook.
2. Elevated PowerShell:

```powershell
cd C:\GeoFooter\AesRibbonHost
.\install.ps1
```

3. Start Outlook.
4. Confirm: **File → Options → Add-ins → COM Add-ins** → **Aliniant AES Ribbon** checked.
5. On **Home**, you should see the **AES** group (separate from any manual “AES (Custom)” macro group — remove the custom one to avoid duplicates).

## Uninstall

```powershell
cd C:\GeoFooter\AesRibbonHost
.\uninstall.ps1
```

## Notes

- Requires **classic** Outlook with COM add-ins (not the Store “new Outlook” web client).
- Must match Outlook bitness. This build targets **x86** (32-bit Outlook). If Outlook is 64-bit, change `PlatformTarget` to `x64` and re-run `install.ps1`.
- VBA should keep creating the **AES** CommandBar (Add-ins tab); the host uses it as the action bridge.
- Service state file: `%LOCALAPPDATA%\GeoFooter\aes_service_state.json` (written by VBA on toggle).
- If COM Add-ins shows “runtime error during loading”, re-run `install.ps1` after a full Outlook quit (including tray icon).
- **Add-in log:** `%LOCALAPPDATA%\GeoFooter\Logs\AesRibbonHost.log`  
  (created when the add-in actually loads into Outlook)
