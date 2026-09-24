Stale root copies of Python modules superseded by C:\GeoFooter\VBA\.
Archived 2026-09-04.

VBA launches the VBA\ copies:
  - MSCANSettings.ResolveSettingsDialogScript tries VBA\ first
  - StartAsyncGeolocationJob shells VBA\geolocate_headers.py
No Python module imports either file; aes_diagnostics_dialog.py accepts
geolocate_headers.py in VBA\ OR root, so it still passes.

Root copies at archive time:
  aes_settings_dialog.py  44372 bytes  2026-09-02 21:32  (pre-dates Sender Status tab)
  geolocate_headers.py   238130 bytes  2026-09-04 10:30  (pre-dates banner + sender status work)
