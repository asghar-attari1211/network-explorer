Quick map launcher (for non-technical users)

- Purpose: generate the interactive site map and open it in your browser with one command.

How to run:

1. Open a terminal in the project folder `/workspaces/RnD`.
2. Run:

```bash
bash run_map.sh
```

Notes:
- The script runs `Route.py` which reads VLAN/LR/OFN/etc. from the Excel files (per the existing configuration) and writes an HTML map to `/workspaces/RnD/Site_Map_<date>.html`.
- If your environment has a GUI, the script will attempt to open the generated HTML with your system default browser via `xdg-open`.
- If you want the script double-clickable in a file manager, make `run_map.sh` executable (`chmod +x run_map.sh`) and double-click to run.
Important: The interactive map fetches its data (JSON assets) via HTTP. The launcher starts a local HTTP server automatically so the browser can load the JSON.

If you'd like, I can make `run_map.sh` executable and run a quick test now.