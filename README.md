# Project Zomboid Workshop Manager

A lightweight Python/Tkinter app for managing Project Zomboid Steam Workshop content. Paste a Workshop link and the app will extract and maintain:

- Dedicated Server Manager helper
- Workshop ID(s)
- Mod ID(s)
- Metadata (Name, PZ Build, Tags, Link)
- Maps vs Mods classification
- Collections with safe add/delete/refresh

All without external dependencies.

## Download and run on Windows (no install)

- If you're on Windows and just want to use the app, you only need the single EXE—no installation required.
- Download the file named `Fiears PZ WS to Server.exe` from this repository and run it:
	- Direct link (main branch):
		- https://github.com/FiearTheWolfdog/PZ-WS-to-Server/raw/main/Fiears%20PZ%20WS%20to%20Server.exe
	- Or check the Releases page if available.
- PLEASE use a dedicated folder for this EXE — it WILL create/update data files next to itself as you use the program.
- The EXE is portable. It will create/update its data files (Settings/Collections/WorkshopMeta/ModIDs/WorkshopIDs) next to the EXE.
- If Windows SmartScreen warns about an unrecognized app:
	- Click "More info" → "Run anyway".

### Update to a new version (Windows EXE)

- Close the app if it’s running.
- Download the new `Fiears PZ WS to Server.exe` and replace your existing EXE file in the same folder.
- Your data files (Settings/Collections/WorkshopMeta/ModIDs/WorkshopIDs) stay in that folder and will be reused automatically.

## Highlights

- One-line files for easy server configs:
	- `WorkshopIDs.txt` and `ModIDs.txt` are kept as a single semicolon-separated line each.
- GUI with dark mode and persisted preferences.
- Paste + Add, duplicate detection, and safe removal.
- Double-click to open items in your browser; quick “Copy Selected Link”.
- Auto-add required mods (skips those lacking Mod IDs).
- Tags parsing with a whitelist filter.
- Map detection (folder parsing + tag fallback) with a dedicated Maps tab.
- Sorting by columns, header right-click sort menu, search filter, and “Sort by Order Added”.
- Drag-select multi-selection and multi-delete for Mods/Maps.
- Collections support:
	- Detect collection URLs, add all children with duplicate skip
	- Collections tab with Delete Selected Collections
	- Refresh Selected Collections syncs additions/removals safely
	- Robust detection: pages with “Required items/mods” are treated as standalone items (not collections)
	- Handles children that expose multiple Mod IDs by prompting for your selection during import and refresh
- About/Info dialog loaded from file, with clickable links.

## Requirements

- Windows users: none if using the EXE (no Python needed)
- Running from source: Python 3.10+ (works with 3.13), no third‑party packages required

## Run from source (optional)

If you prefer running the script directly (or you're on a platform other than Windows):

From this folder in PowerShell:

```powershell
python ".\pz_mod_scraper.py" --gui
```

Tip: There are VS Code tasks included in this workspace you can use to run the GUI.

## Usage

1. Paste a Steam Workshop URL into the input field.
2. Click Add. The app detects whether it’s a mod/map or a collection and acts accordingly.
3. For mods with multiple Mod IDs, you’ll be prompted to choose which to add.
4. For collections, the app imports all children (skipping duplicates). If a child exposes multiple Mod IDs, you’ll choose per-child; the import then continues.
5. Use the tabs:
	 - Mods: regular mods
	 - Maps: items recognized as maps
	 - Collections: manage your collections, delete them, or refresh them
6. Use search and column headers (or right-click header) to sort/filter.
7. Double-click any row to open its Steam page. Use “Copy Selected Link” to copy the URL.
8. Use “Remove Selected” to delete mods/maps from your lists and files.

## Data files

- `WorkshopIDs.txt` — single line, `;`-separated Workshop IDs
- `ModIDs.txt` — single line, `;`-separated Mod IDs
- `WorkshopMeta.json` — metadata cache (names, builds, tags, links, map flags, etc.)
- `Collections.json` — stores collections with:
	- `title`, `url`
	- `items`: known child Workshop IDs
	- `added`: child IDs that this collection actually inserted into your lists (used for safe delete/refresh)
- `Settings.json` — persisted UI settings (dark mode, etc.)
- `AboutInfo.txt` — contents for the Info dialog (supports clickable links)

All files live in this folder alongside the script.

## Notes on detection

- Build parsing: prefers the most recent tag; falls back to description or title when needed.
- Maps are detected via map folder parsing; falls back to presence of a “Map” tag.
- Collection detection uses multiple page markers (Subscribe to all, breadcrumb, ITEMS(N), etc.). Pages that show “Required items/mods” are treated as standalone items.

## Known limitations

- Very large collections that require “View all” pagination aren’t fetched across pages yet.
- If saving JSON files fails (permissions/locks), you’ll see a message—resolve and retry.

## Troubleshooting

- If the GUI doesn’t start, confirm Python is installed and available in PATH. Try:
	- `python --version`
- If links don’t open on click, ensure your default browser is configured: Windows Settings → Apps → Default apps.
- To reset, close the app and delete `WorkshopMeta.json`, `Collections.json`, `WorkshopIDs.txt`, and `ModIDs.txt` (optional). Re-open to rebuild.

## Releases

- Prefer grabbing downloads from the Releases page when available:
	- https://github.com/FiearTheWolfdog/PZ-WS-to-Server/releases
- Download the `Fiears PZ WS to Server.exe` asset. It’s a single portable file—no installer.
- To update: close the app, replace the EXE in your folder, and re-open. Your data files in that folder will be reused.

## Changelog

- 2025-10-06: Info dialog links are now clickable. Collections import/refresh prompt for Mod ID selection when children have multiple IDs.

## License

Personal use. No warranty. This project is not affiliated with The Indie Stone or Valve.
