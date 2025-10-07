#!/usr/bin/env python3
"""
Project Zomboid Workshop/Mod ID scraper with GUI

Two modes:
    - GUI (default): paste a Steam Workshop URL and click ADD to build one-line,
        semicolon-separated lists for Mod IDs and Workshop IDs.
    - CLI: pass URLs as arguments (or use --cli to enter interactive mode).

Output files in this folder:
    - WorkshopIDs.txt — ONE LINE ONLY, semicolon-separated list of workshop item IDs
    - ModIDs.txt      — ONE LINE ONLY, semicolon-separated list of mod IDs

Notes:
    - Mod ID(s) are parsed from the Workshop page description lines like "Mod ID:" or
        "Mod IDs:". If not present, Mod IDs cannot be inferred without downloading the mod.
    - Multiple Mod IDs on a single line are supported (comma/semicolon/space separated).
    - Entries are de-duplicated case-insensitively; display preserves insertion order.
"""

from __future__ import annotations

import html
import os
import re
import sys
import urllib.parse
import urllib.request
from typing import Iterable, List, Optional, Set, Tuple, Dict, Any
import json
import threading
import webbrowser

# GUI imports (standard library)
import tkinter as tk
from tkinter import ttk, messagebox


# App directory (next to the EXE when frozen, next to the .py when running from source)
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

def _path_in_app_dir(name: str) -> str:
    return os.path.join(APP_DIR, name)

# File paths (kept next to the EXE/script so users can edit them)
WORKSHOP_FILE = _path_in_app_dir("WorkshopIDs.txt")
MODS_FILE = _path_in_app_dir("ModIDs.txt")
META_FILE = _path_in_app_dir("WorkshopMeta.json")
ABOUT_FILE = _path_in_app_dir("AboutInfo.txt")
SETTINGS_FILE = _path_in_app_dir("Settings.json")
COLLECTIONS_FILE = _path_in_app_dir("Collections.json")

# Access bundled resources when frozen (PyInstaller provides sys._MEIPASS)
def _resource_path(name: str) -> str:
    base = getattr(sys, '_MEIPASS', None)
    if base:
        return os.path.join(base, name)
    # Fallback to app dir when running from source
    return os.path.join(APP_DIR, name)

# Default About text used to bootstrap a writable AboutInfo.txt on first run
DEFAULT_ABOUT_TEXT = (
    "Project Zomboid Workshop Manager\n\n"
    "Paste a Steam Workshop URL and click ADD.\n\n"
    "Useful links:\n"
    "- Steam Workshop: https://steamcommunity.com/app/108600/workshop/\n"
    "- Project Zomboid: https://projectzomboid.com/\n"
)

# Allowed tags (canonical forms)
ALLOWED_TAGS: List[str] = [
    "Build 40","Build 41","Build 42","Animals","Audio","Balance","Building","Clothing/Armor",
    "Farming","Food","Framework","Hardmode","Interface","Items","Language/Translation","Literature",
    "Map","Military","Misc","Models","Multiplayer","Pop Culture","QoL","Realistic","Silly/Fun",
    "Skills","Textures","Traits","Vehicles","Weapons","WIP",
]
ALLOWED_TAGS_MAP = {t.lower(): t for t in ALLOWED_TAGS}

# In-memory stores (preserve insertion order)
workshop_ids: List[str] = []
mod_ids: List[str] = []
workshop_meta: Dict[str, Dict[str, Any]] = {}
collections_meta: Dict[str, Dict[str, Any]] = {}


class FetchError(Exception):
    pass


def fetch_html(url: str, timeout: int = 20) -> str:
    """Fetch raw HTML for a given URL with a browser-like User-Agent."""
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            charset = resp.headers.get_content_charset() or "utf-8"
            return resp.read().decode(charset, errors="replace")
    except Exception as e:
        raise FetchError(f"Failed to fetch page: {e}") from e


def parse_workshop_id(url: str) -> Optional[str]:
    """Extract the numeric workshop ID from the URL query (id=...)."""
    try:
        parsed = urllib.parse.urlparse(url)
        qs = urllib.parse.parse_qs(parsed.query)
        if "id" in qs and len(qs["id"]) > 0:
            candidate = qs["id"][0]
            if candidate.isdigit():
                return candidate
    except Exception:
        pass

    # Fallback regex
    m = re.search(r"[?&]id=(\d+)", url)
    if m:
        return m.group(1)
    return None


def parse_collection_id(url: str) -> Optional[str]:
    """Extract a collection ID from URL if present (same id= schema)."""
    return parse_workshop_id(url)


def parse_mod_ids_from_html(html_text: str) -> List[str]:
    """Find Mod ID(s) in the page description.

    Heuristics:
      - Look for lines like: "Mod ID: Something" or "Mod IDs: a, b"
      - Accept variations: "ModID", case-insensitive, with ':' or '-'
      - Split multiple IDs by comma/semicolon/whitespace
    """
    # Reduce whitespace to simplify regex matching
    normalized = html.unescape(html_text)

    # The description is often within <div class="workshopItemDescription"> ...
    # but we don't require it; we search the whole doc for resilience.
    patterns = [
        r"(?im)\bMod\s*ID\s*[:\-]\s*([^\r\n<]+)",
        r"(?im)\bMod\s*IDs\s*[:\-]\s*([^\r\n<]+)",
        r"(?im)\bModID\s*[:\-]\s*([^\r\n<]+)",
    ]

    found: List[str] = []
    for pat in patterns:
        for match in re.finditer(pat, normalized):
            raw_line = match.group(1).strip()
            # Remove any HTML tags that slipped in
            raw_line = re.sub(r"<[^>]+>", " ", raw_line)
            # Trim after certain keywords to avoid pulling extra content
            raw_line = re.split(r"\b(Workshop\s*ID|Required|Map|IDs?)\b", raw_line, maxsplit=1)[0].strip()

            # Split into IDs by common separators
            parts = re.split(r"[,;\|/\s]+", raw_line)
            for p in parts:
                pid = p.strip()
                # Basic validation: PZ Mod IDs are usually alnum + _ -
                if pid and re.match(r"^[A-Za-z0-9_\-]+$", pid):
                    found.append(pid)

    # Preserve order but dedupe
    unique: List[str] = []
    seen: Set[str] = set()
    for mid in found:
        if mid.lower() not in seen:
            unique.append(mid)
            seen.add(mid.lower())
    return unique


def _extract_between(text: str, start_pat: str, end_pat: str) -> Optional[str]:
    s = re.search(start_pat, text, flags=re.IGNORECASE | re.DOTALL)
    if not s:
        return None
    start_idx = s.end()
    e = re.search(end_pat, text[start_idx:], flags=re.IGNORECASE | re.DOTALL)
    if not e:
        return None
    return text[start_idx:start_idx + e.start()]


def parse_title_from_html(html_text: str) -> Optional[str]:
    m = re.search(r"<div[^>]*class=\"workshopItemTitle\"[^>]*>(.*?)</div>", html_text, flags=re.DOTALL | re.IGNORECASE)
    if m:
        title = re.sub(r"<[^>]+>", " ", m.group(1)).strip()
        return html.unescape(re.sub(r"\s+", " ", title)) or None
    m = re.search(r"<title>(.*?)</title>", html_text, flags=re.DOTALL | re.IGNORECASE)
    if m:
        title = re.sub(r"\s*::\s*Steam Community\s*$", "", m.group(1).strip())
        return html.unescape(re.sub(r"\s+", " ", title)) or None
    return None


def parse_pz_version_from_html(html_text: str) -> Optional[str]:
    """Heuristically extract the PZ build/version from the description."""
    text = html.unescape(html_text)
    desc = _extract_between(text, r"<div[^>]*class=\"workshopItemDescription\"[^>]*>", r"</div>") or text
    desc_text = re.sub(r"<[^>]+>", " ", desc)
    desc_text = re.sub(r"\s+", " ", desc_text)
    patterns = [
        r"Build\s*(4[12](?:\.\d+){0,2})",
        r"\b(4[12](?:\.\d+){1,2})\b",
        r"\[(4[12](?:\.\d+){1,2})\]",
    ]
    for pat in patterns:
        m = re.search(pat, desc_text, flags=re.IGNORECASE)
        if m:
            return m.group(1)
    m = re.search(r"Build\s*(\d+(?:\.\d+)*)", desc_text, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    return None


def parse_build_from_title(title: Optional[str]) -> Optional[str]:
    if not title:
        return None
    t = title
    # Prefer bracketed hints like [B41/42], [41.78]
    m = re.search(r"\[(?:Build\s*)?(B?\s*\d+(?:[./]\d+)*)\]", t, flags=re.IGNORECASE)
    if m:
        return re.sub(r"\s+", "", m.group(1))
    # Or standalone B41, B42 tokens
    m = re.search(r"\bB\s*(\d+(?:[./]\d+)*)\b", t, flags=re.IGNORECASE)
    if m:
        return re.sub(r"\s+", "", m.group(1))
    # Or explicit Build 41.x in title
    m = re.search(r"Build\s*(\d+(?:\.\d+)*)", t, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    return None


def _extract_build_from_tags(tags: List[str]) -> Optional[str]:
    """Return the highest build number found in tags like 'Build 41', 'Build 42'.
    Returns just the numeric portion (e.g., '42' or '41.78').
    """
    if not tags:
        return None
    numbers: List[Tuple[Tuple[int, ...], str]] = []
    for t in tags:
        m = re.search(r"Build\s*(\d+(?:\.\d+)*)", t, flags=re.IGNORECASE)
        if m:
            s = m.group(1)
            parts = tuple(int(p) for p in s.split('.'))
            numbers.append((parts, s))
    if not numbers:
        return None
    # Pick the maximum by numeric tuple
    numbers.sort()
    return numbers[-1][1]


def read_existing_items(path: str) -> List[str]:
    """Read IDs from a file that may be one-per-line or semicolon-separated (one line)."""
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read().strip()
    if not content:
        return []
    # Support both legacy newline-separated and new semicolon-separated one-line
    if "\n" in content:
        parts = [p.strip() for p in content.splitlines() if p.strip()]
    else:
        parts = [p.strip() for p in content.split(";") if p.strip()]
    # Dedupe preserving order (case-insensitive)
    seen: Set[str] = set()
    result: List[str] = []
    for p in parts:
        key = p.lower()
        if key not in seen:
            result.append(p)
            seen.add(key)
    return result


def write_one_line_semicolon(path: str, items: Iterable[str]) -> None:
    """Overwrite file with ONE LINE of semicolon-separated items."""
    line = ";".join(items)
    with open(path, "w", encoding="utf-8") as f:
        f.write(line + ("\n" if line else ""))


def process_url(url: str) -> Tuple[Optional[str], List[str]]:
    wsid = parse_workshop_id(url)
    try:
        html_text = fetch_html(url)
    except FetchError as e:
        print(f"! Could not fetch page: {e}")
        return wsid, []

    mods = parse_mod_ids_from_html(html_text)
    return wsid, mods


def try_fetch_collection(url: str) -> Optional[Dict[str, Any]]:
    """If URL is a collection, return {id, title, url, items:[wsids]} else None."""
    cid = parse_collection_id(url)
    if not cid:
        return None
    try:
        html_text = fetch_html(url)
    except FetchError:
        return None
    child_ids = parse_collection_children_wsids(html_text, parent_wsid=cid)
    if not child_ids:
        return None
    title = parse_title_from_html(html_text) or f"Collection {cid}"
    return {"id": cid, "title": title, "url": url, "items": child_ids}


def get_meta_for_workshop_id(wsid: str) -> Dict[str, str]:
    """Fetch and parse metadata for a workshop ID; returns dict with title, version, url."""
    url = f"https://steamcommunity.com/sharedfiles/filedetails/?id={wsid}"
    try:
        html_text = fetch_html(url)
    except FetchError:
        return {"title": "(unknown)", "version": "(unknown)", "url": url}
    title = parse_title_from_html(html_text) or "(unknown)"
    # Also parse mod IDs and required workshop IDs (for auto-adding dependencies)
    mods = parse_mod_ids_from_html(html_text)
    requires = []
    try:
        requires = parse_required_wsids_from_html(html_text)
    except Exception:
        pass
    # Parse tags and filter to allowed set
    tags = parse_tags_from_html(html_text)
    # Prefer build from tags (most recent). If missing, fall back to description.
    version = _extract_build_from_tags(tags)
    if not version:
        version = parse_pz_version_from_html(html_text)
    version = version or "(unknown)"
    # Parse map folders to identify map mods
    map_folders = parse_map_folders_from_html(html_text)
    # Fallback: if tag 'Map' is present, treat as map even if map_folders is empty
    is_map = bool(map_folders) or ("Map" in tags)
    return {"title": title, "version": version, "url": url, "mods": mods, "requires": requires, "tags": tags, "map_folders": map_folders, "is_map": is_map}


def parse_required_wsids_from_html(html_text: str, hint_section: bool = True) -> List[str]:
    """Extract required workshop IDs from the Workshop page HTML.

    Strategy:
      - Prefer links under a section that mentions "Required" (e.g., "Required items").
      - Fallback to scanning whole page for sharedfiles/filedetails links.
    """
    ids: List[str] = []
    def extract_ids(text: str) -> List[str]:
        return re.findall(r"sharedfiles/fil[e]?details/\?id=(\d+)", text, flags=re.IGNORECASE)

    text = html.unescape(html_text)
    if hint_section:
        # Look for a window around the phrase "Required" to limit noise
        for m in re.finditer(r"Required\s+(items|mods?)", text, flags=re.IGNORECASE):
            start = max(0, m.start() - 200)
            chunk = text[start:start + 4000]  # scan ahead ~4KB
            ids.extend(extract_ids(chunk))
    if not ids:
        ids = extract_ids(text)
    # Dedupe preserving order
    seen: Set[str] = set()
    out: List[str] = []
    for i in ids:
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out


def get_required_wsids(wsid: str) -> List[str]:
    url = f"https://steamcommunity.com/sharedfiles/filedetails/?id={wsid}"
    try:
        html_text = fetch_html(url)
    except FetchError:
        return []
    return parse_required_wsids_from_html(html_text)


def parse_tags_from_html(html_text: str) -> List[str]:
    text = html.unescape(html_text)
    # Tags often appear in clickable spans/anchors; grab text near a 'tags' container
    # Broad fallback: collect all anchor/span inner texts
    raw = re.findall(r">\s*([^<>\n\r]+?)\s*<", text)
    candidates = [re.sub(r"\s+", " ", t.strip()) for t in raw if t and len(t.strip()) <= 40]
    # Normalize and filter by allowed tags (case-insensitive), preserve order
    seen: Set[str] = set()
    out: List[str] = []
    for c in candidates:
        key = c.lower()
        if key in ALLOWED_TAGS_MAP and key not in seen:
            out.append(ALLOWED_TAGS_MAP[key])
            seen.add(key)
    return out


def parse_map_folders_from_html(html_text: str) -> List[str]:
    text = html.unescape(html_text)
    # Search the description block if available
    desc = _extract_between(text, r"<div[^>]*class=\"workshopItemDescription\"[^>]*>", r"</div>") or text
    # Find lines like "Map Folder: Something"
    lines = re.findall(r"(?im)Map\s*Folder\s*:\s*([^\r\n<]+)", desc)
    folders: List[str] = []
    for l in lines:
        cleaned = re.sub(r"<[^>]+>", " ", l)
        cleaned = cleaned.strip()
        # Some entries may contain comma-separated multiple folder names on one line
        parts = [p.strip() for p in re.split(r",", cleaned) if p.strip()]
        folders.extend(parts if parts else [cleaned])
    # Deduplicate while preserving order
    seen: Set[str] = set()
    out: List[str] = []
    for f in folders:
        k = f.lower()
        if k and k not in seen:
            seen.add(k)
            out.append(f)
    return out


def ensure_in_list(store: List[str], items: Iterable[str]) -> int:
    """Insert items into list if not present (case-insensitive); returns count added."""
    added = 0
    existing_lower = {s.lower() for s in store}
    for it in items:
        if it.lower() not in existing_lower:
            store.append(it)
            existing_lower.add(it.lower())
            added += 1
    return added


def load_existing_to_memory() -> None:
    global workshop_ids, mod_ids, workshop_meta
    global collections_meta
    workshop_ids = read_existing_items(WORKSHOP_FILE)
    mod_ids = read_existing_items(MODS_FILE)
    if os.path.exists(META_FILE):
        try:
            with open(META_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                normalized: Dict[str, Dict[str, Any]] = {}
                for k, v in data.items():
                    if isinstance(v, dict):
                        tags_l = v.get("tags") if isinstance(v.get("tags"), list) else []
                        map_folders_l = v.get("map_folders") if isinstance(v.get("map_folders"), list) else []
                        requires_l = v.get("requires") if isinstance(v.get("requires"), list) else []
                        # Derive is_map if missing, using fallback: tag 'Map' implies map
                        if v.get("is_map") is None:
                            is_map_val = bool(map_folders_l) or ("Map" in tags_l)
                        else:
                            is_map_val = bool(v.get("is_map"))
                        normalized[str(k)] = {
                            "title": v.get("title") or "(unknown)",
                            "version": v.get("version") or "(unknown)",
                            "url": v.get("url") or f"https://steamcommunity.com/sharedfiles/filedetails/?id={k}",
                            "mods": v.get("mods") if isinstance(v.get("mods"), list) else [],
                            "tags": tags_l,
                            # new optional fields preserved if present
                            "requires": requires_l,
                            "map_folders": map_folders_l,
                            "is_map": is_map_val,
                        }
                workshop_meta = normalized
        except Exception:
            workshop_meta = {}
    # Load collections metadata
    collections_meta = {}
    if os.path.exists(COLLECTIONS_FILE):
        try:
            with open(COLLECTIONS_FILE, "r", encoding="utf-8") as f:
                cdata = json.load(f)
            if isinstance(cdata, dict):
                for cid, v in cdata.items():
                    if isinstance(v, dict):
                        collections_meta[str(cid)] = {
                            "title": v.get("title") or f"Collection {cid}",
                            "url": v.get("url") or f"https://steamcommunity.com/sharedfiles/filedetails/?id={cid}",
                            "items": v.get("items") if isinstance(v.get("items"), list) else [],
                            "added": v.get("added") if isinstance(v.get("added"), list) else [],
                        }
        except Exception:
            collections_meta = {}


def save_memory_to_files() -> None:
    write_one_line_semicolon(WORKSHOP_FILE, workshop_ids)
    write_one_line_semicolon(MODS_FILE, mod_ids)
    # Save workshop meta
    try:
        with open(META_FILE, "w", encoding="utf-8") as f:
            json.dump(workshop_meta, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    # Save collections meta (separate try so a failure above doesn't skip this)
    try:
        with open(COLLECTIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(collections_meta, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def parse_collection_children_wsids(html_text: str, parent_wsid: Optional[str] = None) -> List[str]:
    """Extract child workshop item IDs from a Steam Workshop Collection page.
    Best-effort heuristic based on presence of collection containers and links.
    """
    text = html.unescape(html_text)
    # Quick detection of collection containers (avoid weak fallbacks)
    is_collection = bool(
        re.search(r"workshopCollection|collectionChildren|collectionItems|workshopItemCollection|collectionHeader", text, flags=re.IGNORECASE)
        or re.search(r"Subscribe\s+to\s+all|Unsubscribe\s+from\s+all|Save\s+to\s+Collection", text, flags=re.IGNORECASE)
        or re.search(r"ITEMS\s*\(\d+\)", text, flags=re.IGNORECASE)
        or re.search(r"section=collections", text, flags=re.IGNORECASE)
    )
    ids: List[str] = []
    # Link-based extraction
    ids.extend(re.findall(r"sharedfiles/fil[e]?details/\?id=(\d+)", text, flags=re.IGNORECASE))
    # Attribute-based extraction (common on collection grids)
    ids.extend(re.findall(r"data-publishedfileid=\"(\d+)\"", text, flags=re.IGNORECASE))
    # JSON-embedded extraction
    ids.extend(re.findall(r"publishedfileid\"?\s*[:=]\s*\"?(\d+)\"?", text, flags=re.IGNORECASE))
    # Filter out parent id and dedupe
    out: List[str] = []
    seen: Set[str] = set([str(parent_wsid)] if parent_wsid else [])
    for i in ids:
        if i not in seen:
            seen.add(i)
            out.append(i)
    # If the page contains a 'Required items/mods' section, do not treat as collection
    if re.search(r"Required\s+(items|mods?)", text, flags=re.IGNORECASE):
        return []
    # If no collection hint or no children, treat as not a collection
    if not is_collection or len(out) == 0:
        return []
    return out


def build_gui():
    load_existing_to_memory()

    root = tk.Tk()
    root.title("PZ Workshop/Mod ID Builder")
    root.geometry("900x520")
    root.resizable(True, True)

    # Menubar (File, View, About)
    menubar = tk.Menu(root)
    root.config(menu=menubar)
    file_menu = tk.Menu(menubar, tearoff=False)
    view_menu = tk.Menu(menubar, tearoff=False)
    about_menu = tk.Menu(menubar, tearoff=False)
    menubar.add_cascade(label="File", menu=file_menu)
    menubar.add_cascade(label="View", menu=view_menu)
    menubar.add_cascade(label="About", menu=about_menu)

    # Theme / Dark mode
    style = ttk.Style(root)
    original_theme = style.theme_use()
    dark_mode = tk.BooleanVar(value=False)

    # Settings helpers for persisting preferences (e.g., dark mode)
    def read_settings() -> Dict[str, Any]:
        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                return data if isinstance(data, dict) else {}
        except Exception:
            return {}
        return {}

    def write_settings(updates: Dict[str, Any]) -> None:
        try:
            data = read_settings()
            data.update(updates)
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def apply_theme(is_dark: bool):
        try:
            if is_dark:
                style.theme_use("clam")
            else:
                style.theme_use(original_theme)
        except Exception:
            # Fallback to clam if original not available
            style.theme_use("clam")

        if is_dark:
            bg = "#1e1e1e"; fg = "#e0e0e0"; acc = "#0e639c"; entry_bg = "#2d2d2d"; selbg = "#094771"
            status_fg = "#65b665"
        else:
            bg = "#f0f0f0"; fg = "#000000"; acc = "#005fb8"; entry_bg = "#ffffff"; selbg = "#cce6ff"
            status_fg = "#006400"

        root.configure(background=bg)
        # Configure common ttk styles
        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("TButton", padding=4)
        style.configure("TEntry", fieldbackground=entry_bg, foreground=fg)
        style.configure("TSeparator", background=bg)

        # Treeview styling
        style.configure("Treeview", background=bg, fieldbackground=bg, foreground=fg, rowheight=22)
        style.map("Treeview", background=[("selected", selbg)])
        style.configure("Treeview.Heading", background=bg, foreground=fg)

        # Update status label color if already created
        try:
            status_lbl.configure(foreground=status_fg)
        except Exception:
            pass

    def toggle_dark():
        is_dark = bool(dark_mode.get())
        apply_theme(is_dark)
        write_settings({"dark_mode": is_dark})

    # Initialize dark mode from settings before wiring the menu
    _settings = read_settings()
    dark_mode.set(bool(_settings.get("dark_mode", False)))
    view_menu.add_checkbutton(label="Dark Mode", variable=dark_mode, command=toggle_dark)

    # File menu
    file_menu.add_command(label="Exit", command=root.destroy)

    # About -> Info dialog (read-only view of ABOUT_FILE; external editing)
    def open_info_dialog():
        # Ensure AboutInfo.txt exists next to the EXE/script.
        # Seed from bundled resource if available; otherwise use DEFAULT_ABOUT_TEXT.
        if not os.path.exists(ABOUT_FILE):
            try:
                seed_path = _resource_path("AboutInfo.txt")
                if os.path.exists(seed_path):
                    with open(seed_path, "r", encoding="utf-8", errors="ignore") as sf:
                        seed = sf.read()
                else:
                    seed = DEFAULT_ABOUT_TEXT
                with open(ABOUT_FILE, "w", encoding="utf-8") as f:
                    f.write(seed)
            except Exception:
                try:
                    with open(ABOUT_FILE, "w", encoding="utf-8") as f:
                        f.write(DEFAULT_ABOUT_TEXT)
                except Exception:
                    pass

        dlg = tk.Toplevel(root)
        dlg.title("Info")
        dlg.transient(root)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)
        txt = tk.Text(frm, height=14, wrap="word", state="disabled")
        # Theme text widget roughly to match mode
        try:
            if dark_mode.get():
                txt.configure(bg="#1e1e1e", fg="#e0e0e0", insertbackground="#e0e0e0")
            else:
                txt.configure(bg="#ffffff", fg="#000000", insertbackground="#000000")
        except Exception:
            pass
        txt.pack(fill=tk.BOTH, expand=True, pady=(6,6))

        def linkify(content_text: str):
            # Detect http(s) and steam protocol links and tag them as clickable
            try:
                # Choose a link color based on theme
                link_color = "#4ea1ff" if dark_mode.get() else "#0066cc"
                pattern = re.compile(r"(https?://[^\s<>\"]+|steam://[^\s<>\"]+)")
                for i, m in enumerate(pattern.finditer(content_text)):
                    start = m.start()
                    end = m.end()
                    tag = f"link-{i}"
                    start_idx = f"1.0+{start}c"
                    end_idx = f"1.0+{end}c"
                    url = m.group(0)
                    try:
                        txt.tag_add(tag, start_idx, end_idx)
                        txt.tag_config(tag, foreground=link_color, underline=True)
                        txt.tag_bind(tag, "<Button-1>", lambda e, u=url: webbrowser.open(u, new=2))
                        txt.tag_bind(tag, "<Enter>", lambda e: txt.config(cursor="hand2"))
                        txt.tag_bind(tag, "<Leave>", lambda e: txt.config(cursor="xterm"))
                    except Exception:
                        continue
            except Exception:
                pass

        def load_content():
            try:
                with open(ABOUT_FILE, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
            except Exception:
                content = "(Could not read AboutInfo.txt)"
            txt.configure(state="normal")
            txt.delete("1.0", "end")
            txt.insert("1.0", content)
            # Add clickable link tags
            linkify(content)
            txt.configure(state="disabled")

        load_content()
        dlg.geometry("560x380")
    about_menu.add_command(label="Info", command=open_info_dialog)

    main = ttk.Frame(root, padding=12)
    main.pack(fill=tk.BOTH, expand=True)

    url_var = tk.StringVar()
    mods_var = tk.StringVar(value=";".join(mod_ids))
    ws_var = tk.StringVar(value=";".join(workshop_ids))
    status_var = tk.StringVar(value="Paste a Steam Workshop link and click ADD")

    # URL row
    ttk.Label(main, text="Steam Workshop URL:").grid(row=0, column=0, sticky="w")
    url_entry = ttk.Entry(main, textvariable=url_var)
    url_entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))
    url_entry.focus_set()

    # Reusable dialog to choose a single Mod ID when multiple are present
    def choose_mod_id_dialog(options: List[str]) -> Optional[str]:
        if not options:
            return None
        sel: List[Optional[str]] = [options[0]]
        win = tk.Toplevel(root)
        win.title("Select Mod ID")
        win.transient(root)
        win.grab_set()
        ttk.Label(win, text="Multiple Mod IDs found. Select one to add:").pack(anchor="w", padx=10, pady=(10, 6))
        var = tk.StringVar(value=options[0])
        frame = ttk.Frame(win)
        frame.pack(fill=tk.BOTH, expand=True, padx=10)
        for opt in options:
            ttk.Radiobutton(frame, text=opt, value=opt, variable=var).pack(anchor="w")
        btns = ttk.Frame(win)
        btns.pack(fill=tk.X, pady=10)
        def on_ok():
            sel[0] = var.get()
            win.destroy()
        ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.RIGHT, padx=(0,10))
        # If the user clicks the window close [X], treat it like OK to avoid empty selection
        win.protocol("WM_DELETE_WINDOW", on_ok)
        win.update_idletasks()
        win.geometry(f"400x{min(200, 80 + 24*len(options))}")
        win.wait_window()
        return sel[0]

    def on_add_clicked(*_):
        url = url_var.get().strip().strip('"')
        if not url:
            messagebox.showwarning("Missing URL", "Please paste a Steam Workshop URL.")
            return
        # Check if this is a collection URL first
        col_info = try_fetch_collection(url)
        if col_info and isinstance(col_info, dict) and col_info.get("items"):
            status_var.set(f"Detected collection with {len(col_info['items'])} item(s)")
            cid = col_info["id"]
            if cid in collections_meta:
                msg = f"Collection {cid} is already added."
                status_var.set(msg)
                try:
                    messagebox.showinfo("Already added", msg)
                except Exception:
                    pass
            else:
                # Add children items (skip duplicates already in main list)
                added_w_total = 0
                added_m_total = 0
                skipped_total = 0
                actually_added: List[str] = []
                for wid in col_info["items"]:
                    if wid in workshop_ids:
                        skipped_total += 1
                        continue
                    meta = get_meta_for_workshop_id(wid)
                    # If multiple Mod IDs, prompt user to choose one
                    rmods = meta.get("mods", []) if isinstance(meta, dict) else []
                    if isinstance(rmods, list) and len(rmods) > 1:
                        chosen_mid = choose_mod_id_dialog(rmods)
                        rmods = [chosen_mid] if chosen_mid else []
                    ensure_in_list(workshop_ids, [wid])
                    workshop_meta[wid] = meta
                    added = ensure_in_list(mod_ids, rmods)
                    added_w_total += 1
                    added_m_total += added
                    actually_added.append(wid)
                    root.after(0, lambda w=wid, m=meta: upsert_tree_item(w, m))
                # Track collection metadata (items and only those actually added by this collection)
                collections_meta[cid] = {
                    "title": col_info.get("title") or f"Collection {cid}",
                    "url": col_info.get("url"),
                    "items": list(col_info["items"]),
                    "added": actually_added,
                }
                upsert_collection_item(cid, collections_meta[cid])
                ws_var.set(";".join(workshop_ids))
                mods_var.set(";".join(mod_ids))
                save_memory_to_files()
                # Clear URL and report
                url_var.set("")
                try:
                    url_entry.focus_set()
                except Exception:
                    pass
                status_var.set(f"Added collection {cid}: items added {added_w_total}, skipped {skipped_total}, mod IDs added {added_m_total}")
            return
        else:
            status_var.set("Detected standalone Workshop item")
        # Quick duplicate check using URL without fetching page
        wsid_hint = parse_workshop_id(url)
        if wsid_hint and wsid_hint in workshop_ids:
            msg = f"Workshop item {wsid_hint} is already in the list. Skipping."
            status_var.set(msg)
            try:
                messagebox.showinfo("Already added", msg)
            except Exception:
                pass
            return
        wsid, mods = process_url(url)

        # If multiple Mod IDs found, ask the user to select exactly one
        if len(mods) > 1:
            chosen = choose_mod_id_dialog(mods)
            if chosen:
                mods = [chosen]
            else:
                mods = []
        added_ws = ensure_in_list(workshop_ids, [wsid] if wsid else [])
        added_mods = ensure_in_list(mod_ids, mods)

        # Update UI values
        mods_var.set(";".join(mod_ids))
        ws_var.set(";".join(workshop_ids))
        save_memory_to_files()

        if wsid and mods:
            status = f"Added {added_ws} workshop ID and {added_mods} mod ID(s)."
        elif wsid:
            status = f"Added {added_ws} workshop ID. No mod IDs found in page description."
        elif mods:
            status = f"Added {added_mods} mod ID(s). Workshop ID not found in URL."
        else:
            status = "Nothing added. Could not parse Workshop/Mod IDs."
        if added_ws == 0 and (wsid and wsid in workshop_ids):
            status = f"Workshop item {wsid} is already in the list."
        if added_ws == 0 and added_mods == 0 and wsid:
            status = status + " No new IDs added."
        status_var.set(status)

        # Clear URL field after a successful add and refocus for quick entry
        if (added_ws + added_mods) > 0:
            url_var.set("")
            try:
                url_entry.focus_set()
            except Exception:
                pass

        # If a new workshop ID was added, fetch & display its details
        if wsid and added_ws:
            if wsid not in workshop_meta:
                placeholder = {"title": "(loading)", "version": "…", "url": f"https://steamcommunity.com/sharedfiles/filedetails/?id={wsid}", "mods": mods}
                workshop_meta[wsid] = placeholder
            try:
                upsert_tree_item(wsid, workshop_meta.get(wsid, {}))
            except Exception:
                pass
            threading.Thread(target=fetch_and_update_one, args=(wsid,), daemon=True).start()

            # Auto-add required workshops in the background
            def add_requirements(root_wsid: str):
                reqs = get_required_wsids(root_wsid)
                added_count = 0
                skipped_count = 0
                for rid in reqs:
                    if rid in workshop_ids:
                        continue
                    # Fetch metadata and mods first
                    meta = get_meta_for_workshop_id(rid)
                    rmods = meta.get("mods", []) if isinstance(meta, dict) else []
                    if not rmods:
                        skipped_count += 1
                        continue  # skip adding required workshop if no Mod ID is found
                    # Add required workshop ID and its mods
                    ensure_in_list(workshop_ids, [rid])
                    workshop_meta[rid] = meta
                    ensure_in_list(mod_ids, rmods)
                    added_count += 1
                    # Update UI row for this dependency
                    root.after(0, lambda r=rid, m=meta: upsert_tree_item(r, m))
                # Persist updates and report summary
                def finalize():
                    ws_var.set(";".join(workshop_ids))
                    mods_var.set(";".join(mod_ids))
                    save_memory_to_files()
                    if added_count or skipped_count:
                        status_var.set(f"Dependencies — added: {added_count}, skipped (no Mod ID): {skipped_count}")
                root.after(0, finalize)
            threading.Thread(target=add_requirements, args=(wsid,), daemon=True).start()

    add_btn = ttk.Button(main, text="ADD", command=on_add_clicked)
    add_btn.grid(row=0, column=2, sticky="e")

    # Output rows
    ttk.Separator(main, orient="horizontal").grid(row=1, column=0, columnspan=3, pady=10, sticky="ew")

    ttk.Label(main, text="Mod IDs (ONE LINE; semicolon-separated):").grid(row=2, column=0, columnspan=3, sticky="w")
    mods_entry = ttk.Entry(main, textvariable=mods_var)
    mods_entry.grid(row=3, column=0, columnspan=2, sticky="ew", padx=(0, 8))
    def copy_mods():
        root.clipboard_clear()
        root.clipboard_append(mods_var.get())
        status_var.set("Copied Mod IDs to clipboard")
    ttk.Button(main, text="Copy", command=copy_mods).grid(row=3, column=2, sticky="e")

    ttk.Label(main, text="Workshop IDs (ONE LINE; semicolon-separated):").grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))
    ws_entry = ttk.Entry(main, textvariable=ws_var)
    ws_entry.grid(row=5, column=0, columnspan=2, sticky="ew", padx=(0, 8))
    def copy_ws():
        root.clipboard_clear()
        root.clipboard_append(ws_var.get())
        status_var.set("Copied Workshop IDs to clipboard")
    ttk.Button(main, text="Copy", command=copy_ws).grid(row=5, column=2, sticky="e")

    # Details section: list of workshop items with Name, PZ Version, Link
    ttk.Separator(main, orient="horizontal").grid(row=6, column=0, columnspan=4, pady=10, sticky="ew")
    ttk.Label(main, text="Details (from Workshop):").grid(row=7, column=0, sticky="w")
    actions = ttk.Frame(main)
    actions.grid(row=7, column=1, columnspan=3, sticky="e")

    # Tabs for Mods and Maps
    notebook = ttk.Notebook(main)
    notebook.grid(row=8, column=0, columnspan=4, sticky="nsew")

    # Mods tab
    mods_frame = ttk.Frame(notebook)
    notebook.add(mods_frame, text="Mods")
    mods_cols = ("name", "build", "tags", "link")
    mods_tree = ttk.Treeview(mods_frame, columns=mods_cols, show="headings", height=8, selectmode="extended")
    mods_tree.pack(fill=tk.BOTH, expand=True)
    mods_tree._sort_reverse = {}
    for col, text_lbl, width, anchor in [
        ("name", "Name", 320, "w"),
        ("build", "PZ Build", 120, "center"),
        ("tags", "Tags", 240, "w"),
        ("link", "Workshop Link", 300, "w"),
    ]:
        mods_tree.heading(col, text=text_lbl, command=lambda c=col: sort_treeview(mods_tree, c))
        mods_tree.column(col, width=width, anchor=anchor)

    # Maps tab
    maps_frame = ttk.Frame(notebook)
    notebook.add(maps_frame, text="Maps")
    maps_cols = ("name", "build", "tags", "map_folders", "link")
    maps_tree = ttk.Treeview(maps_frame, columns=maps_cols, show="headings", height=8, selectmode="extended")
    maps_tree.pack(fill=tk.BOTH, expand=True)
    maps_tree._sort_reverse = {}
    for col, text_lbl, width, anchor in [
        ("name", "Name", 300, "w"),
        ("build", "PZ Build", 100, "center"),
        ("tags", "Tags", 180, "w"),
        ("map_folders", "Map Folders", 200, "w"),
        ("link", "Workshop Link", 280, "w"),
    ]:
        maps_tree.heading(col, text=text_lbl, command=lambda c=col: sort_treeview(maps_tree, c))
        maps_tree.column(col, width=width, anchor=anchor)

    # Allow the notebook to expand
    main.rowconfigure(8, weight=1)
    main.columnconfigure(1, weight=1)

    # Collections tab
    collections_frame = ttk.Frame(notebook)
    notebook.add(collections_frame, text="Collections")
    collections_cols = ("name", "count", "link")
    collections_tree = ttk.Treeview(collections_frame, columns=collections_cols, show="headings", height=6)
    collections_tree.pack(fill=tk.BOTH, expand=True)
    collections_tree._sort_reverse = {}
    for col, text_lbl, width, anchor in [
        ("name", "Name", 360, "w"),
        ("count", "Items", 80, "center"),
        ("link", "Collection Link", 320, "w"),
    ]:
        collections_tree.heading(col, text=text_lbl, command=lambda c=col: sort_treeview(collections_tree, c))
        collections_tree.column(col, width=width, anchor=anchor)

    # Sorting helpers
    def parse_version_tuple(s: str):
        if not isinstance(s, str):
            return (True, ())
        nums = re.findall(r"\d+", s)
        if not nums:
            return (True, ())  # unknown sorts last
        return (False, tuple(int(n) for n in nums))

    def sort_treeview(tree_widget: ttk.Treeview, col: str, reverse: Optional[bool] = None):
        children = list(tree_widget.get_children(""))
        def value_of(iid: str) -> str:
            try:
                return tree_widget.set(iid, col)
            except Exception:
                return ""
        def key_func(iid: str):
            v = value_of(iid)
            if col == "build":
                return parse_version_tuple(v)
            # Default to lowercase text sort
            return v.lower()
        reverse_map = getattr(tree_widget, "_sort_reverse", {})
        if reverse is None:
            rev = bool(reverse_map.get(col, False))
        else:
            rev = bool(reverse)
        sorted_children = sorted(children, key=key_func, reverse=rev)
        for idx, iid in enumerate(sorted_children):
            tree_widget.move(iid, "", idx)
        # Toggle only when reverse not explicitly provided
        if reverse is None:
            reverse_map[col] = not rev
        else:
            reverse_map[col] = rev
        tree_widget._sort_reverse = reverse_map

    def sort_by_insertion(tree_widget: ttk.Treeview):
        # Order rows according to the global workshop_ids insertion order
        current = list(tree_widget.get_children(""))
        ordered = [wid for wid in workshop_ids if wid in current]
        seen = set(ordered)
        ordered.extend([iid for iid in current if iid not in seen])
        for idx, iid in enumerate(ordered):
            tree_widget.move(iid, "", idx)

    def on_header_right_click(tree_widget: ttk.Treeview, event):
        # Show context menu only when right-clicking on the heading area
        try:
            region = tree_widget.identify_region(event.x, event.y)
            if region != "heading":
                return
            col_id = tree_widget.identify_column(event.x)  # e.g., '#1'
            idx = int(col_id.replace('#', '')) - 1
            cols = tree_widget["columns"]
            if idx < 0 or idx >= len(cols):
                return
            col_name = cols[idx]
            menu = tk.Menu(tree_widget, tearoff=False)
            menu.add_command(label=f"Sort {col_name} Ascending", command=lambda: sort_treeview(tree_widget, col_name, reverse=False))
            menu.add_command(label=f"Sort {col_name} Descending", command=lambda: sort_treeview(tree_widget, col_name, reverse=True))
            menu.add_separator()
            menu.add_command(label="Sort by Order Added", command=lambda: sort_by_insertion(tree_widget))
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                menu.grab_release()
            except Exception:
                pass

    # Click-and-drag multi-select for Mods/Maps
    def setup_drag_select(tree_widget: ttk.Treeview):
        anchor = {"iid": None}
        def on_button1(event):
            region = tree_widget.identify_region(event.x, event.y)
            if region != "cell":
                return
            row = tree_widget.identify_row(event.y)
            if not row:
                return
            anchor["iid"] = row
            tree_widget.focus(row)
            tree_widget.selection_set(row)
        def on_b1_motion(event):
            if not anchor["iid"]:
                return
            region = tree_widget.identify_region(event.x, event.y)
            if region not in ("cell", "tree"):
                return
            row = tree_widget.identify_row(event.y)
            if not row:
                return
            children = list(tree_widget.get_children(""))
            try:
                i0 = children.index(anchor["iid"])
                i1 = children.index(row)
            except ValueError:
                return
            lo, hi = (i0, i1) if i0 <= i1 else (i1, i0)
            sel = children[lo:hi+1]
            tree_widget.selection_set(sel)
        def on_button_release(event):
            # finalize; nothing extra required
            return
        tree_widget.bind("<Button-1>", on_button1, add=True)
        tree_widget.bind("<B1-Motion>", on_b1_motion, add=True)
        tree_widget.bind("<ButtonRelease-1>", on_button_release, add=True)

    setup_drag_select(mods_tree)
    setup_drag_select(maps_tree)

    # Search/filter UI
    search_var = tk.StringVar(value="")
    search_frame = ttk.Frame(actions)
    search_frame.pack(side=tk.RIGHT, padx=(8,0))
    ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=(0,4))
    search_entry = ttk.Entry(search_frame, textvariable=search_var, width=24)
    search_entry.pack(side=tk.LEFT)
    def apply_current_filter():
        term = (search_var.get() or "").strip().lower()
        def text_for_meta(meta: Dict[str, Any]) -> str:
            parts = [
                str(meta.get("title", "")),
                str(meta.get("version", "")),
                ", ".join(meta.get("tags", [])) if isinstance(meta.get("tags"), list) else "",
                str(meta.get("url", "")),
            ]
            if meta.get("is_map") or meta.get("map_folders"):
                parts.append(", ".join(meta.get("map_folders", [])) if isinstance(meta.get("map_folders"), list) else "")
            return " \u2758 ".join([p for p in parts if p])  # use a nice separator
        # Iterate known workshop ids to reliably reattach/detach
        for wid in list(workshop_ids):
            meta = workshop_meta.get(wid, {})
            is_map_item = bool(meta.get("is_map")) or bool(meta.get("map_folders"))
            tw = maps_tree if is_map_item else mods_tree
            if not tw.exists(wid):
                continue
            text_blob = text_for_meta(meta).lower()
            if term and term not in text_blob:
                try:
                    tw.detach(wid)
                except Exception:
                    pass
            else:
                try:
                    tw.move(wid, "", "end")
                except Exception:
                    pass
    def on_search_click():
        apply_current_filter()
        status_var.set("Filter applied")
    def on_search_clear():
        search_var.set("")
        apply_current_filter()
        status_var.set("Filter cleared")
    ttk.Button(search_frame, text="Go", command=on_search_click).pack(side=tk.LEFT, padx=(4,2))
    ttk.Button(search_frame, text="Clear", command=on_search_clear).pack(side=tk.LEFT)
    search_entry.bind("<Return>", lambda e: on_search_click())

    def upsert_tree_item(wsid: str, meta: Dict[str, str]):
        # Decide destination tree based on is_map/map_folders
        is_map = bool(meta.get("is_map")) or bool(meta.get("map_folders"))
        # Remove from the other tree if present to avoid duplicates
        if is_map:
            if mods_tree.exists(wsid):
                mods_tree.delete(wsid)
            values = (
                meta.get("title", "(unknown)"),
                meta.get("version", "(unknown)"),
                "; ".join(meta.get("tags", [])) if isinstance(meta.get("tags"), list) else "",
                "; ".join(meta.get("map_folders", [])) if isinstance(meta.get("map_folders"), list) else "",
                meta.get("url", f"https://steamcommunity.com/sharedfiles/filedetails/?id={wsid}")
            )
            if maps_tree.exists(wsid):
                maps_tree.item(wsid, values=values)
            else:
                maps_tree.insert("", "end", iid=wsid, values=values)
        else:
            if maps_tree.exists(wsid):
                maps_tree.delete(wsid)
            values = (
                meta.get("title", "(unknown)"),
                meta.get("version", "(unknown)"),
                "; ".join(meta.get("tags", [])) if isinstance(meta.get("tags"), list) else "",
                meta.get("url", f"https://steamcommunity.com/sharedfiles/filedetails/?id={wsid}")
            )
            if mods_tree.exists(wsid):
                mods_tree.item(wsid, values=values)
            else:
                mods_tree.insert("", "end", iid=wsid, values=values)
        # Re-apply filter to ensure visibility aligns with current search
        apply_current_filter()

    def fetch_and_update_one(wsid: str):
        meta = get_meta_for_workshop_id(wsid)
        workshop_meta[wsid] = meta
        root.after(0, lambda: (upsert_tree_item(wsid, meta), save_memory_to_files(), status_var.set(f"Updated details for {wsid}")))

    def refresh_all_details():
        if not workshop_ids:
            status_var.set("No workshop IDs to refresh.")
            return
        status_var.set("Refreshing details for all workshop items…")
        def worker():
            for wid in workshop_ids:
                try:
                    meta = get_meta_for_workshop_id(wid)
                    workshop_meta[wid] = meta
                    root.after(0, lambda w=wid, m=meta: upsert_tree_item(w, m))
                except Exception:
                    continue
            root.after(0, lambda: (save_memory_to_files(), status_var.set("Details refreshed.")))
        threading.Thread(target=worker, daemon=True).start()

    def _get_active_selection() -> Optional[str]:
        # Prefer selection in the currently visible tab
        current_tab = notebook.select()
        if current_tab == mods_frame._w:
            sel = mods_tree.selection()
            if sel:
                return sel[0]
        elif current_tab == maps_frame._w:
            sel = maps_tree.selection()
            if sel:
                return sel[0]
        elif current_tab == collections_frame._w:
            sel = collections_tree.selection()
            if sel:
                return sel[0]
        # Fallback: any selection
        sel = mods_tree.selection()
        if sel:
            return sel[0]
        sel = maps_tree.selection()
        if sel:
            return sel[0]
        return None

    def copy_selected_link():
        wsid = _get_active_selection()
        if not wsid:
            return
        link = workshop_meta.get(wsid, {}).get("url", f"https://steamcommunity.com/sharedfiles/filedetails/?id={wsid}")
        root.clipboard_clear()
        root.clipboard_append(link)
        status_var.set("Copied selected link to clipboard")

    ttk.Button(actions, text="Copy Selected Link", command=copy_selected_link).pack(side=tk.RIGHT, padx=(0,8))
    ttk.Button(actions, text="Refresh All Details", command=refresh_all_details).pack(side=tk.RIGHT)

    # Double-click to open link (favor Steam protocol if available)
    def on_tree_double_click(tree_widget, event):
        sel = tree_widget.selection()
        if not sel:
            return
        wsid = sel[0]
        # Try steam protocol first
        steam_url = f"steam://openurl/https://steamcommunity.com/sharedfiles/filedetails/?id={wsid}"
        try:
            webbrowser.open(steam_url, new=2)
        except Exception:
            webbrowser.open(workshop_meta.get(wsid, {}).get("url", f"https://steamcommunity.com/sharedfiles/filedetails/?id={wsid}"), new=2)
    mods_tree.bind("<Double-1>", lambda e: on_tree_double_click(mods_tree, e))
    maps_tree.bind("<Double-1>", lambda e: on_tree_double_click(maps_tree, e))
    collections_tree.bind("<Double-1>", lambda e: on_tree_double_click(collections_tree, e))
    # Right-click header context menus for sorting
    mods_tree.bind("<Button-3>", lambda e: on_header_right_click(mods_tree, e))
    maps_tree.bind("<Button-3>", lambda e: on_header_right_click(maps_tree, e))
    collections_tree.bind("<Button-3>", lambda e: on_header_right_click(collections_tree, e))

    def upsert_collection_item(cid: str, meta: Dict[str, Any]):
        values = (
            meta.get("title", f"Collection {cid}"),
            str(len(meta.get("items", []))) if isinstance(meta.get("items"), list) else "0",
            meta.get("url", f"https://steamcommunity.com/sharedfiles/filedetails/?id={cid}")
        )
        if collections_tree.exists(cid):
            collections_tree.item(cid, values=values)
        else:
            collections_tree.insert("", "end", iid=cid, values=values)

    def remove_workshop_and_mods(wsid: str):
        # Remove from GUI trees
        if mods_tree.exists(wsid):
            mods_tree.delete(wsid)
        if maps_tree.exists(wsid):
            maps_tree.delete(wsid)
        # Remove from memory lists
        try:
            workshop_ids.remove(wsid)
        except ValueError:
            pass
        # Remove associated mod IDs if present in meta
        meta = workshop_meta.pop(wsid, {})
        mods_for_w = meta.get("mods") if isinstance(meta, dict) else None
        if isinstance(mods_for_w, list):
            remove_set = {m.lower() for m in mods_for_w}
            retained: List[str] = []
            for mid in mod_ids:
                if mid.lower() not in remove_set:
                    retained.append(mid)
            mod_ids.clear()
            mod_ids.extend(retained)
            mods_var.set(";".join(mod_ids))
        # Update IDs UI
        ws_var.set(";".join(workshop_ids))
        save_memory_to_files()

    # Remove selected button
    def remove_selected():
        # Determine active tree and selected items
        current_tab = notebook.select()
        sel: List[str] = []
        if current_tab == mods_frame._w:
            sel = list(mods_tree.selection())
        elif current_tab == maps_frame._w:
            sel = list(maps_tree.selection())
        else:
            status_var.set("Switch to Mods or Maps to remove items, or use Delete Selected Collections on Collections tab.")
            return
        if not sel:
            status_var.set("No selection to remove.")
            return
        removed_count = 0
        for wsid in sel:
            remove_workshop_and_mods(wsid)
            removed_count += 1
        status_var.set(f"Removed {removed_count} item(s)")

    ttk.Button(actions, text="Remove Selected", command=remove_selected).pack(side=tk.RIGHT, padx=(0,8))

    def delete_selected_collections():
        sel = collections_tree.selection()
        if not sel:
            status_var.set("No collections selected.")
            return
        # Build index of item->collections to avoid deleting items still in other collections
        # Count only items actually added by each collection
        item_to_collections: Dict[str, int] = {}
        for cid, cmeta in collections_meta.items():
            for wid in cmeta.get("added", []) if isinstance(cmeta.get("added"), list) else []:
                item_to_collections[wid] = item_to_collections.get(wid, 0) + 1
        removed_items = 0
        removed_collections = 0
        for cid in sel:
            cmeta = collections_meta.get(cid, {})
            # Only remove items that this collection actually added
            children_added = cmeta.get("added", []) if isinstance(cmeta.get("added"), list) else []
            for wid in children_added:
                # Skip if this item appears in more than one collection remaining
                count = item_to_collections.get(wid, 0)
                # If this collection is one of the counts, removing it reduces the count by 1
                if count <= 1:
                    if wid in workshop_ids:
                        remove_workshop_and_mods(wid)
                        removed_items += 1
            # Remove the collection row and metadata
            if collections_tree.exists(cid):
                collections_tree.delete(cid)
            if cid in collections_meta:
                del collections_meta[cid]
                removed_collections += 1
        save_memory_to_files()
        status_var.set(f"Removed {removed_collections} collection(s) and {removed_items} item(s)")

    # Collections-specific actions (shown only when Collections tab is active)
    btn_delete_collections = ttk.Button(actions, text="Delete Selected Collections", command=delete_selected_collections)

    def refresh_selected_collections():
        sel = collections_tree.selection()
        if not sel:
            status_var.set("No collections selected.")
            return
        added_items = 0
        removed_items = 0
        updated_collections = 0
        # Build item->collection count map to handle safe removals
        def build_item_to_collections() -> Dict[str, int]:
            m: Dict[str, int] = {}
            for cid0, cmeta0 in collections_meta.items():
                for wid0 in cmeta0.get("items", []) if isinstance(cmeta0.get("items"), list) else []:
                    m[wid0] = m.get(wid0, 0) + 1
            return m
        for cid in sel:
            cmeta = collections_meta.get(cid, {})
            curl = cmeta.get("url")
            if not curl:
                continue
            updated = try_fetch_collection(curl)
            if not updated:
                continue
            latest = updated.get("items", [])
            existing = cmeta.get("items", []) if isinstance(cmeta.get("items"), list) else []
            existing_set = set(existing)
            latest_set = set(latest)
            # Add new items (skip those already present)
            for wid in sorted(latest_set - existing_set):
                if wid in workshop_ids:
                    continue
                meta = get_meta_for_workshop_id(wid)
                rmods = meta.get("mods", []) if isinstance(meta, dict) else []
                if isinstance(rmods, list) and len(rmods) > 1:
                    chosen_mid = choose_mod_id_dialog(rmods)
                    rmods = [chosen_mid] if chosen_mid else []
                ensure_in_list(workshop_ids, [wid])
                workshop_meta[wid] = meta
                ensure_in_list(mod_ids, rmods)
                root.after(0, lambda w=wid, m=meta: upsert_tree_item(w, m))
                added_items += 1
                # Track in 'added'
                cmeta.setdefault("added", [])
                if wid not in cmeta["added"]:
                    cmeta["added"].append(wid)
            # Remove items no longer in collection IF they were added by this collection and not in other collections
            to_remove = [wid for wid in (existing_set - latest_set) if wid in cmeta.get("added", [])]
            if to_remove:
                count_map = build_item_to_collections()
                for wid in to_remove:
                    # After refresh, this collection still counts for wid until we update items; treat safe if count <= 1
                    if count_map.get(wid, 0) <= 1:
                        if wid in workshop_ids:
                            remove_workshop_and_mods(wid)
                            removed_items += 1
                        try:
                            cmeta["added"].remove(wid)
                        except ValueError:
                            pass
            # Update collection metadata with latest items
            cmeta["items"] = list(latest)
            collections_meta[cid] = cmeta
            upsert_collection_item(cid, cmeta)
            updated_collections += 1
        ws_var.set(";".join(workshop_ids))
        mods_var.set(";".join(mod_ids))
        save_memory_to_files()
        status_var.set(f"Collections refreshed: {updated_collections}. Added items: {added_items}. Removed items: {removed_items}.")

    btn_refresh_collections = ttk.Button(actions, text="Refresh Selected Collections", command=refresh_selected_collections)

    def update_actions_visibility(event=None):
        # Show collection-specific buttons only on Collections tab
        current_tab = notebook.select()
        is_col = current_tab == collections_frame._w
        # First hide them
        try:
            btn_delete_collections.pack_forget()
        except Exception:
            pass
        try:
            btn_refresh_collections.pack_forget()
        except Exception:
            pass
        if is_col:
            btn_refresh_collections.pack(side=tk.RIGHT, padx=(0,8))
            btn_delete_collections.pack(side=tk.RIGHT, padx=(0,8))

    notebook.bind("<<NotebookTabChanged>>", update_actions_visibility)
    # Initialize visibility state
    update_actions_visibility()

    # Paste from clipboard and ADD shortcut
    def paste_and_add():
        try:
            clip = root.clipboard_get().strip()
        except Exception:
            clip = ""
        if not clip:
            status_var.set("Clipboard is empty.")
            return
        url_var.set(clip)
        on_add_clicked()
    ttk.Button(main, text="Paste from Clipboard + ADD", command=paste_and_add).grid(row=0, column=3, sticky="e", padx=(8,0))

    # Preload details from cache and schedule fetch for missing
    for wid in workshop_ids:
        meta = workshop_meta.get(wid)
        if meta:
            upsert_tree_item(wid, meta)
        else:
            placeholder = {"title": "(loading)", "version": "…", "url": f"https://steamcommunity.com/sharedfiles/filedetails/?id={wid}"}
            upsert_tree_item(wid, placeholder)
            threading.Thread(target=fetch_and_update_one, args=(wid,), daemon=True).start()

    # Preload collections
    for cid, cmeta in collections_meta.items():
        upsert_collection_item(cid, cmeta)

    # Status bar
    status_lbl = ttk.Label(main, textvariable=status_var, foreground="#006400")
    status_lbl.grid(row=9, column=0, columnspan=3, sticky="w", pady=(12, 0))

    # Grid config
    main.columnconfigure(1, weight=1)

    # Apply initial theme based on saved preference
    apply_theme(bool(dark_mode.get()))

    # Bind Enter key to ADD
    root.bind("<Return>", on_add_clicked)

    root.mainloop()


def interactive_loop():
    """Legacy CLI interactive loop (use --cli)."""
    print("Project Zomboid Workshop/Mod ID scraper (CLI)")
    print("Paste a Steam Workshop URL (or type 'q' to quit).\n")
    load_existing_to_memory()
    while True:
        try:
            url = input("URL: ").strip().strip('"')
        except (EOFError, KeyboardInterrupt):
            print("\nExiting.")
            break
        if not url:
            continue
        if url.lower() in {"q", "quit", "exit"}:
            break

        wsid, mods = process_url(url)
        if len(mods) > 1:
            print("Multiple Mod IDs detected:")
            for i, m in enumerate(mods, 1):
                print(f"  {i}. {m}")
            choice = input("Pick one number to add (or press Enter to skip): ").strip()
            if choice.isdigit() and 1 <= int(choice) <= len(mods):
                mods = [mods[int(choice)-1]]
            else:
                mods = []
        ensure_in_list(workshop_ids, [wsid] if wsid else [])
        ensure_in_list(mod_ids, mods)
        save_memory_to_files()

        print("— Result —")
        print(f"Workshop ID: {wsid or 'not found'}")
        if mods:
            print("Mod ID(s): " + ", ".join(mods))
        else:
            print("Mod ID(s): not found in page description")
        print("Current one-line Mod IDs: " + ";".join(mod_ids))
        print("Current one-line Workshop IDs: " + ";".join(workshop_ids))
        print()


def main(args: List[str]) -> int:
    # Basic arg parsing: --gui (default), --cli, or list of URLs
    if args and args[0] in {"--cli", "-c"}:
        return interactive_loop() or 0
    if args and args[0] in {"--gui", "-g"}:
        build_gui()
        return 0
    if args:
        # Process URLs passed as arguments (CLI batch)
        load_existing_to_memory()
        total_new_w = 0
        total_new_m = 0
        for url in args:
            wsid_hint = parse_workshop_id(url)
            if wsid_hint and wsid_hint in workshop_ids:
                print(f"Skipping {url} — workshop {wsid_hint} already in list.")
                continue
            wsid, mods = process_url(url)
            total_new_w += ensure_in_list(workshop_ids, [wsid] if wsid else [])
            total_new_m += ensure_in_list(mod_ids, mods)
            print(f"Processed: {url}")
            print(f"  Workshop ID: {wsid or 'not found'}")
            print(f"  Mod ID(s): {', '.join(mods) if mods else 'not found'}")
        save_memory_to_files()
        print("\nSummary:")
        print(f"  New workshop IDs added: {total_new_w}")
        print(f"  New mod IDs added: {total_new_m}")
        print("One-line Mod IDs: " + ";".join(mod_ids))
        print("One-line Workshop IDs: " + ";".join(workshop_ids))
        return 0
    else:
        # Default to GUI when no args
        build_gui()
        return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
