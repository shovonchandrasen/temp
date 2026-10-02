"""Flask web server for the NIC Report Generator.

The browser flow is deliberately small and focused:
1. Select the three report files in one action; the browser/server segregate them.
2. Validate the files and review only CSV products needing a manual mapping.
3. Confirm/lock the mappings, then generate the two processed workbooks.

Manual mappings are kept in a local JSON file in this application folder so a
future run can offer to restore them. The original processing script remains
separate and untouched.
"""
import io
import json
import os
import shutil
import sys
import threading
import webbrowser
import zipfile
from pathlib import Path

import pandas as pd
from flask import Flask, abort, jsonify, render_template, request, send_file

BASE_DIR = Path(__file__).resolve().parent
MAIN_DIR = BASE_DIR.parent
sys.path.insert(0, str(BASE_DIR))
from processor import (  # noqa: E402
    UNMATCH_VALUE_COLS,
    classify_file,
    find_unmatched_csv_items,
    get_closing_products,
    get_inventory_products,
    process_files,
    run_checks,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024

UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
MEMORY_FILE = BASE_DIR / "mapping_memory.json"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def _clear_dir(directory: Path):
    directory.mkdir(exist_ok=True)
    for item in directory.iterdir():
        if item.is_file() or item.is_symlink():
            item.unlink()


def _secure(name: str) -> str:
    """Keep uploaded/downloaded names as simple files inside our folders."""
    return os.path.basename(name).replace("..", "_")


def _load_mapping_memory():
    """Read saved manual mappings; a damaged/absent file behaves as empty."""
    try:
        with MEMORY_FILE.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
        mappings = payload.get("mappings", payload) if isinstance(payload, dict) else {}
        if not isinstance(mappings, dict):
            return {}
        clean = {}
        for csv_name, values in mappings.items():
            if not isinstance(values, dict):
                continue
            inv = str(values.get("inventory", "") or "").strip()
            clo = str(values.get("closing", "") or "").strip()
            if inv or clo:
                clean[str(csv_name)] = {"inventory": inv, "closing": clo}
        return clean
    except (OSError, ValueError, TypeError):
        return {}


def _save_mapping_memory(updates):
    """Merge locked mappings into the local memory file atomically.

    A current locked row overwrites both fields, including an intentional blank,
    so changing a previous mapping really updates the saved choice.
    """
    current = _load_mapping_memory()
    lower_to_key = {key.lower(): key for key in current}
    for csv_name, values in (updates or {}).items():
        csv_name = str(csv_name).strip()
        if not csv_name or not isinstance(values, dict):
            continue
        old_key = lower_to_key.get(csv_name.lower())
        if old_key and old_key != csv_name:
            current.pop(old_key, None)
        inv = str(values.get("inventory", "") or "").strip()
        clo = str(values.get("closing", "") or "").strip()
        if inv or clo:
            current[csv_name] = {"inventory": inv, "closing": clo}
        else:
            current.pop(csv_name, None)

    temporary = MEMORY_FILE.with_suffix(".tmp")
    with temporary.open("w", encoding="utf-8") as handle:
        json.dump({"version": 1, "mappings": current}, handle, ensure_ascii=False, indent=2)
    temporary.replace(MEMORY_FILE)
    return current


def _memory_for_items(unmatched):
    memory = _load_mapping_memory()
    by_lower = {key.lower(): value for key, value in memory.items()}
    return {
        item["name"]: by_lower[item["name"].lower()]
        for item in unmatched
        if item["name"].lower() in by_lower
    }


def _find_main_folder_reports():
    """Detect if the 3 required report files are present in the main folder."""
    found = {}
    try:
        for item in sorted(MAIN_DIR.iterdir(), key=lambda p: p.name.lower()):
            if not item.is_file():
                continue
            if item.name.startswith(".") or item.name.startswith("~$"):
                continue
            role = classify_file(item.name)
            if role is None:
                continue
            if role in found:
                return None
            found[role] = item
    except OSError:
        return None

    if set(found.keys()) != {"inventory", "closing", "stock_csv"}:
        return None
    return found


@app.route("/")
def index():
    # A new upload session starts clean; mapping_memory.json is intentionally kept.
    _clear_dir(UPLOAD_DIR)
    _clear_dir(OUTPUT_DIR)

    auto_files = None
    main_reports = _find_main_folder_reports()
    if main_reports:
        auto_files = {}
        slot_map = {"inventory": "inventory", "closing": "closing", "stock_csv": "csv"}
        for role, src_path in main_reports.items():
            safe_name = _secure(src_path.name)
            dest_path = UPLOAD_DIR / safe_name
            shutil.copy2(src_path, dest_path)
            slot_key = slot_map[role]
            auto_files[slot_key] = {
                "name": safe_name,
                "size": dest_path.stat().st_size,
                "fromMainFolder": True,
            }

    return render_template("index.html", auto_files=auto_files)


@app.route("/upload", methods=["POST"])
def upload():
    files = [file for file in request.files.getlist("files") if file and file.filename]
    main_files = [
        name.strip()
        for name in request.form.getlist("main_files")
        if name and name.strip()
    ]
    if len(files) + len(main_files) != 3:
        return jsonify({"ok": False, "error": "Select exactly three report files."}), 400

    saved = {}
    upload_errors = []

    for name in main_files:
        safe_name = _secure(name)
        destination = UPLOAD_DIR / safe_name
        source = MAIN_DIR / safe_name
        if not destination.exists():
            if source.exists() and source.is_file():
                shutil.copy2(source, destination)
            else:
                upload_errors.append(f"File not found in main folder: {safe_name}")
                continue
        role = classify_file(safe_name)
        if role is None:
            upload_errors.append(
                f"Unrecognized file: {safe_name}. Expected Inventory Activity, "
                "Closing Stock, and Store Wise Stock CSV reports."
            )
            continue
        if role in saved:
            upload_errors.append(f"More than one file was supplied for {role}: {safe_name}")
            continue
        saved[role] = (safe_name, str(destination))

    for file in files:
        safe_name = _secure(file.filename)
        destination = UPLOAD_DIR / safe_name
        file.save(destination)
        role = classify_file(safe_name)
        if role is None:
            upload_errors.append(
                f"Unrecognized file: {safe_name}. Expected Inventory Activity, "
                "Closing Stock, and Store Wise Stock CSV reports."
            )
            continue
        if role in saved:
            upload_errors.append(f"More than one file was supplied for {role}: {safe_name}")
            continue
        saved[role] = (safe_name, str(destination))

    if upload_errors or len(saved) != 3:
        message = " | ".join(upload_errors) or "The three report types could not be identified."
        return jsonify({"ok": False, "error": message}), 400

    inv_name, inv_path = saved["inventory"]
    clo_name, clo_path = saved["closing"]
    csv_name, csv_path = saved["stock_csv"]

    checks, meta = run_checks(inv_path, clo_path, csv_path)

    # Full lists are required to decide whether a CSV name is truly unmatched.
    # The UI lists only Excel names that are not already exact CSV names.
    inv_all = []
    clo_all = []
    inv_ui = []
    clo_ui = []
    try:
        csv_frame = pd.read_csv(csv_path)
        csv_frame["Item Name"] = csv_frame["Item Name"].astype(str).str.strip()
        csv_names = {name.lower() for name in csv_frame["Item Name"] if name}
    except Exception as exc:
        csv_names = set()
        checks["CSV has required stock columns"] = (False, str(exc))
        meta["all_passed"] = False

    try:
        inv_all = get_inventory_products(inv_path)
        inv_ui = [product for product in inv_all if product["name"].lower() not in csv_names]
        checks["Inventory Activity product list readable"] = (
            True, f"{len(inv_ui)} mapping candidates of {len(inv_all)} total"
        )
    except Exception as exc:
        checks["Inventory Activity product list readable"] = (False, str(exc))
        meta["all_passed"] = False

    try:
        clo_all = get_closing_products(clo_path)
        clo_ui = [product for product in clo_all if product["name"].lower() not in csv_names]
        checks["Closing Stock product list readable"] = (
            True, f"{len(clo_ui)} mapping candidates of {len(clo_all)} total"
        )
    except Exception as exc:
        checks["Closing Stock product list readable"] = (False, str(exc))
        meta["all_passed"] = False

    unmatched = find_unmatched_csv_items(csv_path, inv_all, clo_all)
    saved_mappings = _memory_for_items(unmatched)

    return jsonify({
        "ok": True,
        "files": {"inventory": inv_name, "closing": clo_name, "csv": csv_name},
        "paths": {"inventory": inv_path, "closing": clo_path, "csv": csv_path},
        "checks": [{"label": key, "ok": value[0], "msg": value[1]} for key, value in checks.items()],
        "all_passed": meta["all_passed"],
        "month_label": meta["month_label"],
        "months": {
            "Inventory Activity": meta["inv_month"][2],
            "Closing Stock": meta["clo_month"][2],
            "Store Wise CSV": meta["csv_month"][2],
        },
        "inventory_products": inv_ui,
        "closing_products": clo_ui,
        "unmatched": unmatched,
        "saved_mappings": saved_mappings,
        "value_cols": UNMATCH_VALUE_COLS,
    })


@app.route("/process", methods=["POST"])
def process():
    data = request.get_json(force=True) or {}
    inv_path = data.get("inventory")
    clo_path = data.get("closing")
    csv_path = data.get("csv")
    inv_mapping = data.get("inv_mapping") or {}
    clo_mapping = data.get("clo_mapping") or {}
    memory_updates = data.get("memory_updates") or {}

    upload_root = UPLOAD_DIR.resolve()
    for path in (inv_path, clo_path, csv_path):
        if not path or not os.path.exists(path):
            return jsonify({"ok": False, "error": "Missing uploaded file."}), 400
        try:
            Path(path).resolve().relative_to(upload_root)
        except ValueError:
            return jsonify({"ok": False, "error": "Invalid uploaded file path."}), 400

    try:
        result = process_files(
            inv_path, clo_path, csv_path, str(OUTPUT_DIR),
            inv_mapping=inv_mapping,
            clo_mapping=clo_mapping,
        )
        saved_memory = _save_mapping_memory(memory_updates)
    except Exception as exc:
        import traceback
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500

    return jsonify({
        "ok": True,
        "outputs": result["outputs"],
        "log": result["log"],
        "saved_mapping_count": len(saved_memory),
    })


@app.route("/download/<path:filename>")
def download_one(filename):
    safe = _secure(filename)
    path = OUTPUT_DIR / safe
    if not path.exists() or not path.is_file():
        abort(404)
    return send_file(str(path), as_attachment=True, download_name=safe)


@app.route("/download-all")
def download_all():
    files = [file for file in OUTPUT_DIR.iterdir() if file.is_file()]
    if not files:
        abort(404)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for file in files:
            archive.write(file, arcname=file.name)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name="processed_reports.zip",
        mimetype="application/zip",
    )


def _open_browser(url, delay=1.5):
    def open_later():
        import time
        time.sleep(delay)
        try:
            webbrowser.open(url)
        except Exception:
            pass

    threading.Thread(target=open_later, daemon=True).start()


if __name__ == "__main__":
    host = "127.0.0.1"
    port = 5000
    url = f"http://{host}:{port}/"
    print("=" * 60)
    print(" NIC Report Generator")
    print(" Server: " + url)
    print(" Keep this window open; press Ctrl+C to stop.")
    print("=" * 60)
    _open_browser(url)
    app.run(host="0.0.0.0", port=port, debug=False)
