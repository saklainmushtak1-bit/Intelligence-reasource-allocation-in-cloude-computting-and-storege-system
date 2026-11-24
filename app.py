import sqlite3
import csv
import os
import shutil
from io import StringIO
from datetime import datetime

import psutil
import pandas as pd
import joblib
from flask import (
    Flask, render_template_string, request, Response, jsonify,
    send_from_directory, redirect, url_for, session
)

app = Flask(__name__)
app.secret_key = "nextdrive_super_secure_academic_key_2026"
DATABASE = "resource_allocation.db"
MODEL_PATH = "cloud_resource_allocator_model.joblib"
STORAGE_DIR = os.path.abspath("cloud_storage")
os.makedirs(STORAGE_DIR, exist_ok=True)
usage_history = []

# Load trained Machine Learning model
ml_model = None
if os.path.exists(MODEL_PATH):
    try:
        ml_model = joblib.load(MODEL_PATH)
        print(f"Loaded Machine Learning Model: {MODEL_PATH}")
    except Exception as e:
        print(f"Error loading model: {e}")


# ---------------- DATABASE & AUTH SETUP ----------------

def init_db():
    with sqlite3.connect(DATABASE) as conn:
        # Users table for User & Admin authentication
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                full_name TEXT,
                plan TEXT DEFAULT '3-Month Free Trial'
            )
        """)

        # Seed default Admin and User accounts if not exists
        conn.execute("""
            INSERT OR IGNORE INTO users (username, password, role, full_name, plan)
            VALUES 
            ('admin', 'admin123', 'admin', 'System Administrator', 'Enterprise Cloud Master'),
            ('user', 'user123', 'user', 'Demo User', '3-Month Free Trial (10 GB)')
        """)

        # Allocation history table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS allocation_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                time TEXT,
                cpu INTEGER,
                ram INTEGER,
                storage INTEGER DEFAULT 50,
                cpu_usage REAL,
                ram_usage REAL,
                storage_usage REAL DEFAULT 0,
                demand_status TEXT,
                demand_class TEXT,
                recommendation TEXT,
                ml_prediction TEXT DEFAULT 'N/A',
                allocated_vms INTEGER DEFAULT 1,
                est_cost_hr REAL DEFAULT 0.05
            )
        """)

        columns = [
            row[1] for row in
            conn.execute("PRAGMA table_info(allocation_history)").fetchall()
        ]

        if "demand_class" not in columns:
            conn.execute("ALTER TABLE allocation_history ADD COLUMN demand_class TEXT DEFAULT 'medium'")
        if "storage" not in columns:
            conn.execute("ALTER TABLE allocation_history ADD COLUMN storage INTEGER DEFAULT 50")
        if "storage_usage" not in columns:
            conn.execute("ALTER TABLE allocation_history ADD COLUMN storage_usage REAL DEFAULT 0")
        if "ml_prediction" not in columns:
            conn.execute("ALTER TABLE allocation_history ADD COLUMN ml_prediction TEXT DEFAULT 'N/A'")
        if "allocated_vms" not in columns:
            conn.execute("ALTER TABLE allocation_history ADD COLUMN allocated_vms INTEGER DEFAULT 1")
        if "est_cost_hr" not in columns:
            conn.execute("ALTER TABLE allocation_history ADD COLUMN est_cost_hr REAL DEFAULT 0.05")


def clean_status(value):
    status = str(value or "").strip().upper()
    while status.endswith(" DEMAND"):
        status = status[:-7].strip()
    if status in ("LOW", "MEDIUM", "HIGH"):
        return status
    return "OTHER"


def get_history():
    with sqlite3.connect(DATABASE) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute("""
            SELECT * FROM allocation_history
            ORDER BY id DESC LIMIT 20
        """).fetchall()

    history = []
    for row in rows:
        item = dict(row)
        status = clean_status(item.get("demand_status"))
        item["demand_status"] = status
        item["demand_class"] = (
            status.lower() if status in ("LOW", "MEDIUM", "HIGH")
            else "other"
        )
        if item.get("storage") is None:
            item["storage"] = 50
        if item.get("storage_usage") is None:
            item["storage_usage"] = 0
        if not item.get("ml_prediction"):
            item["ml_prediction"] = f"{status} Demand (ML)"
        if not item.get("allocated_vms"):
            item["allocated_vms"] = 1 if status == "LOW" else (2 if status == "MEDIUM" else 4)
        if not item.get("est_cost_hr"):
            item["est_cost_hr"] = round(item["allocated_vms"] * 0.048, 3)

        history.append(item)
    return history


def get_users_list():
    with sqlite3.connect(DATABASE) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute("SELECT id, username, role, full_name, plan FROM users").fetchall()
    return [dict(r) for r in rows]


# ---------------- NEXTDRIVE STORAGE HELPER ----------------

def get_file_type_badge(filename):
    ext = os.path.splitext(filename)[1].lower()
    if ext in ['.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp']:
        return ('IMAGE', '#f59e0b', '🖼️', 'images')
    elif ext in ['.mp4', '.mkv', '.avi', '.mov', '.webm', '.flv']:
        return ('VIDEO', '#ef4444', '🎬', 'videos')
    elif ext in ['.mp3', '.wav', '.ogg', '.m4a']:
        return ('AUDIO', '#ec4899', '🎵', 'videos')
    elif ext in ['.pdf', '.doc', '.docx', '.txt']:
        return ('DOCUMENT', '#3b82f6', '📄', 'documents')
    elif ext in ['.csv', '.xlsx', '.xls']:
        return ('SPREADSHEET', '#10b981', '📊', 'documents')
    elif ext in ['.zip', '.tar', '.gz', '.rar', '.7z']:
        return ('ARCHIVE', '#8b5cf6', '📦', 'other')
    elif ext in ['.py', '.json', '.html', '.css', '.js', '.sql', '.cpp', '.java']:
        return ('CODE', '#06b6d4', '💻', 'other')
    return ('FILE', '#64748b', '📁', 'other')


def get_nextdrive_info():
    items = []
    total_bytes = 0
    counts = {'all': 0, 'images': 0, 'videos': 0, 'documents': 0, 'folders': 0}

    if os.path.exists(STORAGE_DIR):
        for fname in sorted(os.listdir(STORAGE_DIR)):
            fpath = os.path.join(STORAGE_DIR, fname)
            mtime = datetime.fromtimestamp(os.path.getmtime(fpath)).strftime("%b %d, %Y • %I:%M %p")
            
            if os.path.isdir(fpath):
                child_count = len(os.listdir(fpath))
                items.append({
                    "name": fname,
                    "is_dir": True,
                    "size_str": f"{child_count} items",
                    "modified": mtime,
                    "category": "FOLDER",
                    "badge_color": "#6366f1",
                    "icon": "📁",
                    "tab": "folders"
                })
                counts['folders'] += 1
                counts['all'] += 1
            elif os.path.isfile(fpath):
                fsize = os.path.getsize(fpath)
                total_bytes += fsize
                category, badge_color, icon, tab_category = get_file_type_badge(fname)
                
                if fsize < 1024:
                    size_str = f"{fsize} B"
                elif fsize < 1024 * 1024:
                    size_str = f"{round(fsize / 1024, 1)} KB"
                elif fsize < 1024 * 1024 * 1024:
                    size_str = f"{round(fsize / (1024 * 1024), 2)} MB"
                else:
                    size_str = f"{round(fsize / (1024 * 1024 * 1024), 2)} GB"

                items.append({
                    "name": fname,
                    "is_dir": False,
                    "size_str": size_str,
                    "size_bytes": fsize,
                    "modified": mtime,
                    "category": category,
                    "badge_color": badge_color,
                    "icon": icon,
                    "tab": tab_category
                })
                counts['all'] += 1
                if tab_category in counts:
                    counts[tab_category] += 1

    quota_gb = 10.0
    quota_bytes = quota_gb * (1024 ** 3)
    used_mb = round(total_bytes / (1024 * 1024), 2)
    used_gb = round(total_bytes / (1024 ** 3), 4)
    free_gb = round(max(0.0, quota_gb - used_gb), 2)
    used_percent = min(100.0, round((total_bytes / quota_bytes) * 100, 2))

    return {
        "file_list": items,
        "total_items": len(items),
        "counts": counts,
        "used_mb": used_mb,
        "used_gb": used_gb,
        "free_gb": free_gb,
        "quota_gb": quota_gb,
        "used_percent": max(0.5, used_percent)
    }


# ---------------- RESOURCE ALLOCATION SUMMARY ----------------

def get_summary():
    with sqlite3.connect(DATABASE) as conn:
        rows = conn.execute("SELECT demand_status FROM allocation_history").fetchall()

    total = len(rows)
    low = medium = high = other = 0

    for row in rows:
        status = clean_status(row[0])
        if status == "LOW":
            low += 1
        elif status == "MEDIUM":
            medium += 1
        elif status == "HIGH":
            high += 1
        else:
            other += 1

    return {
        "total": total,
        "low": low,
        "medium": medium,
        "high": high,
        "other": other
    }


# ---------------- CPU, RAM & STORAGE MONITORING ----------------

def get_usage():
    cpu = psutil.cpu_percent(interval=0.3)
    ram = psutil.virtual_memory().percent
    try:
        disk = psutil.disk_usage('.')
        storage = disk.percent
        storage_total_gb = round(disk.total / (1024**3), 1)
        storage_used_gb = round(disk.used / (1024**3), 1)
        storage_free_gb = round(disk.free / (1024**3), 1)
    except Exception:
        storage = 0.0
        storage_total_gb = storage_used_gb = storage_free_gb = 0.0

    return (
        round(cpu, 1),
        round(ram, 1),
        round(storage, 1),
        storage_total_gb,
        storage_used_gb,
        storage_free_gb
    )


def record_usage(cpu, ram, storage):
    usage_history.append({
        "time": datetime.now().strftime("%H:%M:%S"),
        "cpu": cpu,
        "ram": ram,
        "storage": storage
    })
    if len(usage_history) > 20:
        usage_history.pop(0)


# ---------------- DEMAND ASSESSMENT & ML PREDICTION ----------------

def assess_demand(cpu, ram, storage=50, tasks=40, network_io=250.0):
    ml_tier_label = "MEDIUM"
    ml_confidence = 90.0

    if ml_model is not None:
        try:
            input_df = pd.DataFrame([{
                "cpu_demand": float(cpu),
                "ram_demand": float(ram),
                "storage_demand": float(storage),
                "active_tasks": int(tasks),
                "network_io_mbps": float(network_io)
            }])
            tier_code = ml_model.predict(input_df)[0]
            tier_names = {0: "LOW", 1: "MEDIUM", 2: "HIGH"}
            ml_tier_label = tier_names.get(tier_code, "MEDIUM")
            if hasattr(ml_model, "predict_proba"):
                probs = ml_model.predict_proba(input_df)[0]
                ml_confidence = round(float(probs[tier_code]) * 100, 1)
        except Exception as e:
            print("ML Inference warning:", e)

    if ml_tier_label == "LOW":
        vms = 1
        vm_type = "t3.micro (1 vCPU, 1 GB RAM, NextDrive Standard Volume)"
        hourly_cost = 0.024
        recommendation = (
            f"ML Model: LOW DEMAND (Confidence: {ml_confidence}%). "
            f"Workload is calm. Downscaled cluster to {vms}x {vm_type} for cost savings."
        )
    elif ml_tier_label == "MEDIUM":
        vms = 2
        vm_type = "t3.medium (2 vCPU, 4 GB RAM, NextDrive High-IOPS Volume)"
        hourly_cost = 0.084
        recommendation = (
            f"ML Model: MEDIUM DEMAND (Confidence: {ml_confidence}%). "
            f"Standard traffic detected. Balanced {vms}x {vm_type} nodes active."
        )
    else:
        vms = 4
        vm_type = "t3.xlarge (4 vCPU, 16 GB RAM, NextDrive Auto-Scaled Elastic SSD Volume)"
        hourly_cost = 0.248
        recommendation = (
            f"ML Model: HIGH DEMAND (Confidence: {ml_confidence}%). "
            f"CRITICAL SPIKE detected! Auto-provisioned {vms}x {vm_type} nodes + Expanded NextDrive Cloud Block Storage."
        )

    return (
        ml_tier_label,
        ml_tier_label.lower(),
        recommendation,
        f"{ml_tier_label} (Random Forest: {ml_confidence}%)",
        vms,
        hourly_cost
    )


# ---------------- LOGIN & SIGNUP HTML ----------------

LOGIN_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NextDrive - User & Admin Login</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: #090d16;
    color: #e2e8f0;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
}
.login-card {
    background: #111827;
    border: 1px solid #1f2937;
    border-radius: 20px;
    padding: 36px;
    width: 100%;
    max-width: 440px;
    box-shadow: 0 20px 50px rgba(0,0,0,0.6);
}
.brand-area {
    text-align: center;
    margin-bottom: 24px;
}
.drive-logo {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 48px;
    height: 48px;
    border-radius: 12px;
    background: linear-gradient(135deg, #3b82f6 0%, #10b981 50%, #f59e0b 100%);
    font-size: 26px;
    margin-bottom: 12px;
    box-shadow: 0 4px 14px rgba(59, 130, 246, 0.4);
}
h2 { font-size: 24px; color: #f8fafc; font-weight: 700; }
h2 span { color: #38bdf8; }
.subtext { font-size: 13px; color: #94a3b8; margin-top: 4px; }

.flash-msg {
    padding: 12px;
    border-radius: 8px;
    font-size: 13px;
    margin-bottom: 18px;
    font-weight: 500;
}
.flash-error { background: #7f1d1d; color: #fecaca; border: 1px solid #dc2626; }
.flash-success { background: #064e3b; color: #a7f3d0; border: 1px solid #059669; }

.form-group { margin-bottom: 16px; }
label { display: block; font-size: 13px; color: #cbd5e1; margin-bottom: 6px; font-weight: 500; }
input {
    width: 100%;
    padding: 12px 14px;
    background: #090d16;
    border: 1px solid #374151;
    border-radius: 8px;
    color: #f8fafc;
    font-size: 14px;
}
input:focus { outline: none; border-color: #38bdf8; }

.btn-submit {
    width: 100%;
    padding: 12px;
    background: #2563eb;
    color: white;
    border: none;
    border-radius: 8px;
    font-size: 15px;
    font-weight: 600;
    cursor: pointer;
    margin-top: 10px;
    transition: background 0.2s;
}
.btn-submit:hover { background: #1d4ed8; }

.demo-credentials {
    margin-top: 24px;
    padding-top: 18px;
    border-top: 1px solid #1f2937;
    background: #1a2234;
    padding: 14px;
    border-radius: 10px;
    font-size: 12px;
}
.demo-title { font-weight: 700; color: #f59e0b; margin-bottom: 6px; }
.credential-item { display: flex; justify-content: space-between; margin-bottom: 4px; color: #94a3b8; }
.credential-item strong { color: #f8fafc; }
.quick-btn {
    background: #334155;
    color: #f8fafc;
    border: none;
    padding: 2px 8px;
    border-radius: 4px;
    font-size: 11px;
    cursor: pointer;
}
.quick-btn:hover { background: #475569; }
</style>
</head>
<body>

<div class="login-card">
    <div class="brand-area">
        <div class="drive-logo">▲</div>
        <h2>Next<span>Drive</span> Cloud</h2>
        <p class="subtext">Sign in to access your 10 GB Free Storage & AI Engine</p>
    </div>

    {% if error %}
    <div class="flash-msg flash-error">{{ error }}</div>
    {% endif %}
    {% if success %}
    <div class="flash-msg flash-success">{{ success }}</div>
    {% endif %}

    <form method="POST" action="/login">
        <div class="form-group">
            <label for="username">Username</label>
            <input type="text" id="username" name="username" placeholder="e.g. admin or user" required>
        </div>
        <div class="form-group">
            <label for="password">Password</label>
            <input type="password" id="password" name="password" placeholder="••••••••" required>
        </div>
        <button type="submit" class="btn-submit">Sign In to NextDrive</button>
    </form>

    <div class="demo-credentials">
        <div class="demo-title">🔑 Quick Demo Access (College Submission):</div>
        <div class="credential-item">
            <span>Admin Portal: <strong>admin</strong> / <strong>admin123</strong></span>
            <button class="quick-btn" onclick="fillCredentials('admin', 'admin123')">Fill</button>
        </div>
        <div class="credential-item">
            <span>User Drive: <strong>user</strong> / <strong>user123</strong></span>
            <button class="quick-btn" onclick="fillCredentials('user', 'user123')">Fill</button>
        </div>
    </div>
</div>

<script>
function fillCredentials(u, p) {
    document.getElementById('username').value = u;
    document.getElementById('password').value = p;
}
</script>
</body>
</html>
"""


# ---------------- DASHBOARD HTML (NEXTDRIVE FULL DRIVE SYSTEM) ----------------

HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NextDrive - 10 GB Cloud Storage & Intelligent Resource Allocation</title>

<style>
* { box-sizing: border-box; }

body {
    margin: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    background: #090d16;
    color: #e2e8f0;
}

/* TOP APP BAR (GOOGLE DRIVE STYLE) */
.navbar {
    background: #111827;
    border-bottom: 1px solid #1f2937;
    padding: 12px 28px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    position: sticky;
    top: 0;
    z-index: 100;
}

.brand {
    display: flex;
    align-items: center;
    gap: 12px;
}

.drive-logo {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 40px;
    height: 40px;
    border-radius: 10px;
    background: linear-gradient(135deg, #3b82f6 0%, #10b981 50%, #f59e0b 100%);
    font-size: 22px;
    box-shadow: 0 4px 14px rgba(59, 130, 246, 0.4);
}

.brand h1 {
    margin: 0;
    font-size: 22px;
    font-weight: 700;
    color: #f8fafc;
    letter-spacing: -0.4px;
}

.brand h1 span {
    color: #38bdf8;
}

.free-tier-badge {
    background: rgba(16, 185, 129, 0.15);
    border: 1px solid #10b981;
    color: #34d399;
    font-size: 11px;
    font-weight: 700;
    padding: 3px 9px;
    border-radius: 12px;
}

.trial-badge {
    background: rgba(245, 158, 11, 0.15);
    border: 1px solid #f59e0b;
    color: #fbbf24;
    font-size: 11px;
    font-weight: 700;
    padding: 3px 9px;
    border-radius: 12px;
}

.role-badge-admin {
    background: #7f1d1d;
    border: 1px solid #dc2626;
    color: #fca5a5;
    font-size: 11px;
    font-weight: 700;
    padding: 3px 9px;
    border-radius: 12px;
}

.nav-actions {
    display: flex;
    align-items: center;
    gap: 14px;
}

.btn-upgrade-plan {
    background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
    color: #0f172a;
    font-weight: 700;
    font-size: 12px;
    padding: 7px 14px;
    border-radius: 20px;
    border: none;
    cursor: pointer;
    box-shadow: 0 4px 12px rgba(245, 158, 11, 0.35);
    transition: transform 0.2s;
}
.btn-upgrade-plan:hover { transform: translateY(-1px); }

.badge-ai {
    background: rgba(56, 189, 248, 0.12);
    border: 1px solid rgba(56, 189, 248, 0.4);
    color: #38bdf8;
    font-size: 12px;
    font-weight: 600;
    padding: 6px 12px;
    border-radius: 20px;
}

.user-profile {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 13px;
    color: #cbd5e1;
    background: #1e293b;
    padding: 4px 12px;
    border-radius: 24px;
    border: 1px solid #334155;
}

.avatar {
    width: 28px;
    height: 28px;
    border-radius: 50%;
    background: #3b82f6;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    color: white;
    font-size: 12px;
}

.btn-logout {
    color: #f87171;
    text-decoration: none;
    font-size: 12px;
    font-weight: 600;
    margin-left: 6px;
}
.btn-logout:hover { text-decoration: underline; }

.container {
    max-width: 1240px;
    margin: 22px auto;
    padding: 0 16px;
}

/* NEXTDRIVE GOOGLE DRIVE STYLE SECTION */
.nextdrive-section {
    background: #111827;
    border: 1px solid #1f2937;
    border-radius: 18px;
    padding: 26px;
    margin-bottom: 26px;
    box-shadow: 0 12px 36px rgba(0, 0, 0, 0.5);
}

.drive-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 18px;
    padding-bottom: 20px;
    border-bottom: 1px solid #1f2937;
    margin-bottom: 18px;
}

.drive-title-area {
    display: flex;
    align-items: center;
    gap: 14px;
}

.drive-icon-badge {
    background: #1e293b;
    border: 1px solid #334155;
    width: 48px;
    height: 48px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 24px;
}

.drive-title-area h2 {
    margin: 0;
    font-size: 22px;
    color: #f8fafc;
    font-weight: 700;
}

.drive-title-area p {
    margin: 4px 0 0 0;
    font-size: 13px;
    color: #94a3b8;
}

.drive-controls {
    display: flex;
    align-items: center;
    gap: 12px;
    flex-wrap: wrap;
}

/* 10 GB QUOTA TRACKER */
.storage-quota-card {
    background: #1a2234;
    border: 1px solid #2d3748;
    padding: 10px 16px;
    border-radius: 12px;
    min-width: 240px;
}

.quota-labels {
    display: flex;
    justify-content: space-between;
    font-size: 12px;
    font-weight: 600;
    color: #e2e8f0;
    margin-bottom: 6px;
}

.quota-labels .used-highlight {
    color: #38bdf8;
}

.progress-track {
    background: #0f172a;
    border-radius: 999px;
    height: 8px;
    overflow: hidden;
    border: 1px solid #334155;
}

.progress-fill {
    height: 100%;
    background: linear-gradient(90deg, #10b981 0%, #38bdf8 70%, #818cf8 100%);
    border-radius: 999px;
    transition: width 0.3s ease;
}

.quota-sub {
    font-size: 11px;
    color: #94a3b8;
    margin-top: 5px;
    display: flex;
    justify-content: space-between;
}

/* BUTTONS */
.btn-action {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 10px 18px;
    border-radius: 24px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.2s ease;
    border: none;
    text-decoration: none;
}

.btn-upload-file {
    background: #2563eb;
    color: white;
    box-shadow: 0 4px 14px rgba(37, 99, 235, 0.4);
}
.btn-upload-file:hover {
    background: #1d4ed8;
    transform: translateY(-1px);
}

.btn-new-folder {
    background: #1e293b;
    color: #f8fafc;
    border: 1px solid #334155;
}
.btn-new-folder:hover {
    background: #334155;
    color: #38bdf8;
}

.btn-new-file {
    background: #1e293b;
    color: #f8fafc;
    border: 1px solid #334155;
}
.btn-new-file:hover {
    background: #334155;
    color: #34d399;
}

/* SECTION TABS (GOOGLE DRIVE FILTER TABS) */
.drive-tabs {
    display: flex;
    gap: 10px;
    margin-bottom: 18px;
    border-bottom: 1px solid #1f2937;
    padding-bottom: 12px;
    overflow-x: auto;
}

.tab-btn {
    background: #1a2234;
    border: 1px solid #2d3748;
    color: #94a3b8;
    padding: 8px 16px;
    border-radius: 20px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.2s;
    display: inline-flex;
    align-items: center;
    gap: 6px;
}

.tab-btn:hover {
    background: #243248;
    color: #f8fafc;
}

.tab-btn.active {
    background: #2563eb;
    color: white;
    border-color: #2563eb;
}

.tab-count {
    background: rgba(0,0,0,0.3);
    padding: 1px 7px;
    border-radius: 10px;
    font-size: 11px;
}

/* FILE EXPLORER TABLE */
.drive-table-container {
    overflow-x: auto;
}

.drive-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}

.drive-table th {
    text-align: left;
    padding: 14px 16px;
    color: #94a3b8;
    font-weight: 600;
    border-bottom: 1px solid #1f2937;
    text-transform: uppercase;
    font-size: 11px;
    letter-spacing: 0.5px;
}

.drive-table td {
    padding: 14px 16px;
    border-bottom: 1px solid rgba(31, 41, 55, 0.8);
    color: #cbd5e1;
}

.drive-table tr:hover {
    background: rgba(30, 41, 59, 0.45);
}

.file-name-cell {
    display: flex;
    align-items: center;
    gap: 12px;
    font-weight: 500;
    color: #f1f5f9;
}

.file-icon {
    font-size: 20px;
}

.badge-tag {
    font-size: 10px;
    font-weight: 700;
    padding: 3px 8px;
    border-radius: 6px;
    color: white;
    letter-spacing: 0.3px;
}

.action-btn {
    text-decoration: none;
    font-weight: 600;
    font-size: 12px;
    padding: 6px 12px;
    border-radius: 6px;
    display: inline-block;
    transition: background 0.15s;
}

.btn-download {
    color: #38bdf8;
    background: rgba(56, 189, 248, 0.1);
    margin-right: 8px;
}
.btn-download:hover { background: rgba(56, 189, 248, 0.2); }

.btn-delete {
    color: #f87171;
    background: rgba(248, 113, 113, 0.1);
}
.btn-delete:hover { background: rgba(248, 113, 113, 0.2); }

/* STATS & TELEMETRY */
.cards {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    margin-bottom: 22px;
}

.card, .panel {
    background: #111827;
    border: 1px solid #1f2937;
    padding: 20px;
    border-radius: 16px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.25);
}

.panel {
    margin-bottom: 24px;
}

.card h3, .panel h2 {
    margin-top: 0;
    font-size: 13px;
    color: #94a3b8;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.panel h2 {
    font-size: 17px;
    color: #f1f5f9;
    border-bottom: 1px solid #1f2937;
    padding-bottom: 12px;
    margin-bottom: 16px;
    text-transform: none;
}

.number {
    font-size: 28px;
    font-weight: 700;
    color: #f8fafc;
    margin: 8px 0;
}

.low { color: #38bdf8; }
.medium { color: #fbbf24; }
.high { color: #f87171; }

.sub-info {
    font-size: 12px;
    color: #64748b;
}

.warning {
    display: none;
    background: #7f1d1d;
    color: #fecaca;
    padding: 12px 16px;
    border-radius: 8px;
    margin-bottom: 20px;
    font-weight: 600;
    border: 1px solid #dc2626;
}

.chart-wrap {
    height: 250px;
    margin-top: 15px;
}

canvas {
    width: 100% !important;
    height: 100% !important;
}

.form-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 16px;
}

label {
    display: block;
    margin-bottom: 6px;
    font-weight: 500;
    font-size: 13px;
    color: #cbd5e1;
}

input, select {
    width: 100%;
    padding: 11px 13px;
    background: #090d16;
    border: 1px solid #374151;
    border-radius: 8px;
    color: #f8fafc;
    font-size: 14px;
}

input:focus, select:focus {
    outline: none;
    border-color: #38bdf8;
}

.btn-primary {
    margin-top: 18px;
    padding: 12px 24px;
    border: none;
    border-radius: 8px;
    background: #2563eb;
    color: white;
    font-weight: 600;
    font-size: 14px;
    cursor: pointer;
    transition: background 0.2s;
}

.btn-primary:hover {
    background: #1d4ed8;
}

.result {
    padding: 18px;
    margin-top: 20px;
    background: #090d16;
    border-left: 4px solid #38bdf8;
    border-radius: 8px;
}

.result-metrics {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 12px;
    margin-top: 12px;
    padding-top: 12px;
    border-top: 1px solid #1f2937;
}

.res-box {
    background: #111827;
    padding: 10px;
    border-radius: 8px;
    text-align: center;
    border: 1px solid #1f2937;
}

.res-box span {
    display: block;
    font-size: 11px;
    color: #94a3b8;
    text-transform: uppercase;
}

.res-box strong {
    font-size: 16px;
    color: #38bdf8;
}

.export-link {
    display: inline-block;
    margin-bottom: 16px;
    background: #1f2937;
    color: #f8fafc;
    padding: 8px 14px;
    border-radius: 6px;
    font-size: 13px;
    text-decoration: none;
    font-weight: 500;
}

.export-link:hover {
    background: #374151;
}

.note {
    color: #64748b;
    font-size: 12px;
    margin-top: 8px;
}

/* PRICING MODAL STYLES */
.modal-overlay {
    display: none;
    position: fixed;
    top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(0, 0, 0, 0.75);
    backdrop-filter: blur(6px);
    z-index: 1000;
    align-items: center;
    justify-content: center;
}

.modal-card {
    background: #111827;
    border: 1px solid #374151;
    border-radius: 20px;
    max-width: 800px;
    width: 90%;
    padding: 32px;
    box-shadow: 0 20px 50px rgba(0,0,0,0.6);
    position: relative;
}

.modal-close {
    position: absolute;
    top: 20px;
    right: 24px;
    background: #1e293b;
    border: 1px solid #334155;
    color: #94a3b8;
    width: 32px;
    height: 32px;
    border-radius: 50%;
    cursor: pointer;
    font-size: 16px;
    display: flex;
    align-items: center;
    justify-content: center;
}
.modal-close:hover { color: white; background: #334155; }

.pricing-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 20px;
    margin-top: 24px;
}

.plan-card {
    background: #1a2234;
    border: 1px solid #2d3748;
    border-radius: 16px;
    padding: 24px;
    text-align: center;
    position: relative;
    transition: transform 0.2s, border-color 0.2s;
}
.plan-card:hover {
    transform: translateY(-4px);
    border-color: #38bdf8;
}
.plan-card.featured {
    border: 2px solid #f59e0b;
    background: #1f2a40;
}
.plan-ribbon {
    position: absolute;
    top: -12px;
    left: 50%;
    transform: translateX(-50%);
    background: #f59e0b;
    color: #0f172a;
    font-size: 10px;
    font-weight: 800;
    padding: 3px 12px;
    border-radius: 12px;
    text-transform: uppercase;
}

.plan-price {
    font-size: 32px;
    font-weight: 800;
    color: #f8fafc;
    margin: 12px 0 6px 0;
}
.plan-price span {
    font-size: 14px;
    font-weight: 500;
    color: #94a3b8;
}

.plan-features {
    list-style: none;
    padding: 0;
    margin: 18px 0;
    text-align: left;
    font-size: 12px;
    color: #cbd5e1;
}
.plan-features li {
    margin-bottom: 8px;
    display: flex;
    align-items: center;
    gap: 8px;
}

.btn-subscribe {
    width: 100%;
    padding: 10px;
    border-radius: 8px;
    font-weight: 700;
    border: none;
    cursor: pointer;
    font-size: 13px;
}
.btn-sub-free { background: #334155; color: #f8fafc; }
.btn-sub-monthly { background: #2563eb; color: white; }
.btn-sub-yearly { background: #f59e0b; color: #0f172a; }

footer {
    text-align: center;
    padding: 26px;
    color: #64748b;
    font-size: 13px;
    border-top: 1px solid #111827;
}
</style>
</head>

<body>

<!-- NAVIGATION BAR -->
<nav class="navbar">
    <div class="brand">
        <div class="drive-logo">▲</div>
        <h1>Next<span>Drive</span></h1>
        <span class="free-tier-badge">10 GB FREE</span>
        <span class="trial-badge">🎉 3 MONTHS FREE TRIAL</span>
        {% if current_user.role == 'admin' %}
        <span class="role-badge-admin">⚡ ADMIN CONSOLE</span>
        {% endif %}
    </div>
    <div class="nav-actions">
        <button class="btn-upgrade-plan" onclick="openPricingModal()">⭐ Upgrade Plan</button>
        <span class="badge-ai">🤖 Random Forest Model (96.5% Acc)</span>
        <div class="user-profile">
            <div class="avatar">{{ current_user.username[0].upper() }}</div>
            <span><strong>{{ current_user.username }}</strong> ({{ current_user.role }})</span>
            <a href="/logout" class="btn-logout" title="Sign out">Logout ⎋</a>
        </div>
    </div>
</nav>

<main class="container">

    <div id="usageWarning" class="warning" role="alert">
        CRITICAL ALERT: System Telemetry shows resources exceeding 80% saturation threshold!
    </div>

    <!-- ADMIN USER MANAGEMENT PANEL (ONLY VISIBLE TO ADMIN) -->
    {% if current_user.role == 'admin' %}
    <section class="panel" style="border-left: 4px solid #ef4444;">
        <h2>🛡️ Administrator Management Console</h2>
        <p style="color: #94a3b8; font-size: 13px;">Manage connected tenant accounts, user storage allocations, and cloud subscription tiers.</p>
        <div style="overflow-x: auto;">
            <table class="drive-table">
                <thead>
                    <tr>
                        <th>User ID</th>
                        <th>Username</th>
                        <th>Role</th>
                        <th>Full Name</th>
                        <th>Active Cloud Subscription</th>
                        <th>Access Level</th>
                    </tr>
                </thead>
                <tbody>
                    {% for u in users_list %}
                    <tr>
                        <td>#{{ u.id }}</td>
                        <td><strong>{{ u.username }}</strong></td>
                        <td>
                            <span class="badge-tag" style="background: {% if u.role == 'admin' %}#ef4444{% else %}#10b981{% endif %};">
                                {{ u.role|upper }}
                            </span>
                        </td>
                        <td>{{ u.full_name }}</td>
                        <td>{{ u.plan }}</td>
                        <td>{% if u.role == 'admin' %}Full Root Control{% else %}Standard 10 GB Storage{% endif %}</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </section>
    {% endif %}

    <!-- 1. NEXTDRIVE GOOGLE DRIVE STYLE STORAGE AT THE VERY TOP (10 GB FREE) -->
    <section class="nextdrive-section">
        <div class="drive-header">
            <div class="drive-title-area">
                <div class="drive-icon-badge">☁️</div>
                <div>
                    <h2>My NextDrive Cloud Storage</h2>
                    <p>High-Speed Cloud Storage Bucket & Elastic File System • <strong>Trial: 90 Days Remaining Free</strong></p>
                </div>
            </div>

            <div class="drive-controls">
                <!-- 10 GB Quota Tracker -->
                <div class="storage-quota-card">
                    <div class="quota-labels">
                        <span>Free Storage Used</span>
                        <span class="used-highlight">{{ nextdrive.used_mb }} MB / {{ nextdrive.quota_gb }} GB</span>
                    </div>
                    <div class="progress-track">
                        <div class="progress-fill" style="width: {{ nextdrive.used_percent }}%;"></div>
                    </div>
                    <div class="quota-sub">
                        <span>{{ nextdrive.free_gb }} GB Free</span>
                        <span>{{ nextdrive.used_percent }}% Used</span>
                    </div>
                </div>

                <!-- Google Drive Style Upload Button Form -->
                <form action="/storage/upload" method="POST" enctype="multipart/form-data" style="display:inline-block;">
                    <input type="file" id="fileUploadInput" name="file" required style="display:none;" onchange="this.form.submit()">
                    <button type="button" class="btn-action btn-upload-file" onclick="document.getElementById('fileUploadInput').click()">
                        <span>➕</span> Upload File
                    </button>
                </form>

                <!-- New Folder Button -->
                <button type="button" class="btn-action btn-new-folder" onclick="promptNewFolder()">
                    <span>📁</span> New Folder
                </button>

                <!-- New File Generator Button -->
                <button type="button" class="btn-action btn-new-file" onclick="promptNewFile()">
                    <span>📝</span> New Doc / Note
                </button>
            </div>
        </div>

        <!-- DRIVE FILTER TABS (VIDEOS, IMAGES, DOCUMENTS, FOLDERS, ALL) -->
        <div class="drive-tabs">
            <button class="tab-btn active" onclick="filterDrive('all', this)">
                📁 All Items <span class="tab-count">{{ nextdrive.counts.all }}</span>
            </button>
            <button class="tab-btn" onclick="filterDrive('images', this)">
                🖼️ Images <span class="tab-count">{{ nextdrive.counts.images }}</span>
            </button>
            <button class="tab-btn" onclick="filterDrive('videos', this)">
                🎬 Videos & Media <span class="tab-count">{{ nextdrive.counts.videos }}</span>
            </button>
            <button class="tab-btn" onclick="filterDrive('documents', this)">
                📄 Documents & Sheets <span class="tab-count">{{ nextdrive.counts.documents }}</span>
            </button>
            <button class="tab-btn" onclick="filterDrive('folders', this)">
                🗂️ Folders <span class="tab-count">{{ nextdrive.counts.folders }}</span>
            </button>
        </div>

        <!-- Google Drive Style File Explorer Table -->
        <div class="drive-table-container">
            <table class="drive-table">
                <thead>
                    <tr>
                        <th>Item Name</th>
                        <th>Type</th>
                        <th>Size</th>
                        <th>Uploaded / Modified</th>
                        <th style="text-align: right;">Actions</th>
                    </tr>
                </thead>
                <tbody id="driveTableBody">
                    {% for item in nextdrive.file_list %}
                    <tr class="drive-row" data-tab="{{ item.tab }}">
                        <td>
                            <div class="file-name-cell">
                                <span class="file-icon">{{ item.icon }}</span>
                                <span>{{ item.name }}</span>
                            </div>
                        </td>
                        <td>
                            <span class="badge-tag" style="background: {{ item.badge_color }};">
                                {{ item.category }}
                            </span>
                        </td>
                        <td>{{ item.size_str }}</td>
                        <td>{{ item.modified }}</td>
                        <td style="text-align: right;">
                            {% if not item.is_dir %}
                            <a href="/storage/download/{{ item.name }}" class="action-btn btn-download">⬇ Download</a>
                            {% endif %}
                            <a href="/storage/delete/{{ item.name }}" class="action-btn btn-delete">🗑 Delete</a>
                        </td>
                    </tr>
                    {% else %}
                    <tr>
                        <td colspan="5" style="text-align: center; padding: 36px; color: #64748b;">
                            Your 10 GB NextDrive is ready! Click <strong>Upload File</strong> or <strong>New Folder</strong> to start.
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </section>

    <!-- 2. TELEMETRY & CLUSTER METRICS -->
    <section class="cards">
        <div class="card">
            <h3>Total Allocations</h3>
            <div class="number">{{ summary.total }}</div>
            <div class="sub-info">History Records Logged</div>
        </div>

        <div class="card">
            <h3>Host CPU Utilization</h3>
            <div id="liveCpu" class="number">{{ cpu_usage }}%</div>
            <div class="sub-info">Host Processor Telemetry</div>
        </div>

        <div class="card">
            <h3>Host RAM Utilization</h3>
            <div id="liveRam" class="number">{{ ram_usage }}%</div>
            <div class="sub-info">Physical Memory Load</div>
        </div>

        <div class="card">
            <h3>Physical Server Disk</h3>
            <div id="liveStorage" class="number">{{ storage_usage }}%</div>
            <div id="liveStorageDetail" class="sub-info">{{ storage_used_gb }} GB / {{ storage_total_gb }} GB Used</div>
        </div>
    </section>

    <!-- 3. LIVE MULTI-RESOURCE GRAPH -->
    <section class="panel">
        <h2>Live Multi-Resource Saturation Graph</h2>
        <div class="chart-wrap">
            <canvas id="usageChart"></canvas>
        </div>
        <p class="note" id="lastUpdated">Streaming live telemetry...</p>
    </section>

    <!-- 4. ML RESOURCE ALLOCATION SIMULATOR -->
    <section class="panel">
        <h2>Simulate Workload & Predict Allocation (ML Classifier)</h2>
        <p style="color: #94a3b8; font-size: 14px; margin-top: -6px;">
            Input anticipated workload demands. The trained Machine Learning algorithm will predict the optimal scaling tier, VM cluster size, and hourly operating cost.
        </p>

        <form action="/allocate" method="POST">
            <div class="form-grid">
                <div>
                    <label for="cpu">Projected CPU Demand (%)</label>
                    <input type="number" id="cpu" name="cpu" min="1" max="100" value="65" required>
                </div>
                <div>
                    <label for="ram">Projected RAM Demand (%)</label>
                    <input type="number" id="ram" name="ram" min="1" max="100" value="70" required>
                </div>
                <div>
                    <label for="storage">NextDrive Storage Demand (%)</label>
                    <input type="number" id="storage" name="storage" min="1" max="100" value="55" required>
                </div>
                <div>
                    <label for="tasks">Active Concurrent Tasks</label>
                    <input type="number" id="tasks" name="tasks" min="1" max="500" value="60" required>
                </div>
                <div>
                    <label for="network_io">Network I/O Throughput (Mbps)</label>
                    <input type="number" id="network_io" name="network_io" min="1" max="2000" value="350" required>
                </div>
            </div>
            <button type="submit" class="btn-primary">⚡ Run ML Inference & Allocate Resources</button>
        </form>

        {% if result %}
        <div class="result">
            <h3 style="margin-top:0; color:#f8fafc;">
                Allocation Result: 
                <span class="{{ result.demand_class }}">
                    {{ result.demand_status }} DEMAND
                </span>
            </h3>
            <p style="color: #cbd5e1; margin-bottom: 8px;">{{ result.recommendation }}</p>
            
            <div class="result-metrics">
                <div class="res-box">
                    <span>Algorithm</span>
                    <strong>Random Forest</strong>
                </div>
                <div class="res-box">
                    <span>Provisioned VMs</span>
                    <strong>{{ result.vms }} Instance(s)</strong>
                </div>
                <div class="res-box">
                    <span>Hourly Cost</span>
                    <strong>${{ result.cost }}/hr</strong>
                </div>
                <div class="res-box">
                    <span>Storage Tier</span>
                    <strong>NextDrive Auto-Scale SSD</strong>
                </div>
            </div>
        </div>
        {% endif %}
    </section>

    <!-- 5. HISTORICAL RECORDS TABLE -->
    <section class="panel">
        <h2>Persistent Allocation Log</h2>
        <a class="export-link" href="/export">📥 Export Allocation Dataset (CSV)</a>

        <div style="overflow-x: auto;">
            <table class="drive-table">
                <thead>
                    <tr>
                        <th>Timestamp</th>
                        <th>CPU / RAM / Storage</th>
                        <th>Actual Usage</th>
                        <th>ML Decision</th>
                        <th>VM Cluster</th>
                        <th>Est. Cost</th>
                        <th>Recommendation Details</th>
                    </tr>
                </thead>
                <tbody>
                    {% for row in history %}
                    <tr>
                        <td>{{ row.time }}</td>
                        <td>{{ row.cpu }}% / {{ row.ram }}% / {{ row.storage }}%</td>
                        <td>{{ row.cpu_usage }}% / {{ row.ram_usage }}% / {{ row.storage_usage }}%</td>
                        <td>
                            <span class="{{ row.demand_class }}">
                                {{ row.demand_status }}
                            </span>
                        </td>
                        <td>{{ row.allocated_vms }} VM(s)</td>
                        <td>${{ row.est_cost_hr }}/hr</td>
                        <td style="max-width:320px; font-size:12px; color:#cbd5e1;">{{ row.recommendation }}</td>
                    </tr>
                    {% else %}
                    <tr>
                        <td colspan="7">No simulation logs found yet. Run an allocation above to record history.</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </section>

</main>

<!-- PRICING & SUBSCRIPTION MODAL -->
<div id="pricingModal" class="modal-overlay" onclick="closePricingModal(event)">
    <div class="modal-card" onclick="event.stopPropagation()">
        <button class="modal-close" onclick="closePricingModal()">&times;</button>
        <div style="text-align: center;">
            <span class="trial-badge">🎉 3 MONTHS COMPLIMENTARY TRIAL</span>
            <h2 style="font-size: 26px; margin: 10px 0 6px 0; color: #f8fafc;">NextDrive Cloud Storage Plans</h2>
            <p style="color: #94a3b8; font-size: 14px; margin: 0;">Enjoy 100% free access for the first 3 months. Choose your continuing plan below.</p>
        </div>

        <div class="pricing-grid">
            <!-- Free Trial Plan -->
            <div class="plan-card">
                <h3 style="color: #38bdf8; margin: 0 0 6px 0;">3-Month Trial</h3>
                <div class="plan-price">₹0 <span>/ 3 months</span></div>
                <p style="font-size: 11px; color: #34d399; margin: 0;">Active for all new users</p>
                <ul class="plan-features">
                    <li>✅ 10 GB High-Speed Storage</li>
                    <li>✅ Video & Image Media Hub</li>
                    <li>✅ Unlimited Upload & Download</li>
                    <li>✅ AI Resource Autoscaler</li>
                </ul>
                <button class="btn-subscribe btn-sub-free" disabled>Current Active Plan</button>
            </div>

            <!-- Monthly Plan: ₹199 / Month -->
            <div class="plan-card">
                <h3 style="color: #f8fafc; margin: 0 0 6px 0;">Pro Monthly</h3>
                <div class="plan-price">₹199 <span>/ month</span></div>
                <p style="font-size: 11px; color: #94a3b8; margin: 0;">Billed monthly after 3 months</p>
                <ul class="plan-features">
                    <li>✅ 50 GB Elastic Cloud Storage</li>
                    <li>✅ High-Speed Video Streaming</li>
                    <li>✅ Multi-VM Cluster Priority</li>
                    <li>✅ 24/7 Cloud Backup & Sync</li>
                </ul>
                <button class="btn-subscribe btn-sub-monthly" onclick="alert('Thank you! ₹199/month subscription activated starting after your 3-month free trial.')">Subscribe ₹199/mo</button>
            </div>

            <!-- Yearly Plan: ₹1999 / Year -->
            <div class="plan-card featured">
                <div class="plan-ribbon">Best Value (Save ₹389)</div>
                <h3 style="color: #f59e0b; margin: 0 0 6px 0;">Pro Annual</h3>
                <div class="plan-price">₹1,999 <span>/ year</span></div>
                <p style="font-size: 11px; color: #34d399; margin: 0;">Effectively ₹166/month</p>
                <ul class="plan-features">
                    <li>✅ 100 GB High-Performance Storage</li>
                    <li>✅ Dedicated Autoscaling Instance</li>
                    <li>✅ Enterprise 99.99% Uptime SLA</li>
                    <li>✅ Priority Student & Dev Support</li>
                </ul>
                <button class="btn-subscribe btn-sub-yearly" onclick="alert('Thank you! ₹1,999/year subscription activated starting after your 3-month free trial.')">Subscribe ₹1,999/yr</button>
            </div>
        </div>
    </div>
</div>

<footer>
    NextDrive: Intelligent Cloud Resource Allocation & 10 GB Storage System &copy; 2026
</footer>

<script>
// Modal Controller
function openPricingModal() {
    document.getElementById('pricingModal').style.display = 'flex';
}

function closePricingModal(event) {
    document.getElementById('pricingModal').style.display = 'none';
}

// Filter Drive Items by Tab
function filterDrive(tab, btn) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');

    const rows = document.querySelectorAll('.drive-row');
    rows.forEach(r => {
        if (tab === 'all' || r.getAttribute('data-tab') === tab) {
            r.style.display = '';
        } else {
            r.style.display = 'none';
        }
    });
}

// Prompt to create New Folder
function promptNewFolder() {
    const name = prompt("Enter new folder name:", "New Project Folder");
    if (name && name.trim()) {
        const form = document.createElement('form');
        form.method = 'POST';
        form.action = '/storage/create-folder';
        const input = document.createElement('input');
        input.type = 'hidden';
        input.name = 'folder_name';
        input.value = name.trim();
        form.appendChild(input);
        document.body.appendChild(form);
        form.submit();
    }
}

// Prompt to create New File / Note
function promptNewFile() {
    const name = prompt("Enter new file name with extension (e.g. notes.txt, document.doc, test.py):", "notes.txt");
    if (name && name.trim()) {
        const form = document.createElement('form');
        form.method = 'POST';
        form.action = '/storage/create-file';
        const input = document.createElement('input');
        input.type = 'hidden';
        input.name = 'file_name';
        input.value = name.trim();
        form.appendChild(input);
        document.body.appendChild(form);
        form.submit();
    }
}

const canvas = document.getElementById('usageChart');
const ctx = canvas.getContext('2d');
const warningBox = document.getElementById('usageWarning');
const lastUpdatedEl = document.getElementById('lastUpdated');

const liveCpu = document.getElementById('liveCpu');
const liveRam = document.getElementById('liveRam');
const liveStorage = document.getElementById('liveStorage');
const liveStorageDetail = document.getElementById('liveStorageDetail');

let historyData = {{ usage_history|tojson }};

function drawChart() {
    const width = canvas.parentElement.clientWidth;
    const height = 250;
    canvas.width = width;
    canvas.height = height;

    ctx.clearRect(0, 0, width, height);

    ctx.strokeStyle = '#1f2937';
    ctx.lineWidth = 1;
    for (let i = 0; i <= 4; i++) {
        const y = 30 + (i * 45);
        ctx.beginPath();
        ctx.moveTo(40, y);
        ctx.lineTo(width - 20, y);
        ctx.stroke();

        ctx.fillStyle = '#64748b';
        ctx.font = '11px sans-serif';
        ctx.fillText((100 - i * 25) + '%', 10, y + 4);
    }

    if (historyData.length < 2) return;

    const maxPoints = 20;
    const stepX = (width - 70) / (maxPoints - 1);

    function plotLine(key, color) {
        ctx.strokeStyle = color;
        ctx.lineWidth = 2.5;
        ctx.beginPath();

        historyData.forEach((pt, idx) => {
            const x = 50 + (idx * stepX);
            const val = pt[key] || 0;
            const y = 210 - (val / 100 * 180);

            if (idx === 0) ctx.moveTo(x, y);
            else ctx.lineTo(x, y);
        });
        ctx.stroke();
    }

    plotLine('cpu', '#38bdf8');      // Sky Blue: CPU
    plotLine('ram', '#4ade80');      // Neon Green: RAM
    plotLine('storage', '#c084fc');  // Purple: Storage

    ctx.font = '12px sans-serif';
    ctx.fillStyle = '#38bdf8';
    ctx.fillText('● CPU (%)', width - 240, 20);
    ctx.fillStyle = '#4ade80';
    ctx.fillText('● RAM (%)', width - 160, 20);
    ctx.fillStyle = '#c084fc';
    ctx.fillText('● Storage (%)', width - 75, 20);
}

function fetchTelemetry() {
    fetch('/api/usage')
        .then(res => res.json())
        .then(data => {
            historyData.push(data);
            if (historyData.length > 20) historyData.shift();

            liveCpu.textContent = data.cpu + '%';
            liveRam.textContent = data.ram + '%';
            liveStorage.textContent = data.storage + '%';
            if (data.storage_used_gb) {
                liveStorageDetail.textContent = data.storage_used_gb + ' GB / ' + data.storage_total_gb + ' GB Used';
            }

            if (data.cpu >= 80 || data.ram >= 80 || data.storage >= 80) {
                warningBox.style.display = 'block';
            } else {
                warningBox.style.display = 'none';
            }

            lastUpdatedEl.textContent = 'Telemetry synced at ' + data.time;
            drawChart();
        })
        .catch(err => console.log('Telemetry sync error:', err));
}

window.addEventListener('resize', drawChart);
drawChart();
setInterval(fetchTelemetry, 3000);
</script>
</body>
</html>
"""


# ---------------- ROUTES & AUTHENTICATION ----------------

def get_current_user():
    if "user_id" in session:
        with sqlite3.connect(DATABASE) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT * FROM users WHERE id = ?", (session["user_id"],)).fetchone()
            if row:
                return dict(row)
    return None


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        with sqlite3.connect(DATABASE) as conn:
            conn.row_factory = sqlite3.Row
            user = conn.execute(
                "SELECT * FROM users WHERE username = ? AND password = ?",
                (username, password)
            ).fetchone()

        if user:
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            return redirect(url_for("home"))
        else:
            return render_template_string(LOGIN_HTML, error="Invalid Username or Password. Please try again.")

    return render_template_string(LOGIN_HTML)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def render_dashboard(result=None):
    current_user = get_current_user()
    if not current_user:
        return redirect(url_for("login"))

    cpu_usage, ram_usage, storage_usage, storage_total_gb, storage_used_gb, storage_free_gb = get_usage()
    record_usage(cpu_usage, ram_usage, storage_usage)
    nextdrive_info = get_nextdrive_info()
    users_list = get_users_list() if current_user["role"] == "admin" else []

    return render_template_string(
        HTML,
        current_user=current_user,
        users_list=users_list,
        cpu_usage=cpu_usage,
        ram_usage=ram_usage,
        storage_usage=storage_usage,
        storage_total_gb=storage_total_gb,
        storage_used_gb=storage_used_gb,
        storage_free_gb=storage_free_gb,
        nextdrive=nextdrive_info,
        history=get_history(),
        usage_history=usage_history,
        summary=get_summary(),
        result=result
    )


@app.route("/")
def home():
    if "user_id" not in session:
        return redirect(url_for("login"))
    return render_dashboard()


@app.route("/storage/upload", methods=["POST"])
def upload_file():
    if "user_id" not in session:
        return redirect(url_for("login"))
    if "file" in request.files:
        f = request.files["file"]
        if f.filename:
            f.save(os.path.join(STORAGE_DIR, f.filename))
    return redirect(url_for("home"))


@app.route("/storage/create-folder", methods=["POST"])
def create_folder():
    if "user_id" not in session:
        return redirect(url_for("login"))
    folder_name = request.form.get("folder_name", "").strip()
    if folder_name:
        folder_path = os.path.join(STORAGE_DIR, folder_name)
        os.makedirs(folder_path, exist_ok=True)
    return redirect(url_for("home"))


@app.route("/storage/create-file", methods=["POST"])
def create_file():
    if "user_id" not in session:
        return redirect(url_for("login"))
    file_name = request.form.get("file_name", "").strip()
    if file_name:
        file_path = os.path.join(STORAGE_DIR, file_name)
        if not os.path.exists(file_path):
            with open(file_path, "w") as f:
                f.write(f"Created by {session.get('username', 'user')} on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} in NextDrive Cloud Storage.\n")
    return redirect(url_for("home"))


@app.route("/storage/download/<path:filename>")
def download_file(filename):
    if "user_id" not in session:
        return redirect(url_for("login"))
    return send_from_directory(STORAGE_DIR, filename, as_attachment=True)


@app.route("/storage/delete/<path:filename>")
def delete_file(filename):
    if "user_id" not in session:
        return redirect(url_for("login"))
    fpath = os.path.join(STORAGE_DIR, filename)
    if os.path.exists(fpath):
        if os.path.isdir(fpath):
            shutil.rmtree(fpath)
        else:
            os.remove(fpath)
    return redirect(url_for("home"))


@app.route("/allocate", methods=["POST"])
def allocate():
    if "user_id" not in session:
        return redirect(url_for("login"))
    try:
        cpu = int(request.form.get("cpu", "50"))
        ram = int(request.form.get("ram", "50"))
        storage = int(request.form.get("storage", "50"))
        tasks = int(request.form.get("tasks", "40"))
        network_io = float(request.form.get("network_io", "250.0"))
    except (TypeError, ValueError):
        return "Please enter valid numerical input values.", 400

    cpu_usage, ram_usage, storage_usage, _, _, _ = get_usage()
    record_usage(cpu_usage, ram_usage, storage_usage)

    status, demand_class, recommendation, ml_str, vms, hourly_cost = assess_demand(
        cpu, ram, storage, tasks, network_io
    )

    with sqlite3.connect(DATABASE) as conn:
        conn.execute("""
            INSERT INTO allocation_history (
                time, cpu, ram, storage, cpu_usage, ram_usage, storage_usage,
                demand_status, demand_class, recommendation,
                ml_prediction, allocated_vms, est_cost_hr
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            cpu,
            ram,
            storage,
            cpu_usage,
            ram_usage,
            storage_usage,
            status,
            demand_class,
            recommendation,
            ml_str,
            vms,
            hourly_cost
        ))

    result = {
        "demand_status": status,
        "demand_class": demand_class,
        "recommendation": recommendation,
        "vms": vms,
        "cost": hourly_cost
    }

    return render_dashboard(result)


@app.route("/api/usage")
def api_usage():
    cpu_usage, ram_usage, storage_usage, storage_total_gb, storage_used_gb, storage_free_gb = get_usage()
    record_usage(cpu_usage, ram_usage, storage_usage)

    return jsonify({
        "time": datetime.now().strftime("%H:%M:%S"),
        "cpu": cpu_usage,
        "ram": ram_usage,
        "storage": storage_usage,
        "storage_total_gb": storage_total_gb,
        "storage_used_gb": storage_used_gb,
        "storage_free_gb": storage_free_gb
    })


@app.route("/export")
def export_csv():
    if "user_id" not in session:
        return redirect(url_for("login"))
    rows = get_history()
    output = StringIO()
    writer = csv.writer(output)

    writer.writerow([
        "Timestamp", "CPU Demand", "RAM Demand", "Storage Demand",
        "Actual CPU Usage", "Actual RAM Usage", "Actual Storage Usage",
        "Demand Status", "Allocated VMs", "Est Hourly Cost", "Recommendation"
    ])

    for item in rows:
        writer.writerow([
            item["time"],
            item["cpu"],
            item["ram"],
            item.get("storage", 50),
            item["cpu_usage"],
            item["ram_usage"],
            item.get("storage_usage", 0),
            item["demand_status"],
            item.get("allocated_vms", 1),
            item.get("est_cost_hr", 0.05),
            item["recommendation"]
        ])

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={
            "Content-Disposition":
                "attachment; filename=cloud_resource_allocation_history.csv"
        }
    )


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
