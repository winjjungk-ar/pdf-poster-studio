CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS recent_projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_path TEXT NOT NULL,
    opened_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS project_settings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_name TEXT NOT NULL,
    pdf_path TEXT,
    output_mode TEXT DEFAULT 'no_cut',
    split_layout TEXT DEFAULT 'left_right',
    scale_percent REAL DEFAULT 96.0,
    overlap_mm REAL DEFAULT 5.0,
    safe_margin_mm REAL DEFAULT 5.0,
    add_cut_guides INTEGER DEFAULT 1,
    add_tile_numbers INTEGER DEFAULT 1,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS patch_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    version TEXT NOT NULL,
    title TEXT NOT NULL,
    applied_at TEXT NOT NULL,
    status TEXT NOT NULL,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_time TEXT NOT NULL,
    event_type TEXT NOT NULL,
    message TEXT NOT NULL
);