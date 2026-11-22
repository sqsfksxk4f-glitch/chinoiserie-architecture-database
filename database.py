import sqlite3
import json
import os
import uuid
import shutil
import copy
import pandas as pd

DB_DIR = os.path.join(os.path.dirname(__file__), "data")
DB_PATH = os.path.join(DB_DIR, "architecture.db")
IMAGE_DIR = os.path.join(DB_DIR, "images")
IMAGE_TRASH_DIR = os.path.join(IMAGE_DIR, "trash")


def _ensure_image_dirs():
    os.makedirs(IMAGE_DIR, exist_ok=True)
    os.makedirs(IMAGE_TRASH_DIR, exist_ok=True)

BUILDING_TYPES = ["园林", "民居", "寺庙", "宫殿", "衙署", "塔", "桥", "祠堂", "书院", "现代中式", "其他"]

FIELDS = [
    ("name", "建筑名称（中文）", "TEXT NOT NULL"),
    ("name_en", "英文名称", "TEXT"),
    ("name_official", "官方语言名称", "TEXT"),
    ("country", "国家", "TEXT"),
    ("city", "城市", "TEXT"),
    ("year_built", "建造时间", "TEXT"),
    ("demolition_year", "拆除时间", "TEXT"),
    ("restoration_year", "修复时间", "TEXT"),
    ("quantity", "数量规模", "INTEGER DEFAULT 1"),
    ("building_type", "功能", "TEXT"),
    ("architect", "设计师", "TEXT"),
    ("latitude", "纬度", "REAL"),
    ("longitude", "经度", "REAL"),
    ("material", "结构材料", "TEXT"),
    ("area_sqm", "建筑面积(㎡)", "REAL"),
    ("protection_level", "保护级别", "TEXT"),
    ("description", "特征描述", "TEXT"),
    ("image_path", "图片路径", "TEXT"),
    ("image_urls", "建筑图片链接", "TEXT"),
    ("search_time", "检索时间", "TEXT"),
    ("search_keywords", "检索词", "TEXT"),
    ("condition", "建筑状态", "TEXT"),
    ("spread_method", "传播方式", "TEXT"),
    ("notes", "备注", "TEXT"),
]

CONDITION_OPTIONS = ["良好", "破损", "倒塌"]

FIELD_NAMES = [f[0] for f in FIELDS]
FIELD_LABELS = [f[1] for f in FIELDS]


def get_connection():
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db():
    conn = get_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS cases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            name_en TEXT,
            name_official TEXT,
            country TEXT,
            city TEXT,
            year_built TEXT,
            demolition_year TEXT,
            restoration_year TEXT,
            quantity INTEGER DEFAULT 1,
            building_type TEXT,
            architect TEXT,
            latitude REAL,
            longitude REAL,
            material TEXT,
            area_sqm REAL,
            protection_level TEXT,
            description TEXT,
            image_path TEXT,
            image_urls TEXT,
            search_time TEXT,
            search_keywords TEXT,
            condition TEXT,
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS building_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type_name TEXT UNIQUE
        )
    """)
    for bt in BUILDING_TYPES:
        conn.execute("INSERT OR IGNORE INTO building_types (type_name) VALUES (?)", (bt,))
    conn.commit()
    conn.close()
    migrate_db_v2()
    migrate_db_v3()
    migrate_db_v4()
    migrate_db_v5()
    migrate_db_v6()
    migrate_db_v7()
    migrate_db_v8()
    migrate_db_v9()
    migrate_db_v10()
    migrate_db_v11()
    migrate_db_v12()


def migrate_db_v10():
    """Add screen_engine column to scholar_search_results."""
    conn = get_connection()
    try:
        conn.execute("ALTER TABLE scholar_search_results ADD COLUMN screen_engine TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()


def migrate_db_v11():
    """Create verified_cases table (same schema as raw_cases)."""
    conn = get_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS verified_cases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            raw_case_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            name_en TEXT,
            name_official TEXT,
            country TEXT,
            city TEXT,
            construction_year TEXT,
            demolition_year TEXT,
            restoration_year TEXT,
            quantity INTEGER DEFAULT 1,
            function TEXT,
            designer TEXT,
            latitude REAL,
            longitude REAL,
            search_time TEXT,
            search_keywords TEXT,
            image_urls TEXT,
            source_urls TEXT,
            authenticity_status TEXT DEFAULT 'raw',
            authenticity_score REAL,
            authenticity_notes TEXT,
            building_status TEXT,
            image_path TEXT,
            spread_method TEXT,
            image_metadata TEXT,
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()


def migrate_db_v12():
    """Add tags column to scholar_search_results."""
    conn = get_connection()
    try:
        conn.execute("ALTER TABLE scholar_search_results ADD COLUMN tags TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()


def update_scholar_result_tags(result_ids, tags):
    """Set tags on scholar search results by ID."""
    conn = get_connection()
    if isinstance(result_ids, int):
        result_ids = [result_ids]
    for rid in result_ids:
        conn.execute("UPDATE scholar_search_results SET tags=? WHERE id=?", (tags, rid))
    conn.commit()
    conn.close()


def mark_scholar_results_tags(links, titles, tag):
    """Batch update tags for scholar results by link/title."""
    if not links and not titles:
        return
    conn = get_connection()
    for link in links:
        conn.execute("UPDATE scholar_search_results SET tags=? WHERE link=?", (tag, link))
    for title in titles:
        conn.execute("UPDATE scholar_search_results SET tags=? WHERE title=?", (tag, title))
    conn.commit()
    conn.close()


def get_all_cases(search=None, filters=None, order_by=None):
    conn = get_connection()
    query = "SELECT * FROM cases WHERE 1=1"
    params = []
    if search:
        query += " AND (name LIKE ? OR name_en LIKE ? OR name_official LIKE ? OR country LIKE ? OR city LIKE ? OR architect LIKE ? OR description LIKE ?)"
        s = f"%{search}%"
        params.extend([s, s, s, s, s, s, s])
    if filters:
        for field, value in filters.items():
            if value:
                query += f" AND {field} = ?"
                params.append(value)
    if order_by:
        field, direction = order_by
        if field in FIELD_NAMES:
            query += f" ORDER BY {field} {direction}"
        else:
            query += " ORDER BY updated_at DESC"
    else:
        query += " ORDER BY sort_order ASC, updated_at DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_case_by_id(case_id):
    conn = get_connection()
    row = conn.execute("SELECT * FROM cases WHERE id = ?", (case_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def insert_case(data):
    conn = get_connection()
    fields = [f for f in FIELD_NAMES if f in data]
    placeholders = ", ".join(["?" for _ in fields])
    columns = ", ".join(fields)
    values = [data[f] for f in fields]
    conn.execute(f"INSERT INTO cases ({columns}) VALUES ({placeholders})", values)
    conn.commit()
    new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.close()
    return new_id


def update_case(case_id, data):
    conn = get_connection()
    sets = ", ".join([f"{k} = ?" for k in data])
    values = list(data.values()) + [case_id]
    conn.execute(f"UPDATE cases SET {sets}, updated_at = CURRENT_TIMESTAMP WHERE id = ?", values)
    conn.commit()
    conn.close()


def delete_case(case_id):
    conn = get_connection()
    conn.execute("DELETE FROM cases WHERE id = ?", (case_id,))
    conn.commit()
    conn.close()


def import_from_dataframe(df):
    conn = get_connection()
    col_map = {label: name for name, label, *_ in FIELDS}
    imported = 0
    for _, row in df.iterrows():
        data = {}
        for df_col in df.columns:
            if df_col in col_map:
                data[col_map[df_col]] = str(row[df_col]) if pd.notna(row[df_col]) else None
            elif df_col in FIELD_NAMES:
                data[df_col] = str(row[df_col]) if pd.notna(row[df_col]) else None
        if data.get("name"):
            insert_case(data)
            imported += 1
    conn.close()
    return imported


def _parse_coordinates(val):
    """Parse a coordinate string into (latitude, longitude).
    Handles: '53.9,27.5', '53.9°N, 27.5°E', '53.9, 27.5', '53.9/27.5'.
    Returns (lat, lng) or (None, None).
    """
    import re
    # Remove degree symbols and N/S/E/W markers
    cleaned = val.replace("°", " ").replace("N", "").replace("S", "-").replace("E", "").replace("W", "-")
    # Try to split by common separators
    parts = re.split(r"[,，/\s]+", cleaned)
    parts = [p.strip() for p in parts if p.strip()]
    nums = []
    for p in parts:
        try:
            nums.append(float(p))
        except ValueError:
            continue
    if len(nums) >= 2:
        return nums[0], nums[1]
    if len(nums) == 1:
        return nums[0], None
    return None, None


def import_raw_from_dataframe(df, column_mapping, image_data=None,
                               merge_extra_to_notes=False, update_existing=False):
    """Import DataFrame rows into raw_cases table with flexible column mapping.

    Args:
        df: pandas DataFrame
        column_mapping: dict {excel_col: db_field or None or '__NOTES__'}
        image_data: dict {row_index: [image_bytes, ...]} from Excel extraction
        merge_extra_to_notes: if True, merge unmapped column values into authenticity_notes
        update_existing: if True, update existing cases matched by name instead of creating duplicates

    Returns:
        (imported_count, updated_count, skipped_count, error_list)
    """
    name_col = None
    source_cols = []
    notes_cols = []
    for excel_col, db_field in column_mapping.items():
        if db_field == "name":
            name_col = excel_col
        elif db_field == "source_urls":
            source_cols.append(excel_col)
        elif db_field == "__NOTES__":
            notes_cols.append(excel_col)

    imported = 0
    updated = 0
    skipped = 0
    errors = []
    imported_ids = []

    for idx, (_, row) in enumerate(df.iterrows()):
        # Check if the name column is empty
        name_val = None
        if name_col and name_col in row:
            name_val = row[name_col]
        name_empty = name_val is None or (isinstance(name_val, float) and pd.isna(name_val)) or str(name_val).strip() == ""

        data = {}
        extra_notes_parts = []

        for excel_col, db_field in column_mapping.items():
            if db_field is None:
                if merge_extra_to_notes:
                    val = row[excel_col]
                    if not (isinstance(val, float) and pd.isna(val)) and str(val).strip() != "":
                        extra_notes_parts.append(f"[{excel_col}]: {val}")
                continue

            if db_field == "__NOTES__":
                val = row[excel_col]
                if not (isinstance(val, float) and pd.isna(val)) and str(val).strip() != "":
                    extra_notes_parts.append(f"[{excel_col}]: {val}")
                continue

            if db_field == "__COORDINATES__":
                val = row[excel_col]
                if not (isinstance(val, float) and pd.isna(val)) and str(val).strip() != "":
                    lat, lng = _parse_coordinates(str(val).strip())
                    if lat is not None:
                        data["latitude"] = lat
                    if lng is not None:
                        data["longitude"] = lng
                continue

            val = row[excel_col]
            if isinstance(val, float) and pd.isna(val):
                continue

            val_str = str(val).strip()
            if val_str == "" or "DISPIMG" in val_str:
                continue

            data[db_field] = val_str

        # Handle empty-name rows
        is_placeholder = False
        if name_empty:
            if data.get("country"):
                data["name"] = ""
                if "quantity" not in data:
                    data["quantity"] = count_raw_cases_by_country(data["country"])
                is_placeholder = True
            elif not data:
                skipped += 1
                continue
            else:
                data["name"] = ""

        if extra_notes_parts:
            existing = data.get("authenticity_notes", "")
            if existing:
                extra_notes_parts.insert(0, existing)
            data["authenticity_notes"] = " | ".join(extra_notes_parts)

        # Merge source_urls from multiple columns
        if len(source_cols) > 1:
            source_parts = []
            for sc in source_cols:
                val = row[sc]
                if not (isinstance(val, float) and pd.isna(val)) and str(val).strip() != "":
                    source_parts.append(str(val).strip())
            if source_parts:
                data["source_urls"] = "\n".join(source_parts)

        # Handle rows with empty name
        if not data.get("name") and not is_placeholder:
            skipped += 1
            continue

        try:
            existing_id = None
            if update_existing:
                if data.get("name"):
                    existing_id = _find_raw_case_by_name(data["name"])
                elif is_placeholder and data.get("country"):
                    # For placeholder records, match by country
                    existing_id = _find_raw_case_by_country(data["country"])

            if existing_id:
                update_raw_case(existing_id, data)
                case_id = existing_id
                # Also add new source URLs (skip duplicates)
                if data.get("source_urls"):
                    existing_urls = set(
                        s["url"] for s in get_sources_for_raw_case(case_id)
                    )
                    for url in data["source_urls"].split("\n"):
                        url = url.strip()
                        if url and url not in existing_urls:
                            try:
                                insert_case_source(case_id, url)
                            except Exception:
                                pass
                updated += 1
            else:
                case_id = insert_raw_case(data, track_undo=False)
                imported_ids.append(case_id)

                # Insert source URLs into case_sources table
                if data.get("source_urls"):
                    for url in data["source_urls"].split("\n"):
                        url = url.strip()
                        if url:
                            try:
                                insert_case_source(case_id, url)
                            except Exception:
                                pass
                imported += 1

            if image_data:
                excel_row_idx = idx + 1  # +1 for header row
                if excel_row_idx in image_data:
                    img_refs = []
                    for img_bytes in image_data[excel_row_idx]:
                        try:
                            img_id = insert_case_image(case_id, img_bytes, source_url="Excel嵌入图片")
                            img_refs.append(f"dbimg://{img_id}")
                        except Exception as e:
                            errors.append(f"行{idx+2}图片: {e}")
                    if img_refs:
                        update_raw_case(case_id, {"image_path": json.dumps(img_refs, ensure_ascii=False)})
        except Exception as e:
            errors.append(f"行{idx+2}: {e}")

    if imported_ids:
        push_undo("import_batch", {"ids": imported_ids})
    return imported, updated, skipped, errors


def _find_raw_case_by_name(name):
    """Return raw_case id if a case with this name exists, else None."""
    conn = get_connection()
    row = conn.execute(
        "SELECT id FROM raw_cases WHERE name = ? LIMIT 1", (name,)
    ).fetchone()
    conn.close()
    return row[0] if row else None


def count_raw_cases_by_country(country):
    """Count how many raw_cases have this country (with non-empty name)."""
    conn = get_connection()
    row = conn.execute(
        "SELECT COUNT(*) FROM raw_cases WHERE country = ? AND name IS NOT NULL AND name != ''",
        (country,)
    ).fetchone()
    conn.close()
    return row[0] if row else 0


def _find_raw_case_by_country(country):
    """Return raw_case id of a country placeholder (name empty, same country), else None."""
    conn = get_connection()
    row = conn.execute(
        "SELECT id FROM raw_cases WHERE (name IS NULL OR name = '') AND country = ? LIMIT 1",
        (country,)
    ).fetchone()
    conn.close()
    return row[0] if row else None


def export_all_to_dataframe():
    import pandas as pd
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM cases ORDER BY id", conn)
    conn.close()
    col_rename = {name: label for name, label, *_ in FIELDS}
    col_rename["id"] = "ID"
    col_rename["created_at"] = "创建时间"
    col_rename["updated_at"] = "更新时间"
    df.rename(columns=col_rename, inplace=True)
    return df


def get_stats_by_field(field):
    conn = get_connection()
    rows = conn.execute(
        f"SELECT {field}, COUNT(*) as cnt FROM cases WHERE {field} IS NOT NULL AND {field} != '' GROUP BY {field} ORDER BY cnt DESC"
    ).fetchall()
    conn.close()
    return [(r[0], r[1]) for r in rows]


def get_distinct_values(table, column):
    """Return sorted list of distinct non-null, non-empty values for a column."""
    conn = get_connection()
    rows = conn.execute(
        f"SELECT DISTINCT {column} FROM {table} "
        f"WHERE {column} IS NOT NULL AND {column} != '' "
        f"ORDER BY {column}"
    ).fetchall()
    conn.close()
    return [r[0] for r in rows]


def find_duplicates():
    conn = get_connection()
    rows = conn.execute("""
        SELECT name, country, COUNT(*) as cnt, GROUP_CONCAT(id) as ids
        FROM cases
        GROUP BY name, country
        HAVING cnt > 1
    """).fetchall()
    conn.close()
    return [{"name": r[0], "country": r[1], "count": r[2], "ids": [int(x) for x in r[3].split(",")]} for r in rows]


def find_raw_duplicates():
    conn = get_connection()
    rows = conn.execute("""
        SELECT name, country, COUNT(*) as cnt, GROUP_CONCAT(id) as ids
        FROM raw_cases
        WHERE name IS NOT NULL AND name != ''
        GROUP BY name, country
        HAVING cnt > 1
    """).fetchall()
    conn.close()
    return [{"name": r[0], "country": r[1], "count": r[2], "ids": [int(x) for x in r[3].split(",")]} for r in rows]


def find_raw_duplicates_all_fields():
    """Find groups of raw_cases where ALL fields are identical."""
    conn = get_connection()
    # Get all column names except id and timestamps
    all_cols = [r[1] for r in conn.execute("PRAGMA table_info(raw_cases)").fetchall()]
    cmp_cols = [c for c in all_cols if c not in ("id", "created_at", "updated_at")]
    cols_sql = ", ".join(cmp_cols)
    group_sql = ", ".join(cmp_cols)
    rows = conn.execute(f"""
        SELECT {cols_sql}, COUNT(*) as cnt, GROUP_CONCAT(id) as ids
        FROM raw_cases
        GROUP BY {group_sql}
        HAVING cnt > 1
    """).fetchall()
    conn.close()
    results = []
    for r in rows:
        d = {}
        for i, col in enumerate(cmp_cols):
            d[col] = r[i]
        d["count"] = r[len(cmp_cols)]
        d["ids"] = [int(x) for x in r[len(cmp_cols) + 1].split(",")]
        results.append(d)
    return results


def find_duplicates_by_name_country():
    """Find groups of raw_cases where name AND country match (case-insensitive).
    Returns [{name, country, count, ids: [int, ...]}, ...] for groups with 2+ records."""
    conn = get_connection()
    rows = conn.execute("""
        SELECT name, country, COUNT(*) as cnt, GROUP_CONCAT(id) as ids
        FROM raw_cases
        WHERE name IS NOT NULL AND name != '' AND country IS NOT NULL AND country != ''
        GROUP BY LOWER(name), LOWER(country)
        HAVING cnt > 1
        ORDER BY cnt DESC
    """).fetchall()
    conn.close()
    results = []
    for r in rows:
        results.append({
            "name": r[0],
            "country": r[1],
            "count": r[2],
            "ids": [int(x) for x in r[3].split(",")]
        })
    return results


def find_verified_duplicates_all_fields():
    """Find groups of verified_cases where ALL fields are identical."""
    conn = get_connection()
    all_cols = [r[1] for r in conn.execute("PRAGMA table_info(verified_cases)").fetchall()]
    cmp_cols = [c for c in all_cols if c not in ("id", "created_at", "updated_at")]
    cols_sql = ", ".join(cmp_cols)
    group_sql = ", ".join(cmp_cols)
    rows = conn.execute(f"""
        SELECT {cols_sql}, COUNT(*) as cnt, GROUP_CONCAT(id) as ids
        FROM verified_cases
        GROUP BY {group_sql}
        HAVING cnt > 1
    """).fetchall()
    conn.close()
    results = []
    for r in rows:
        d = {}
        for i, col in enumerate(cmp_cols):
            d[col] = r[i]
        d["count"] = r[len(cmp_cols)]
        d["ids"] = [int(x) for x in r[len(cmp_cols) + 1].split(",")]
        results.append(d)
    return results


def find_verified_duplicates_by_name_country():
    """Find groups of verified_cases where name AND country match (case-insensitive)."""
    conn = get_connection()
    rows = conn.execute("""
        SELECT name, country, COUNT(*) as cnt, GROUP_CONCAT(id) as ids
        FROM verified_cases
        WHERE name IS NOT NULL AND name != '' AND country IS NOT NULL AND country != ''
        GROUP BY LOWER(name), LOWER(country)
        HAVING cnt > 1
        ORDER BY cnt DESC
    """).fetchall()
    conn.close()
    results = []
    for r in rows:
        results.append({
            "name": r[0],
            "country": r[1],
            "count": r[2],
            "ids": [int(x) for x in r[3].split(",")]
        })
    return results


def smart_merge_duplicates(dry_run=False):
    """Auto-merge raw_cases groups with same name+country.
    Smart rules per field: same value→keep, empty→skip, one-has-value→use it,
    different-values→join with '; ', links→concatenate all.
    If dry_run=True, returns planned merges without executing.
    Returns (merged_count, total_secondary_deleted)."""
    groups = find_duplicates_by_name_country()
    if not groups:
        return 0, 0

    raw_field_names = [f[0] for f in RAW_FIELDS]
    merged_groups = 0
    total_deleted = 0

    for g in groups:
        ids = g["ids"]
        if len(ids) < 2:
            continue
        records = []
        for rid in ids:
            r = get_raw_case_by_id(rid)
            if r:
                records.append(r)
        if len(records) < 2:
            continue

        primary_id = records[0]["id"]
        secondary_ids = [r["id"] for r in records[1:]]

        field_updates = {}
        for field in raw_field_names:
            vals = []
            for r in records:
                v = r.get(field)
                if v is not None and str(v).strip() != "" and str(v).strip().lower() != "none":
                    s = str(v).strip()
                    if s not in vals:
                        vals.append(s)
            if not vals:
                field_updates[field] = None
            elif len(vals) == 1:
                field_updates[field] = vals[0]
            else:
                field_updates[field] = "; ".join(vals)

        if not dry_run:
            merge_multiple_raw_cases(primary_id, secondary_ids, field_updates)
        merged_groups += 1
        total_deleted += len(secondary_ids)

    return merged_groups, total_deleted


def get_missing_info():
    conn = get_connection()
    results = []
    required_fields = ["name", "country", "city", "year_built", "building_type"]
    for field in required_fields:
        rows = conn.execute(
            f"SELECT id, name FROM cases WHERE {field} IS NULL OR {field} = ''"
        ).fetchall()
        for r in rows:
            label = {name: lbl for name, lbl, *_ in FIELDS}[field]
            results.append({"id": r[0], "name": r[1], "missing_field": field, "missing_label": label})
    conn.close()
    return results


# ============================================================
#  Phase 1: Data organization refactoring — new tables & CRUD
# ============================================================

RAW_FIELDS = [
    ("name", "建筑名称（中文）", "TEXT NOT NULL"),
    ("name_en", "英文名称", "TEXT"),
    ("name_official", "官方语言名称", "TEXT"),
    ("country", "国家", "TEXT"),
    ("city", "城市", "TEXT"),
    ("construction_year", "建造时间", "TEXT"),
    ("demolition_year", "拆除时间", "TEXT"),
    ("restoration_year", "修复时间", "TEXT"),
    ("quantity", "数量规模", "INTEGER DEFAULT 1"),
    ("function", "功能", "TEXT"),
    ("designer", "设计师", "TEXT"),
    ("latitude", "纬度", "REAL"),
    ("longitude", "经度", "REAL"),
    ("search_time", "检索时间", "TEXT"),
    ("search_keywords", "检索词", "TEXT"),
    ("image_urls", "图片链接", "TEXT"),
    ("source_urls", "来源链接", "TEXT"),
    ("authenticity_status", "真实性状态", "TEXT DEFAULT 'raw'"),
    ("authenticity_score", "可信度", "REAL"),
    ("authenticity_notes", "评估备注", "TEXT"),
    ("building_status", "建筑状态", "TEXT"),
    ("image_path", "图片路径", "TEXT"),
    ("spread_method", "传播方式", "TEXT"),
    ("notes", "备注", "TEXT"),
]

RAW_FIELD_NAMES = [f[0] for f in RAW_FIELDS]
RAW_FIELD_LABELS = [f[1] for f in RAW_FIELDS]


def get_raw_case_extra_columns():
    """Return column names from the raw_cases table that are not in RAW_FIELD_NAMES."""
    conn = get_connection()
    rows = conn.execute("PRAGMA table_info(raw_cases)").fetchall()
    conn.close()
    all_cols = [r[1] for r in rows]
    return [c for c in all_cols if c not in RAW_FIELD_NAMES]


def add_raw_field(field_name, field_label, field_type="TEXT"):
    """Dynamically add a new column to raw_cases and register it in RAW_FIELDS."""
    conn = get_connection()
    try:
        conn.execute(f"ALTER TABLE raw_cases ADD COLUMN {field_name} {field_type}")
    except sqlite3.OperationalError:
        conn.close()
        return False  # column already exists
    conn.commit()
    conn.close()
    RAW_FIELDS.append((field_name, field_label, field_type))
    RAW_FIELD_NAMES.append(field_name)
    RAW_FIELD_LABELS.append(field_label)
    return True


def add_literature_column(col_name, col_type="TEXT"):
    """Dynamically add a new column to the literature table and register it."""
    conn = get_connection()
    try:
        conn.execute(f"ALTER TABLE literature ADD COLUMN {col_name} {col_type}")
    except sqlite3.OperationalError:
        conn.close()
        return False
    conn.commit()
    conn.close()
    if col_name not in _LITERATURE_COLS:
        _LITERATURE_COLS.append(col_name)
    return True


def delete_raw_field(field_name):
    """Drop a column from raw_cases and remove it from RAW_FIELDS."""
    conn = get_connection()
    try:
        conn.execute(f"ALTER TABLE raw_cases DROP COLUMN {field_name}")
    except sqlite3.OperationalError:
        conn.close()
        return False
    conn.commit()
    conn.close()
    # Remove from RAW_FIELDS
    for i, (fn, fl, ft) in enumerate(RAW_FIELDS):
        if fn == field_name:
            RAW_FIELDS.pop(i)
            break
    if field_name in RAW_FIELD_NAMES:
        RAW_FIELD_NAMES.remove(field_name)
    if field_name in RAW_FIELD_LABELS:
        RAW_FIELD_LABELS.remove(field_name)
    return True


def rename_raw_field(old_name, new_name):
    """Rename a column in raw_cases (SQLite 3.25+)."""
    conn = get_connection()
    try:
        conn.execute(f"ALTER TABLE raw_cases RENAME COLUMN {old_name} TO {new_name}")
    except sqlite3.OperationalError:
        conn.close()
        return False
    conn.commit()
    conn.close()
    for i, (fn, fl, ft) in enumerate(RAW_FIELDS):
        if fn == old_name:
            RAW_FIELDS[i] = (new_name, fl, ft)
            break
    if old_name in RAW_FIELD_NAMES:
        RAW_FIELD_NAMES[RAW_FIELD_NAMES.index(old_name)] = new_name
    if old_name in RAW_FIELD_LABELS:
        # Labels don't change on rename
        pass
    return True

ACADEMIC_DOMAINS = {".edu", ".ac.", ".ac.uk", ".edu.cn", ".edu.sg", ".edu.hk"}
GOV_DOMAINS = {".gov", ".gov.cn", ".mil", ".go.kr", ".go.jp"}
NEWS_DOMAINS = {"news.", "bbc.", "cnn.", "reuters.", "nytimes.", "theguardian.", "wsj.",
                "washingtonpost.", "bloomberg.", "apnews.", "usatoday.", "latimes.",
                "chinadaily.", "scmp.", "straitstimes.", "nikkei.", "japantimes.",
                "koreatimes.", "thanhnien.", "vnexpress."}


def _source_credibility(url):
    """Classify a URL's credibility: academic / government / news / blog / unknown."""
    if not url:
        return "unknown"
    from urllib.parse import urlparse
    domain = urlparse(url).netloc.lower()
    for kw in ACADEMIC_DOMAINS:
        if kw in domain:
            return "academic"
    for kw in GOV_DOMAINS:
        if kw in domain:
            return "government"
    for kw in NEWS_DOMAINS:
        if kw in domain:
            return "news"
    if any(kw in domain for kw in ("blog", "wordpress", "medium", "tumblr", "weibo", "zhihu")):
        return "blog"
    return "unknown"


def migrate_db_v2():
    """Create new tables and alter existing cases table for Phase 1 refactoring."""
    conn = get_connection()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS raw_cases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            name_en TEXT,
            name_official TEXT,
            country TEXT,
            city TEXT,
            construction_year TEXT,
            demolition_year TEXT,
            restoration_year TEXT,
            quantity INTEGER DEFAULT 1,
            function TEXT,
            designer TEXT,
            latitude REAL,
            longitude REAL,
            search_time TEXT,
            search_keywords TEXT,
            image_urls TEXT,
            source_urls TEXT,
            authenticity_status TEXT DEFAULT 'raw',
            authenticity_score REAL,
            authenticity_notes TEXT,
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS case_sources (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            raw_case_id INTEGER NOT NULL,
            url TEXT NOT NULL,
            credibility TEXT DEFAULT 'unknown',
            title TEXT,
            fetch_time TEXT,
            FOREIGN KEY (raw_case_id) REFERENCES raw_cases(id) ON DELETE CASCADE
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS field_sources (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            raw_case_id INTEGER,
            case_id INTEGER,
            field_name TEXT NOT NULL,
            field_value TEXT,
            source_url TEXT,
            source_type TEXT DEFAULT 'ai',
            confidence REAL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS case_images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            raw_case_id INTEGER NOT NULL,
            filename TEXT,
            image_data BLOB,
            file_path TEXT,
            source_url TEXT,
            author TEXT,
            image_time TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (raw_case_id) REFERENCES raw_cases(id) ON DELETE CASCADE
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS image_trash (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            raw_case_id INTEGER NOT NULL,
            filename TEXT,
            image_data BLOB,
            file_path TEXT,
            source_url TEXT,
            author TEXT,
            image_time TEXT,
            deleted_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (raw_case_id) REFERENCES raw_cases(id) ON DELETE CASCADE
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS verification_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            raw_case_id INTEGER NOT NULL,
            action TEXT NOT NULL,
            details TEXT,
            result TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (raw_case_id) REFERENCES raw_cases(id) ON DELETE CASCADE
        )
    """)

    # Add new columns to existing cases table (ignore if already exist)
    alter_columns = [
        ("demolition_year", "TEXT"),
        ("restoration_year", "TEXT"),
        ("quantity", "INTEGER DEFAULT 1"),
        ("latitude", "REAL"),
        ("longitude", "REAL"),
        ("image_urls", "TEXT"),
        ("search_time", "TEXT"),
        ("search_keywords", "TEXT"),
        ("condition", "TEXT"),
        ("verification_status", "TEXT DEFAULT 'unverified'"),
        ("completeness_score", "REAL DEFAULT 0"),
        ("source_credibility", "TEXT DEFAULT 'unknown'"),
        ("raw_case_ids", "TEXT"),
        ("name_en", "TEXT"),
        ("name_official", "TEXT"),
        ("spread_method", "TEXT"),
        ("sort_order", "REAL DEFAULT 0"),
    ]
    for col_name, col_def in alter_columns:
        try:
            conn.execute(f"ALTER TABLE cases ADD COLUMN {col_name} {col_def}")
        except sqlite3.OperationalError:
            pass  # column already exists

    # Add new columns to raw_cases table (ignore if already exist)
    raw_alter_columns = [
        ("name_en", "TEXT"),
        ("name_official", "TEXT"),
        ("building_status", "TEXT"),
        ("image_path", "TEXT"),
        ("spread_method", "TEXT"),
        ("sort_order", "REAL DEFAULT 0"),
        ("image_metadata", "TEXT"),
    ]
    for col_name, col_def in raw_alter_columns:
        try:
            conn.execute(f"ALTER TABLE raw_cases ADD COLUMN {col_name} {col_def}")
        except sqlite3.OperationalError:
            pass  # column already exists

    # Migrate image tables to support file_path
    for img_table in ("case_images", "image_trash"):
        try:
            conn.execute(f"ALTER TABLE {img_table} ADD COLUMN file_path TEXT")
        except sqlite3.OperationalError:
            pass

    # Migrate existing raw_cases.source_urls into case_sources table
    _migrate_source_urls_to_case_sources(conn)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS literature (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            authors TEXT,
            pub_date TEXT,
            source_url TEXT,
            source_engine TEXT,
            content TEXT,
            summary TEXT,
            file_path TEXT,
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    try:
        conn.execute("ALTER TABLE literature ADD COLUMN pub_date TEXT")
    except sqlite3.OperationalError:
        pass

    for col in ("journal", "city", "country", "translator", "publisher"):
        try:
            conn.execute(f"ALTER TABLE literature ADD COLUMN {col} TEXT")
        except sqlite3.OperationalError:
            pass

    conn.execute("""
        CREATE TABLE IF NOT EXISTS filtered_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            query TEXT NOT NULL,
            title TEXT NOT NULL,
            snippet TEXT,
            link TEXT,
            relevance_reason TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS raw_search_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            query TEXT NOT NULL,
            engine TEXT NOT NULL,
            title TEXT,
            snippet TEXT,
            link TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    try:
        conn.execute("ALTER TABLE raw_search_results ADD COLUMN status TEXT DEFAULT 'active'")
    except sqlite3.OperationalError:
        pass
    for col in ["authors", "year", "publication", "pdf_url", "cited_by"]:
        try:
            conn.execute(f"ALTER TABLE raw_search_results ADD COLUMN {col} TEXT DEFAULT ''")
        except sqlite3.OperationalError:
            pass

    conn.execute("""
        CREATE TABLE IF NOT EXISTS dedup_removed (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            link TEXT NOT NULL,
            query TEXT,
            removed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Migrate old serper_results table to filtered_results
    old_exists = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='serper_results'"
    ).fetchone()
    if old_exists:
        new_exists = conn.execute(
            "SELECT COUNT(*) FROM filtered_results"
        ).fetchone()[0]
        if new_exists == 0:
            conn.execute(
                "INSERT INTO filtered_results (id, query, title, snippet, link, relevance_reason, created_at) "
                "SELECT id, query, title, snippet, link, relevance_reason, created_at FROM serper_results"
            )
        conn.execute("DROP TABLE serper_results")

    conn.commit()
    _ensure_image_dirs()
    conn.close()


def _migrate_source_urls_to_case_sources(conn):
    """One-time migration: parse raw_cases.source_urls and insert into case_sources."""
    rows = conn.execute(
        "SELECT id, source_urls FROM raw_cases WHERE source_urls IS NOT NULL AND source_urls != ''"
    ).fetchall()
    migrated_count = 0
    for raw_case_id, source_urls in rows:
        existing = set(
            r[0] for r in conn.execute(
                "SELECT url FROM case_sources WHERE raw_case_id = ?", (raw_case_id,)
            ).fetchall()
        )
        urls = []
        text = source_urls.strip()
        if text.startswith("["):
            try:
                urls = json.loads(text)
            except json.JSONDecodeError:
                urls = text.split("\n")
        else:
            urls = text.split("\n")
        for url in urls:
            url = url.strip()
            if url and url not in existing:
                cred = _source_credibility(url)
                conn.execute(
                    "INSERT INTO case_sources (raw_case_id, url, credibility, fetch_time) VALUES (?, ?, ?, datetime('now'))",
                    (raw_case_id, url, cred))
                existing.add(url)
        if urls:
            migrated_count += 1
    if migrated_count:
        print(f"[migration] Migrated source_urls for {migrated_count} raw cases.")


def migrate_db_v3():
    """Create scholar_search_results table and migrate data from raw_search_results."""
    conn = get_connection()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS scholar_search_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            query TEXT NOT NULL,
            title TEXT,
            authors TEXT DEFAULT '',
            year TEXT DEFAULT '',
            publication TEXT DEFAULT '',
            snippet TEXT,
            link TEXT,
            pdf_url TEXT DEFAULT '',
            cited_by TEXT DEFAULT '',
            status TEXT DEFAULT 'active',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Migrate existing scholar data from raw_search_results
    existing = conn.execute(
        "SELECT COUNT(*) FROM scholar_search_results"
    ).fetchone()[0]
    if existing == 0:
        scholar_rows = conn.execute(
            "SELECT query, title, authors, year, publication, snippet, link, "
            "pdf_url, cited_by, status, created_at "
            "FROM raw_search_results WHERE engine='scholar'"
        ).fetchall()
        if scholar_rows:
            conn.executemany(
                "INSERT INTO scholar_search_results "
                "(query, title, authors, year, publication, snippet, link, pdf_url, cited_by, status, created_at) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                scholar_rows
            )
            conn.execute("DELETE FROM raw_search_results WHERE engine='scholar'")
            print(f"[migration v3] Migrated {len(scholar_rows)} scholar results to scholar_search_results")

    conn.commit()
    conn.close()


def migrate_db_v4():
    """Add ds_reason column to scholar_search_results."""
    conn = get_connection()
    try:
        conn.execute("ALTER TABLE scholar_search_results ADD COLUMN ds_reason TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass  # column already exists
    conn.commit()
    conn.close()


def migrate_db_v5():
    """Add performance indexes."""
    conn = get_connection()
    indexes = [
        "CREATE INDEX IF NOT EXISTS idx_raw_search_query_engine ON raw_search_results(query, engine)",
        "CREATE INDEX IF NOT EXISTS idx_raw_search_link ON raw_search_results(link)",
        "CREATE INDEX IF NOT EXISTS idx_filtered_results_link ON filtered_results(link)",
        "CREATE INDEX IF NOT EXISTS idx_filtered_results_query ON filtered_results(query)",
        "CREATE INDEX IF NOT EXISTS idx_dedup_removed_link ON dedup_removed(link)",
        "CREATE INDEX IF NOT EXISTS idx_scholar_search_query ON scholar_search_results(query)",
        "CREATE INDEX IF NOT EXISTS idx_scholar_search_link ON scholar_search_results(link)",
    ]
    for sql in indexes:
        try:
            conn.execute(sql)
        except sqlite3.OperationalError as e:
            print(f"[migrate_v5] {e}")
    conn.commit()
    conn.close()


def migrate_db_v6():
    """Add relevance_reason column to raw_search_results and backfill from filtered_results."""
    conn = get_connection()
    try:
        conn.execute("ALTER TABLE raw_search_results ADD COLUMN relevance_reason TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass  # column already exists
    # Backfill reasons from filtered_results by matching link
    conn.execute(
        "UPDATE raw_search_results SET relevance_reason = ("
        "  SELECT f.relevance_reason FROM filtered_results f "
        "  WHERE f.link = raw_search_results.link AND f.relevance_reason != '' "
        "  LIMIT 1"
        ") WHERE relevance_reason = ''"
    )
    conn.commit()
    conn.close()


def migrate_db_v7():
    """Add author and image_time columns to case_images and image_trash, plus image_metadata to raw_cases."""
    conn = get_connection()
    for table in ("case_images", "image_trash"):
        for col_name in ("author", "image_time"):
            try:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} TEXT")
            except sqlite3.OperationalError:
                pass
    try:
        conn.execute("ALTER TABLE raw_cases ADD COLUMN image_metadata TEXT")
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()


def migrate_db_v8():
    """Add notes column to raw_cases."""
    conn = get_connection()
    try:
        conn.execute("ALTER TABLE raw_cases ADD COLUMN notes TEXT")
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()


def migrate_db_v9():
    """Create operation_log table for auditing data changes in search dialogs."""
    conn = get_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS operation_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            page TEXT NOT NULL,
            action TEXT NOT NULL,
            detail TEXT,
            count INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()


def log_operation(page, action, detail="", count=0):
    """Record an operation to the audit log."""
    conn = get_connection()
    conn.execute(
        "INSERT INTO operation_log (page, action, detail, count) VALUES (?, ?, ?, ?)",
        (page, action, detail, count))
    conn.commit()
    conn.close()


def get_recent_operations(limit=20):
    """Get recent operation log entries."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM operation_log ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# --- Raw Cases CRUD ---

def get_all_raw_cases(status=None, search=None, filters=None, order_by=None):
    conn = get_connection()
    query = "SELECT * FROM raw_cases WHERE 1=1"
    params = []
    if status:
        query += " AND authenticity_status = ?"
        params.append(status)
    if search:
        query += " AND (name LIKE ? OR name_en LIKE ? OR name_official LIKE ? OR country LIKE ? OR city LIKE ?)"
        s = f"%{search}%"
        params.extend([s, s, s, s, s])
    if filters:
        valid_cols = _get_all_raw_case_columns(conn)
        raw_field_map = {"country": "country", "function": "function",
                         "building_type": "function", "spread_method": "spread_method",
                         "construction_year": "construction_year", "building_status": "building_status"}
        for key, val in filters.items():
            col = raw_field_map.get(key, key)
            if col in valid_cols:
                query += f" AND {col} = ?"
                params.append(val)
    if order_by:
        field, direction = order_by
        raw_field_map = {"name": "name", "country": "country", "city": "city",
                         "function": "function", "construction_year": "construction_year",
                         "spread_method": "spread_method", "building_status": "building_status"}
        col = raw_field_map.get(field, field)
        query += f" ORDER BY {col} {direction}"
    else:
        query += " ORDER BY sort_order ASC, updated_at DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_raw_case_by_id(case_id):
    conn = get_connection()
    row = conn.execute("SELECT * FROM raw_cases WHERE id = ?", (case_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def insert_raw_case(data, track_undo=True):
    conn = get_connection()
    valid_fields = _get_all_raw_case_columns(conn)
    fields = [f for f in valid_fields if f in data]
    placeholders = ", ".join(["?" for _ in fields])
    columns = ", ".join(fields)
    values = [data[f] for f in fields]
    conn.execute(f"INSERT INTO raw_cases ({columns}) VALUES ({placeholders})", values)
    conn.commit()
    last_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    if track_undo:
        push_undo("insert", {"id": last_id, "name": data.get("name", "")})
    conn.close()
    return last_id


def _get_all_raw_case_columns(conn=None):
    """Return all actual column names of raw_cases table (code-defined + dynamic)."""
    should_close = conn is None
    if conn is None:
        conn = get_connection()
    rows = conn.execute("PRAGMA table_info(raw_cases)").fetchall()
    if should_close:
        conn.close()
    return [r[1] for r in rows]


def update_raw_case(case_id, data):
    conn = get_connection()
    old = dict(conn.execute("SELECT * FROM raw_cases WHERE id = ?", (case_id,)).fetchone())
    old_values = {}
    for k in data:
        if k in old and old[k] != data[k]:
            old_values[k] = old[k]

    valid_fields = _get_all_raw_case_columns(conn)
    filtered = {k: v for k, v in data.items() if k in valid_fields}
    sets = ", ".join([f"{k} = ?" for k in filtered])
    values = list(filtered.values()) + [case_id]
    conn.execute(f"UPDATE raw_cases SET {sets}, updated_at = CURRENT_TIMESTAMP WHERE id = ?", values)
    conn.commit()

    if old_values:
        push_undo("update", {"id": case_id, "name": old.get("name", ""), "old_values": old_values})
    conn.close()


def merge_raw_cases(primary_id, secondary_id, field_choices=None):
    """Merge secondary into primary. field_choices: {field: value} overrides for merged data.
    Move related data from secondary to primary, then delete secondary."""
    conn = get_connection()
    primary = dict(conn.execute("SELECT * FROM raw_cases WHERE id = ?", (primary_id,)).fetchone())
    secondary = dict(conn.execute("SELECT * FROM raw_cases WHERE id = ?", (secondary_id,)).fetchone())

    if field_choices is None:
        # Default: fill empty fields in primary with secondary's values
        updates = {}
        for key, val in secondary.items():
            if key in ("id", "created_at", "updated_at"):
                continue
            primary_val = primary.get(key)
            if (primary_val is None or primary_val == "" or primary_val == "None") and val is not None and val != "":
                updates[key] = val
    else:
        updates = field_choices

    if updates:
        sets = ", ".join(f"{k} = ?" for k in updates)
        values = list(updates.values()) + [primary_id]
        conn.execute(f"UPDATE raw_cases SET {sets}, updated_at = CURRENT_TIMESTAMP WHERE id = ?", values)

    # Move related data from secondary to primary
    for table, id_col in [("case_sources", "raw_case_id"), ("verification_log", "raw_case_id"),
                           ("case_images", "raw_case_id"), ("field_sources", "raw_case_id")]:
        conn.execute(f"UPDATE {table} SET {id_col} = ? WHERE {id_col} = ?",
                     (primary_id, secondary_id))

    # Update cases table: record both raw case IDs for traceability
    conn.execute(
        "UPDATE cases SET raw_case_ids = ? || ',' || ? WHERE raw_case_ids = ?",
        (str(primary_id), str(secondary_id), str(secondary_id)))
    conn.execute(
        "UPDATE cases SET raw_case_ids = raw_case_ids || ',' || ? "
        "WHERE raw_case_ids = ? AND raw_case_ids NOT LIKE '%' || ? || '%'",
        (str(secondary_id), str(primary_id), str(secondary_id)))

    # Save undo data before deleting secondary
    primary_old = {k: primary[k] for k in updates} if updates else {}
    push_undo("merge", {
        "primary_id": primary_id, "primary_old": primary_old,
        "secondary": secondary,
    })

    # Delete secondary
    conn.execute("DELETE FROM raw_cases WHERE id = ?", (secondary_id,))

    conn.commit()
    filled = list(updates.keys()) if updates else []
    conn.close()
    return filled


def merge_multiple_raw_cases(primary_id, secondary_ids, field_updates=None):
    """Merge N secondary cases into one primary.

    field_updates: {field: value} to set on primary before absorbing secondaries.
    All related records (sources, images, logs) are moved to primary.
    Each secondary is deleted after merge.
    Returns the list of field names that were updated on primary.
    """
    conn = get_connection()
    primary = dict(conn.execute("SELECT * FROM raw_cases WHERE id = ?", (primary_id,)).fetchone())
    filled = []

    if field_updates:
        sets = ", ".join(f"{k} = ?" for k in field_updates)
        values = list(field_updates.values()) + [primary_id]
        conn.execute(f"UPDATE raw_cases SET {sets}, updated_at = CURRENT_TIMESTAMP WHERE id = ?", values)
        filled = list(field_updates.keys())

    for secondary_id in secondary_ids:
        secondary = dict(conn.execute("SELECT * FROM raw_cases WHERE id = ?", (secondary_id,)).fetchone())

        for table, id_col in [("case_sources", "raw_case_id"), ("verification_log", "raw_case_id"),
                               ("case_images", "raw_case_id"), ("field_sources", "raw_case_id")]:
            conn.execute(f"UPDATE {table} SET {id_col} = ? WHERE {id_col} = ?",
                         (primary_id, secondary_id))

        conn.execute(
            "UPDATE cases SET raw_case_ids = ? || ',' || ? WHERE raw_case_ids = ?",
            (str(primary_id), str(secondary_id), str(secondary_id)))
        conn.execute(
            "UPDATE cases SET raw_case_ids = raw_case_ids || ',' || ? "
            "WHERE raw_case_ids = ? AND raw_case_ids NOT LIKE '%' || ? || '%'",
            (str(secondary_id), str(primary_id), str(secondary_id)))

        push_undo("merge", {
            "primary_id": primary_id,
            "primary_old": {k: primary[k] for k in field_updates} if field_updates else {},
            "secondary": secondary,
        })

        conn.execute("DELETE FROM raw_cases WHERE id = ?", (secondary_id,))

    conn.commit()
    conn.close()
    return filled


def merge_verified_cases(primary_id, secondary_ids, field_updates=None):
    """直接合并核验表记录：更新主记录字段，删除被合并的记录。"""
    conn = get_connection()
    if field_updates:
        sets = ", ".join(f"{k} = ?" for k in field_updates)
        values = list(field_updates.values()) + [primary_id]
        conn.execute(f"UPDATE verified_cases SET {sets}, updated_at = CURRENT_TIMESTAMP WHERE id = ?", values)

    for sid in secondary_ids:
        conn.execute("DELETE FROM verified_cases WHERE id = ?", (sid,))

    conn.commit()
    conn.close()
    return list(field_updates.keys()) if field_updates else []


def delete_raw_case(case_id):
    conn = get_connection()
    # Save full case data for undo
    case = dict(conn.execute("SELECT * FROM raw_cases WHERE id = ?", (case_id,)).fetchone())
    sources = [dict(r) for r in conn.execute("SELECT * FROM case_sources WHERE raw_case_id = ?", (case_id,)).fetchall()]
    logs = [dict(r) for r in conn.execute("SELECT * FROM verification_log WHERE raw_case_id = ?", (case_id,)).fetchall()]
    push_undo("delete", {"case": case, "sources": sources, "logs": logs})

    conn.execute("DELETE FROM raw_cases WHERE id = ?", (case_id,))
    conn.commit()
    conn.close()


def set_raw_case_status(case_id, status):
    conn = get_connection()
    old = conn.execute("SELECT authenticity_status, name FROM raw_cases WHERE id = ?", (case_id,)).fetchone()
    old_status = old[0] if old else None
    old_name = old[1] if old else ""
    conn.execute("UPDATE raw_cases SET authenticity_status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                 (status, case_id))
    conn.commit()
    if old_status and old_status != status:
        push_undo("status_change", {"id": case_id, "name": old_name, "old_status": old_status})
    conn.close()


def update_list_order(table, ordered_ids):
    """Batch update sort_order for a list of IDs in the given table.

    Args:
        table: "raw_cases" or "cases"
        ordered_ids: list of case IDs in desired display order
    """
    conn = get_connection()
    for i, case_id in enumerate(ordered_ids):
        conn.execute(
            f"UPDATE {table} SET sort_order = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            (i + 1, case_id))
    conn.commit()
    conn.close()


# --- Case Sources CRUD ---

def insert_case_source(raw_case_id, url, title=None, credibility=None):
    if credibility is None:
        credibility = _source_credibility(url)
    conn = get_connection()
    conn.execute(
        "INSERT INTO case_sources (raw_case_id, url, title, credibility, fetch_time) VALUES (?, ?, ?, ?, datetime('now'))",
        (raw_case_id, url, title, credibility))
    conn.commit()
    conn.close()


def get_sources_for_raw_case(raw_case_id):
    conn = get_connection()
    rows = conn.execute("SELECT * FROM case_sources WHERE raw_case_id = ? ORDER BY id", (raw_case_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_case_source(source_id, data):
    conn = get_connection()
    sets = ", ".join([f"{k} = ?" for k in data])
    values = list(data.values()) + [source_id]
    conn.execute(f"UPDATE case_sources SET {sets} WHERE id = ?", values)
    conn.commit()
    conn.close()


def delete_case_source(source_id):
    conn = get_connection()
    conn.execute("DELETE FROM case_sources WHERE id = ?", (source_id,))
    conn.commit()
    conn.close()


# --- Field Sources CRUD ---

def insert_field_source(raw_case_id=None, case_id=None, field_name=None, field_value=None,
                        source_url=None, source_type="ai", confidence=None):
    conn = get_connection()
    conn.execute(
        "INSERT INTO field_sources (raw_case_id, case_id, field_name, field_value, source_url, source_type, confidence) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        (raw_case_id, case_id, field_name, field_value, source_url, source_type, confidence))
    conn.commit()
    conn.close()


def get_field_sources(raw_case_id=None, case_id=None):
    conn = get_connection()
    if raw_case_id:
        rows = conn.execute(
            "SELECT * FROM field_sources WHERE raw_case_id = ? ORDER BY id", (raw_case_id,)).fetchall()
    elif case_id:
        rows = conn.execute(
            "SELECT * FROM field_sources WHERE case_id = ? ORDER BY id", (case_id,)).fetchall()
    else:
        rows = []
    conn.close()
    return [dict(r) for r in rows]


# --- Verification Log CRUD ---

def insert_verification_log(raw_case_id, action, details=None, result=None):
    conn = get_connection()
    conn.execute(
        "INSERT INTO verification_log (raw_case_id, action, details, result) VALUES (?, ?, ?, ?)",
        (raw_case_id, action, details, result))
    conn.commit()
    conn.close()


def get_verification_logs(raw_case_id):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM verification_log WHERE raw_case_id = ? ORDER BY created_at DESC", (raw_case_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# --- Verified Cases CRUD ---

def get_all_verified_cases(status=None, search=None, filters=None, order_by=None):
    """Return all cases from verified_cases table with optional filters."""
    conn = get_connection()
    query = "SELECT * FROM verified_cases WHERE 1=1"
    params = []
    if status:
        query += " AND authenticity_status = ?"
        params.append(status)
    else:
        query += " AND authenticity_status != 'raw'"
    if search:
        query += (" AND (name LIKE ? OR name_en LIKE ? OR name_official LIKE ? "
                  "OR country LIKE ? OR city LIKE ? OR designer LIKE ?)")
        s = f"%{search}%"
        params.extend([s, s, s, s, s, s])
    if filters:
        for field, value in filters.items():
            if value:
                query += f" AND {field} = ?"
                params.append(value)
    if order_by:
        field, direction = order_by
        if field in RAW_FIELD_NAMES:
            query += f" ORDER BY {field} {direction}"
        else:
            query += " ORDER BY updated_at DESC"
    else:
        query += " ORDER BY updated_at DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_verified_case_by_id(case_id):
    """Get a single verified case by id."""
    conn = get_connection()
    row = conn.execute("SELECT * FROM verified_cases WHERE id = ?", (case_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_verified_case_by_raw_id(raw_case_id):
    """Get a verified case by its original raw_case_id."""
    conn = get_connection()
    row = conn.execute(
        "SELECT * FROM verified_cases WHERE raw_case_id = ?", (raw_case_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def insert_verified_case(data):
    """Insert a new verified case. Returns the new id."""
    conn = get_connection()
    fields = [f for f in RAW_FIELD_NAMES + ["raw_case_id"] if f in data]
    placeholders = ", ".join(["?" for _ in fields])
    columns = ", ".join(fields)
    values = [data[f] for f in fields]
    conn.execute(f"INSERT INTO verified_cases ({columns}) VALUES ({placeholders})", values)
    conn.commit()
    new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.close()
    return new_id


def update_verified_case(case_id, data):
    """Update a verified case by id."""
    conn = get_connection()
    sets = ", ".join([f"{k} = ?" for k in data])
    values = list(data.values()) + [case_id]
    conn.execute(f"UPDATE verified_cases SET {sets}, updated_at = CURRENT_TIMESTAMP WHERE id = ?", values)
    conn.commit()
    conn.close()


def delete_verified_case(case_id):
    """Delete a verified case by id."""
    conn = get_connection()
    conn.execute("DELETE FROM verified_cases WHERE id = ?", (case_id,))
    conn.commit()
    conn.close()


def copy_raw_to_verified(raw_case_id):
    """Copy a raw_case into verified_cases. Returns the verified case id.
    If already exists, returns the existing id."""
    existing = get_verified_case_by_raw_id(raw_case_id)
    if existing:
        return existing["id"]
    raw = get_raw_case_by_id(raw_case_id)
    if not raw:
        return None
    data = {k: v for k, v in raw.items()
            if k in RAW_FIELD_NAMES and k != "id"}
    data["raw_case_id"] = raw_case_id
    return insert_verified_case(data)


def set_verified_case_status(case_id, status):
    """Update authenticity_status on a verified case."""
    for i in range(10):
        conn = get_connection()
        try:
            conn.execute(
                "UPDATE verified_cases SET authenticity_status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                (status, case_id))
            conn.commit()
            conn.close()
            return
        except sqlite3.OperationalError as e:
            conn.close()
            if "locked" in str(e).lower() and i < 9:
                import time
                time.sleep(0.3 * (i + 1))
            else:
                raise




# --- Support Functions ---

def calculate_completeness(case_id=None, raw_case_id=None):
    """Calculate completeness score (0-1) based on how many fields are filled."""
    if case_id:
        case = get_case_by_id(case_id)
        if not case:
            return 0
        check_fields = ["name", "country", "city", "year_built", "demolition_year",
                        "restoration_year", "quantity", "building_type", "architect",
                        "latitude", "longitude", "material", "area_sqm",
                        "protection_level", "description"]
        filled = sum(1 for f in check_fields if case.get(f) is not None)
        return round(filled / len(check_fields), 2)
    elif raw_case_id:
        case = get_raw_case_by_id(raw_case_id)
        if not case:
            return 0
        check_fields = ["name", "country", "city", "construction_year", "function",
                        "designer", "latitude", "longitude", "search_keywords", "source_urls"]
        filled = sum(1 for f in check_fields if case.get(f) is not None)
        return round(filled / len(check_fields), 2)
    return 0


def find_similar_raw_cases(name, country=None):
    """Find raw_cases with similar name in the same country for cross-validation."""
    conn = get_connection()
    query = "SELECT * FROM raw_cases WHERE name LIKE ?"
    params = [f"%{name}%"]
    if country:
        query += " AND country = ?"
        params.append(country)
    query += " ORDER BY id"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def _get_verified_or_raw(raw_case_id):
    """Return the verified case dict if it exists, otherwise the raw case dict."""
    verified = get_verified_case_by_raw_id(raw_case_id)
    if verified:
        return verified
    return get_raw_case_by_id(raw_case_id)


def promote_raw_to_case(raw_case_id):
    """Copy a verified case (or raw case as fallback) into the main cases table.
    If a case record already exists for this raw_case_id, update it instead of inserting."""
    src = _get_verified_or_raw(raw_case_id)
    if not src:
        return None

    data = {
        "name": src.get("name"),
        "name_en": src.get("name_en"),
        "name_official": src.get("name_official"),
        "country": src.get("country"),
        "city": src.get("city"),
        "year_built": src.get("construction_year"),
        "demolition_year": src.get("demolition_year"),
        "restoration_year": src.get("restoration_year"),
        "quantity": src.get("quantity"),
        "building_type": src.get("function"),
        "architect": src.get("designer"),
        "latitude": src.get("latitude"),
        "longitude": src.get("longitude"),
        "image_urls": src.get("image_urls"),
        "image_path": src.get("image_path"),
        "search_time": src.get("search_time"),
        "search_keywords": src.get("search_keywords"),
        "notes": src.get("authenticity_notes"),
    }
    # Ensure name is never None (NOT NULL constraint)
    if data.get("name") is None:
        data["name"] = ""

    # Check if case already exists for this raw_case_id
    existing = _find_case_by_raw_id(raw_case_id)
    if existing:
        update_case(existing["id"], data)
        case_id = existing["id"]
    else:
        case_id = insert_case(data)

    credibility = "unknown"
    sources = get_sources_for_raw_case(raw_case_id)
    if sources:
        creds = [s["credibility"] for s in sources]
        if "academic" in creds:
            credibility = "academic"
        elif "government" in creds:
            credibility = "government"
        elif "news" in creds:
            credibility = "news"
    source_ids = str(raw_case_id)

    completeness = calculate_completeness(case_id=case_id)

    conn = get_connection()
    conn.execute(
        "UPDATE cases SET verification_status='verified', completeness_score=?, "
        "source_credibility=?, raw_case_ids=? WHERE id=?",
        (completeness, credibility, source_ids, case_id))
    conn.commit()
    conn.close()

    # Mark verified case as verified if it exists; also mark raw case for backwards compatibility
    verified = get_verified_case_by_raw_id(raw_case_id)
    if verified:
        set_verified_case_status(verified["id"], "verified")
    set_raw_case_status(raw_case_id, "verified")
    insert_verification_log(raw_case_id, "promote", f"Promoted to case #{case_id}", "success")

    return case_id


def sync_raw_to_case(raw_case_id):
    """Copy the latest verified (or raw) field values to the linked cases record."""
    src = _get_verified_or_raw(raw_case_id)
    if not src:
        return None
    case = _find_case_by_raw_id(raw_case_id)
    if not case:
        return None
    updates = {}
    _map = {
        "name": "name", "name_en": "name_en", "name_official": "name_official",
        "country": "country", "city": "city",
        "construction_year": "year_built", "demolition_year": "demolition_year",
        "restoration_year": "restoration_year", "quantity": "quantity",
        "function": "building_type", "designer": "architect",
        "latitude": "latitude", "longitude": "longitude",
        "image_path": "image_path", "image_urls": "image_urls",
        "building_status": "condition", "spread_method": "spread_method",
    }
    for raw_key, case_key in _map.items():
        src_val = src.get(raw_key)
        case_val = case.get(case_key)
        if str(src_val if src_val is not None else "") != str(case_val if case_val is not None else ""):
            updates[case_key] = src_val
    if updates:
        conn = get_connection()
        sets = ", ".join(f"{k}=?" for k in updates)
        vals = list(updates.values()) + [case["id"]]
        conn.execute(f"UPDATE cases SET {sets} WHERE id=?", vals)
        conn.commit()
        conn.close()
    return case["id"]


def _find_case_by_raw_id(raw_case_id):
    """Find the cases record that links to a given raw_case_id (exact match)."""
    target = str(raw_case_id)
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM cases WHERE raw_case_ids IS NOT NULL AND raw_case_ids != ''"
    ).fetchall()
    conn.close()
    for r in rows:
        ids = [x.strip() for x in (r["raw_case_ids"] or "").split(",")]
        if target in ids:
            return dict(r)
    return None


def get_display_name(case, lang="中文"):
    """Return the building name in the requested language, falling back to Chinese."""
    if lang == "English":
        return case.get("name_en") or case.get("name", "")
    elif lang == "官方语言":
        return case.get("name_official") or case.get("name", "")
    else:
        return case.get("name", "")


# --- Case Images CRUD ---

def _image_ext(image_data):
    """Detect image file extension from binary header. Returns '.jpg' as default."""
    if image_data[:4] == b'\x89PNG':
        return '.png'
    if image_data[:4] == b'RIFF':
        return '.webp'
    if image_data[:3] == b'GIF':
        return '.gif'
    if image_data[:2] == b'BM':
        return '.bmp'
    return '.jpg'


def insert_case_image(raw_case_id, image_data, filename=None, source_url=None, author=None, image_time=None):
    """Save image to filesystem, store metadata in DB. Returns image id."""
    _ensure_image_dirs()
    ext = _image_ext(image_data)
    file_name = f"{uuid.uuid4().hex}{ext}"
    file_path = os.path.join(IMAGE_DIR, file_name)
    with open(file_path, "wb") as f:
        f.write(image_data)
    conn = get_connection()
    conn.execute(
        "INSERT INTO case_images (raw_case_id, filename, image_data, file_path, source_url, author, image_time) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (raw_case_id, filename or file_name, b"", file_name, source_url, author, image_time))
    conn.commit()
    img_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.close()
    return img_id


def update_case_image(image_id, data):
    """Update metadata columns (author, image_time, filename, source_url) for an image."""
    conn = get_connection()
    valid = {"author", "image_time", "filename", "source_url"}
    filtered = {k: v for k, v in data.items() if k in valid}
    if not filtered:
        conn.close()
        return
    sets = ", ".join([f"{k} = ?" for k in filtered])
    values = list(filtered.values()) + [image_id]
    conn.execute(f"UPDATE case_images SET {sets} WHERE id = ?", values)
    conn.commit()
    conn.close()


def get_image_metadata(raw_case_id, src):
    """Get {author, image_time} for a specific image src from raw_cases.image_metadata JSON."""
    conn = get_connection()
    row = conn.execute("SELECT image_metadata FROM raw_cases WHERE id = ?", (raw_case_id,)).fetchone()
    conn.close()
    if row and row["image_metadata"]:
        try:
            all_meta = json.loads(row["image_metadata"])
            return all_meta.get(src, {})
        except (json.JSONDecodeError, TypeError):
            pass
    return {}


def set_image_metadata(raw_case_id, src, author, image_time):
    """Save {author, image_time} for a specific image src to raw_cases.image_metadata JSON."""
    conn = get_connection()
    row = conn.execute("SELECT image_metadata FROM raw_cases WHERE id = ?", (raw_case_id,)).fetchone()
    all_meta = {}
    if row and row["image_metadata"]:
        try:
            all_meta = json.loads(row["image_metadata"])
        except (json.JSONDecodeError, TypeError):
            all_meta = {}
    all_meta[src] = {"author": author, "image_time": image_time}
    conn.execute("UPDATE raw_cases SET image_metadata = ? WHERE id = ?",
                 (json.dumps(all_meta, ensure_ascii=False), raw_case_id))
    conn.commit()
    conn.close()


def get_case_image(image_id):
    """Return {id, raw_case_id, filename, image_data, file_path, source_url} or None.
    Reads image_data from filesystem; falls back to DB BLOB for legacy images."""
    conn = get_connection()
    row = conn.execute("SELECT * FROM case_images WHERE id = ?", (image_id,)).fetchone()
    conn.close()
    if not row:
        return None
    rec = dict(row)
    fp = rec.get("file_path")
    if fp:
        fpath = os.path.join(IMAGE_DIR, fp)
        if os.path.exists(fpath):
            with open(fpath, "rb") as f:
                rec["image_data"] = f.read()
    # Legacy: image_data already in BLOB, file_path is NULL
    return rec


def get_images_for_raw_case(raw_case_id):
    """Return list of image metadata (without file contents) for a raw case."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT id, raw_case_id, filename, file_path, source_url, created_at, author, image_time FROM case_images WHERE raw_case_id = ? ORDER BY id",
        (raw_case_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_images_for_raw_case(raw_case_id):
    """Delete all image files and records for a raw case."""
    conn = get_connection()
    rows = conn.execute("SELECT file_path FROM case_images WHERE raw_case_id = ?", (raw_case_id,)).fetchall()
    for r in rows:
        if r["file_path"]:
            _remove_file(IMAGE_DIR, r["file_path"])
            _remove_file(IMAGE_TRASH_DIR, r["file_path"])
    conn.execute("DELETE FROM case_images WHERE raw_case_id = ?", (raw_case_id,))
    conn.commit()
    conn.close()


def delete_case_image(image_id):
    """Delete a single image file and record by id."""
    conn = get_connection()
    row = conn.execute("SELECT file_path FROM case_images WHERE id = ?", (image_id,)).fetchone()
    if row and row["file_path"]:
        _remove_file(IMAGE_DIR, row["file_path"])
        _remove_file(IMAGE_TRASH_DIR, row["file_path"])
    conn.execute("DELETE FROM case_images WHERE id = ?", (image_id,))
    conn.commit()
    conn.close()


def soft_delete_case_image(image_id):
    """Move an image file to trash dir. Returns trash record id, or None."""
    _ensure_image_dirs()
    conn = get_connection()
    row = conn.execute("SELECT raw_case_id, filename, file_path, source_url, author, image_time FROM case_images WHERE id = ?",
                       (image_id,)).fetchone()
    if not row:
        conn.close()
        return None
    fp = row["file_path"]
    # Move file to trash
    if fp:
        src = os.path.join(IMAGE_DIR, fp)
        dst = os.path.join(IMAGE_TRASH_DIR, fp)
        if os.path.exists(src):
            shutil.move(src, dst)
    conn.execute(
        "INSERT INTO image_trash (raw_case_id, filename, image_data, file_path, source_url, author, image_time) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (row["raw_case_id"], row["filename"], b"", fp, row["source_url"], row["author"], row["image_time"]))
    trash_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.execute("DELETE FROM case_images WHERE id = ?", (image_id,))

    # Keep max 10 trash entries per raw_case
    keep = 10
    conn.execute(
        "DELETE FROM image_trash WHERE raw_case_id = ? AND id NOT IN ("
        "SELECT id FROM image_trash WHERE raw_case_id = ? ORDER BY deleted_at DESC LIMIT ?)",
        (row["raw_case_id"], row["raw_case_id"], keep))

    conn.commit()
    conn.close()
    return trash_id


def undo_delete_image(raw_case_id):
    """Restore the most recently deleted image for a raw_case. Returns new image id or None."""
    _ensure_image_dirs()
    conn = get_connection()
    row = conn.execute(
        "SELECT id, filename, file_path, source_url, author, image_time FROM image_trash WHERE raw_case_id = ? ORDER BY deleted_at DESC LIMIT 1",
        (raw_case_id,)).fetchone()
    if not row:
        conn.close()
        return None
    fp = row["file_path"]
    # Move file back from trash
    if fp:
        src = os.path.join(IMAGE_TRASH_DIR, fp)
        dst = os.path.join(IMAGE_DIR, fp)
        if os.path.exists(src):
            shutil.move(src, dst)
    conn.execute(
        "INSERT INTO case_images (raw_case_id, filename, image_data, file_path, source_url, author, image_time) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (raw_case_id, row["filename"], b"", fp, row["source_url"], row["author"], row["image_time"]))
    new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.execute("DELETE FROM image_trash WHERE id = ?", (row["id"],))
    conn.commit()
    conn.close()
    return new_id


def _remove_file(directory, file_name):
    """Safely remove a file from a directory."""
    path = os.path.join(directory, file_name)
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass


# ============================================================
#  Undo / Redo Stack
# ============================================================

_undo_stack = []
_redo_stack = []
_MAX_UNDO = 10


def push_undo(action, data):
    """Push an undoable action onto the stack. Clears redo stack on new action."""
    _redo_stack.clear()
    _undo_stack.append({"action": action, "data": copy.deepcopy(data)})
    if len(_undo_stack) > _MAX_UNDO:
        _undo_stack.pop(0)


def pop_undo():
    """Pop the most recent undo action. Returns (action, data) or None."""
    if not _undo_stack:
        return None
    return _undo_stack.pop()


def get_undo_count():
    return len(_undo_stack)


def get_redo_count():
    return len(_redo_stack)


def _execute_reverse(entry):
    """Execute the reverse of an action. Returns (description, reverse_entry) or (None, None)."""
    action = entry["action"]
    data = entry["data"]
    conn = get_connection()
    reverse_entry = None

    try:
        if action == "insert":
            # Undo insert: read full case before deleting, then delete
            row = conn.execute("SELECT * FROM raw_cases WHERE id = ?", (data["id"],)).fetchone()
            if row:
                case = dict(row)
                sources = [dict(s) for s in conn.execute(
                    "SELECT * FROM case_sources WHERE raw_case_id = ?", (data["id"],)).fetchall()]
                logs = [dict(l) for l in conn.execute(
                    "SELECT * FROM verification_log WHERE raw_case_id = ?", (data["id"],)).fetchall()]
                # Remove id from case dict (it will be auto-generated on re-insert)
                case.pop("id", None)
                reverse_entry = {"action": "delete", "data": {"case": case, "sources": sources, "logs": logs}}
            conn.execute("DELETE FROM raw_cases WHERE id = ?", (data["id"],))
            conn.commit()
            return (f"撤销添加: {data.get('name', '')}", reverse_entry)

        elif action == "delete":
            # Undo delete: re-insert the deleted case
            restored = data["case"]
            fields = [f for f in _get_all_raw_case_columns(conn) if f in restored]
            placeholders = ", ".join(["?" for _ in fields])
            columns = ", ".join(fields)
            values = [restored.get(f) for f in fields]
            conn.execute(f"INSERT INTO raw_cases ({columns}) VALUES ({placeholders})", values)
            new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

            for src in data.get("sources", []):
                conn.execute(
                    "INSERT INTO case_sources (raw_case_id, url, title, credibility, fetch_time) VALUES (?, ?, ?, ?, ?)",
                    (new_id, src["url"], src.get("title"), src.get("credibility"), src.get("fetch_time")))
            for log in data.get("logs", []):
                conn.execute(
                    "INSERT INTO verification_log (raw_case_id, action, details, result, created_at) VALUES (?, ?, ?, ?, ?)",
                    (new_id, log["action"], log.get("details"), log.get("result"), log.get("created_at")))
            conn.commit()
            reverse_entry = {"action": "insert", "data": {"id": new_id, "name": restored.get("name", "")}}
            return (f"撤销删除: {restored.get('name', '')}", reverse_entry)

        elif action == "update":
            case_id = data["id"]
            old_values = data["old_values"]
            # Read current values before undoing (for redo)
            row = conn.execute("SELECT * FROM raw_cases WHERE id = ?", (case_id,)).fetchone()
            current_values = {}
            if row:
                for k in old_values:
                    current_values[k] = row[k]
            sets = ", ".join(f"{k} = ?" for k in old_values)
            values = list(old_values.values()) + [case_id]
            conn.execute(f"UPDATE raw_cases SET {sets}, updated_at = CURRENT_TIMESTAMP WHERE id = ?", values)
            conn.commit()
            if current_values:
                reverse_entry = {"action": "update", "data": {"id": case_id, "name": data.get("name", ""),
                                                               "old_values": current_values}}
            return (f"撤销修改: {data.get('name', '')}", reverse_entry)

        elif action == "status_change":
            # Read current status before undoing
            row = conn.execute("SELECT authenticity_status FROM raw_cases WHERE id = ?",
                               (data["id"],)).fetchone()
            current_status = row["authenticity_status"] if row else None
            conn.execute("UPDATE raw_cases SET authenticity_status = ? WHERE id = ?",
                         (data["old_status"], data["id"]))
            conn.commit()
            if current_status is not None:
                reverse_entry = {"action": "status_change", "data": {"id": data["id"], "name": data.get("name", ""),
                                                                      "old_status": current_status}}
            return (f"撤销状态变更: {data.get('name', '')}", reverse_entry)

        elif action == "merge":
            secondary = data["secondary"]
            primary_id = data["primary_id"]
            # Re-insert secondary
            sec_fields = [f for f in _get_all_raw_case_columns(conn) if f in secondary]
            sec_ph = ", ".join(["?" for _ in sec_fields])
            sec_cols = ", ".join(sec_fields)
            sec_vals = [secondary.get(f) for f in sec_fields]
            conn.execute(f"INSERT INTO raw_cases ({sec_cols}) VALUES ({sec_ph})", sec_vals)
            new_sec_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

            # Revert primary to old values
            if data.get("primary_old"):
                old = data["primary_old"]
                p_sets = ", ".join(f"{k} = ?" for k in old)
                p_vals = list(old.values()) + [primary_id]
                conn.execute(f"UPDATE raw_cases SET {p_sets} WHERE id = ?", p_vals)

            # Move back related data
            for table, id_col in [("case_sources", "raw_case_id"), ("verification_log", "raw_case_id"),
                                   ("case_images", "raw_case_id"), ("field_sources", "raw_case_id")]:
                conn.execute(f"UPDATE {table} SET {id_col} = ? WHERE {id_col} = ? AND id IN (SELECT id FROM {table} WHERE {id_col} = ? ORDER BY id DESC LIMIT (SELECT COUNT(*) FROM {table} WHERE {id_col} = ?))",
                             (new_sec_id, primary_id))
            conn.commit()
            # Redo entry for merge: store the restored secondary id so redo can re-merge
            reverse_entry = {"action": "merge_undo", "data": {"secondary_id": new_sec_id, "primary_id": primary_id,
                                                               "secondary": secondary, "primary_old": data.get("primary_old")}}
            return (f"撤销合并: {secondary.get('name', '')}", reverse_entry)

        elif action == "merge_undo":
            # Redo merge: delete the restored secondary and re-merge into primary
            sec_id = data["secondary_id"]
            primary_id = data["primary_id"]
            # Read current primary values before merge
            row = conn.execute("SELECT * FROM raw_cases WHERE id = ?", (primary_id,)).fetchone()
            primary_old = dict(row) if row else {}
            primary_old.pop("id", None)
            # Read secondary before deleting
            sec_row = conn.execute("SELECT * FROM raw_cases WHERE id = ?", (sec_id,)).fetchone()
            secondary_data = dict(sec_row) if sec_row else {}
            secondary_data.pop("id", None)
            # Move secondary's related data to primary
            for table, id_col in [("case_sources", "raw_case_id"), ("verification_log", "raw_case_id"),
                                   ("case_images", "raw_case_id"), ("field_sources", "raw_case_id")]:
                conn.execute(f"UPDATE {table} SET {id_col} = ? WHERE {id_col} = ?",
                             (primary_id, sec_id))
            # Merge secondary fields into primary (non-empty values overwrite)
            for k, v in secondary_data.items():
                if v and not primary_old.get(k):
                    conn.execute(f"UPDATE raw_cases SET {k} = ? WHERE id = ?", (v, primary_id))
            # Delete the secondary
            conn.execute("DELETE FROM raw_cases WHERE id = ?", (sec_id,))
            conn.commit()
            reverse_entry = {"action": "merge", "data": {"secondary": secondary_data, "primary_id": primary_id,
                                                          "primary_old": primary_old}}
            return (f"恢复合并: {secondary_data.get('name', '')}", reverse_entry)

        elif action == "import_batch":
            ids = data["ids"]
            for i in ids:
                conn.execute("DELETE FROM raw_cases WHERE id = ?", (i,))
            conn.commit()
            # No redo for import batch (data is gone)
            return (f"撤销导入: {len(ids)} 条记录", None)
    finally:
        conn.close()
    return (None, None)


def undo_last():
    """Undo the most recent operation. Returns description string or None."""
    entry = pop_undo()
    if not entry:
        return None
    desc, reverse_entry = _execute_reverse(entry)
    if reverse_entry:
        _redo_stack.append(reverse_entry)
        if len(_redo_stack) > _MAX_UNDO:
            _redo_stack.pop(0)
    return desc


def redo_last():
    """Redo the most recently undone operation. Returns description string or None."""
    if not _redo_stack:
        return None
    entry = _redo_stack.pop()
    desc, reverse_entry = _execute_reverse(entry)
    if reverse_entry:
        _undo_stack.append(reverse_entry)
        if len(_undo_stack) > _MAX_UNDO:
            _undo_stack.pop(0)
    return desc


# =============================================================================
#  Literature CRUD
# =============================================================================

_LITERATURE_COLS = [
    "id", "title", "authors", "pub_date", "source_url", "source_engine",
    "content", "summary", "file_path", "notes", "created_at",
    "journal", "city", "country", "translator", "publisher",
]


def _get_literature_cols(conn):
    actual = set(r[1] for r in conn.execute("PRAGMA table_info(literature)").fetchall())
    return [c for c in _LITERATURE_COLS if c in actual and c != "id"]


def get_all_literature(search=None, order_by="created_at DESC"):
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    if search:
        q = f"%{search}%"
        rows = conn.execute(
            f"SELECT * FROM literature WHERE title LIKE ? OR authors LIKE ? OR "
            f"summary LIKE ? OR notes LIKE ? ORDER BY {order_by}",
            (q, q, q, q),
        ).fetchall()
    else:
        rows = conn.execute(
            f"SELECT * FROM literature ORDER BY {order_by}"
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_literature_by_id(lit_id):
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT * FROM literature WHERE id = ?", (lit_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def insert_literature(data):
    conn = get_connection()
    cols = _get_literature_cols(conn)
    filtered = {k: v for k, v in data.items() if k in cols and k != "id"}
    if not filtered:
        conn.close()
        return None
    placeholders = ", ".join(["?"] * len(filtered))
    columns = ", ".join(filtered.keys())
    cur = conn.execute(
        f"INSERT INTO literature ({columns}) VALUES ({placeholders})",
        list(filtered.values()),
    )
    conn.commit()
    lid = cur.lastrowid
    conn.close()
    return lid


def update_literature(lit_id, data):
    conn = get_connection()
    cols = _get_literature_cols(conn)
    filtered = {k: v for k, v in data.items() if k in cols and k != "id"}
    if not filtered:
        conn.close()
        return
    set_clause = ", ".join(f"{k} = ?" for k in filtered)
    conn.execute(
        f"UPDATE literature SET {set_clause} WHERE id = ?",
        list(filtered.values()) + [lit_id],
    )
    conn.commit()
    conn.close()


def delete_literature(lit_id):
    conn = get_connection()
    row = conn.execute(
        "SELECT file_path FROM literature WHERE id = ?", (lit_id,)
    ).fetchone()
    if row and row[0]:
        fp = os.path.join(DB_DIR, "literature", row[0])
        if os.path.exists(fp):
            os.remove(fp)
    conn.execute("DELETE FROM literature WHERE id = ?", (lit_id,))
    conn.commit()
    conn.close()


def search_literature_by_title(title):
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT * FROM literature WHERE title = ?", (title,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_lit_matches(match_type, matches):
    """Save a new batch of literature match results. Returns batch_id.
    match_type: 'main' or 'raw'.
    matches: list of (lit_record, matched_title_or_record, similarity)."""
    from datetime import datetime
    conn = get_connection()
    conn.execute(
        "CREATE TABLE IF NOT EXISTS lit_match_log ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT,"
        "match_type TEXT NOT NULL,"
        "lit_title TEXT,"
        "matched_title TEXT,"
        "similarity REAL,"
        "lit_authors TEXT,"
        "lit_pub_date TEXT,"
        "batch_id TEXT,"
        "created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)"
    )
    try:
        conn.execute("ALTER TABLE lit_match_log ADD COLUMN batch_id TEXT")
    except sqlite3.OperationalError:
        pass
    batch_id = datetime.now().strftime("%m-%d %H:%M:%S")
    for lit, matched, sim in matches:
        matched_title = matched if isinstance(matched, str) else matched.get("title", "")
        conn.execute(
            "INSERT INTO lit_match_log (match_type, lit_title, matched_title, similarity, lit_authors, lit_pub_date, batch_id) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (match_type, lit.get("title", ""), matched_title, sim,
             lit.get("authors", "") or "", lit.get("pub_date", "") or "", batch_id)
        )
    conn.commit()
    conn.close()
    return batch_id


def get_lit_match_batches(match_type):
    """Get list of {batch_id, count, created_at} for a match type, newest first."""
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("ALTER TABLE lit_match_log ADD COLUMN batch_id TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        rows = conn.execute(
            "SELECT batch_id, COUNT(*) as cnt, MAX(created_at) as created_at "
            "FROM lit_match_log WHERE match_type = ? AND batch_id IS NOT NULL "
            "GROUP BY batch_id ORDER BY created_at DESC",
            (match_type,)
        ).fetchall()
        return [dict(r) for r in rows]
    except sqlite3.OperationalError:
        return []
    finally:
        conn.close()


def get_lit_matches(match_type, batch_id=None):
    """Get saved literature match results. If batch_id given, filter by it."""
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("ALTER TABLE lit_match_log ADD COLUMN batch_id TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        if batch_id:
            rows = conn.execute(
                "SELECT * FROM lit_match_log WHERE match_type = ? AND batch_id = ? ORDER BY similarity DESC",
                (match_type, batch_id)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM lit_match_log WHERE match_type = ? ORDER BY similarity DESC",
                (match_type,)
            ).fetchall()
        return [dict(r) for r in rows]
    except sqlite3.OperationalError:
        return []
    finally:
        conn.close()


def has_lit_matches(match_type):
    """Check if there are saved matches for this type (with valid batch_id)."""
    conn = get_connection()
    try:
        conn.execute("ALTER TABLE lit_match_log ADD COLUMN batch_id TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cnt = conn.execute(
            "SELECT COUNT(*) FROM lit_match_log WHERE match_type = ? AND batch_id IS NOT NULL",
            (match_type,)
        ).fetchone()[0]
        return cnt > 0
    except sqlite3.OperationalError:
        return False
    finally:
        conn.close()


def clear_lit_matches(match_type):
    """Delete all match records for a given type."""
    conn = get_connection()
    conn.execute("DELETE FROM lit_match_log WHERE match_type = ?", (match_type,))
    conn.commit()
    conn.close()


# ---- Serper search results ----

def insert_filtered_results(query, results):
    """Bulk insert Serper search results. Also updates relevance_reason in raw_search_results."""
    if not results:
        return
    conn = get_connection()
    conn.executemany(
        "INSERT INTO filtered_results (query, title, snippet, link, relevance_reason) "
        "VALUES (?, ?, ?, ?, ?)",
        [(query, r["title"], r.get("snippet", ""), r.get("link", ""),
          r.get("relevance_reason", "")) for r in results]
    )
    # Also persist reason in raw_search_results so it survives dedup deletes
    for r in results:
        link = r.get("link", "")
        reason = r.get("relevance_reason", "")
        if link and reason:
            conn.execute(
                "UPDATE raw_search_results SET relevance_reason=? "
                "WHERE link=? AND (relevance_reason IS NULL OR relevance_reason='')",
                (reason, link)
            )
    conn.commit()
    conn.close()


def insert_raw_results(query, engine, results):
    """Save raw search results before AI filtering."""
    if not results:
        return
    conn = get_connection()
    conn.executemany(
        "INSERT INTO raw_search_results (query, engine, title, snippet, link, "
        "authors, year, publication, pdf_url, cited_by) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        [(query, engine, r["title"], r.get("snippet", ""), r.get("link", ""),
          r.get("authors", ""), str(r.get("year", "")),
          r.get("publication", ""), r.get("pdfUrl", ""),
          str(r.get("citedBy", "")))
         for r in results]
    )
    conn.commit()
    conn.close()


def get_raw_result_queries():
    """Get distinct search queries for raw results."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT DISTINCT query, engine, COUNT(*) as cnt, MAX(created_at) as last_time "
        "FROM raw_search_results GROUP BY query, engine ORDER BY last_time DESC"
    ).fetchall()
    conn.close()
    return [{"query": r[0], "engine": r[1], "count": r[2], "last_time": r[3]} for r in rows]


def insert_dedup_removed(links, query=None):
    """Record links that were removed by deduplication."""
    if not links:
        return
    conn = get_connection()
    conn.executemany(
        "INSERT OR IGNORE INTO dedup_removed (link, query) VALUES (?, ?)",
        [(link, query) for link in links]
    )
    conn.commit()
    conn.close()


def mark_raw_results_dedup_status(link, dedup_status):
    """Update dedup_status for records in raw_search_results matching a link."""
    if not link or not dedup_status:
        return
    conn = get_connection()
    conn.execute(
        "UPDATE raw_search_results SET dedup_status=? WHERE link=? AND dedup_status=''",
        (dedup_status, link)
    )
    conn.commit()
    conn.close()


def delete_dedup_removed_by_links(links):
    """Remove entries from dedup_removed by links (for undo)."""
    if not links:
        return
    conn = get_connection()
    for link in links:
        conn.execute("DELETE FROM dedup_removed WHERE link=?", (link,))
    conn.commit()
    conn.close()


def get_dedup_removed_links():
    """Return set of all dedup-removed links."""
    conn = get_connection()
    rows = conn.execute("SELECT link FROM dedup_removed").fetchall()
    conn.close()
    return {r[0] for r in rows}


def get_raw_results_by_query(query, engine):
    """Get raw search results for a specific query and engine,
    with selection status and AI reason from raw_search_results,
    and dedup removal status from dedup_removed."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT r.id, r.title, r.snippet, r.link, r.created_at, "
        "  r.authors, r.year, r.publication, r.pdf_url, r.cited_by, "
        "  CASE WHEN r.status = 'cleaned' THEN '已清洗' "
        "       WHEN r.status = 'deduped' THEN '已去重' "
        "       WHEN r.status = '转为文献' THEN '转为文献' "
        "       WHEN d.link IS NOT NULL THEN '已筛选(去重移除)' "
        "       WHEN r.relevance_reason != '' THEN '已筛选' "
        "       ELSE '未选中' END as filter_status, "
        "  COALESCE(r.relevance_reason, '') as reason, "
        "  COALESCE(r.status, 'active') as clean_status "
        "FROM raw_search_results r "
        "LEFT JOIN dedup_removed d ON r.link = d.link "
        "WHERE r.query=? AND r.engine=? "
        "ORDER BY r.id",
        (query, engine)
    ).fetchall()
    conn.close()
    return [{"id": r[0], "title": r[1], "snippet": r[2], "link": r[3],
             "created_at": r[4], "authors": r[5], "year": r[6],
             "publication": r[7], "pdf_url": r[8], "cited_by": r[9],
             "status": r[10], "reason": r[11], "clean_status": r[12]} for r in rows]


def get_raw_result_by_link(link):
    """Get a single raw search result by link. Returns dict or None."""
    if not link:
        return None
    conn = get_connection()
    row = conn.execute(
        "SELECT id, title, snippet, link, authors, year, publication, pdf_url, cited_by, engine "
        "FROM raw_search_results WHERE link = ? LIMIT 1",
        (link,)
    ).fetchone()
    conn.close()
    if not row:
        return None
    return {"id": row[0], "title": row[1], "snippet": row[2], "link": row[3],
            "authors": row[4], "year": row[5], "publication": row[6],
            "pdf_url": row[7], "cited_by": row[8], "engine": row[9]}


def mark_raw_result_status(link, status):
    """Set the status column for a raw search result by link."""
    conn = get_connection()
    conn.execute(
        "UPDATE raw_search_results SET status=? WHERE link=?",
        (status, link)
    )
    conn.commit()
    conn.close()


def mark_raw_results_cleaned(link, query=None):
    """Mark a raw result as cleaned (manually removed)."""
    conn = get_connection()
    conn.execute(
        "UPDATE raw_search_results SET status='cleaned' WHERE link=?",
        (link,)
    )
    conn.commit()
    conn.close()


def mark_raw_results_batch_cleaned(links, query=None):
    """Batch mark raw results as cleaned."""
    if not links:
        return
    conn = get_connection()
    for link in links:
        conn.execute(
            "UPDATE raw_search_results SET status='cleaned' WHERE link=?",
            (link,)
        )
    conn.commit()
    conn.close()


def unmark_raw_results_cleaned(links, titles):
    """Revert cleaned status back to active (for undo)."""
    if not links and not titles:
        return
    conn = get_connection()
    for link in links:
        conn.execute(
            "UPDATE raw_search_results SET status='active' WHERE link=? AND status='cleaned'",
            (link,)
        )
    for title in titles:
        conn.execute(
            "UPDATE raw_search_results SET status='active' WHERE title=? AND status='cleaned'",
            (title,)
        )
    conn.commit()
    conn.close()


def mark_raw_results_batch_deduped(links, query=None):
    """Batch mark raw results as deduped."""
    if not links:
        return
    conn = get_connection()
    for link in links:
        conn.execute(
            "UPDATE raw_search_results SET status='deduped' WHERE link=?",
            (link,)
        )
    conn.commit()
    conn.close()


def mark_raw_results_batch_cleaned_by_titles(titles, query=None):
    """Batch mark raw results as cleaned, matching by title (fallback for empty links)."""
    if not titles:
        return
    conn = get_connection()
    for title in titles:
        conn.execute(
            "UPDATE raw_search_results SET status='cleaned' WHERE title=? AND status IN ('active','ds_selected')",
            (title,)
        )
    conn.commit()
    conn.close()


def mark_raw_results_batch_deduped_by_titles(titles, query=None):
    """Batch mark raw results as deduped, matching by title (fallback for empty links)."""
    if not titles:
        return
    conn = get_connection()
    for title in titles:
        conn.execute(
            "UPDATE raw_search_results SET status='deduped' WHERE title=? AND status IN ('active','ds_selected')",
            (title,)
        )
    conn.commit()
    conn.close()


def mark_raw_results_batch_ds_selected(links, titles, query=None):
    """Mark raw results as DS-selected (passed DeepSeek screening).
    Uses links first, falls back to titles for items with empty links."""
    if not links and not titles:
        return
    conn = get_connection()
    for link in links:
        conn.execute(
            "UPDATE raw_search_results SET status='ds_selected' WHERE link=?",
            (link,)
        )
    for title in titles:
        conn.execute(
            "UPDATE raw_search_results SET status='ds_selected' WHERE title=? AND status='active'",
            (title,)
        )
    conn.commit()
    conn.close()


def delete_raw_results_by_ids(ids):
    """Hard delete raw search results by IDs."""
    if not ids:
        return
    conn = get_connection()
    placeholders = ",".join("?" for _ in ids)
    conn.execute(f"DELETE FROM raw_search_results WHERE id IN ({placeholders})", ids)
    conn.commit()
    conn.close()


# ---- Scholar search results (literature screening) ----

def insert_scholar_results(query, results, screen_engine=""):
    """Save raw Scholar search results to scholar_search_results table."""
    if not results:
        return
    conn = get_connection()
    conn.executemany(
        "INSERT INTO scholar_search_results (query, title, authors, year, publication, "
        "snippet, link, pdf_url, cited_by, screen_engine) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        [(query, r["title"], r.get("authors", ""), str(r.get("year", "")),
          r.get("publication", ""), r.get("snippet", ""), r.get("link", ""),
          r.get("pdfUrl", ""), str(r.get("citedBy", "")), screen_engine)
         for r in results]
    )
    conn.commit()
    conn.close()


def get_scholar_result_queries():
    """Get distinct search queries (grouped by query + screen_engine) from scholar_search_results."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT query, screen_engine, COUNT(*) as cnt, MAX(created_at) as last_time "
        "FROM scholar_search_results GROUP BY query, screen_engine ORDER BY last_time DESC"
    ).fetchall()
    conn.close()
    return [{"query": r[0], "screen_engine": r[1], "count": r[2], "last_time": r[3]} for r in rows]


def get_scholar_results_by_query(query, screen_engine=None):
    """Get scholar search results for a specific query and optional screen_engine."""
    conn = get_connection()
    if screen_engine:
        rows = conn.execute(
            "SELECT id, title, authors, year, publication, snippet, link, "
            "pdf_url, cited_by, status, gemini_reason, created_at, screen_engine, tags "
            "FROM scholar_search_results WHERE query=? AND screen_engine=? ORDER BY id",
            (query, screen_engine)
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT id, title, authors, year, publication, snippet, link, "
            "pdf_url, cited_by, status, gemini_reason, created_at, screen_engine, tags "
            "FROM scholar_search_results WHERE query=? ORDER BY id",
            (query,)
        ).fetchall()
    conn.close()
    return [{"id": r[0], "title": r[1], "authors": r[2], "year": r[3],
             "publication": r[4], "snippet": r[5], "link": r[6],
             "pdf_url": r[7], "cited_by": r[8], "clean_status": r[9],
             "gemini_reason": r[10], "created_at": r[11], "screen_engine": r[12],
             "tags": r[13]} for r in rows]


def mark_scholar_results_gemini_selected(links, titles, reasons=None):
    """Mark scholar results as gemini_selected (passed Gemini screening).
    Optionally store screening reasons.
    reasons: dict mapping link/title -> reason string.
    """
    if not links and not titles:
        return
    conn = get_connection()
    for link in links:
        reason = (reasons or {}).get(link, "") if reasons else ""
        conn.execute(
            "UPDATE scholar_search_results SET status='gemini_selected', gemini_reason=? WHERE link=?",
            (reason, link)
        )
    for title in titles:
        reason = (reasons or {}).get(title, "") if reasons else ""
        conn.execute(
            "UPDATE scholar_search_results SET status='gemini_selected', gemini_reason=? "
            "WHERE title=? AND status='active'",
            (reason, title)
        )
    conn.commit()
    conn.close()


def mark_scholar_results_status(links, titles, status):
    """Batch update status for scholar results (cleaned/deduped).
    Uses links first, falls back to titles for items with empty links."""
    if not links and not titles:
        return
    conn = get_connection()
    for link in links:
        conn.execute(
            "UPDATE scholar_search_results SET status=? WHERE link=?",
            (status, link)
        )
    for title in titles:
        conn.execute(
            "UPDATE scholar_search_results SET status=? WHERE title=? "
            "AND status IN ('active','gemini_selected')",
            (status, title)
        )
    conn.commit()
    conn.close()


def restore_raw_results_by_ids(ids):
    """Restore raw results from cleaned/deduped back to active.
    Also removes from dedup_removed table."""
    if not ids:
        return
    conn = get_connection()
    for id_ in ids:
        conn.execute(
            "UPDATE raw_search_results SET status='active' WHERE id=? AND status IN ('cleaned','deduped')",
            (id_,)
        )
    conn.commit()
    conn.close()


def restore_raw_results_full(ids, links, queries):
    """Full restore: set status to active, remove from dedup_removed,
    and re-insert into filtered_results if previously there.
    queries maps id -> query_label."""
    if not ids:
        return
    conn = get_connection()
    for id_ in ids:
        conn.execute(
            "UPDATE raw_search_results SET status='active' WHERE id=? AND status IN ('cleaned','deduped')",
            (id_,)
        )
    for link in links:
        if link:
            conn.execute("DELETE FROM dedup_removed WHERE link=?", (link,))
    conn.commit()
    conn.close()


def restore_scholar_results_by_ids(ids):
    """Restore scholar results from cleaned/deduped back to active."""
    if not ids:
        return
    conn = get_connection()
    for id_ in ids:
        conn.execute(
            "UPDATE scholar_search_results SET status='active' WHERE id=? AND status IN ('cleaned','deduped')",
            (id_,)
        )
    conn.commit()
    conn.close()


def delete_scholar_results_by_ids(ids):
    """Hard delete scholar search results by IDs."""
    if not ids:
        return
    conn = get_connection()
    placeholders = ",".join("?" for _ in ids)
    conn.execute(f"DELETE FROM scholar_search_results WHERE id IN ({placeholders})", ids)
    conn.commit()
    conn.close()


def get_scholar_results_missing_reason():
    """Get gemini_selected scholar results with empty gemini_reason."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT id, title, authors, year, publication, snippet "
        "FROM scholar_search_results "
        "WHERE status='gemini_selected' AND (gemini_reason IS NULL OR gemini_reason='') "
        "ORDER BY id"
    ).fetchall()
    conn.close()
    return [{"id": r[0], "title": r[1], "authors": r[2], "year": r[3],
             "publication": r[4], "snippet": r[5]} for r in rows]


def update_scholar_result_reason(record_id, reason):
    """Update gemini_reason for a single scholar result."""
    conn = get_connection()
    conn.execute(
        "UPDATE scholar_search_results SET gemini_reason=? WHERE id=?",
        (reason, record_id)
    )
    conn.commit()
    conn.close()


def get_all_filtered_results(search=None):
    """Get all Serper results, optionally filtered."""
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    if search:
        rows = conn.execute(
            "SELECT * FROM filtered_results WHERE query LIKE ? OR title LIKE ? "
            "ORDER BY id DESC",
            (f"%{search}%", f"%{search}%")
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM filtered_results ORDER BY id DESC"
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_filtered_results_by_ids(ids):
    """Delete filtered_results rows matching the given primary key ids."""
    if not ids:
        return
    conn = get_connection()
    placeholders = ",".join(["?" for _ in ids])
    conn.execute(f"DELETE FROM filtered_results WHERE id IN ({placeholders})", ids)
    conn.commit()
    conn.close()


def delete_filtered_results_by_links(links):
    """Delete filtered_results rows matching the given links."""
    if not links:
        return
    conn = get_connection()
    placeholders = ",".join(["?" for _ in links])
    conn.execute(f"DELETE FROM filtered_results WHERE link IN ({placeholders})", links)
    conn.commit()
    conn.close()


def delete_filtered_result(result_id):
    conn = get_connection()
    conn.execute("DELETE FROM filtered_results WHERE id = ?", (result_id,))
    conn.commit()
    conn.close()


def mark_raw_results_cleaned_by_links(links):
    """Mark raw_search_results as 'cleaned' by link."""
    if not links:
        return
    conn = get_connection()
    for link in links:
        if link:
            conn.execute(
                "UPDATE raw_search_results SET status='cleaned' WHERE link=? AND status='active'",
                (link,)
            )
    conn.commit()
    conn.close()


def clear_filtered_results():
    conn = get_connection()
    conn.execute("DELETE FROM filtered_results")
    conn.commit()
    conn.close()


def get_filtered_queries():
    """Get distinct queries with result counts."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT DISTINCT query, COUNT(*) as cnt FROM filtered_results "
        "GROUP BY query ORDER BY MAX(id) DESC"
    ).fetchall()
    conn.close()
    return [(r[0], r[1]) for r in rows]


def lookup_search_keywords_for_url(url):
    """Look up search keywords from raw_search_results for a given URL.
    Returns the query string that was used to find this URL, or None."""
    if not url:
        return None
    conn = get_connection()
    row = conn.execute(
        "SELECT DISTINCT query FROM raw_search_results WHERE link = ? LIMIT 1",
        (url,)
    ).fetchone()
    # Also try scholar_search_results
    if not row:
        row = conn.execute(
            "SELECT DISTINCT query FROM scholar_search_results WHERE link = ? LIMIT 1",
            (url,)
        ).fetchone()
    conn.close()
    return row[0] if row else None


def batch_update_search_keywords():
    """Update search_keywords for raw cases by looking up their source URLs
    in raw_search_results. Only updates when a match is found.
    Imported cases typically don't have URLs in the main list so they're unaffected."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT id, source_urls, search_keywords FROM raw_cases "
        "WHERE source_urls IS NOT NULL AND source_urls != ''"
    ).fetchall()
    updated = 0
    for case_id, source_urls_json, current_kw in rows:
        try:
            urls = json.loads(source_urls_json)
        except (json.JSONDecodeError, TypeError):
            urls = [source_urls_json]
        if not isinstance(urls, list):
            urls = [urls] if urls else []
        for url in urls:
            if not url:
                continue
            q = conn.execute(
                "SELECT DISTINCT query FROM raw_search_results WHERE link = ? LIMIT 1",
                (url,)
            ).fetchone()
            if not q:
                q = conn.execute(
                    "SELECT DISTINCT query FROM scholar_search_results WHERE link = ? LIMIT 1",
                    (url,)
                ).fetchone()
            if q and q[0] != current_kw:
                conn.execute(
                    "UPDATE raw_cases SET search_keywords = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                    (q[0], case_id)
                )
                updated += 1
                break
    conn.commit()
    conn.close()
    return updated

