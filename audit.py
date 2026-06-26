import sqlite3
import json
from datetime import datetime
from pathlib import Path

DB_PATH = Path("audit.db")


def _conn():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with _conn() as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS conversions (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                ts           TEXT    NOT NULL,
                ip           TEXT,
                invoice_no   TEXT,
                original_qty INTEGER,
                final_qty    INTEGER,
                deleted_qty  INTEGER,
                deleted_items TEXT,
                iva_rate     REAL,
                tasa_bcv     REAL,
                total_usd    REAL,
                total_bs     REAL
            )
        """)


def log_conversion(ip, invoice_no, original_qty, final_items, iva_rate, tasa_bcv):
    final_qty   = len(final_items)
    deleted_qty = original_qty - final_qty
    subtotal    = round(sum(i["total"] for i in final_items), 2)
    iva_amount  = round(subtotal * iva_rate, 2)
    total_usd   = round(subtotal + iva_amount, 2)
    total_bs    = round(total_usd * tasa_bcv, 2)

    # Store only codes of deleted items (we don't have the original list here,
    # so we just record counts; deleted_items is stored as JSON from frontend)
    with _conn() as con:
        con.execute("""
            INSERT INTO conversions
              (ts, ip, invoice_no, original_qty, final_qty, deleted_qty,
               deleted_items, iva_rate, tasa_bcv, total_usd, total_bs)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
            ip, invoice_no, original_qty, final_qty, deleted_qty,
            None,  # filled by caller if available
            iva_rate, tasa_bcv, total_usd, total_bs,
        ))


def log_conversion_full(ip, invoice_no, original_qty, final_items,
                        deleted_items_json, iva_rate, tasa_bcv):
    final_qty   = len(final_items)
    deleted_qty = original_qty - final_qty
    subtotal    = round(sum(i["total"] for i in final_items), 2)
    iva_amount  = round(subtotal * iva_rate, 2)
    total_usd   = round(subtotal + iva_amount, 2)
    total_bs    = round(total_usd * tasa_bcv, 2)

    with _conn() as con:
        con.execute("""
            INSERT INTO conversions
              (ts, ip, invoice_no, original_qty, final_qty, deleted_qty,
               deleted_items, iva_rate, tasa_bcv, total_usd, total_bs)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
            ip, invoice_no, original_qty, final_qty, deleted_qty,
            deleted_items_json,
            iva_rate, tasa_bcv, total_usd, total_bs,
        ))


def get_all(limit=500):
    with _conn() as con:
        rows = con.execute(
            "SELECT * FROM conversions ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
    return [dict(r) for r in rows]
