import sys
import os
import csv
import shutil
import sqlite3
import zipfile
import logging
import json
import re
from copy import copy
from datetime import datetime, date, timedelta
from pathlib import Path

from PySide6.QtCore import Qt, QRectF, QDate, QDateTime, QUrl
from PySide6.QtGui import QAction, QPixmap, QPainter, QPen, QBrush, QFont, QColor, QStandardItem, QStandardItemModel, QTextDocument, QDesktopServices
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFormLayout, QLineEdit, QTextEdit, QPushButton, QLabel,
    QComboBox, QTableWidget, QTableWidgetItem, QMessageBox,
    QFileDialog, QDateEdit, QSpinBox, QDoubleSpinBox, QGroupBox,
    QSplitter, QListWidget, QStackedWidget, QDialog, QDialogButtonBox,
    QHeaderView, QAbstractItemView, QCheckBox, QScrollArea, QTabWidget, QListWidgetItem,
    QGridLayout, QDateTimeEdit, QInputDialog
)
from PySide6.QtPrintSupport import QPrinter, QPrintDialog

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
)
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font
from openpyxl.cell.cell import MergedCell
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn



APP_NAME = "HSE Management System"
APP_VERSION = "1.4.19"

APP_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / APP_NAME
DB_DIR = APP_DIR / "database"
ATTACH_DIR = APP_DIR / "attachments"
REPORT_DIR = APP_DIR / "reports"
BACKUP_DIR = APP_DIR / "backups"
LOG_DIR = APP_DIR / "logs"

for p in [APP_DIR, DB_DIR, ATTACH_DIR, REPORT_DIR, BACKUP_DIR, LOG_DIR]:
    p.mkdir(parents=True, exist_ok=True)

DB_FILE = DB_DIR / "hse.db"

logging.basicConfig(
    filename=LOG_DIR / "application.log",
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s"
)


# ============================================================
# DATABASE
# ============================================================

class Database:

    def __init__(self):
        self.conn = sqlite3.connect(DB_FILE)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.create_tables()

    def create_tables(self):

        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );

        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT,
            name TEXT NOT NULL,
            designation TEXT,
            company TEXT,
            department TEXT,
            contact TEXT,
            active INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            company_type TEXT,
            contact_person TEXT,
            contact_number TEXT,
            status TEXT DEFAULT 'Active'
        );

        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project TEXT NOT NULL,
            area TEXT,
            location TEXT,
            description TEXT,
            active INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS observations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number TEXT UNIQUE,
            obs_date TEXT,
            obs_time TEXT,
            project TEXT,
            company TEXT,
            location TEXT,
            area TEXT,
            responsible TEXT,
            designation TEXT,
            employee_id TEXT,
            observer TEXT,
            observer_designation TEXT,
            observer_id TEXT,
            observer_company TEXT,
            obs_type TEXT,
            category TEXT,
            subcategory TEXT,
            observation TEXT,
            immediate_action TEXT,
            corrective_action TEXT,
            preventive_action TEXT,
            priority TEXT,
            target_date TEXT,
            status TEXT DEFAULT 'Open',
            closeout_date TEXT,
            closed_by TEXT,
            verified_by TEXT,
            verification_date TEXT,
            closeout_comments TEXT,
            verification TEXT,
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS observation_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            observation_id INTEGER,
            file_path TEXT,
            attachment_type TEXT,
            FOREIGN KEY(observation_id)
                REFERENCES observations(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS incident_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            incident_id INTEGER NOT NULL,
            file_path TEXT,
            attachment_type TEXT DEFAULT 'Evidence',
            remark TEXT DEFAULT '',
            FOREIGN KEY(incident_id)
                REFERENCES incidents(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS audit_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            audit_id INTEGER NOT NULL,
            file_path TEXT,
            attachment_type TEXT DEFAULT 'Evidence',
            FOREIGN KEY(audit_id)
                REFERENCES audits(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS capa_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            capa_id INTEGER NOT NULL,
            file_path TEXT,
            attachment_type TEXT DEFAULT 'Evidence',
            FOREIGN KEY(capa_id)
                REFERENCES capa(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS incidents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number TEXT UNIQUE,
            incident_date TEXT,
            incident_time TEXT,
            location TEXT,
            project TEXT,
            company TEXT,
            department TEXT,
            activity TEXT,
            incident_type TEXT,
            person_involved TEXT,
            employee_id TEXT,
            designation TEXT,
            supervisor TEXT,
            witnesses TEXT,
            description TEXT,
            immediate_action TEXT,
            consequences TEXT,
            potential_consequences TEXT,
            equipment TEXT,
            investigation_method TEXT,
            why1 TEXT,
            why2 TEXT,
            why3 TEXT,
            why4 TEXT,
            why5 TEXT,
            direct_cause TEXT,
            contributing_factors TEXT,
            root_cause TEXT,
            corrective_action TEXT,
            preventive_action TEXT,
            status TEXT DEFAULT 'Open',
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS audits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number TEXT UNIQUE,
            audit_date TEXT,
            audit_type TEXT,
            standard TEXT,
            project TEXT,
            location TEXT,
            department TEXT,
            auditor TEXT,
            lead_auditor TEXT,
            auditee TEXT,
            scope TEXT,
            objective TEXT,
            criteria TEXT,
            start_time TEXT,
            end_time TEXT,
            status TEXT DEFAULT 'Open',
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS audit_findings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            audit_id INTEGER,
            clause TEXT,
            sub_clause TEXT,
            requirement TEXT,
            finding_type TEXT,
            observation TEXT,
            evidence TEXT,
            risk_impact TEXT,
            corrective_action TEXT,
            responsible TEXT,
            target_date TEXT,
            status TEXT DEFAULT 'Open',
            verification TEXT,
            closeout_evidence TEXT,
            FOREIGN KEY(audit_id)
                REFERENCES audits(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS capa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number TEXT UNIQUE,
            source TEXT,
            reference_number TEXT,
            finding TEXT,
            root_cause TEXT,
            corrective_action TEXT,
            preventive_action TEXT,
            responsible TEXT,
            priority TEXT,
            target_date TEXT,
            status TEXT DEFAULT 'Open',
            verification TEXT,
            closeout_evidence TEXT,
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT,
            subcategory TEXT,
            active INTEGER DEFAULT 1
        );
        """)

        # Training module tables are additive; existing HSE tables and records are unchanged.
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS training_courses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_name TEXT NOT NULL,
            designation TEXT DEFAULT '',
            department TEXT DEFAULT '',
            mandatory INTEGER DEFAULT 1,
            provider_type TEXT DEFAULT 'Internal',
            validity_months INTEGER DEFAULT 0,
            requirement_source TEXT DEFAULT 'Company',
            certificate_required INTEGER DEFAULT 0,
            active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS training_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_db_id INTEGER NOT NULL,
            course_id INTEGER NOT NULL,
            training_date TEXT NOT NULL,
            expiry_date TEXT DEFAULT '',
            provider TEXT DEFAULT '',
            certificate_no TEXT DEFAULT '',
            attachment_path TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(employee_db_id) REFERENCES employees(id),
            FOREIGN KEY(course_id) REFERENCES training_courses(id)
        );
        CREATE INDEX IF NOT EXISTS idx_training_records_employee_course
            ON training_records(employee_db_id, course_id, training_date);
        """)

        # Training-only employee fields; additive migrations preserve all existing data.
        employee_columns = {r["name"] for r in self.conn.execute("PRAGMA table_info(employees)").fetchall()}
        if "date_joining" not in employee_columns:
            self.conn.execute("ALTER TABLE employees ADD COLUMN date_joining TEXT DEFAULT ''")
        if "safety_induction_date" not in employee_columns:
            self.conn.execute("ALTER TABLE employees ADD COLUMN safety_induction_date TEXT DEFAULT ''")
        # Older versions offered "Either"; the current Training Type choices are Internal or External.
        self.conn.execute("UPDATE training_courses SET provider_type='Internal' WHERE lower(COALESCE(provider_type,''))='either'")

        # Backward-compatible migrations for existing installations.
        self.conn.execute("CREATE TABLE IF NOT EXISTS incident_attachments (id INTEGER PRIMARY KEY AUTOINCREMENT, incident_id INTEGER NOT NULL, file_path TEXT, attachment_type TEXT DEFAULT 'Evidence', FOREIGN KEY(incident_id) REFERENCES incidents(id) ON DELETE CASCADE)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS audit_attachments (id INTEGER PRIMARY KEY AUTOINCREMENT, audit_id INTEGER NOT NULL, file_path TEXT, attachment_type TEXT DEFAULT 'Evidence', FOREIGN KEY(audit_id) REFERENCES audits(id) ON DELETE CASCADE)")
        self.conn.execute("CREATE TABLE IF NOT EXISTS capa_attachments (id INTEGER PRIMARY KEY AUTOINCREMENT, capa_id INTEGER NOT NULL, file_path TEXT, attachment_type TEXT DEFAULT 'Evidence', FOREIGN KEY(capa_id) REFERENCES capa(id) ON DELETE CASCADE)")
        # Newer incident-investigation procedure data is stored as JSON so existing
        # installations keep all original columns/features without a destructive migration.
        incident_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(incidents)").fetchall()}
        if "investigation_details" not in incident_cols:
            self.conn.execute("ALTER TABLE incidents ADD COLUMN investigation_details TEXT DEFAULT ''")
        # Incident/RCA extensions are additive only; existing data remains intact.
        incident_extra = {
            "incident_title": "TEXT DEFAULT ''",
            "permit_number": "TEXT DEFAULT ''",
            "people_details": "TEXT DEFAULT ''",
            "equipment_details": "TEXT DEFAULT ''",
            "environmental_category": "TEXT DEFAULT ''",
            "environmental_quantity": "TEXT DEFAULT ''",
            "property_name": "TEXT DEFAULT ''",
            "property_damage_cost": "TEXT DEFAULT ''",
            "organization": "TEXT DEFAULT ''",
            "procedure_reference": "TEXT DEFAULT ''",
            "report_reference": "TEXT DEFAULT ''",
            "capa_required": "INTEGER DEFAULT 0",
            "capa_number": "TEXT DEFAULT ''"
        }
        for col, definition in incident_extra.items():
            if col not in incident_cols:
                self.conn.execute(f"ALTER TABLE incidents ADD COLUMN {col} {definition}")
        # Add attachment remarks without replacing any existing attachment data.
        incident_attachment_cols = [r[1] for r in self.conn.execute("PRAGMA table_info(incident_attachments)").fetchall()]
        if "remark" not in incident_attachment_cols:
            self.conn.execute("ALTER TABLE incident_attachments ADD COLUMN remark TEXT DEFAULT ''")

        # Audit module extensions are additive and isolated from all other modules.
        audit_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(audits)").fetchall()}
        audit_extra = {
            "title": "TEXT DEFAULT ''",
            "created_by": "TEXT DEFAULT ''",
            "updated_at": "TEXT DEFAULT ''",
            "submitted_at": "TEXT DEFAULT ''",
            "auditor_position": "TEXT DEFAULT ''",
            "auditor_date": "TEXT DEFAULT ''",
            "reviewer": "TEXT DEFAULT ''",
            "reviewer_position": "TEXT DEFAULT ''",
            "reviewer_date": "TEXT DEFAULT ''",
            "approver": "TEXT DEFAULT ''",
            "approver_position": "TEXT DEFAULT ''",
            "approver_date": "TEXT DEFAULT ''"
        }
        for col, definition in audit_extra.items():
            if col not in audit_cols:
                self.conn.execute(f"ALTER TABLE audits ADD COLUMN {col} {definition}")

        finding_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(audit_findings)").fetchall()}
        finding_extra = {
            "finding_number": "INTEGER DEFAULT 1",
            "finding_standard": "TEXT DEFAULT ''",
            "finding_detail": "TEXT DEFAULT ''",
            "responsible_person_id": "INTEGER",
            "location": "TEXT DEFAULT ''",
            "close_out_evidence": "TEXT DEFAULT ''",
            "created_at": "TEXT DEFAULT ''",
            "updated_at": "TEXT DEFAULT ''"
        }
        for col, definition in finding_extra.items():
            if col not in finding_cols:
                self.conn.execute(f"ALTER TABLE audit_findings ADD COLUMN {col} {definition}")

        audit_attachment_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(audit_attachments)").fetchall()}
        if "finding_id" not in audit_attachment_cols:
            self.conn.execute("ALTER TABLE audit_attachments ADD COLUMN finding_id INTEGER")
        if "uploaded_by" not in audit_attachment_cols:
            self.conn.execute("ALTER TABLE audit_attachments ADD COLUMN uploaded_by TEXT DEFAULT ''")
        if "uploaded_at" not in audit_attachment_cols:
            self.conn.execute("ALTER TABLE audit_attachments ADD COLUMN uploaded_at TEXT DEFAULT ''")

        self.conn.execute("""CREATE TABLE IF NOT EXISTS audit_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT, audit_id INTEGER NOT NULL, action TEXT,
            user_name TEXT, details TEXT, timestamp TEXT,
            FOREIGN KEY(audit_id) REFERENCES audits(id) ON DELETE CASCADE)""")

        # Audit planning and checklist tables are additive and isolated to the Audit module.
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS audit_schedule (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            procedure_name TEXT NOT NULL,
            area TEXT DEFAULT '',
            risk_level TEXT DEFAULT 'Medium',
            frequency TEXT DEFAULT 'Annual',
            planned_date TEXT DEFAULT '',
            owner TEXT DEFAULT '',
            status TEXT DEFAULT 'Planned',
            linked_audit_id INTEGER,
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT '',
            FOREIGN KEY(linked_audit_id) REFERENCES audits(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS audit_checklist_responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            schedule_id INTEGER NOT NULL,
            item_no INTEGER NOT NULL,
            procedure_name TEXT DEFAULT '',
            iso_clause TEXT DEFAULT '',
            checklist_item TEXT NOT NULL,
            result TEXT DEFAULT 'N/A',
            comments TEXT DEFAULT '',
            evidence_path TEXT DEFAULT '',
            updated_at TEXT DEFAULT '',
            UNIQUE(schedule_id, item_no),
            FOREIGN KEY(schedule_id) REFERENCES audit_schedule(id) ON DELETE CASCADE
        );
        """)
        # Additional internal-audit traceability fields (additive, Audit module only).
        schedule_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(audit_schedule)").fetchall()}
        if "area" not in schedule_cols:
            self.conn.execute("ALTER TABLE audit_schedule ADD COLUMN area TEXT DEFAULT ''")
        response_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(audit_checklist_responses)").fetchall()}
        for col, definition in {"responsible": "TEXT DEFAULT ''", "target_date": "TEXT DEFAULT ''"}.items():
            if col not in response_cols:
                self.conn.execute(f"ALTER TABLE audit_checklist_responses ADD COLUMN {col} {definition}")
        finding_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(audit_findings)").fetchall()}
        for col, definition in {"source_schedule_id": "INTEGER", "source_item_no": "INTEGER", "procedure_name": "TEXT DEFAULT ''", "capa_required": "INTEGER DEFAULT 0", "capa_number": "TEXT DEFAULT ''"}.items():
            if col not in finding_cols:
                self.conn.execute(f"ALTER TABLE audit_findings ADD COLUMN {col} {definition}")
        response_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(audit_checklist_responses)").fetchall()}
        for col, definition in {"capa_required": "INTEGER DEFAULT 0", "capa_number": "TEXT DEFAULT ''"}.items():
            if col not in response_cols:
                self.conn.execute(f"ALTER TABLE audit_checklist_responses ADD COLUMN {col} {definition}")

        # Audit-only additive workflow tables. Existing records and CAPA tables are preserved.
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS audit_procedure_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT, procedure_id TEXT DEFAULT '', title TEXT NOT NULL,
            document_number TEXT DEFAULT '', issuing_organization TEXT DEFAULT '', site_contract TEXT DEFAULT '',
            revision TEXT DEFAULT '', effective_date TEXT DEFAULT '', review_date TEXT DEFAULT '',
            clause_references TEXT DEFAULT '', owner TEXT DEFAULT '', applicability TEXT DEFAULT '',
            approval_status TEXT DEFAULT 'Draft checklist pending review', source_file TEXT DEFAULT '',
            checklist_json TEXT DEFAULT '[]', revision_history TEXT DEFAULT '', active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS audit_notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT, schedule_id INTEGER NOT NULL, document_path TEXT DEFAULT '',
            recipients TEXT DEFAULT '', revision INTEGER DEFAULT 1, issued_at TEXT DEFAULT '', issued_by TEXT DEFAULT '',
            FOREIGN KEY(schedule_id) REFERENCES audit_schedule(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS audit_finding_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT, finding_id INTEGER NOT NULL, action TEXT NOT NULL,
            old_status TEXT DEFAULT '', new_status TEXT DEFAULT '', comments TEXT DEFAULT '', evidence_path TEXT DEFAULT '',
            user_name TEXT DEFAULT '', timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(finding_id) REFERENCES audit_findings(id) ON DELETE CASCADE
        );
        """)
        for col, definition in {
            "severity": "TEXT DEFAULT 'Medium'", "capa_reason": "TEXT DEFAULT ''",
            "verification_status": "TEXT DEFAULT 'Pending Verification'", "auditee_response": "TEXT DEFAULT ''",
            "closure_date": "TEXT DEFAULT ''", "finding_reference": "TEXT DEFAULT ''",
            "closeout_approver": "TEXT DEFAULT ''", "closeout_approval_status": "TEXT DEFAULT 'Pending Approval'"
        }.items():
            if col not in finding_cols:
                self.conn.execute(f"ALTER TABLE audit_findings ADD COLUMN {col} {definition}")
        response_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(audit_checklist_responses)").fetchall()}
        for col, definition in {
            "severity": "TEXT DEFAULT 'Medium'", "capa_reason": "TEXT DEFAULT ''",
            "finding_classification": "TEXT DEFAULT 'Minor Nonconformity'",
            "auditee_response": "TEXT DEFAULT ''", "procedure_revision": "TEXT DEFAULT ''"
        }.items():
            if col not in response_cols:
                self.conn.execute(f"ALTER TABLE audit_checklist_responses ADD COLUMN {col} {definition}")
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_findings_schedule_item ON audit_findings(source_schedule_id, source_item_no)")
        self.conn.commit()

        # Observation multi-action + CAPA tracking link (additive only).
        observation_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(observations)").fetchall()}
        if "capa_required" not in observation_cols:
            self.conn.execute("ALTER TABLE observations ADD COLUMN capa_required INTEGER DEFAULT 0")
        if "capa_number" not in observation_cols:
            self.conn.execute("ALTER TABLE observations ADD COLUMN capa_number TEXT DEFAULT ''")
        if "corrective_actions_json" not in observation_cols:
            self.conn.execute("ALTER TABLE observations ADD COLUMN corrective_actions_json TEXT DEFAULT '[]'")

        capa_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(capa)").fetchall()}
        capa_extra = {
            "actions_json": "TEXT DEFAULT '[]'",
            "approval_status": "TEXT DEFAULT 'Pending Approval'",
            "approved_by": "TEXT DEFAULT ''",
            "approved_date": "TEXT DEFAULT ''",
            "completion_pct": "INTEGER DEFAULT 0",
            "action_status": "TEXT DEFAULT 'Open'",
            "observation_text": "TEXT DEFAULT ''",
        }
        for col, definition in capa_extra.items():
            if col not in capa_cols:
                self.conn.execute(f"ALTER TABLE capa ADD COLUMN {col} {definition}")

        self.conn.commit()

    def execute(self, sql, params=()):
        cur = self.conn.cursor()
        cur.execute(sql, params)
        self.conn.commit()
        return cur

    def fetchall(self, sql, params=()):
        return self.conn.execute(sql, params).fetchall()

    def fetchone(self, sql, params=()):
        return self.conn.execute(sql, params).fetchone()

    def setting(self, key, default=""):
        row = self.fetchone(
            "SELECT value FROM settings WHERE key=?",
            (key,)
        )
        return row["value"] if row else default

    def set_setting(self, key, value):
        self.execute(
            """
            INSERT INTO settings(key,value)
            VALUES(?,?)
            ON CONFLICT(key)
            DO UPDATE SET value=excluded.value
            """,
            (key, value)
        )


db = Database()


# ============================================================
# HELPERS
# ============================================================

OBS_TYPES = [
    "Walkthrough Observation",
    "Routine Inspection",
    "Client Observation",
    "Internal Walkthrough",
    "Unsafe Act",
    "Unsafe Condition",
    "Good Observation",
    "Good Practice",
    "Positive Observation",
    "Environmental Observation",
    "Near Miss",
    "Other"
]

CATEGORIES = [
    "PPE",
    "Work at Height",
    "Excavation",
    "Electrical Safety",
    "Lifting Operations",
    "Scaffolding",
    "Confined Space",
    "Fire Safety",
    "Housekeeping",
    "Slip Trip Fall",
    "Chemical Safety",
    "Environmental",
    "Traffic Management",
    "Plant and Machinery",
    "Manual Handling",
    "Emergency Preparedness",
    "Heat Stress",
    "Welfare",
    "Dropped Objects",
    "Tools and Equipment",
    "Permit to Work",
    "Barricading",
    "Material Storage",
    "Access and Egress",
    "Gas Testing",
    "Hot Work",
    "Noise",
    "Dust",
    "Waste Management",
    "Environmental Spill",
    "Other"
]

PRIORITIES = ["Low", "Medium", "High", "Critical"]

STATUSES = [
    "Open",
    "In Progress",
    "Pending Verification",
    "Closed",
    "Cancelled"
]

INCIDENT_TYPES = [
    "Fatality",
    "Lost Time Injury",
    "Medical Treatment Case",
    "Restricted Work Case",
    "First Aid Case",
    "Near Miss",
    "Property Damage",
    "Environmental Incident",
    "Fire Incident",
    "Vehicle Incident",
    "Equipment Damage",
    "Chemical Spill",
    "Other"
]

INVESTIGATION_METHODS = [
    "Root Cause Analysis",
    "5 Why Analysis",
    "Fishbone / Ishikawa",
    "ICAM",
    "Barrier Analysis",
    "Bow-Tie Analysis",
    "Fault Tree Analysis",
    "Causal Tree",
    "Other"
]

AUDIT_TYPES = [
    "Internal Audit",
    "External Audit",
    "Client Audit",
    "Certification Audit",
    "Surveillance Audit",
    "Regulatory Audit",
    "Project Audit",
    "Supplier Audit",
    "Other"
]

FINDING_TYPES = [
    "Positive Observation",
    "Good Practice",
    "Conformity",
    "Opportunity for Improvement",
    "Observation",
    "Minor Nonconformity",
    "Major Nonconformity",
    "Environmental Finding",
    "Legal / Compliance Finding",
    "Other"
]

CAPA_SOURCES = [
    "HSE Inspection",
    "Incident",
    "Near Miss",
    "Audit",
    "Client Observation",
    "Regulatory Inspection",
    "Environmental Incident",
    "Management Review",
    "Employee Observation",
    "Other"
]


def next_number(prefix, table):
    year = datetime.now().year
    row = db.fetchone(
        f"""
        SELECT number FROM {table}
        WHERE number LIKE ?
        ORDER BY id DESC LIMIT 1
        """,
        (f"{prefix}-{year}-%",)
    )

    if not row:
        n = 1
    else:
        try:
            n = int(row["number"].split("-")[-1]) + 1
        except Exception:
            n = 1

    return f"{prefix}-{year}-{n:05d}"


def today():
    return datetime.now().strftime("%Y-%m-%d")


def now_time():
    return datetime.now().strftime("%H:%M")


def safe(value):
    return "" if value is None else str(value)


def xml_safe(value):
    """Remove characters that are illegal in XML/Office Open XML files."""
    text = safe(value)
    return "".join(ch for ch in text if ch in "\t\n\r" or ord(ch) >= 32)


def attachment_names(table_name, record_id):
    return "; ".join(Path(safe(r["file_path"])).name for r in attachment_rows(table_name, record_id))


def attachment_export_map(table_name):
    return {
        "observations": "observation_attachments",
        "incidents": "incident_attachments",
        "audits": "audit_attachments",
        "capa": "capa_attachments",
    }.get(table_name)


def overdue(target, status):
    if not target or status in ("Closed", "Cancelled"):
        return False
    try:
        return date.fromisoformat(target) < date.today()
    except Exception:
        return False


def parse_corrective_actions(source):
    """Return a clean list of corrective-action strings from JSON, row, or legacy text."""
    if source is None:
        return []
    if isinstance(source, (list, tuple)):
        return [safe(x).strip() for x in source if safe(x).strip()]
    if isinstance(source, dict) or hasattr(source, "keys"):
        try:
            raw = source["corrective_actions_json"] if "corrective_actions_json" in source.keys() else ""
        except Exception:
            raw = ""
        actions = parse_corrective_actions(raw)
        if actions:
            return actions
        try:
            legacy = source["corrective_action"] if "corrective_action" in source.keys() else ""
        except Exception:
            legacy = ""
        return parse_corrective_actions(legacy)
    text_value = safe(source).strip()
    if not text_value:
        return []
    if text_value.startswith("["):
        try:
            data = json.loads(text_value)
            if isinstance(data, list):
                out = []
                for item in data:
                    if isinstance(item, dict):
                        val = safe(item.get("action") or item.get("text") or "").strip()
                    else:
                        val = safe(item).strip()
                    if val:
                        out.append(val)
                if out:
                    return out
        except Exception:
            pass
    lines = []
    for line in text_value.replace("\r", "\n").split("\n"):
        line = line.strip()
        if not line:
            continue
        if len(line) > 2 and line[0].isdigit() and line[1] in ".)":
            line = line[2:].strip()
        elif ". " in line[:4] and line.split(".", 1)[0].isdigit():
            line = line.split(".", 1)[1].strip()
        lines.append(line)
    return lines or ([text_value] if text_value else [])


def format_corrective_actions_display(actions, numbered=False):
    actions = parse_corrective_actions(actions)
    if not actions:
        return ""
    if numbered:
        return "\n".join(f"{i}. {a}" for i, a in enumerate(actions, 1))
    return "\n".join(actions)


def parse_capa_actions(source):
    """CAPA action items: [{action, evidence: [paths], has_evidence: bool}]"""
    if source is None:
        return []
    if isinstance(source, list):
        items = source
    else:
        text_value = safe(source).strip()
        if not text_value:
            return []
        try:
            items = json.loads(text_value)
        except Exception:
            items = [{"action": a, "evidence": [], "has_evidence": False} for a in parse_corrective_actions(text_value)]
    out = []
    for item in items or []:
        if isinstance(item, dict):
            action = safe(item.get("action") or item.get("text") or "").strip()
            evidence = item.get("evidence") or item.get("attachments") or []
            if isinstance(evidence, str):
                evidence = [evidence] if evidence.strip() else []
            evidence = [safe(p) for p in evidence if safe(p).strip()]
            has_ev = bool(item.get("has_evidence")) or any(Path(p).exists() for p in evidence)
            if action:
                out.append({"action": action, "evidence": evidence, "has_evidence": has_ev})
        else:
            action = safe(item).strip()
            if action:
                out.append({"action": action, "evidence": [], "has_evidence": False})
    return out


def capa_actions_completion(actions):
    items = parse_capa_actions(actions)
    total = len(items)
    if total == 0:
        return 0, 0, 0, "Open"
    done = sum(1 for x in items if x.get("has_evidence") or x.get("evidence"))
    pct = int(round(100 * done / total))
    if done == 0:
        status = "Open"
    elif done < total:
        status = "Pending"
    else:
        status = "CAPA Action Complete"
    return done, total, pct, status


def sync_observation_from_capa(capa_row):
    """When CAPA tracking completes/approves, push status back to the linked observation only."""
    ref = safe(capa_row["reference_number"] if "reference_number" in capa_row.keys() else "")
    if not ref:
        return
    obs = db.fetchone("SELECT * FROM observations WHERE number=?", (ref,))
    if not obs:
        return
    done, total, pct, action_status = capa_actions_completion(
        capa_row["actions_json"] if "actions_json" in capa_row.keys() else "[]"
    )
    approval = safe(capa_row["approval_status"] if "approval_status" in capa_row.keys() else "")
    new_status = safe(obs["status"])
    comments = safe(obs["closeout_comments"]) if "closeout_comments" in obs.keys() else ""
    note = f"CAPA {safe(capa_row['number'])}: {action_status} ({pct}%)"
    if approval:
        note += f" | Approval: {approval}"
    if note not in comments:
        comments = (comments + "\n" + note).strip() if comments else note
    if action_status == "CAPA Action Complete" and approval in ("Approved", "CAPA Approved"):
        new_status = "Closed"
    elif action_status == "CAPA Action Complete":
        new_status = "Pending Verification"
    elif action_status == "Pending":
        new_status = "In Progress"
    db.execute(
        "UPDATE observations SET status=?, closeout_comments=? WHERE id=?",
        (new_status, comments, obs["id"])
    )


# ============================================================
# DOCUMENT / ATTACHMENT HELPERS
# ============================================================

ISO_CLAUSES = {
    "ISO 45001": {
        "4": ["4.1 Understanding the organization and its context",
              "4.2 Understanding the needs and expectations of workers and other interested parties",
              "4.3 Determining the scope of the OH&S management system",
              "4.4 OH&S management system"],
        "5": ["5.1 Leadership and commitment", "5.2 OH&S policy",
              "5.3 Organizational roles, responsibilities and authorities",
              "5.4 Consultation and participation of workers"],
        "6": ["6.1 Actions to address risks and opportunities",
              "6.1.2 Hazard identification and assessment of risks and opportunities",
              "6.1.3 Determination of legal requirements and other requirements",
              "6.2 OH&S objectives and planning to achieve them"],
        "7": ["7.1 Resources", "7.2 Competence", "7.3 Awareness",
              "7.4 Communication", "7.5 Documented information"],
        "8": ["8.1 Operational planning and control", "8.1.2 Eliminating hazards and reducing OH&S risks",
              "8.2 Emergency preparedness and response"],
        "9": ["9.1 Monitoring, measurement, analysis and performance evaluation",
              "9.2 Internal audit", "9.3 Management review"],
        "10": ["10.1 General", "10.2 Incident, nonconformity and corrective action",
               "10.3 Continual improvement"],
    },
    "ISO 14001": {
        "4": ["4.1 Understanding the organization and its context",
              "4.2 Understanding the needs and expectations of interested parties",
              "4.3 Determining the scope of the environmental management system",
              "4.4 Environmental management system"],
        "5": ["5.1 Leadership and commitment", "5.2 Environmental policy",
              "5.3 Organizational roles, responsibilities and authorities"],
        "6": ["6.1 Actions to address risks and opportunities",
              "6.1.2 Environmental aspects", "6.1.3 Compliance obligations",
              "6.1.4 Planning action", "6.2 Environmental objectives and planning"],
        "7": ["7.1 Resources", "7.2 Competence", "7.3 Awareness",
              "7.4 Communication", "7.5 Documented information"],
        "8": ["8.1 Operational planning and control", "8.2 Emergency preparedness and response"],
        "9": ["9.1 Monitoring, measurement, analysis and evaluation",
              "9.2 Internal audit", "9.3 Management review"],
        "10": ["10.1 General", "10.2 Nonconformity and corrective action",
               "10.3 Continual improvement"],
    }
}


def company_name():
    return db.setting("company_name", "") or APP_NAME


def document_prefix():
    return db.setting("document_prefix", "HSE") or "HSE"


def document_number(kind):
    prefix = document_prefix()
    return next_number(f"{prefix}-{kind}", {
        "INC": "incidents",
        "OBS": "observations",
        "AUD": "audits",
        "CAPA": "capa"
    }.get(kind, "incidents"))


def attachment_rows(table_name, record_id):
    allowed = {
        "observation_attachments": "observation_id",
        "incident_attachments": "incident_id",
        "audit_attachments": "audit_id",
        "capa_attachments": "capa_id"
    }
    if table_name not in allowed:
        return []
    return db.fetchall(
        f"SELECT * FROM {table_name} WHERE {allowed[table_name]}=? ORDER BY id",
        (record_id,)
    )


def copy_attachments(paths, record_number, table_name, record_id, attachment_type="Evidence"):
    column = {
        "observation_attachments": "observation_id",
        "incident_attachments": "incident_id",
        "audit_attachments": "audit_id",
        "capa_attachments": "capa_id"
    }[table_name]
    saved = []
    for path in paths or []:
        try:
            source = Path(path)
            if not source.exists():
                continue
            destination = ATTACH_DIR / f"{record_number}_{source.name}"
            counter = 1
            while destination.exists():
                destination = ATTACH_DIR / f"{record_number}_{counter}_{source.name}"
                counter += 1
            shutil.copy2(source, destination)
            db.execute(
                f"INSERT INTO {table_name} ({column},file_path,attachment_type) VALUES(?,?,?)",
                (record_id, str(destination), attachment_type)
            )
            saved.append(str(destination))
        except Exception:
            logging.exception("Failed to save attachment %s", path)
    return saved


def add_docx_cell_shading(cell, fill="17365D"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_docx_cell_borders(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if edge in kwargs:
            edge_data = kwargs.get(edge)
            tag = "w:{}".format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["val", "sz", "space", "color"]:
                if key in edge_data:
                    element.set(qn("w:{}".format(key)), str(edge_data[key]))


def style_docx_table(table, header=True, font_size=7):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = False
    for row_index, row in enumerate(table.rows):
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_docx_cell_borders(
                cell,
                top={"val": "single", "sz": 6, "color": "808080"},
                bottom={"val": "single", "sz": 6, "color": "808080"},
                left={"val": "single", "sz": 6, "color": "808080"},
                right={"val": "single", "sz": 6, "color": "808080"},
            )
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.space_before = Pt(0)
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
            if header and row_index == 0:
                add_docx_cell_shading(cell)
                for run in cell.paragraphs[0].runs:
                    run.font.bold = True
                    run.font.color.rgb = __import__("docx").shared.RGBColor(255,255,255)


def report_logo_path():
    """Resolve the current Settings logo reliably for Word, PDF and Excel reports."""
    configured=safe(db.setting("company_logo",""))
    candidates=[]
    if configured:
        cp=Path(configured)
        candidates.extend([cp, APP_DIR / cp, APP_DIR / cp.name, ATTACH_DIR / cp.name])
    candidates += sorted(ATTACH_DIR.glob("company_logo.*"))
    seen=set()
    for p in candidates:
        try:
            p=Path(p)
            if not p.is_absolute():
                p=(APP_DIR / p)
            if p.exists() and p.is_file():
                key=str(p.resolve())
                if key not in seen:
                    seen.add(key); return p.resolve()
        except Exception:
            logging.exception("Unable to resolve report logo")
    return None

def add_docx_report_header(document,title,kind="REPORT",report_no=""):
    section=document.sections[0]; header=section.header
    usable=(section.page_width-section.left_margin-section.right_margin)/914400.0
    if usable < 9.0:
        widths=[Inches(1.20),Inches(4.05),Inches(2.20)]
    else:
        widths=[Inches(1.55),Inches(5.55),Inches(3.2)]
    table=header.add_table(rows=1,cols=3,width=Inches(sum(w.inches for w in widths))); table.autofit=False
    for cell,w in zip(table.rows[0].cells,widths): cell.width=w; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    prefix=document_prefix(); report_no=report_no or f"{prefix}-{kind}-REPORT-{datetime.now().strftime('%Y%m%d%H%M%S')}"; doc_no=f"{prefix}-{kind}"; project=db.setting("project_name","")
    left,mid,right=table.rows[0].cells
    lp=left.paragraphs[0]; lp.alignment=WD_ALIGN_PARAGRAPH.CENTER; logo=report_logo_path()
    if logo:
        try: lp.add_run().add_picture(str(logo),width=Inches(1.15))
        except Exception: lp.add_run(xml_safe(company_name())).bold=True
    else: lp.add_run(xml_safe(company_name())).bold=True
    mp=mid.paragraphs[0]; mp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=mp.add_run(xml_safe(title)); r.bold=True; r.font.size=Pt(14)
    r=mp.add_run("\n"+xml_safe(company_name())); r.bold=True; r.font.size=Pt(9)
    rp=right.paragraphs[0]; rp.alignment=WD_ALIGN_PARAGRAPH.LEFT
    for label,val in [("Report No.",report_no),("Document No.",doc_no),("Project",project)]:
        rr=rp.add_run(f"{label}: {xml_safe(val)}\n"); rr.font.size=Pt(8); rr.bold=(label=="Report No.")
    for cell in table.rows[0].cells:
        set_docx_cell_borders(cell,top={"val":"single","sz":8,"color":"808080"},bottom={"val":"single","sz":8,"color":"808080"},left={"val":"single","sz":8,"color":"808080"},right={"val":"single","sz":8,"color":"808080"})
    document.add_paragraph().paragraph_format.space_after=Pt(0)

def add_docx_header(document):
    add_docx_report_header(document,"HSE REPORT")


def add_docx_footer(document):
    footer_text = db.setting("report_footer", "")
    if not footer_text:
        return
    p = document.sections[0].footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(xml_safe(footer_text)).font.size = Pt(8)


def add_docx_title(document, title, number=""):
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(xml_safe(title))
    r.bold = True
    r.font.size = Pt(18)
    if number:
        p2 = document.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(xml_safe(f"Document No.: {number}"))
        r2.bold = True
        r2.font.size = Pt(10)

def save_docx_validated(document, path):
    """Write a Word report atomically and validate the OOXML ZIP before replacing the target."""
    target=Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temp=target.with_name(target.name + ".tmp.docx")
    try:
        document.save(str(temp))
        with zipfile.ZipFile(temp, "r") as z:
            bad=z.testzip()
            if bad:
                raise ValueError(f"Generated Word report is corrupt: {bad}")
        os.replace(temp, target)
    finally:
        temp.unlink(missing_ok=True)



def add_docx_kv_table(document, pairs):
    table = document.add_table(rows=0, cols=2)
    for key, value in pairs:
        cells = table.add_row().cells
        cells[0].text = xml_safe(key)
        cells[1].text = xml_safe(value)
        cells[0].paragraphs[0].runs[0].bold = True
    style_docx_table(table, header=False)
    return table


def add_docx_boxed_section(document, title, text, font_size=8):
    if not safe(text):
        return
    document.add_heading(title, level=2)
    t=document.add_table(rows=1, cols=1)
    c=t.cell(0,0); c.text=xml_safe(text)
    c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    style_docx_table(t, header=False, font_size=font_size)
    for p in c.paragraphs:
        p.paragraph_format.space_after=Pt(0)
    return t


def add_docx_attachments(document, rows):
    document.add_heading("Evidence / Attachments", level=2)
    if not rows:
        document.add_paragraph("No attachments recorded.")
        return
    for i,row in enumerate(rows,1):
        path=Path(safe(row["file_path"])); remark=safe(row["remark"]) if "remark" in row.keys() else ""
        document.add_paragraph(f"{i}. {path.name} | Type: {safe(row['attachment_type'])} | Remark: {remark}")
        if path.exists() and path.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
            try:
                p=document.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(str(path), width=Inches(2.0), height=Inches(2.0))
            except Exception:
                pass
        elif path.exists() and path.suffix.lower()==".pdf":
            document.add_paragraph("PDF attachment: " + path.name)
    

def add_docx_signatures(document):
    document.add_paragraph()
    document.add_heading("Review and Approval", level=2)
    table = document.add_table(rows=2, cols=3)
    labels = ["Prepared / HSE Report Made By", "Reviewed By", "Approved By"]
    for i, label in enumerate(labels):
        table.cell(0, i).text = label
        table.cell(1, i).text = "\n\n____________________________\nSignature: __________________\nDate: ______________________"
    style_docx_table(table, header=True)


def generate_incident_docx(incident_id, path):
    row=db.fetchone("SELECT * FROM incidents WHERE id=?",(incident_id,))
    if not row: raise ValueError("Incident record not found.")
    doc=Document()
    # Use a controlled professional report layout.  The report contains several
    # investigation tables and evidence fields, so landscape prevents table
    # compression while preserving the existing report content.
    set_docx_landscape(doc)
    add_docx_report_header(doc,"ACCIDENT / INCIDENT INVESTIGATION REPORT","INC",row["number"])
    title=safe(row["incident_title"]) if "incident_title" in row.keys() else ""
    add_docx_kv_table(doc,[
        ("Incident Title",title),("Permit Number",safe(row["permit_number"]) if "permit_number" in row.keys() else ""),("Incident Date",row["incident_date"]),("Incident Time",row["incident_time"]),
        ("Company",row["company"] or company_name()),("Project",row["project"]),("Location",row["location"]),
        ("Department",row["department"]),("Activity",row["activity"]),("Incident Type",row["incident_type"]),
        ("Investigation Method",row["investigation_method"]),("Status",row["status"]),
        ("Organisation",row["organization"]),("Report / Reference No.",row["report_reference"]),
        ("CAPA Tracking Reference", (safe(row["capa_number"]) + " — Accident Corrective Action Plan") if ("capa_number" in row.keys() and safe(row["capa_number"])) else "Not linked")
    ])
    try: details=json.loads(safe(row["investigation_details"]) or "{}")
    except Exception: details={}
    def j(col):
        try:return json.loads(safe(row[col]) or "[]")
        except Exception:return []
    def section_table(title, headers, data):
        if not data:return
        doc.add_heading(title,2); t=doc.add_table(rows=1,cols=len(headers))
        for i,h in enumerate(headers): t.rows[0].cells[i].text=h
        for vals in data:
            c=t.add_row().cells
            for i,v in enumerate(vals): c[i].text=xml_safe(v)
        style_docx_table(t,font_size=8)
    people=j("people_details"); equipment=j("equipment_details")
    section_table("People Involved",["Person Involved","Role","Designation","ID"],[[x.get("name",""),x.get("role",""),x.get("designation",""),x.get("id","")] for x in people])
    section_table("Equipment",["Equipment","ID","Specification"],[[x.get("name",""),x.get("id",""),x.get("specification","")] for x in equipment])
    for key,titleh,headers in [("environmental_items","Environmental Information",["Category","Quantity / Unit"]),("property_items","Property / Asset",["Property / Asset Name","Cost of Damage"]),("procedure_items","Procedure / Reference",["Reference","Details"])]:
        items=details.get(key,[]) or []
        vals=[]
        for x in items:
            if key=="environmental_items": vals.append([x.get("category",""),x.get("quantity","")])
            elif key=="property_items": vals.append([x.get("name",""),x.get("cost","")])
            else: vals.append([x.get("reference",""),x.get("details","")])
        section_table(titleh,headers,vals)
    add_docx_boxed_section(doc,"Incident Description",row["description"],8)
    add_docx_boxed_section(doc,"Immediate Response / Action",row["immediate_action"],8)
    incident_actions = details.get("corrective_actions", []) or []
    if incident_actions:
        doc.add_heading("Corrective Action Plan", 2)
        action_table = doc.add_table(rows=1, cols=4)
        for i, heading in enumerate(["Corrective Action", "Responsible Person", "Target Date", "Status"]):
            action_table.rows[0].cells[i].text = heading
        for action in incident_actions:
            cells = action_table.add_row().cells
            values = [action.get("action", ""), action.get("responsible", ""), action.get("target_date", ""), action.get("status", "Open")]
            for i, value in enumerate(values):
                cells[i].text = xml_safe(value)
        style_docx_table(action_table, header=True, font_size=8)
    # Witness statements, including optional attachment and remark for each statement.
    witnesses=details.get("witness_statements",[]) or []
    if witnesses:
        doc.add_heading("Witness Statements",2)
        for i,w in enumerate(witnesses,1):
            doc.add_paragraph(f"Witness {i}: {xml_safe(w.get('name',''))} | Remark: {xml_safe(w.get('remark',''))}")
            doc.add_paragraph(xml_safe(w.get("statement","")))
            if w.get("attachment"): doc.add_paragraph("Statement Attachment: "+xml_safe(w.get("attachment")))
    if row["investigation_method"]=="ICAM":
        doc.add_heading("ICAM ANALYSIS",2)
        def icam_text(title,key):
            v=details.get(key,"")
            if v: add_docx_boxed_section(doc,title,v,8)
        icam_text("Severity", "icam_severity")
        def icam_table(title,headers,key):
            vals=details.get(key,[]) or []
            if not vals:return
            doc.add_heading(title,2)
            t=doc.add_table(rows=1,cols=len(headers))
            for i,h in enumerate(headers): t.rows[0].cells[i].text=h
            for x in vals:
                c=t.add_row().cells
                mapping={
                    "Name":"name","Position":"position","Department":"department","Investigation Role":"role",
                    "Date":"date","Time":"time","Event / Description":"event",
                    "Factor":"factor","What Happened":"what_happened","Event / Condition":"event_condition",
                    "Control Measure Description":"control_measure","Responsible Person":"responsible","Open / Close":"status","Target Date":"target_date",
                }
                for i,h in enumerate(headers): c[i].text=xml_safe(x.get(mapping.get(h,h.lower().replace(" ","_")),""))
            style_docx_table(t,header=True,font_size=7)
        icam_table("Investigation Team",["Name","Position","Department","Investigation Role"],"icam_team")
        icam_table("EVENT TIMELINE",["Date","Time","Event / Description"],"icam_timeline_events")
        for key,title in [("icam_a","A. DEFENCE / CONTROL FACTORS"),("icam_ita","B. INDIVIDUAL / TEAM ACTIONS (ITA)"),("icam_tec","C. TASK / ENVIRONMENTAL CONDITIONS (TEC)"),("icam_of","D. ORGANISATIONAL FACTORS (OF)")]:
            icam_table(title,["Factor","What Happened","Event / Condition"],key)
        icam_table("CONTROL MEASURES",["Factor","Control Measure Description","Responsible Person","Open / Close","Target Date"],"icam_control_measures")
        icam_text("Recommendations","icam_recommendations")
        summary=[]
        for key,label in [("icam_a","A"),("icam_ita","ITA"),("icam_tec","TEC"),("icam_of","OF")]:
            for x in details.get(key,[]) or []:
                summary.append([label,x.get("factor",""),x.get("what_happened",""),x.get("event_condition","")])
        if summary:
            doc.add_heading("ICAM Category Summary",2)
            st=doc.add_table(rows=1,cols=4)
            for i,h in enumerate(["ICAM Category","Selected Factor","What Happened","Event / Condition"]): st.rows[0].cells[i].text=h
            for vals in summary:
                c=st.add_row().cells
                for i,v in enumerate(vals): c[i].text=xml_safe(v)
            style_docx_table(st,header=True,font_size=7)
    if row["investigation_method"]=="Fishbone / Ishikawa":
        fish=details.get("fishbone",[]) or []
        doc.add_heading("Fishbone / Ishikawa Analysis",2)
        # Keep each Fishbone cause attachment beside the exact cause that it supports.
        if fish:
            incident_attach_map={Path(safe(a["file_path"])).name:Path(safe(a["file_path"])) for a in attachment_rows("incident_attachments",incident_id)}
            ft=doc.add_table(rows=1,cols=5)
            for i,h in enumerate(["Factor","Suggested Cause","Specific Cause / Explanation","Evidence / Finding","Attachment"]): ft.rows[0].cells[i].text=h
            for x in fish:
                cells=ft.add_row().cells
                vals=[x.get("category",""),x.get("cause",""),x.get("detail",""),x.get("evidence","")]
                for i,v in enumerate(vals): cells[i].text=xml_safe(v)
                att_names=x.get("attachments",[]) or []
                if not att_names: cells[4].text="None"
                else:
                    cells[4].text=""
                    for nm in att_names:
                        fp=incident_attach_map.get(Path(nm).name, ATTACH_DIR / Path(nm).name)
                        if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
                            try:
                                pp=cells[4].paragraphs[0] if not cells[4].paragraphs[0].runs else cells[4].add_paragraph()
                                pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; pp.add_run().add_picture(str(fp),width=Inches(2.0),height=Inches(2.0))
                                cells[4].add_paragraph(xml_safe(fp.name))
                            except Exception: cells[4].add_paragraph(xml_safe(fp.name))
                        else:
                            cells[4].add_paragraph(xml_safe(fp.name))
            style_docx_table(ft,font_size=7)
        team=details.get("investigation_team",[]) or []
        section_table("Investigation Team",["Name","Position","Department","Role in Investigation"],[[x.get("name",""),x.get("position",""),x.get("department",""),x.get("role","")] for x in team])
        timeline=details.get("timeline_events",[]) or []
        section_table("Timeline of Events",["Time","Event / What Happened"],[[x.get("time",""),x.get("event","")] for x in timeline])
        if details.get("problem_statement"):
            doc.add_heading("Problem Statement",2); doc.add_paragraph(xml_safe(details.get("problem_statement")))
        ver=details.get("root_cause_verification",[]) or []
        section_table("Root Cause Verification",["Potential Cause","Evidence / Finding","Verified?","Responsible"],[[x.get("cause",""),x.get("evidence",""),x.get("verified",""),x.get("responsible","")] for x in ver])
        if details.get("effectiveness_verification"):
            doc.add_heading("Effectiveness Verification",2); doc.add_paragraph(xml_safe(details.get("effectiveness_verification")))
        if details.get("lessons_learned"):
            doc.add_heading("Lessons Learned",2); doc.add_paragraph(xml_safe(details.get("lessons_learned")))
        if details.get("conclusion"):
            doc.add_heading("Conclusion",2); doc.add_paragraph(xml_safe(details.get("conclusion")))
        # Cause-and-effect summary showing the Fishbone factors leading to the incident/problem.
        doc.add_heading("Cause-and-Effect Diagram",2)
        factor_groups={}
        for x in fish:
            factor_groups.setdefault(x.get("category","Other"),[]).append(x.get("cause","") or x.get("detail",""))
        diagram_rows=[["Factor","Possible Causes","Effect / Problem"]]
        for cat in ["People","Machine / Equipment","Method","Material","Measurement","Environment","Other"]:
            vals=factor_groups.get(cat,[])
            if vals: diagram_rows.append([cat,"; ".join(v for v in vals if v),safe(row["incident_title"]) or "INCIDENT / PROBLEM"])
        section_table("Fishbone Cause-and-Effect",["Main Factor","Possible Causes","Effect / Problem"],diagram_rows[1:] if len(diagram_rows)>1 else [])
    vals=details.get("direct_causes",[])
    vals=vals if isinstance(vals,list) else ([vals] if vals else [])
    if not vals and row["direct_cause"]: vals=[row["direct_cause"]]
    add_docx_boxed_section(doc,"Direct Cause", "\n".join(f"{i}. {v}" for i,v in enumerate(vals,1)),8)
    for h,key,fallback in [("Contributing Factors","contributing_factors_list",row["contributing_factors"]),("Root Cause","root_causes_list",row["root_cause"])]:
        vals=details.get(key,[]); vals=vals if isinstance(vals,list) else ([vals] if vals else [])
        if not vals and fallback: vals=[fallback]
        doc.add_heading(h,2)
        for i,v in enumerate(vals,1): doc.add_paragraph(f"{i}. {xml_safe(v)}")
    if row["investigation_method"]=="5 Why Analysis" or any(row[f"why{i}"] for i in range(1,6)):
        section_table("5 Why Analysis",["Step","Analysis"],[[f"Why {i}",row[f"why{i}"] or "N/A"] for i in range(1,6)])
    actions=details.get("corrective_actions",[]) or []
    section_table("Corrective Action Plan",["Action","Responsible","Target Date","Status","Closeout Evidence"],[[a.get("action",""),a.get("responsible",""),a.get("target_date",""),a.get("status","Open"),"; ".join(a.get("attachments",[])) or "None"] for a in actions])
    add_docx_attachments(doc,attachment_rows("incident_attachments",incident_id))
    sig=details.get("approval_signatures",{}) if isinstance(details.get("approval_signatures",{}),dict) else {}
    doc.add_paragraph()
    doc.add_heading("Review and Approval / Sign-off", level=2)
    st=doc.add_table(rows=4, cols=3); st.style="Table Grid"; st.autofit=False
    for i,label in enumerate(["Prepared by","Reviewed by","Approved by"]):
        st.cell(0,i).text=label
        x=sig.get(label.lower().replace(" ","_"),{}) or {}
        st.cell(1,i).text=f"Name: {x.get('name','')}"
        st.cell(2,i).text=f"Position: {x.get('position','')}"
        st.cell(3,i).text=f"Signature: {x.get('signature','')}\nDate: {x.get('date','')}"
    style_docx_table(st,header=True,font_size=8)
    add_docx_footer(doc); save_docx_validated(doc,path)


OBSERVATION_EXPORT_FIELDS = [
    ("S.No.", "__serial__"), ("Observation No.", "number"), ("Observation Date", "obs_date"),
    ("Observation Time", "obs_time"), ("Project", "project"), ("Company", "company"),
    ("Location", "location"), ("Area", "area"), ("Responsible Person", "responsible"),
    ("Designation", "designation"), ("Employee ID", "employee_id"), ("Observer", "observer"),
    ("Observer Designation", "observer_designation"), ("Observer ID", "observer_id"),
    ("Observer Company", "observer_company"), ("Observation Type", "obs_type"),
    ("Category", "category"), ("Subcategory", "subcategory"),
    ("Observation Detail", "observation"),
    ("Action Taken", "__corrective_actions__"),
    ("Priority", "priority"), ("Target Date", "target_date"), ("Status", "status"),
    ("CAPA Required", "__capa_required__"), ("CAPA No.", "capa_number"),
    ("Closeout Date", "closeout_date"), ("Closed By", "closed_by"),
    ("Verified By", "verified_by"), ("Verification Date", "verification_date"),
    ("Closeout Comments", "closeout_comments"), ("Verification", "verification"),
    ("Evidence", "__attachments__"),
]


def observation_export_rows(record_ids=None):
    params=[]
    where=""
    if record_ids:
        placeholders=",".join("?" for _ in record_ids)
        where=f" WHERE id IN ({placeholders})"
        params=list(record_ids)
    # Current records only; deleted observations can never contribute to the serial sequence.
    rows=db.fetchall(f"SELECT * FROM observations{where} ORDER BY id DESC", params)
    columns=[x[0] for x in OBSERVATION_EXPORT_FIELDS]
    data=[]
    for serial, row in enumerate(rows, 1):
        values=[]
        for _, key in OBSERVATION_EXPORT_FIELDS:
            if key == "__serial__":
                value=serial
            elif key == "__attachments__":
                value=attachment_names("observation_attachments", row["id"])
            elif key == "__corrective_actions__":
                value=format_corrective_actions_display(row)
            elif key == "__capa_required__":
                try:
                    value = "Yes" if int(row["capa_required"] or 0) else "No"
                except Exception:
                    value = "No"
            else:
                try:
                    value=row[key]
                except Exception:
                    value=""
            values.append(xml_safe(value))
        data.append(values)
    return columns, data, rows


def observation_word_sequence():
    return [
        "S.No.", "Observation Date", "Location", "Observation Type", "Category",
        "Observation Detail", "Action Taken", "Responsible", "Evidence", "Status"
    ]


def observation_attachment_images(record_id):
    result=[]
    for a in attachment_rows("observation_attachments", record_id):
        p=Path(safe(a["file_path"]))
        if p.exists() and p.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp"}:
            result.append((p, safe(a["attachment_type"])))
    return result

def set_docx_landscape(document):
    section=document.sections[0]
    section.orientation=WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height=section.page_height, section.page_width
    section.left_margin=Inches(0.35)
    section.right_margin=Inches(0.35)
    section.top_margin=Inches(0.45)
    section.bottom_margin=Inches(0.45)


def add_observation_attachments_to_docx(document, observation_rows):
    """Add compact attachment information without creating a second wide table."""
    document.add_heading("Attachments / Evidence", level=2)
    table = document.add_table(rows=1, cols=3)
    hdr = table.rows[0].cells
    hdr[0].text = "S.No."
    hdr[1].text = "Observation No."
    hdr[2].text = "Attachment / Evidence"
    any_attachment = False
    for serial, obs in enumerate(observation_rows, 1):
        attachments = attachment_rows("observation_attachments", obs["id"])
        if not attachments:
            continue
        any_attachment = True
        cells = table.add_row().cells
        cells[0].text = str(serial)
        cells[1].text = safe(obs["number"])
        for idx, a in enumerate(attachments):
            path = Path(safe(a["file_path"]))
            if idx:
                cells[2].add_paragraph()
            p = cells[2].paragraphs[-1]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(path.name)
            run.font.size = Pt(7)
            if path.exists() and path.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp"}:
                try:
                    p.add_run("  ")
                    p.add_run().add_picture(str(path), width=Inches(0.55))
                except Exception:
                    pass
    if not any_attachment:
        cells = table.add_row().cells
        cells[0].merge(cells[2])
        cells[0].text = "No attachments recorded."
    style_docx_table(table, header=True, font_size=7)


def generate_observation_docx(path, record_ids=None):
    columns, data, rows = observation_export_rows(record_ids)
    if not rows:
        raise ValueError("There is no observation data to export.")
    doc = Document()
    set_docx_landscape(doc)
    # Keep the report compact and professional: one table, one row per current observation.
    add_docx_report_header(doc,"HSE OBSERVATION REPORT","OBS",f"{document_prefix()}-OBS-REPORT-{datetime.now().strftime('%Y%m%d%H%M%S')}")
    pg=doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    pg.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    pg.paragraph_format.space_after=Pt(2)

    preferred = observation_word_sequence()
    idx = {name: columns.index(name) for name in preferred if name in columns}
    table = doc.add_table(rows=1, cols=len(preferred))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for c, name in enumerate(preferred):
        table.rows[0].cells[c].text = name

    for values, row in zip(data[:1000], rows[:1000]):
        cells = table.add_row().cells
        attachments = attachment_rows("observation_attachments", row["id"])
        for c, name in enumerate(preferred):
            cell=cells[c]
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if name == "Evidence":
                cell.text=""
                if not attachments:
                    cell.paragraphs[0].add_run("None").font.size=Pt(5.2)
                else:
                    for ai,a in enumerate(attachments):
                        p2=cell.paragraphs[0] if ai==0 else cell.add_paragraph()
                        p2.paragraph_format.space_after=Pt(0)
                        p2.paragraph_format.space_before=Pt(0)
                        path2=Path(safe(a["file_path"]))
                        if path2.exists() and path2.suffix.lower() in {".png",".jpg",".jpeg",".bmp"}:
                            try:
                                p2.add_run().add_picture(str(path2), width=Inches(0.38))
                            except Exception:
                                pass
                        if not (path2.exists() and path2.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp"}):
                            run=p2.add_run(" " + path2.name)
                            run.font.size=Pt(4.8)
            elif name == "Observation Type":
                cell.text=xml_safe(row["obs_type"])
            elif name == "Observation Detail":
                cell.text=xml_safe(row["observation"])
            elif name == "Action Taken":
                cell.text=xml_safe(format_corrective_actions_display(row))
            elif name == "Responsible":
                cell.text=xml_safe(row["responsible"])
            elif name == "Category":
                cell.text=xml_safe(row["category"])
            elif name == "S.No.":
                cell.text=str(values[idx[name]])
            else:
                cell.text=xml_safe(values[idx[name]])

    # A4 landscape: total 10.95 inches, safely inside the 11.19-inch usable width.
    widths=[0.34,0.70,0.92,0.88,0.95,1.92,1.45,0.90,2.10,0.79]
    for row_i,rowx in enumerate(table.rows):
        trPr=rowx._tr.get_or_add_trPr()
        cant_split=OxmlElement("w:cantSplit"); trPr.append(cant_split)
        _set_docx_row_height(rowx, 2160, "exact")  # 1.5 inch for every observation-table row.
        for i,cell in enumerate(rowx.cells):
            cell.width=Inches(widths[i])
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p2 in cell.paragraphs:
                p2.paragraph_format.space_after=Pt(0)
                p2.paragraph_format.space_before=Pt(0)
                p2.paragraph_format.line_spacing=0.85
                for run in p2.runs:
                    run.font.size=Pt(8)
    style_docx_table(table, header=True, font_size=8)
    add_docx_footer(doc)
    save_docx_validated(doc, path)


def _set_docx_cell_margins(cell, top=40, start=50, bottom=40, end=50):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def _set_docx_row_height(row, height_twips, rule="atLeast"):
    trPr = row._tr.get_or_add_trPr()
    trHeight = trPr.find(qn("w:trHeight"))
    if trHeight is None:
        trHeight = OxmlElement("w:trHeight")
        trPr.append(trHeight)
    trHeight.set(qn("w:val"), str(height_twips))
    trHeight.set(qn("w:hRule"), rule)


def _card_cell_text(cell, label, value, label_width=None, font_size=6.4):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 0.85
    r = p.add_run(xml_safe(label))
    r.bold = True
    r.font.size = Pt(font_size)
    if value:
        r2 = p.add_run("  " + xml_safe(value))
        r2.font.size = Pt(font_size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _set_docx_cell_margins(cell)


def _add_card_image_evidence(cell, attachments):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    image_count = 0
    non_images = []
    for a in attachments:
        fp = Path(safe(a["file_path"]))
        if not fp.exists():
            continue
        if fp.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp"}:
            try:
                if image_count:
                    p.add_run("   ")
                run = p.add_run()
                run.add_picture(str(fp), width=Inches(0.72))
                image_count += 1
            except Exception:
                non_images.append(fp.name)
        else:
            non_images.append(fp.name)
    if image_count == 0 and not non_images:
        p.add_run("No evidence attached.").font.size = Pt(6)
    elif non_images:
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.space_before = Pt(0)
        r = p2.add_run("Files: " + "; ".join(non_images))
        r.font.size = Pt(5.6)
    return image_count


def generate_observation_card_docx(observation_ids, path):
    """Create one professional STOP Card per selected observation.
    Each card occupies exactly one A4 portrait page. Multiple selected observations
    are placed in the same Word file, one card per page.
    """
    if isinstance(observation_ids, (int, str)):
        observation_ids = [int(observation_ids)]
    ids = []
    for value in observation_ids or []:
        try:
            iv = int(value)
            if iv not in ids:
                ids.append(iv)
        except Exception:
            continue
    if not ids:
        raise ValueError("No observation was selected.")

    rows = []
    for oid in ids:
        row = db.fetchone("SELECT * FROM observations WHERE id=?", (oid,))
        if row:
            rows.append(row)
    if not rows:
        raise ValueError("The selected observation(s) no longer exist.")

    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.left_margin = Inches(0.12)
    section.right_margin = Inches(0.12)
    section.top_margin = Inches(0.12)
    section.bottom_margin = Inches(0.12)
    section.header_distance = Inches(0.03)
    section.footer_distance = Inches(0.03)

    def add_border(cell, color="777777", sz=5):
        set_docx_cell_borders(cell,
            top={"val":"single","sz":sz,"color":color},
            bottom={"val":"single","sz":sz,"color":color},
            left={"val":"single","sz":sz,"color":color},
            right={"val":"single","sz":sz,"color":color})

    def add_text(cell, text, size=7.0, bold=False, align=None):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 0.82
        if align is not None:
            p.alignment = align
        r = p.add_run(xml_safe(text) if text else "")
        r.font.size = Pt(size)
        r.bold = bold
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _set_docx_cell_margins(cell, 18, 25, 18, 25)

    def add_section_heading(cell, title, size=7.0):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(xml_safe(title))
        r.bold = True
        r.font.size = Pt(size)
        add_docx_cell_shading(cell, "E7EDF7")
        _set_docx_cell_margins(cell, 18, 25, 18, 25)
        add_border(cell)

    for card_index, row in enumerate(rows):
        if card_index:
            doc.add_page_break()

        # Header: use the complete A4 width rather than the reduced sample size.
        header = doc.add_table(rows=1, cols=3)
        header.autofit = False
        header.alignment = WD_TABLE_ALIGNMENT.CENTER
        hw = [1.05, 5.70, 1.28]
        for i, w in enumerate(hw):
            header.rows[0].cells[i].width = Inches(w)
            _set_docx_cell_margins(header.rows[0].cells[i], 20, 30, 20, 30)
            add_border(header.rows[0].cells[i], sz=6)

        c = header.rows[0].cells[0]
        c.text = ""
        logo_path = report_logo_path()
        if logo_path and Path(logo_path).exists():
            try:
                c.paragraphs[0].add_run().add_picture(logo_path, width=Inches(0.82))
            except Exception:
                add_text(c, company_name(), 7.5, True, WD_ALIGN_PARAGRAPH.CENTER)
        else:
            add_text(c, company_name(), 7.5, True, WD_ALIGN_PARAGRAPH.CENTER)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        c = header.rows[0].cells[1]
        c.text = ""
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("HSE OBSERVATION AND STOP CARD")
        r.bold = True; r.font.size = Pt(13)
        company_line = c.add_paragraph()
        company_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        company_line.paragraph_format.space_after = Pt(0)
        company_run = company_line.add_run(xml_safe(company_name()))
        company_run.bold = True; company_run.font.size = Pt(8)
        project_name = safe(db.setting("project_name", ""))
        if project_name:
            project_line = c.add_paragraph()
            project_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
            project_line.paragraph_format.space_after = Pt(0)
            project_run = project_line.add_run("Project: " + xml_safe(project_name))
            project_run.font.size = Pt(6.5)

        c = header.rows[0].cells[2]
        c.text = ""
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("Document Number\n"); r.bold = True; r.font.size = Pt(6.5)
        r = p.add_run(xml_safe(row["number"])); r.bold = True; r.font.size = Pt(8)

        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(1)
        spacer.paragraph_format.space_before = Pt(0)

        outer = doc.add_table(rows=1, cols=2)
        outer.autofit = False
        outer.alignment = WD_TABLE_ALIGNMENT.CENTER
        outer_widths = [3.95, 3.96]
        for i, w in enumerate(outer_widths):
            outer.rows[0].cells[i].width = Inches(w)
            outer.rows[0].cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            _set_docx_cell_margins(outer.rows[0].cells[i], 12, 14, 12, 14)
            add_border(outer.rows[0].cells[i], sz=6)

        # LEFT SIDE ---------------------------------------------------------
        left = outer.rows[0].cells[0]
        left.text = ""
        lt = left.add_table(rows=1, cols=2)
        lt.autofit = False
        lt.alignment = WD_TABLE_ALIGNMENT.CENTER
        lt.columns[0].width = Inches(1.30)
        lt.columns[1].width = Inches(2.57)
        lt.rows[0].cells[0].merge(lt.rows[0].cells[1])
        add_section_heading(lt.rows[0].cells[0], "OBSERVATION / STOP CARD", 7.5)

        fields = [
            ("Originator's Name", row["observer"]),
            ("Company", row["observer_company"] or row["company"]),
            ("Observer Designation", row["observer_designation"]),
            ("Time and Date", f'{safe(row["obs_date"])}  {safe(row["obs_time"])}'.strip()),
            ("Location", row["location"]),
            ("Project / Area", " / ".join(x for x in [safe(row["project"]), safe(row["area"])] if x)),
            ("Reported / Responsible", row["responsible"]),
            ("Status", row["status"]),
            ("Target Date", row["target_date"]),
        ]
        for label, value in fields:
            cells = lt.add_row().cells
            cells[0].width = Inches(1.30); cells[1].width = Inches(2.57)
            add_text(cells[0], label, 6.1, True)
            add_text(cells[1], value, 6.5)
            add_docx_cell_shading(cells[0], "F0F0F0")
            add_border(cells[0]); add_border(cells[1])

        cells = lt.add_row().cells
        cells[0].merge(cells[1])
        add_section_heading(cells[0], "Observation Type", 7.0)
        p = cells[0].add_paragraph()
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 0.75
        for i, typ in enumerate(OBS_TYPES):
            mark = "☑" if typ == safe(row["obs_type"]) else "☐"
            rr = p.add_run(f"{mark} {typ}")
            rr.font.size = Pt(5.9)
            if i < len(OBS_TYPES)-1:
                rr.add_text("   ")

        cells = lt.add_row().cells
        cells[0].merge(cells[1])
        add_section_heading(cells[0], "Observation Details", 7.0)
        p = cells[0].add_paragraph(xml_safe(row["observation"]) or "No observation detail recorded.")
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0); p.paragraph_format.line_spacing = 0.82
        for rr in p.runs: rr.font.size = Pt(7.2)

        cells = lt.add_row().cells
        cells[0].merge(cells[1])
        add_section_heading(cells[0], "Action Taken / Corrective Action", 7.0)
        p = cells[0].add_paragraph(xml_safe(format_corrective_actions_display(row)) or "No action recorded.")
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0); p.paragraph_format.line_spacing = 0.82
        for rr in p.runs: rr.font.size = Pt(7.2)

        cells = lt.add_row().cells
        cells[0].merge(cells[1])
        add_section_heading(cells[0], "Related Factors / Category", 7.0)
        p = cells[0].add_paragraph()
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
        related = safe(row["category"])
        if safe(row["subcategory"]):
            related += " / " + safe(row["subcategory"])
        rr = p.add_run(("☑ " + related) if related else "☐ Not specified")
        rr.font.size = Pt(6.7)

        cells = lt.add_row().cells
        cells[0].merge(cells[1])
        add_section_heading(cells[0], "Evidence", 7.0)
        _add_card_image_evidence(cells[0], attachment_rows("observation_attachments", row["id"]))

        # RIGHT SIDE --------------------------------------------------------
        right = outer.rows[0].cells[1]
        right.text = ""
        rt = right.add_table(rows=1, cols=2)
        rt.autofit = False
        rt.alignment = WD_TABLE_ALIGNMENT.CENTER
        rt.columns[0].width = Inches(1.12)
        rt.columns[1].width = Inches(2.72)
        rt.rows[0].cells[0].merge(rt.rows[0].cells[1])
        add_section_heading(rt.rows[0].cells[0], "HAZARD IDENTIFICATION CARD", 7.5)

        selected_cat = safe(row["category"])
        selected_subcat = safe(row["subcategory"])
        selected_priority = safe(row["priority"])
        selected_status = safe(row["status"])

        hazard_cells = rt.add_row().cells
        hazard_cells[0].width = Inches(1.12); hazard_cells[1].width = Inches(2.72)
        add_text(hazard_cells[0], "Hazard Category", 6.2, True, WD_ALIGN_PARAGRAPH.CENTER)
        add_docx_cell_shading(hazard_cells[0], "F0F0F0")
        hazard_cells[1].text = ""
        # Three compact columns: every menu category remains visible without using the sample card.
        grid = hazard_cells[1].add_table(rows=0, cols=3)
        grid.autofit = False
        for col in grid.columns:
            col.width = Inches(0.88)
        n = len(CATEGORIES)
        per_col = (n + 2) // 3
        for rix in range(per_col):
            cells3 = grid.add_row().cells
            for j in range(3):
                idx_cat = rix + j * per_col
                text_cat = CATEGORIES[idx_cat] if idx_cat < n else ""
                add_text(cells3[j], (("☑ " if text_cat == selected_cat else "☐ ") + text_cat) if text_cat else "", 5.0)
                if text_cat == selected_cat:
                    for run in cells3[j].paragraphs[0].runs:
                        run.bold = True
        if selected_subcat:
            p = hazard_cells[1].add_paragraph()
            p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
            rr = p.add_run("Subcategory: " + selected_subcat); rr.bold = True; rr.font.size = Pt(5.8)
        add_border(hazard_cells[0]); add_border(hazard_cells[1])

        def add_card_checklist(label, options, selected_value, font_size=5.9):
            cells = rt.add_row().cells
            add_text(cells[0], label, 6.2, True, WD_ALIGN_PARAGRAPH.CENTER)
            add_docx_cell_shading(cells[0], "F0F0F0")
            cells[1].text = ""
            p = cells[1].paragraphs[0]
            p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0); p.paragraph_format.line_spacing = 0.76
            for option in options:
                mark = "☑" if option == selected_value else "☐"
                rr = p.add_run(f"{mark} {option}   ")
                rr.font.size = Pt(font_size)
            add_border(cells[0]); add_border(cells[1])

        add_card_checklist("Priority", PRIORITIES, selected_priority, 6.0)
        add_card_checklist("Status", STATUSES, selected_status, 6.0)

        action_cells = rt.add_row().cells
        add_text(action_cells[0], "Action / Control", 6.2, True, WD_ALIGN_PARAGRAPH.CENTER)
        add_docx_cell_shading(action_cells[0], "F0F0F0")
        add_text(action_cells[1], format_corrective_actions_display(row) or "No action recorded.", 6.2)
        add_border(action_cells[0]); add_border(action_cells[1])

        # Keep the two nested tables visually compact and prevent row splitting.
        for nested in (lt, rt, grid):
            for rr in nested.rows:
                trPr = rr._tr.get_or_add_trPr()
                cant = OxmlElement("w:cantSplit")
                trPr.append(cant)
                for cell in rr.cells:
                    _set_docx_cell_margins(cell, 14, 20, 14, 20)

        # Keep the whole outer card on one page where possible.
        for rr in outer.rows:
            trPr = rr._tr.get_or_add_trPr()
            trPr.append(OxmlElement("w:cantSplit"))

    add_docx_footer(doc)
    save_docx_validated(doc, path)

def generate_table_docx(table_name, path, record_ids=None):
    if table_name == "observations":
        return generate_observation_docx(path, record_ids=record_ids)
    columns, data = MainWindow.table_data_static(table_name, record_ids=record_ids)
    if not columns:
        raise ValueError("There is no data to export.")
    doc = Document()
    add_docx_header(doc)
    add_docx_title(doc, f"{table_name.replace('_', ' ').title()} Report")
    p = doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    table = doc.add_table(rows=1, cols=len(columns))
    for i, col in enumerate(columns):
        table.rows[0].cells[i].text = xml_safe(col)
    if table_name == "incidents" and "Attachments" in columns:
        incident_rows = db.fetchall("SELECT * FROM incidents ORDER BY id DESC")
        if record_ids:
            wanted=set(record_ids); incident_rows=[r for r in incident_rows if r["id"] in wanted]
        attach_col=columns.index("Attachments")
        for r in incident_rows[:1000]:
            cells=table.add_row().cells
            # Populate the same register fields as table_data_static, but replace the
            # attachment-count/text cell with the actual evidence files.
            for i,col in enumerate(columns):
                if col != "Attachments":
                    cells[i].text=xml_safe(r[col] if col in r.keys() else "")
            attachments=attachment_rows("incident_attachments",r["id"])
            if not attachments:
                cells[attach_col].text="No attachment"
            else:
                cells[attach_col].text=""
                for a in attachments:
                    fp=Path(safe(a["file_path"]))
                    remark=safe(a["remark"]) if "remark" in a.keys() else ""
                    if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
                        try:
                            pp=cells[attach_col].paragraphs[0] if not cells[attach_col].paragraphs[0].runs else cells[attach_col].add_paragraph()
                            pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
                            pp.add_run().add_picture(str(fp),width=Inches(1.25))
                            cap=cells[attach_col].add_paragraph(fp.name + (f" | {remark}" if remark else ""))
                            cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
                        except Exception:
                            cells[attach_col].add_paragraph(fp.name + (f" | {remark}" if remark else ""))
                    else:
                        cells[attach_col].add_paragraph(fp.name + (f" | {remark}" if remark else ""))
    else:
        for row in data[:1000]:
            cells = table.add_row().cells
            for i, value in enumerate(row):
                cells[i].text = xml_safe(value)
    style_docx_table(table)
    add_docx_footer(doc)
    save_docx_validated(doc, path)


class PieChartWidget(QWidget):
    def __init__(self, values=None, parent=None):
        super().__init__(parent)
        self.values = values or {}
        self.setMinimumHeight(260)

    def set_values(self, values):
        self.values = values or {}
        self.update()

    def paintEvent(self, event):
        painter=QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect=self.rect().adjusted(20,20,-20,-20)
        total=sum(max(0,int(v)) for v in self.values.values())
        if total<=0:
            painter.drawText(rect,Qt.AlignmentFlag.AlignCenter,"No observation data")
            painter.end(); return
        center=rect.center()
        radius=min(rect.width(),rect.height())//3
        start=0
        palette=[Qt.GlobalColor.darkBlue,Qt.GlobalColor.darkGreen,Qt.GlobalColor.darkRed,
                 Qt.GlobalColor.darkCyan,Qt.GlobalColor.darkMagenta,Qt.GlobalColor.darkYellow,
                 Qt.GlobalColor.gray]
        for i,(label,value) in enumerate(self.values.items()):
            value=max(0,int(value))
            span=int(round(360*16*value/total))
            painter.setBrush(QBrush(palette[i%len(palette)]))
            painter.setPen(QPen(Qt.GlobalColor.white,1))
            painter.drawPie(QRectF(center.x()-radius,center.y()-radius,2*radius,2*radius),start,span)
            start += span
        y=25
        x=rect.left()
        painter.setFont(QFont("Arial",9))
        for i,(label,value) in enumerate(self.values.items()):
            painter.setBrush(QBrush(palette[i%len(palette)]))
            painter.drawRect(x,y,14,14)
            painter.setPen(QPen(Qt.GlobalColor.black))
            painter.drawText(x+20,y+12,f"{label}: {value}")
            y += 20
        painter.end()


class BarChartWidget(QWidget):
    def __init__(self, values=None, parent=None):
        super().__init__(parent)
        self.values=values or {}
        self.setMinimumHeight(260)

    def set_values(self, values):
        self.values=values or {}
        self.update()

    def paintEvent(self,event):
        painter=QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect=self.rect().adjusted(45,20,-20,-45)
        if not self.values:
            painter.drawText(rect,Qt.AlignmentFlag.AlignCenter,"No data")
            painter.end(); return
        maxv=max([int(v) for v in self.values.values()] or [1])
        count=len(self.values)
        barw=max(18,int(rect.width()/max(count,1)*0.65))
        gap=max(8,int(rect.width()/max(count,1)*0.35))
        x=rect.left()
        colors=[Qt.GlobalColor.darkBlue,Qt.GlobalColor.darkGreen,Qt.GlobalColor.darkRed,
                Qt.GlobalColor.darkCyan,Qt.GlobalColor.darkMagenta,Qt.GlobalColor.darkYellow]
        for i,(label,value) in enumerate(self.values.items()):
            value=int(value)
            h=int(rect.height()*value/maxv) if maxv else 0
            y=rect.bottom()-h
            painter.setBrush(QBrush(colors[i%len(colors)]))
            painter.setPen(QPen(Qt.GlobalColor.black))
            painter.drawRect(x,y,barw,h)
            painter.setPen(QPen(Qt.GlobalColor.black))
            painter.drawText(x,y-5,str(value))
            painter.save()
            painter.translate(x+barw/2,rect.bottom()+12)
            painter.rotate(-45)
            painter.drawText(0,0,str(label)[:16])
            painter.restore()
            x += barw+gap
        painter.drawLine(rect.left(),rect.bottom(),rect.right(),rect.bottom())
        painter.end()


# ============================================================
# MAIN WINDOW
# ============================================================

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle(
            f"{APP_NAME} - {APP_VERSION}"
        )

        self.resize(1400, 850)

        self.setStyleSheet("""
        QMainWindow {
            background: #F4F6F8;
        }

        QLabel {
            color: #202020;
        }

        QPushButton {
            background: #17365D;
            color: white;
            border: none;
            padding: 9px 14px;
            border-radius: 5px;
            min-height: 20px;
        }

        QPushButton:hover {
            background: #245486;
        }

        QLineEdit, QTextEdit, QComboBox, QDateEdit {
            background: white;
            border: 1px solid #B8C2CC;
            border-radius: 4px;
            padding: 7px;
        }

        QTableWidget {
            background: white;
            gridline-color: #D5DADF;
        }

        QHeaderView::section {
            background: #17365D;
            color: white;
            padding: 7px;
            border: none;
        }

        QListWidget {
            background: #102A43;
            color: white;
            border: none;
        }

        QListWidget::item {
            padding: 14px;
        }

        QListWidget::item:selected {
            background: #245486;
        }

        QGroupBox {
            font-weight: bold;
            border: 1px solid #CCD3DA;
            margin-top: 10px;
            padding: 10px;
        }
        """)

        self.build_ui()
        self.dashboard()

    # --------------------------------------------------------

    def build_ui(self):

        central = QWidget()
        self.setCentralWidget(central)

        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)

        self.nav = QListWidget()

        modules = [
            "Dashboard",
            "HSE Inspection Register",
            "Incident Investigation",
            "Audit Register",
            "CAPA Tracking Register",
            "Permit to Work (PTW)",
            "Training",
            "Reports",
            "Master Data",
            "Settings"
        ]

        self.nav.addItems(modules)
        self.nav.setFixedWidth(245)
        self.nav.currentRowChanged.connect(
            self.navigation_changed
        )

        layout.addWidget(self.nav)

        self.stack = QStackedWidget()
        layout.addWidget(self.stack, 1)

    # --------------------------------------------------------

    def navigation_changed(self, row):

        if row == 0:
            self.dashboard()
        elif row == 1:
            self.observations()
        elif row == 2:
            self.incidents()
        elif row == 3:
            self.audits()
        elif row == 4:
            self.capa()
        elif row == 5:
            w, layout = self.page("Permit to Work Management")
            layout.addWidget(PTWModule(self, db), 1)
        elif row == 6:
            self.training()
        elif row == 7:
            self.reports()
        elif row == 8:
            self.master_data()
        elif row == 9:
            self.settings_page()

    # --------------------------------------------------------

    def page(self, title):

        w = QWidget()
        layout = QVBoxLayout(w)

        title_label = QLabel(title)
        title_label.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            color: #17365D;
            padding: 8px;
        """)

        layout.addWidget(title_label)

        self.stack.addWidget(w)
        self.stack.setCurrentWidget(w)

        return w, layout

    # ========================================================
    # DASHBOARD
    # ========================================================


    def dashboard(self):
        w,layout=self.page("HSE Dashboard")
        cards=QHBoxLayout()
        values=[
            ("Observations","SELECT COUNT(*) c FROM observations"),
            ("Open Observations","SELECT COUNT(*) c FROM observations WHERE status NOT IN ('Closed','Cancelled')"),
            ("Overdue","SELECT COUNT(*) c FROM observations WHERE status NOT IN ('Closed','Cancelled') AND target_date < date('now')"),
            ("Incidents","SELECT COUNT(*) c FROM incidents"),
            ("Audits","SELECT COUNT(*) c FROM audits"),
            ("CAPA","SELECT COUNT(*) c FROM capa")
        ]
        for name,sql in values:
            value=db.fetchone(sql)["c"]
            box=QGroupBox(name); bl=QVBoxLayout(box)
            label=QLabel(str(value)); label.setStyleSheet("font-size:28px;font-weight:bold;color:#17365D;")
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            bl.addWidget(label); cards.addWidget(box)
        layout.addLayout(cards)
        layout.addWidget(QLabel(f"<b>Company:</b> {company_name()}    <b>Project:</b> {db.setting('project_name','')}"))
        selector=QComboBox(); selector.addItems(["HSE Home Dashboard","Inspection Dashboard","Incident Dashboard"]); layout.addWidget(selector)
        def switch_dashboard(text):
            if text=="Inspection Dashboard": self.observations()
            elif text=="Incident Dashboard": self.incidents()
        selector.currentTextChanged.connect(switch_dashboard)

        charts=QHBoxLayout()
        pie=PieChartWidget()
        obs_types=db.fetchall("SELECT obs_type,COUNT(*) c FROM observations GROUP BY obs_type ORDER BY obs_type COLLATE NOCASE")
        pie.set_values({safe(r["obs_type"]) or "Unspecified":r["c"] for r in obs_types})
        pie_box=QGroupBox("Observation Type Distribution (no categories merged)"); pl=QVBoxLayout(pie_box); pl.addWidget(pie)
        bar=BarChartWidget()
        stats={"Observations":db.fetchone("SELECT COUNT(*) c FROM observations")["c"],
               "Incidents":db.fetchone("SELECT COUNT(*) c FROM incidents")["c"],
               "Audits":db.fetchone("SELECT COUNT(*) c FROM audits")["c"],
               "CAPA":db.fetchone("SELECT COUNT(*) c FROM capa")["c"]}
        bar.set_values(stats)
        bar_box=QGroupBox("HSE Register Summary"); bl=QVBoxLayout(bar_box); bl.addWidget(bar)
        charts.addWidget(pie_box,1); charts.addWidget(bar_box,1); layout.addLayout(charts)

        # A clear, complete category register. Every category is a separate row, including zero counts.
        cat_box=QGroupBox("HSE Observation Categories — click a row to view that category")
        cat_layout=QVBoxLayout(cat_box)
        cat_table=QTableWidget()
        categories=sorted(set(CATEGORIES), key=lambda x: x.lower())
        counts={safe(r["category"]):int(r["c"]) for r in db.fetchall("SELECT category,COUNT(*) c FROM observations GROUP BY category")}
        cat_table.setColumnCount(3)
        cat_table.setHorizontalHeaderLabels(["Category","Observations","Open / Active"])
        cat_table.setRowCount(len(categories))
        cat_table.horizontalHeader().setSectionResizeMode(0,QHeaderView.Stretch)
        cat_table.horizontalHeader().setSectionResizeMode(1,QHeaderView.ResizeToContents)
        cat_table.horizontalHeader().setSectionResizeMode(2,QHeaderView.ResizeToContents)
        cat_table.setAlternatingRowColors(True)
        cat_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        cat_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        for r,cat in enumerate(categories):
            total=counts.get(cat,0)
            active=db.fetchone("SELECT COUNT(*) c FROM observations WHERE category=? AND status NOT IN ('Closed','Cancelled')",(cat,))["c"]
            cat_table.setItem(r,0,QTableWidgetItem(cat))
            cat_table.setItem(r,1,QTableWidgetItem(str(total)))
            cat_table.setItem(r,2,QTableWidgetItem(str(active)))
        cat_table.setMinimumHeight(min(330, max(150, len(categories)*26+35)))
        def open_category(row,_column):
            item=cat_table.item(row,0)
            if item:
                self.observation_dashboard_filter=item.text()
                self.observations()
        cat_table.cellDoubleClicked.connect(open_category)
        cat_layout.addWidget(cat_table)
        layout.addWidget(cat_box,1)

        recent_inc=QGroupBox("Recent Incidents"); ril=QVBoxLayout(recent_inc); it=QTableWidget(); it.setColumnCount(5); it.setHorizontalHeaderLabels(["ID / Number","Date","Type","Status","Completion"]); ir=db.fetchall("SELECT id,number,incident_date,incident_type,status,incident_title,description,investigation_method,direct_cause,root_cause FROM incidents ORDER BY id DESC LIMIT 10"); it.setRowCount(len(ir))
        for rr,x in enumerate(ir):
            total=10; done=sum(bool(safe(x[k])) for k in ["incident_date","incident_time"] if k in x.keys())
            done += sum(bool(safe(x[k])) for k in ["location","project","incident_type","description","investigation_method","direct_cause","root_cause"] if k in x.keys())
            pct=round(done*100/total)
            for cc,v in enumerate([f"{safe(x['id'])} / {safe(x['number'])}",x['incident_date'],x['incident_type'],x['status'],f"{pct}%"]): it.setItem(rr,cc,QTableWidgetItem(safe(v)))
        ril.addWidget(it); layout.addWidget(recent_inc)

        recent=QGroupBox("Recent Observations")
        recent_layout=QVBoxLayout(recent)
        table=QTableWidget(); table.setColumnCount(8)
        table.setHorizontalHeaderLabels(["S.No.","Observation No.","Date","Type","Category","Location","Status","Action Taken"])
        table.horizontalHeader().setSectionResizeMode(7,QHeaderView.Stretch)
        for c in range(7): table.horizontalHeader().setSectionResizeMode(c,QHeaderView.ResizeToContents)
        rows=db.fetchall("SELECT id,number,obs_date,obs_type,category,location,status,corrective_action FROM observations ORDER BY id DESC LIMIT 10")
        table.setRowCount(len(rows))
        for r,row in enumerate(rows):
            vals=[str(r+1),row["number"],row["obs_date"],row["obs_type"],row["category"],row["location"],row["status"],format_corrective_actions_display(row)]
            for c,v in enumerate(vals):
                item=QTableWidgetItem(safe(v)); item.setToolTip(safe(v)); table.setItem(r,c,item)
        table.setAlternatingRowColors(True); table.setWordWrap(True)
        recent_layout.addWidget(table)
        layout.addWidget(recent,1)


    def observations(self):
        w, layout = self.page("HSE Inspection Register")
        toolbar = QHBoxLayout()
        search = QLineEdit()
        search.setPlaceholderText("Search observation no., location, responsible person, description, action...")
        toolbar.addWidget(search)
        clear_filter=QPushButton("Clear Category Filter")
        toolbar.addWidget(clear_filter)
        add_button = QPushButton("+ New Observation")
        attachment_button = QPushButton("Attachments")
        delete_button = QPushButton("Delete Selected")
        export_button = QPushButton("Export CSV")
        excel_button = QPushButton("Export Excel")
        pdf_button = QPushButton("Export PDF")
        word_button = QPushButton("Export Word")
        card_button = QPushButton("Draft STOP Card(s)")
        for b in [add_button, attachment_button, delete_button, export_button, excel_button, pdf_button, word_button, card_button]:
            toolbar.addWidget(b)
        layout.addLayout(toolbar)

        table = QTableWidget()
        headers = ["ID", "S.No.", "Observation No.", "Observation Date", "Observation Type", "Observation Detail", "Status",
                   "Location", "Corrective Action(s)", "Target Date", "CAPA Required", "CAPA No.", "Attachments"]
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setColumnHidden(0, True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        table.setWordWrap(True)
        table.setAlternatingRowColors(True)
        layout.addWidget(table)

        def load():
            term = search.text().strip()
            category_filter = getattr(self, "observation_dashboard_filter", "")
            query = "SELECT * FROM observations"
            clauses=[]; params=[]
            if term:
                clauses.append("(number LIKE ? OR location LIKE ? OR responsible LIKE ? OR observation LIKE ? OR corrective_action LIKE ? OR category LIKE ? OR COALESCE(capa_number,'') LIKE ?)")
                like = f"%{term}%"
                params.extend([like, like, like, like, like, like, like])
            if category_filter:
                clauses.append("category=?"); params.append(category_filter)
            if clauses:
                query += " WHERE " + " AND ".join(clauses)
            query += " ORDER BY id DESC"
            rows = db.fetchall(query, tuple(params))
            table.setRowCount(len(rows))
            for r, row in enumerate(rows):
                attachments = attachment_names("observation_attachments", row["id"])
                actions_text = format_corrective_actions_display(row)
                try:
                    capa_req = "Yes" if int(row["capa_required"] or 0) else "No"
                except Exception:
                    capa_req = "No"
                try:
                    capa_no = safe(row["capa_number"]) or "—"
                except Exception:
                    capa_no = "—"
                values = [
                    row["id"], r + 1, row["number"], row["obs_date"], row["obs_type"], row["observation"], row["status"],
                    row["location"], actions_text, row["target_date"], capa_req, capa_no,
                    attachments or "No attachment"
                ]
                for c, value in enumerate(values):
                    item = QTableWidgetItem(safe(value))
                    item.setToolTip(safe(value))
                    table.setItem(r, c, item)
                if overdue(row["target_date"], row["status"]):
                    table.item(r, 9).setForeground(Qt.GlobalColor.red)
                    table.item(r, 6).setForeground(Qt.GlobalColor.red)
                if capa_req == "Yes":
                    table.item(r, 10).setForeground(Qt.GlobalColor.darkBlue)
                table.setRowHeight(r, max(48, 18 + 16 * max(1, len(parse_corrective_actions(row)))))

        def selected_observation_ids(require_one=False):
            indexes = table.selectionModel().selectedRows(0) if table.selectionModel() else []
            ids = []
            for index in indexes:
                item = table.item(index.row(), 0)
                if item:
                    try:
                        oid = int(item.text())
                        if oid not in ids:
                            ids.append(oid)
                    except Exception:
                        pass
            if not ids:
                current = table.currentRow()
                if current >= 0 and table.item(current, 0):
                    try:
                        ids = [int(table.item(current, 0).text())]
                    except Exception:
                        ids = []
            if require_one and len(ids) != 1:
                QMessageBox.warning(self, "Observation", "Select exactly one observation for this action.")
                return []
            if not ids:
                QMessageBox.warning(self, "Observation", "Select at least one observation first.")
            return ids

        def selected_observation_id():
            ids = selected_observation_ids(require_one=True)
            return ids[0] if ids else None

        def draft_selected_card():
            obs_ids = selected_observation_ids()
            if not obs_ids:
                return
            records = []
            for obs_id in obs_ids:
                record = db.fetchone("SELECT number FROM observations WHERE id=?", (obs_id,))
                if record:
                    records.append(record)
            if not records:
                QMessageBox.warning(self, "Observation", "The selected observation(s) could not be found.")
                return
            if len(records) == 1:
                default_name = f"{safe(records[0]['number'])}_Observation_STOP_Card.docx"
            else:
                default_name = "Selected_Observations_STOP_Cards.docx"
            path, _ = QFileDialog.getSaveFileName(
                self, "Draft Observation STOP Card(s)", default_name, "Word Documents (*.docx)"
            )
            if not path:
                return
            try:
                generate_observation_card_docx(obs_ids, path)
                self.show_export_success(path)
            except Exception as e:
                logging.exception("Observation card report failed")
                QMessageBox.critical(self, "Report Error", str(e))

        def manage_attachments():
            obs_id = selected_observation_id()
            if obs_id is None:
                return
            record = db.fetchone("SELECT number FROM observations WHERE id=?", (obs_id,))
            if not record:
                return
            dialog = QDialog(self)
            dialog.setWindowTitle(f"Attachments - {record['number']}")
            dialog.resize(760, 480)
            v = QVBoxLayout(dialog)
            v.addWidget(QLabel(f"<b>Observation:</b> {safe(record['number'])}"))
            listw = QListWidget()
            v.addWidget(listw, 1)
            def refresh_files():
                listw.clear()
                rows = attachment_rows("observation_attachments", obs_id)
                for a in rows:
                    p = Path(safe(a["file_path"]))
                    listw.addItem(f"{p.name}  |  {safe(a['attachment_type'])}")
                if not rows:
                    listw.addItem("No attachments recorded.")
            def add_files():
                paths, _ = QFileDialog.getOpenFileNames(
                    dialog, "Add Observation Evidence", "",
                    "Evidence Files (*.png *.jpg *.jpeg *.bmp *.pdf *.doc *.docx *.xls *.xlsx *.txt *.mp4 *.avi *.mov);;All Files (*)"
                )
                if paths:
                    copy_attachments(paths, safe(record["number"]), "observation_attachments", obs_id)
                    refresh_files(); load()
            add = QPushButton("Add Attachment")
            close = QPushButton("Close")
            rowbtn = QHBoxLayout(); rowbtn.addWidget(add); rowbtn.addStretch(); rowbtn.addWidget(close); v.addLayout(rowbtn)
            add.clicked.connect(add_files); close.clicked.connect(dialog.accept)
            refresh_files(); dialog.exec()

        def delete_selected():
            obs_id = selected_observation_id()
            if obs_id is None:
                return
            record = db.fetchone("SELECT number FROM observations WHERE id=?", (obs_id,))
            if not record:
                return
            answer = QMessageBox.question(
                self, "Confirm Delete",
                f"Delete observation {record['number']} and its evidence?\n\nThis cannot be undone."
            )
            if answer != QMessageBox.Yes:
                return
            for a in attachment_rows("observation_attachments", obs_id):
                try:
                    Path(safe(a["file_path"])).unlink(missing_ok=True)
                except Exception:
                    logging.exception("Unable to remove observation attachment")
            db.execute("DELETE FROM observations WHERE id=?", (obs_id,))
            load()

        def clear_category_filter():
            self.observation_dashboard_filter=""
            search.clear()
            load()
        clear_filter.clicked.connect(clear_category_filter)
        search.textChanged.connect(load)
        add_button.clicked.connect(lambda: self.observation_form(load))
        attachment_button.clicked.connect(manage_attachments)
        delete_button.clicked.connect(delete_selected)
        export_button.clicked.connect(lambda: self.export_csv("observations"))
        excel_button.clicked.connect(lambda: self.export_excel("observations"))
        pdf_button.clicked.connect(lambda: self.export_pdf("observations"))
        word_button.clicked.connect(lambda: self.export_docx("observations"))
        card_button.clicked.connect(draft_selected_card)
        load()


    def observation_form(self, refresh):
        dialog = QDialog(self)
        dialog.setWindowTitle("New HSE Observation / Inspection")
        dialog.resize(980, 820)
        dialog.setMinimumSize(900, 700)
        outer = QVBoxLayout(dialog)

        header = QLabel("HSE Inspection Observation")
        header.setStyleSheet("font-size:20px;font-weight:bold;color:#17365D;padding:6px;")
        outer.addWidget(header)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        content = QWidget()
        form = QFormLayout(content)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        edits = {}

        def edit(name):
            e = QLineEdit()
            edits[name] = e
            form.addRow(name.replace("_", " ").title() + ":", e)
            return e

        observation_date = QDateEdit(); observation_date.setCalendarPopup(True); observation_date.setDate(datetime.now().date())
        form.addRow("Observation Date:", observation_date)

        for name in ["project", "company", "location", "area", "responsible", "designation", "employee_id",
                     "observer", "observer_designation", "observer_id", "observer_company"]:
            edit(name)
        edits["project"].setText(safe(db.setting("project_name", "")))
        edits["company"].setText(safe(db.setting("company_name", "")))

        obs_type = QComboBox(); obs_type.addItems(OBS_TYPES)
        form.addRow("Observation Type:", obs_type)
        category = QComboBox(); category.addItems(CATEGORIES)
        form.addRow("Category:", category)
        edit("subcategory")

        observation = QTextEdit(); observation.setMinimumHeight(100)
        form.addRow("Observation Description:", observation)

        actions_box = QGroupBox("Corrective Action(s) — add one or more")
        actions_layout = QVBoxLayout(actions_box)
        action_rows = []

        def add_action_row(text=""):
            row_w = QWidget()
            row_l = QHBoxLayout(row_w)
            row_l.setContentsMargins(0, 0, 0, 0)
            idx = QLabel(str(len(action_rows) + 1) + ".")
            idx.setFixedWidth(28)
            te = QTextEdit()
            te.setPlaceholderText("Corrective action required")
            te.setFixedHeight(54)
            if text:
                te.setPlainText(text)
            rem = QPushButton("Remove")
            rem.setFixedWidth(80)
            row_l.addWidget(idx)
            row_l.addWidget(te, 1)
            row_l.addWidget(rem)
            actions_layout.insertWidget(max(0, actions_layout.count() - 1), row_w)
            entry = {"widget": row_w, "edit": te, "label": idx}
            action_rows.append(entry)

            def remove():
                if len(action_rows) <= 1:
                    te.clear()
                    return
                action_rows.remove(entry)
                row_w.deleteLater()
                for i, e in enumerate(action_rows, 1):
                    e["label"].setText(str(i) + ".")
            rem.clicked.connect(remove)

        add_btn = QPushButton("+ Add Corrective Action")
        add_btn.clicked.connect(lambda: add_action_row())
        actions_layout.addWidget(add_btn)
        add_action_row()
        form.addRow(actions_box)

        priority = QComboBox(); priority.addItems(PRIORITIES)
        form.addRow("Priority:", priority)
        target = QDateEdit(); target.setCalendarPopup(True); target.setDate(datetime.now().date())
        form.addRow("Target Completion Date:", target)
        status = QComboBox(); status.addItems(STATUSES)
        form.addRow("Status:", status)

        capa_required = QComboBox()
        capa_required.addItems(["No", "Yes"])
        capa_required.setToolTip(
            "Select Yes when formal CAPA tracking is required. All observation details and corrective actions "
            "will be sent to the CAPA Tracking Register."
        )
        form.addRow("CAPA Required?:", capa_required)
        capa_hint = QLabel("If Yes, observation details and every corrective action are created in CAPA Tracking Register.")
        capa_hint.setStyleSheet("color:#245486;font-size:11px;")
        capa_hint.setWordWrap(True)
        form.addRow("", capa_hint)

        attachment_paths = []
        attachment_label = QLabel("No evidence files selected.")
        attachment_label.setWordWrap(True)
        attach_button = QPushButton("Attach Evidence Files")
        attach_button.setMinimumHeight(36)

        def choose_files():
            paths, _ = QFileDialog.getOpenFileNames(
                dialog, "Select Evidence Files", "",
                "Evidence Files (*.png *.jpg *.jpeg *.bmp *.pdf *.doc *.docx *.xls *.xlsx *.txt *.mp4 *.avi *.mov);;All Files (*)"
            )
            if paths:
                attachment_paths.clear()
                attachment_paths.extend(paths)
                attachment_label.setText("\n".join(Path(p).name for p in paths))

        attach_button.clicked.connect(choose_files)
        form.addRow("Evidence / Attachments:", attach_button)
        form.addRow("Selected Files:", attachment_label)

        scroll_area.setWidget(content)
        outer.addWidget(scroll_area, 1)

        buttons = QDialogButtonBox()
        save_button = buttons.addButton("Save Observation", QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = buttons.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
        save_button.setMinimumHeight(40)
        cancel_button.setMinimumHeight(40)
        outer.addWidget(buttons)
        cancel_button.clicked.connect(dialog.reject)

        def collect_actions():
            return [e["edit"].toPlainText().strip() for e in action_rows if e["edit"].toPlainText().strip()]

        def save():
            if not edits["location"].text().strip():
                QMessageBox.warning(dialog, "Required", "Location is required.")
                return
            if not observation.toPlainText().strip():
                QMessageBox.warning(dialog, "Required", "Observation description is required.")
                return
            actions = collect_actions()
            if not actions:
                QMessageBox.warning(dialog, "Required", "Add at least one corrective action.")
                return
            try:
                number = next_number("HSE-OBS", "observations")
                need_capa = capa_required.currentText() == "Yes"
                actions_json = json.dumps([{"action": a} for a in actions], ensure_ascii=False)
                corrective_text = format_corrective_actions_display(actions)
                db.execute("""
                    INSERT INTO observations (
                        number, obs_date, obs_time, project, company, location, area,
                        responsible, designation, employee_id, observer, observer_designation,
                        observer_id, observer_company, obs_type, category, subcategory,
                        observation, corrective_action, corrective_actions_json, priority, target_date, status,
                        capa_required, capa_number, created_at
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    number, observation_date.date().toString("yyyy-MM-dd"), now_time(), edits["project"].text(), edits["company"].text(),
                    edits["location"].text(), edits["area"].text(), edits["responsible"].text(),
                    edits["designation"].text(), edits["employee_id"].text(), edits["observer"].text(),
                    edits["observer_designation"].text(), edits["observer_id"].text(),
                    edits["observer_company"].text(), obs_type.currentText(), category.currentText(),
                    edits["subcategory"].text(), observation.toPlainText(), corrective_text, actions_json,
                    priority.currentText(), target.date().toString("yyyy-MM-dd"), status.currentText(),
                    1 if need_capa else 0, "", datetime.now().isoformat()
                ))
                row = db.fetchone("SELECT id FROM observations WHERE number=?", (number,))
                if row and attachment_paths:
                    copy_attachments(attachment_paths, number, "observation_attachments", row["id"])

                capa_number = ""
                if need_capa and row:
                    capa_number = next_number("HSE-CAPA", "capa")
                    finding_text = (
                        f"Observation {number}\n"
                        f"Type: {obs_type.currentText()} | Category: {category.currentText()}\n"
                        f"Location: {edits['location'].text()}\n\n"
                        f"{observation.toPlainText().strip()}"
                    )
                    capa_actions = [{"action": a, "evidence": [], "has_evidence": False} for a in actions]
                    db.execute("""
                        INSERT INTO capa (
                            number, source, reference_number, finding, root_cause, corrective_action,
                            preventive_action, responsible, priority, target_date, status,
                            verification, closeout_evidence, created_at,
                            actions_json, approval_status, approved_by, approved_date,
                            completion_pct, action_status, observation_text
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (
                        capa_number,
                        "HSE Inspection",
                        number,
                        finding_text,
                        "",
                        corrective_text,
                        "",
                        edits["responsible"].text().strip(),
                        priority.currentText(),
                        target.date().toString("yyyy-MM-dd"),
                        "Open",
                        "",
                        "",
                        datetime.now().isoformat(),
                        json.dumps(capa_actions, ensure_ascii=False),
                        "Pending Approval",
                        "",
                        "",
                        0,
                        "Open",
                        observation.toPlainText().strip(),
                    ))
                    db.execute("UPDATE observations SET capa_number=? WHERE id=?", (capa_number, row["id"]))

                dialog.accept()
                refresh()
                if need_capa and capa_number:
                    QMessageBox.information(
                        self, "CAPA Tracking",
                        f"Observation {number} saved.\n\n"
                        f"CAPA Required = Yes\n"
                        f"{len(actions)} corrective action(s) sent to CAPA Tracking Register as {capa_number}."
                    )
            except Exception as e:
                logging.exception("Observation save failed")
                QMessageBox.critical(dialog, "Save Error", f"Unable to save observation.\n\n{e}")

        save_button.clicked.connect(save)
        dialog.exec()


    def incidents(self):
        w, layout = self.page("Accident / Incident Investigation")
        banner=QHBoxLayout(); logo=QLabel(); logo.setFixedSize(80,60); logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_path=db.setting("company_logo","")
        if logo_path and Path(logo_path).exists():
            logo.setPixmap(QPixmap(logo_path).scaled(70,55,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation))
        banner.addWidget(logo); banner.addWidget(QLabel(f"<b>{safe(company_name())}</b><br>Accident / Incident Investigation Register"),1); layout.addLayout(banner)
        toolbar=QHBoxLayout()
        buttons=[("+ New Incident", None),("Edit Selected",None),("Delete Selected",None),("Export CSV",lambda:self.export_csv("incidents")),( "Export Excel",lambda:self.export_excel("incidents")),( "Export PDF",lambda:self.export_pdf("incidents")),( "Export Word",lambda:self.export_docx("incidents")),( "Professional Report",None)]
        btns={}
        for text,fn in buttons:
            b=QPushButton(text); toolbar.addWidget(b); btns[text]=b
            if fn:b.clicked.connect(fn)
        layout.addLayout(toolbar)
        table=QTableWidget(); headers=["ID","Number","Date","Type","Location","Project","Method","Status","Completion","Attachments","Professional Report"]
        table.setColumnCount(len(headers)); table.setHorizontalHeaderLabels(headers); table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents); table.horizontalHeader().setStretchLastSection(True); table.setSelectionBehavior(QAbstractItemView.SelectRows); layout.addWidget(table)
        def load():
            rows=db.fetchall("SELECT * FROM incidents ORDER BY id DESC"); table.setRowCount(len(rows))
            for r,row in enumerate(rows):
                # Show the actual evidence file names in the register instead of an
                # attachment count, so the user can immediately see what is attached.
                attached_names=attachment_names("incident_attachments",row["id"]) or "No attachment"
                
                total=10; done=sum(bool(safe(row[k])) for k in ["incident_date","incident_time","location","project","incident_type","description","immediate_action","investigation_method","direct_cause","root_cause"])
                pct=round(done*100/total)
                try:
                    report_details=json.loads(safe(row["investigation_details"]) or "{}")
                except Exception:
                    report_details={}
                vals=[row["id"],row["number"],row["incident_date"],row["incident_type"],row["location"],row["project"],row["investigation_method"],row["status"],f"{pct}%",attached_names,"Available" if report_details.get("professional_report_generated") else "Not Generated"]
                for c,v in enumerate(vals):table.setItem(r,c,QTableWidgetItem(safe(v)))
        def professional():
            incident_id=self.selected_id(table,"Professional Report")
            if incident_id is None:return
            number=table.item(table.currentRow(),1).text(); path,_=QFileDialog.getSaveFileName(self,"Save Investigation Report",f"{number}_Investigation_Report.docx","Word Document (*.docx)")
            if not path:return
            try:
                generate_incident_docx(incident_id,path)
                if not Path(path).exists() or Path(path).stat().st_size < 1000:
                    raise ValueError("The Investigation Professional Report was not created correctly.")
                rr=db.fetchone("SELECT investigation_details FROM incidents WHERE id=?",(incident_id,))
                try: details=json.loads(safe(rr["investigation_details"]) or "{}") if rr else {}
                except Exception: details={}
                details["professional_report_generated"]=True; details["professional_report_path"]=str(path)
                db.execute("UPDATE incidents SET investigation_details=? WHERE id=?",(json.dumps(details,ensure_ascii=False),incident_id))
                self.show_export_success(path); load()
            except Exception as e: logging.exception("Incident Word report failed"); QMessageBox.critical(self,"Report Error",str(e))
        btns["+ New Incident"].clicked.connect(lambda:self.incident_form(load))
        def edit_selected():
            rid=self.selected_id(table,"Edit Incident")
            if rid is None:return
            rr=db.fetchone("SELECT status FROM incidents WHERE id=?",(rid,))
            if rr and safe(rr["status"]).lower()=="closed": QMessageBox.information(self,"Submitted","This report has been submitted and cannot be edited."); return
            self.incident_form(load,rid)
        btns["Edit Selected"].clicked.connect(edit_selected)
        btns["Delete Selected"].clicked.connect(lambda:self.delete_selected_record(table,"incidents","incident_attachments","Incident",load))
        btns["Professional Report"].clicked.connect(professional)
        load()


    def incident_form(self, refresh, incident_id=None):
        existing = db.fetchone("SELECT * FROM incidents WHERE id=?", (incident_id,)) if incident_id else None
        if incident_id and not existing: return
        dialog=QDialog(self); dialog.setWindowTitle("Edit Accident / Incident Investigation" if existing else "New Accident / Incident Investigation"); dialog.resize(1200,850); dialog.setMinimumSize(1100,750)
        outer=QVBoxLayout(dialog); title=QLabel("ACCIDENT / INCIDENT INVESTIGATION"); title.setStyleSheet("font-size:20px;font-weight:bold;color:#17365D;padding:6px;"); outer.addWidget(title)
        scroll=QScrollArea(); scroll.setWidgetResizable(True); content=QWidget(); form=QFormLayout(content); form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow); edits={}
        def add_edit(name,label=None,value=""):
            e=QLineEdit(safe(value)); edits[name]=e; form.addRow(label or name.replace("_"," ").title()+":",e); return e
        add_edit("incident_title","Incident Title",existing["incident_title"] if existing and "incident_title" in existing.keys() else "")
        add_edit("permit_number","Permit Number",existing["permit_number"] if existing and "permit_number" in existing.keys() else "")
        for name in ["incident_time","location","project","company","department","activity"]:
            default_value = existing[name] if existing else (db.setting("project_name", "") if name == "project" else db.setting("company_name", "") if name == "company" else "")
            add_edit(name,value=default_value)
        for name in ["supervisor","witnesses"]: add_edit(name,value=existing[name] if existing else "")
        incident_type=QComboBox(); incident_type.addItems(INCIDENT_TYPES); form.addRow("Type of Incident:",incident_type)
        description=QTextEdit(); description.setMinimumHeight(100); description.setPlainText(safe(existing["description"]) if existing else ""); form.addRow("Incident Description:",description)
        immediate=QTextEdit(); immediate.setMinimumHeight(70); immediate.setPlainText(safe(existing["immediate_action"]) if existing else ""); form.addRow("Immediate Action:",immediate)
        # People involved
        people_box=QGroupBox("People Involved"); people_layout=QVBoxLayout(people_box); people_rows=[]
        def add_person(data=None):
            roww=QHBoxLayout(); n=QLineEdit(safe((data or {}).get("name",""))); role=QComboBox(); role.addItems(["Injured Person","Involved Person","Witness","Other"]); role.setCurrentText(safe((data or {}).get("role","Involved Person"))); d=QLineEdit(safe((data or {}).get("designation",""))); i=QLineEdit(safe((data or {}).get("id",""))); rem=QPushButton("Remove")
            n.setPlaceholderText("Person involved"); d.setPlaceholderText("Designation"); i.setPlaceholderText("ID"); roww.addWidget(n,2); roww.addWidget(role,2); roww.addWidget(d,2); roww.addWidget(i,1); roww.addWidget(rem); wrap=QWidget(); wrap.setLayout(roww); people_layout.addWidget(wrap); people_rows.append((wrap,n,role,d,i)); rem.clicked.connect(lambda: (people_rows.remove(next(x for x in people_rows if x[0] is wrap)), wrap.deleteLater()))
        addp=QPushButton("+ Add Person"); people_layout.addWidget(addp); addp.clicked.connect(lambda:add_person())
        if existing:
            try:
                for x in json.loads(safe(existing["people_details"]) or "[]"): add_person(x)
            except Exception: pass
        if not people_rows: add_person()
        form.addRow(people_box)
        # Equipment
        equipment_box=QGroupBox("Equipment"); equipment_layout=QVBoxLayout(equipment_box); equipment_rows=[]
        def add_equipment(data=None):
            roww=QHBoxLayout(); n=QLineEdit(safe((data or {}).get("name", ""))); i=QLineEdit(safe((data or {}).get("id", ""))); sp=QLineEdit(safe((data or {}).get("specification", ""))); rem=QPushButton("Remove")
            n.setPlaceholderText("Equipment"); i.setPlaceholderText("ID"); sp.setPlaceholderText("Specification"); roww.addWidget(n,2); roww.addWidget(i,1); roww.addWidget(sp,3); roww.addWidget(rem); wrap=QWidget(); wrap.setLayout(roww); equipment_layout.addWidget(wrap); equipment_rows.append((wrap,n,i,sp)); rem.clicked.connect(lambda: (equipment_rows.remove(next(x for x in equipment_rows if x[0] is wrap)), wrap.deleteLater()))
        adde=QPushButton("+ Add Equipment"); equipment_layout.addWidget(adde); adde.clicked.connect(lambda:add_equipment())
        if existing:
            try:
                for x in json.loads(safe(existing["equipment_details"]) or "[]"): add_equipment(x)
            except Exception: pass
        if not equipment_rows: add_equipment()
        form.addRow(equipment_box)
        # Repeatable incident detail sections
        ENV_CATEGORIES=["Spillage","Oil Spill","Chemical Spill","Fuel Spill","Diesel Spill","Hydrocarbon Spill","Water Discharge","Sewage Discharge","Air Emission","Dust Emission","Smoke / Fumes","Gas Release","Waste","Hazardous Waste","Non-Hazardous Waste","Soil Contamination","Water Contamination","Noise","Odour","Radiation","Resource Consumption","Energy Use","Water Use","Other"]
        def repeat_section(title, fields, old_key):
            box=QGroupBox(title); lay=QVBoxLayout(box); rows=[]
            def add(data=None):
                data=data or {}; h=QHBoxLayout(); widgets=[]
                for key,label in fields:
                    if key=="category":
                        w=QComboBox(); w.addItems(ENV_CATEGORIES); w.setCurrentText(safe(data.get(key,"")))
                    else:
                        w=QLineEdit(safe(data.get(key,""))); w.setPlaceholderText(label)
                    widgets.append(w); h.addWidget(w,3 if key in {"details","name"} else 2)
                rm=QPushButton("Remove"); h.addWidget(rm); wrap=QWidget(); wrap.setLayout(h); lay.addWidget(wrap); rows.append((wrap,widgets)); rm.clicked.connect(lambda: (rows.remove(next(x for x in rows if x[0] is wrap)),wrap.deleteLater()))
            plus=QPushButton("+ Add"); lay.addWidget(plus); plus.clicked.connect(lambda:add())
            vals=[]
            if existing:
                try: vals=json.loads(safe(existing["investigation_details"]) or "{}").get(old_key,[])
                except Exception: vals=[]
            for v in vals: add(v)
            if not rows: add()
            form.addRow(box); return rows
        environmental_rows=repeat_section("Environmental",[("category","Category (e.g. Spillage)"),("quantity","Quantity / Unit")],"environmental_items")
        property_rows=repeat_section("Property / Asset",[("name","Property / Asset Name"),("cost","Cost of Damage")],"property_items")
        procedure_rows=repeat_section("Procedure / Reference",[("reference","Procedure / Reference"),("details","Details")],"procedure_items")
        add_edit("organization","Organisation",existing["organization"] if existing else "")
        add_edit("report_reference","Report / Reference No.",existing["report_reference"] if existing else "")
        method=QComboBox(); method.addItems(INVESTIGATION_METHODS)
        if existing: method.setCurrentText(safe(existing["investigation_method"]))
        # Lock the selected investigation method so the method-specific content cannot
        # be changed accidentally.  Unlock is explicit and reversible.
        method_row=QHBoxLayout(); method_row.addWidget(method,1)
        method_lock=QPushButton("Lock Method"); method_lock.setMinimumHeight(34); method_lock.setMinimumWidth(125)
        method_state=QLabel("Unlocked")
        method_state.setStyleSheet("font-weight:bold;color:#8A6D1D;")
        method_row.addWidget(method_lock); method_row.addWidget(method_state)
        method_wrap=QWidget(); method_wrap.setLayout(method_row); form.addRow("Investigation Method:",method_wrap)
        method_locked=[False]
        def toggle_method_lock():
            method_locked[0]=not method_locked[0]
            method.setEnabled(not method_locked[0])
            if method_locked[0]:
                method_lock.setText("Unlock Method"); method_state.setText("Locked"); method_state.setStyleSheet("font-weight:bold;color:#B00020;")
            else:
                method_lock.setText("Lock Method"); method_state.setText("Unlocked"); method_state.setStyleSheet("font-weight:bold;color:#8A6D1D;")
        method_lock.clicked.connect(toggle_method_lock)
        procedure_box=QGroupBox("Investigation Procedure"); procedure_layout=QVBoxLayout(procedure_box); procedure_form=QFormLayout(); procedure_layout.addLayout(procedure_form); form.addRow(procedure_box); proc_widgets=[]
        def clear_proc():
            while procedure_form.count():
                item=procedure_form.takeAt(0); w=item.widget()
                if w:w.deleteLater()
            proc_widgets.clear()
        def add_proc(label,name,multi=True,value=""):
            w=QTextEdit() if multi else QLineEdit(); w.setPlainText(safe(value)) if multi else w.setText(safe(value));
            if multi:w.setMinimumHeight(65)
            proc_widgets.append((name,w)); procedure_form.addRow(label+":",w); return w
        def build_proc(_=None):
            clear_proc(); m=method.currentText(); old={}
            try: old=json.loads(safe(existing["investigation_details"]) or "{}") if existing else {}
            except Exception: old={}
            if m=="ICAM":
                icam_box=QGroupBox("ICAM Analysis"); icam_lay=QVBoxLayout(icam_box)
                icam_lay.addWidget(QLabel("Select applicable ICAM factors. Record what happened and the related event/condition. Control measures are generated automatically from the selected factors."))
                severity_box=QGroupBox("Severity")
                severity_lay=QVBoxLayout(severity_box)
                severity=QComboBox()
                severity_items=[
                    ("Low", "Minor consequence; first-aid level or low impact."),
                    ("Moderate", "Recordable / moderate consequence or significant operational impact."),
                    ("Major", "Serious injury/illness, major damage or significant business impact."),
                    ("Critical", "Life-threatening or very serious consequence; major loss potential."),
                    ("Catastrophic", "Fatality, permanent disabling outcome or catastrophic loss potential."),
                    ("Not Yet Determined", "Use when available evidence is insufficient to determine severity."),
                ]
                severity.addItems([x[0] for x in severity_items]); severity.setCurrentText(safe(old.get("icam_severity","Not Yet Determined")) or "Not Yet Determined")
                severity_lay.addWidget(severity)
                severity_guide=QLabel(); severity_guide.setWordWrap(True)
                severity_lay.addWidget(severity_guide)
                def update_severity_guide(_=None):
                    txt=dict(severity_items).get(severity.currentText(),""); severity_guide.setText("Guidance: "+txt)
                severity.currentTextChanged.connect(update_severity_guide); update_severity_guide()
                icam_lay.addWidget(severity_box); proc_widgets.append(("icam_severity",severity))
                def icam_txt(label,key,height=60):
                    w=QTextEdit(); w.setMinimumHeight(height); w.setPlainText(safe(old.get(key,""))); icam_lay.addWidget(QLabel(label)); icam_lay.addWidget(w); proc_widgets.append(("icam_text_"+key,w)); return w

                control_rebuild_callback=[None]
                def icam_repeat(title,key,fields, factor_options=None):
                    box=QGroupBox(title); lay=QVBoxLayout(box); rows=[]
                    def add(data=None):
                        data=data or {}; h=QHBoxLayout(); widgets=[]
                        for fld,label in fields:
                            if fld=="factor" and factor_options:
                                w=QComboBox(); w.addItems(factor_options); w.setCurrentText(safe(data.get(fld,"")) or factor_options[0])
                            else:
                                w=QLineEdit(safe(data.get(fld,""))); w.setPlaceholderText(label)
                            widgets.append(w); h.addWidget(w,3 if fld in {"what_happened","event_condition","description"} else 2)
                        rm=QPushButton("Remove"); h.addWidget(rm); wrap=QWidget(); wrap.setLayout(h); lay.addWidget(wrap); rows.append((wrap,)+tuple(widgets))
                        rm.clicked.connect(lambda:remove(wrap))
                        if key in {"icam_a","icam_ita","icam_tec","icam_of"} and control_rebuild_callback[0]:
                            control_rebuild_callback[0]()
                    def remove(wrap):
                        item=next((x for x in rows if x[0] is wrap),None)
                        if item:
                            rows.remove(item); wrap.deleteLater()
                            if key in {"icam_a","icam_ita","icam_tec","icam_of"} and control_rebuild_callback[0]:
                                control_rebuild_callback[0]()
                    plus=QPushButton("+ Add"); lay.addWidget(plus); plus.clicked.connect(lambda:add())
                    for x in (old.get(key,[]) or []): add(x)
                    if not rows: add()
                    icam_lay.addWidget(box); proc_widgets.append(("icam_repeat_"+key,rows)); return rows

                team_rows=icam_repeat("Investigation Team","icam_team",[("name","Name"),("position","Position"),("department","Department"),("role","Investigation Role")])
                timeline_rows=icam_repeat("EVENT TIMELINE - Add multiple events", "icam_timeline_events",[("date","Date"),("time","Time"),("event","Event / Description")])
                factor_sets={
                    "icam_a":["Physical Barriers","Engineering Controls","Isolation / LOTO","Guards","Alarms","Warning Systems","PPE","Procedures","Permits","Supervision","Inspections","Monitoring","Emergency Controls","Other"],
                    "icam_ita":["Error / Mistake","Slip / Lapse","Rule / Procedure Deviation","Decision-Making","Communication","Situational Awareness","Competency","Teamwork","Coordination","Supervision","Fatigue","Distraction","Workload","Other"],
                    "icam_tec":["Task Design","Task Complexity","Equipment Design","Equipment Condition","Workplace Layout","Access / Egress","Tools","Workload","Time Pressure","Staffing","Weather","Temperature","Lighting","Noise","Housekeeping","Ergonomics","Fatigue","Distraction","Other"],
                    "icam_of":["Leadership","Management Systems","Safety Culture","Planning","Resource Allocation","Risk Management","Procedures","Training Systems","Competency Management","Supervision","Maintenance Management","Contractor Management","Communication","Change Management / MOC","Procurement","Design","Inspection","Audit","Performance Monitoring","Lessons Learned","Previous Incidents / Findings","Other"]
                }
                a_rows=icam_repeat("A. DEFENCE / CONTROL FACTORS","icam_a",[("factor","Factor"),("what_happened","What Happened"),("event_condition","Event / Condition")],factor_sets["icam_a"])
                ita_rows=icam_repeat("B. INDIVIDUAL / TEAM ACTIONS (ITA)","icam_ita",[("factor","Factor"),("what_happened","What Happened"),("event_condition","Event / Condition")],factor_sets["icam_ita"])
                tec_rows=icam_repeat("C. TASK / ENVIRONMENTAL CONDITIONS (TEC)","icam_tec",[("factor","Factor"),("what_happened","What Happened"),("event_condition","Event / Condition")],factor_sets["icam_tec"])
                of_rows=icam_repeat("D. ORGANISATIONAL FACTORS (OF)","icam_of",[("factor","Factor"),("what_happened","What Happened"),("event_condition","Event / Condition")],factor_sets["icam_of"])
                # Control measures are always a one-to-one match with the currently
                # displayed A/B/C/D attribution-factor rows. The selected factor is
                # fetched directly from the attribution row; the user only enters the
                # control description, responsible person, status and target date.
                control_box=QGroupBox("CONTROL MEASURES - Matched to ICAM Attribution Factors")
                control_lay=QVBoxLayout(control_box); control_rows=[]
                control_header=QLabel("Factor | Control Measure Description | Responsible Person | Open / Close | Target Date")
                control_header.setStyleSheet("font-weight:bold;color:#17365D;")
                control_lay.addWidget(control_header)
                def selected_factor_records():
                    out=[]
                    for key,rows in [("icam_a",a_rows),("icam_ita",ita_rows),("icam_tec",tec_rows),("icam_of",of_rows)]:
                        for item in rows:
                            factor=item[1].currentText().strip() if isinstance(item[1],QComboBox) else item[1].text().strip()
                            if factor: out.append((key,item,factor))
                    return out
                def rebuild_controls():
                    # Preserve values by the actual attribution-row object so adding
                    # another factor never steals or duplicates another row's data.
                    current_values={}
                    for item in control_rows:
                        _,source_item,_,cm,rp,status,target=item
                        current_values[id(source_item)]={
                            "control_measure":cm.text().strip(),
                            "responsible":rp.text().strip(),
                            "status":status.currentText(),
                            "target_date":target.text().strip(),
                        }
                    saved_controls=old.get("icam_control_measures",[]) or []
                    saved_by_factor={}
                    for x in saved_controls:
                        if isinstance(x,dict):
                            saved_by_factor.setdefault(safe(x.get("factor")),[]).append(x)
                    for item in list(control_rows): item[0].deleteLater()
                    control_rows.clear()
                    saved_occurrence={}
                    for _,source_item,factor in selected_factor_records():
                        data=current_values.get(id(source_item))
                        if data is None:
                            idx=saved_occurrence.get(factor,0)
                            matches=saved_by_factor.get(factor,[])
                            data=matches[idx] if idx < len(matches) else {}
                            saved_occurrence[factor]=idx+1
                        h=QHBoxLayout()
                        fl=QLabel(factor); fl.setMinimumWidth(190); fl.setToolTip("Fetched directly from the selected ICAM attribution factor")
                        cm=QLineEdit(safe(data.get("control_measure",""))); cm.setPlaceholderText("Control Measure Description")
                        rp=QLineEdit(safe(data.get("responsible",""))); rp.setPlaceholderText("Responsible Person")
                        status=QComboBox(); status.addItems(["Open","Closed"]); status.setCurrentText(safe(data.get("status",data.get("close_out","Open"))) or "Open")
                        target=QLineEdit(safe(data.get("target_date",data.get("close_out","")) if data.get("target_date","") or data.get("close_out","") else "")); target.setPlaceholderText("Target Date")
                        h.addWidget(fl,2); h.addWidget(cm,3); h.addWidget(rp,2); h.addWidget(status,1); h.addWidget(target,2)
                        wrap=QWidget(); wrap.setLayout(h); control_lay.addWidget(wrap)
                        control_rows.append((wrap,source_item,factor,cm,rp,status,target))
                icam_lay.addWidget(control_box)
                control_rebuild_callback[0]=rebuild_controls
                for rows in (a_rows,ita_rows,tec_rows,of_rows):
                    for item in rows:
                        if isinstance(item[1],QComboBox): item[1].currentTextChanged.connect(lambda _=None: rebuild_controls())
                rebuild_controls()
                proc_widgets.extend([("icam_control_rows",control_rows)])
                icam_txt("Recommendations","icam_recommendations",70)
                procedure_form.addRow(icam_box)
            elif m=="5 Why Analysis":
                for i in range(1,6): add_proc(f"Why {i}",f"why{i}",False,old.get(f"why{i}",""))
            elif m=="Fishbone / Ishikawa":
                # Structured Fishbone / Ishikawa RCA workspace. The category lists and
                # guidance below are based on the supplied Fishbone RCA structure.
                fish_box=QGroupBox("Fishbone / Ishikawa Analysis")
                fish_lay=QVBoxLayout(fish_box)
                fish_lay.addWidget(QLabel("Identify possible causes under each factor. Select a suggested cause or choose Other and add a specific cause."))
                fish_rows=[]
                fish_options={
                    "People": ["Lack of training","Expired / not evident training","Lack of experience","Fatigue","Human error","Poor communication","Inadequate supervision","Lack of competency","Other"],
                    "Machine / Equipment": ["Equipment failure","Poor maintenance","Incorrect settings","Defective component","Lack of guarding","Calibration problem","Other"],
                    "Method": ["Incorrect procedure","SOP not available","SOP not followed","Poor work planning","Inadequate risk assessment","Incorrect sequence of work","Other"],
                    "Material": ["Wrong material","Defective material","Poor quality","Incorrect specification","Contamination","Improper storage","Other"],
                    "Measurement": ["Incorrect measurement","Uncalibrated instrument","Wrong inspection method","Incorrect data","Inadequate monitoring","Other"],
                    "Environment": ["Temperature","Humidity","Lighting","Noise","Dust","Poor housekeeping","Congestion","Weather conditions","Other"],
                }
                saved_fish=old.get("fishbone",{}) if isinstance(old.get("fishbone",{}),dict) else {}
                for cat in fish_options:
                    for item in (saved_fish.get(cat,[]) or []):
                        pass
                fish_headers=QLabel("Factor | Suggested Cause | Specific Cause / Explanation | Evidence / Finding | Attachment")
                fish_headers.setStyleSheet("font-weight:bold;color:#17365D;")
                fish_lay.addWidget(fish_headers)
                fish_rows_box=QVBoxLayout()
                fish_lay.addLayout(fish_rows_box)
                # Keep every newly-added Fishbone cause inside this dedicated container.
                # This prevents rows from being inserted into the following Fishbone sections.
                def add_fish_row(category="People", cause="", detail="", evidence="", attachments=None):
                    attachments=list(attachments or [])
                    h=QHBoxLayout(); cat=QComboBox(); cat.addItems(list(fish_options.keys())+["Other"]); cat.setCurrentText(category if category in fish_options else "Other")
                    cause_cb=QComboBox(); cause_cb.setEditable(True); cause_cb.addItems(fish_options.get(cat.currentText(),["Other"])); cause_cb.setCurrentText(cause)
                    detail_e=QLineEdit(detail); detail_e.setPlaceholderText("Specific cause / explanation")
                    evidence_e=QLineEdit(evidence); evidence_e.setPlaceholderText("Evidence / finding")
                    attach_files=[]; attach_files.extend(attachments)
                    attach_lab=QLabel("; ".join(Path(x).name for x in attach_files) if attach_files else "No attachment"); attach_lab.setWordWrap(True)
                    attach=QPushButton("Add Attachment")
                    rm=QPushButton("Remove")
                    h.addWidget(cat,2); h.addWidget(cause_cb,3); h.addWidget(detail_e,3); h.addWidget(evidence_e,3); h.addWidget(attach,2); h.addWidget(rm,1)
                    wrap=QWidget(); wrap.setLayout(h); fish_rows_box.addWidget(wrap); fish_rows.append((wrap,cat,cause_cb,detail_e,evidence_e,attach_files,attach_lab))
                    # Put the filename label immediately below this cause row.
                    fish_rows_box.addWidget(attach_lab)
                    attach_lab.setStyleSheet("color:#555;font-size:11px;padding-left:8px;")
                    def pick_fish_attachment():
                        paths,_=QFileDialog.getOpenFileNames(dialog,"Select Fishbone Cause Evidence","","Evidence Files (*)")
                        if paths:
                            attach_files.clear(); attach_files.extend(paths); attach_lab.setText("; ".join(Path(x).name for x in attach_files))
                    attach.clicked.connect(pick_fish_attachment)
                    def refresh_causes(_=None):
                        current=cause_cb.currentText(); cause_cb.blockSignals(True); cause_cb.clear(); cause_cb.addItems(fish_options.get(cat.currentText(),["Other"])); cause_cb.setCurrentText(current); cause_cb.blockSignals(False)
                    cat.currentTextChanged.connect(refresh_causes)
                    rm.clicked.connect(lambda: remove_fish_row(wrap,attach_lab))
                def remove_fish_row(wrap,attach_lab=None):
                    item=next((x for x in fish_rows if x[0] is wrap),None)
                    if item:
                        fish_rows.remove(item); wrap.deleteLater()
                        if attach_lab: attach_lab.deleteLater()
                saved_rows=[]
                for cat,items in saved_fish.items():
                    for x in items or []:
                        if isinstance(x,dict): saved_rows.append((cat,x.get("cause",""),x.get("detail",""),x.get("evidence",""),x.get("attachments",[]) or []))
                if saved_rows:
                    for x in saved_rows: add_fish_row(*x)
                else:
                    add_fish_row()
                add_fish=QPushButton("+ Add Fishbone Cause")
                add_fish.setMinimumHeight(36)
                fish_lay.addWidget(add_fish)
                # QAbstractButton.clicked emits a boolean; use a no-argument lambda so
                # that the signal value cannot be mistaken for the cause category.
                add_fish.clicked.connect(lambda checked=False: add_fish_row())
                fish_lay.addWidget(QLabel("Fishbone factors: People, Machine / Equipment, Method, Material, Measurement and Environment. Use Add for additional applicable causes."))
                # Supporting Fishbone RCA sections from the supplied RCA flow. These are
                # only shown when Fishbone is selected, so other investigation methods keep
                # their existing UI unchanged.
                fish_problem=QTextEdit(); fish_problem.setMinimumHeight(70); fish_problem.setPlainText(safe(old.get("problem_statement",""))); fish_lay.addWidget(QLabel("Problem Statement (what went wrong; do not put the suspected root cause here)")); fish_lay.addWidget(fish_problem)
                team_box=QGroupBox("Investigation Team"); team_lay=QVBoxLayout(team_box); team_rows=[]
                def add_team(data=None):
                    data=data or {}; h=QHBoxLayout(); n=QLineEdit(safe(data.get("name",""))); pos=QLineEdit(safe(data.get("position",""))); dep=QLineEdit(safe(data.get("department",""))); role=QLineEdit(safe(data.get("role",""))); rm=QPushButton("Remove")
                    n.setPlaceholderText("Name"); pos.setPlaceholderText("Position"); dep.setPlaceholderText("Department"); role.setPlaceholderText("Role in Investigation")
                    h.addWidget(n,2); h.addWidget(pos,2); h.addWidget(dep,2); h.addWidget(role,2); h.addWidget(rm); w=QWidget(); w.setLayout(h); team_lay.addWidget(w); team_rows.append((w,n,pos,dep,role)); rm.clicked.connect(lambda:remove_team(w))
                def remove_team(w):
                    item=next((x for x in team_rows if x[0] is w),None)
                    if item: team_rows.remove(item); w.deleteLater()
                team_plus=QPushButton("+ Add Investigation Team Member"); team_lay.addWidget(team_plus); team_plus.clicked.connect(lambda:add_team())
                for x in (old.get("investigation_team",[]) or []): add_team(x)
                if not team_rows: add_team()
                fish_lay.addWidget(team_box)
                timeline_box=QGroupBox("Timeline of Events"); timeline_lay=QVBoxLayout(timeline_box); timeline_rows=[]
                def add_timeline(data=None):
                    data=data or {}; h=QHBoxLayout(); tm=QLineEdit(safe(data.get("time",""))); ev=QLineEdit(safe(data.get("event",""))); rm=QPushButton("Remove"); tm.setPlaceholderText("Time"); ev.setPlaceholderText("Event / What happened")
                    h.addWidget(tm,1); h.addWidget(ev,5); h.addWidget(rm); w=QWidget(); w.setLayout(h); timeline_lay.addWidget(w); timeline_rows.append((w,tm,ev)); rm.clicked.connect(lambda:remove_timeline(w))
                def remove_timeline(w):
                    item=next((x for x in timeline_rows if x[0] is w),None)
                    if item: timeline_rows.remove(item); w.deleteLater()
                tl_plus=QPushButton("+ Add Timeline Event"); timeline_lay.addWidget(tl_plus); tl_plus.clicked.connect(lambda:add_timeline())
                for x in (old.get("timeline_events",[]) or []): add_timeline(x)
                if not timeline_rows: add_timeline()
                fish_lay.addWidget(timeline_box)
                verify_box=QGroupBox("Root Cause Verification"); verify_lay=QVBoxLayout(verify_box); verify_rows=[]
                def add_verify(data=None):
                    data=data or {}; h=QHBoxLayout(); cause=QLineEdit(safe(data.get("cause",""))); evidence=QLineEdit(safe(data.get("evidence",""))); verified=QComboBox(); verified.addItems(["Yes","No","Pending"]); verified.setCurrentText(safe(data.get("verified","Pending")) or "Pending"); person=QLineEdit(safe(data.get("responsible",""))); rm=QPushButton("Remove")
                    cause.setPlaceholderText("Potential Cause"); evidence.setPlaceholderText("Evidence / Investigation Finding"); person.setPlaceholderText("Responsible for Verification")
                    h.addWidget(cause,2); h.addWidget(evidence,3); h.addWidget(verified,1); h.addWidget(person,2); h.addWidget(rm); w=QWidget(); w.setLayout(h); verify_lay.addWidget(w); verify_rows.append((w,cause,evidence,verified,person)); rm.clicked.connect(lambda:remove_verify(w))
                def remove_verify(w):
                    item=next((x for x in verify_rows if x[0] is w),None)
                    if item: verify_rows.remove(item); w.deleteLater()
                vr_plus=QPushButton("+ Add Root Cause Verification"); verify_lay.addWidget(vr_plus); vr_plus.clicked.connect(lambda:add_verify())
                for x in (old.get("root_cause_verification",[]) or []): add_verify(x)
                if not verify_rows: add_verify()
                fish_lay.addWidget(verify_box)
                fish_effective=QTextEdit(); fish_effective.setMinimumHeight(60); fish_effective.setPlainText(safe(old.get("effectiveness_verification",""))); fish_lay.addWidget(QLabel("Effectiveness Verification")); fish_lay.addWidget(fish_effective)
                fish_lessons=QTextEdit(); fish_lessons.setMinimumHeight(60); fish_lessons.setPlainText(safe(old.get("lessons_learned",""))); fish_lay.addWidget(QLabel("Lessons Learned")); fish_lay.addWidget(fish_lessons)
                fish_conclusion=QTextEdit(); fish_conclusion.setMinimumHeight(60); fish_conclusion.setPlainText(safe(old.get("conclusion",""))); fish_lay.addWidget(QLabel("Conclusion")); fish_lay.addWidget(fish_conclusion)
                proc_widgets.extend([
                    ("fish_problem_statement", fish_problem),
                    ("fish_investigation_team", team_rows),
                    ("fish_timeline_events", timeline_rows),
                    ("fish_root_cause_verification", verify_rows),
                    ("fish_effectiveness_verification", fish_effective),
                    ("fish_lessons_learned", fish_lessons),
                    ("fish_conclusion", fish_conclusion),
                ])
                # IMPORTANT: the Fishbone workspace belongs to the investigation
                # procedure container.  Putting it directly on the outer form caused
                # old Fishbone panels to remain when another method was selected.
                procedure_form.addRow(fish_box)
                # Preserve any legacy Fishbone text when an older record is edited.
                if not saved_rows:
                    legacy_map={"People":"people","Machine / Equipment":"machine","Method":"method","Material":"material","Environment":"environment","Measurement":"management"}
                    for cat,key in legacy_map.items():
                        legacy=old.get(key,"")
                        if legacy: add_fish_row(cat,"Other",legacy,"")
                proc_widgets.append(("fishbone_rows", fish_rows))
            elif m=="Barrier Analysis":
                for lab,key in [("Hazard / Threat","hazard"),("Top Event","top_event"),("Required Barriers","required_barriers"),("Failed / Missing Barriers","failed_barriers"),("Recovery / Mitigation","mitigation")]: add_proc(lab,key,value=old.get(key,""))
            elif m=="Bow-Tie Analysis":
                for lab,key in [("Threats","threats"),("Top Event","top_event"),("Preventive Barriers","preventive_barriers"),("Consequences","consequences"),("Mitigative Barriers","mitigative_barriers")]: add_proc(lab,key,value=old.get(key,""))
            elif m=="Fault Tree Analysis":
                for lab,key in [("Top Event","top_event"),("Intermediate Events","intermediate_events"),("Basic Events","basic_events"),("Logic / Gate Analysis","logic_analysis")]: add_proc(lab,key,value=old.get(key,""))
            elif m=="Causal Tree":
                for lab,key in [("Event","event"),("Causal Sequence","causal_sequence"),("Direct Causes","direct_causes_method"),("Contributing Causes","contributing_causes"),("Root Causes","root_causes_method")]: add_proc(lab,key,value=old.get(key,""))
            else: add_proc("Investigation Procedure / Notes","summary",value=old.get("summary",""))
        method.currentTextChanged.connect(build_proc); build_proc()
        # Multiple RCA fields
        def multi_cause_group(title,key):
            box=QGroupBox(title); lay=QVBoxLayout(box); rows=[]
            def add(value=""):
                h=QHBoxLayout(); e=QLineEdit(safe(value)); e.setPlaceholderText(title); rm=QPushButton("Remove"); h.addWidget(e,1); h.addWidget(rm); wrap=QWidget(); wrap.setLayout(h); lay.addWidget(wrap); rows.append((wrap,e)); rm.clicked.connect(lambda: (rows.remove(next(x for x in rows if x[0] is wrap)),wrap.deleteLater()))
            plus=QPushButton("+ Add"); lay.addWidget(plus); plus.clicked.connect(lambda:add());
            vals=[]
            if existing:
                try: vals=json.loads(safe(existing["investigation_details"]) or "{}").get(key,[])
                except Exception: vals=[]
            if isinstance(vals,str): vals=[vals]
            for v in vals: add(v)
            if not rows: add()
            form.addRow(box); return rows
        direct_rows=multi_cause_group("Direct Cause", "direct_causes")
        contrib_rows=multi_cause_group("Contributing Factors", "contributing_factors_list")
        root_rows=multi_cause_group("Root Cause", "root_causes_list")
        # Corrective action plan
        ca_box=QGroupBox("Corrective Action Plan"); ca_lay=QVBoxLayout(ca_box); ca_rows=[]
        def add_ca(data=None):
            data=data or {}; h=QHBoxLayout(); action=QLineEdit(safe(data.get("action",""))); resp=QLineEdit(safe(data.get("responsible",""))); target=QLineEdit(safe(data.get("target_date",""))); status=QComboBox(); status.addItems(["Open","Closed"]); status.setCurrentText(safe(data.get("status","Open")) or "Open"); attach=[]; lab=QLabel("No closeout evidence"); attach_btn=QPushButton("Attach"); rem=QPushButton("Remove")
            action.setPlaceholderText("Action"); resp.setPlaceholderText("Responsible Person"); target.setPlaceholderText("Target Date");
            h.addWidget(action,3); h.addWidget(resp,2); h.addWidget(target,1); h.addWidget(status,1); h.addWidget(attach_btn); h.addWidget(rem); wrap=QWidget(); wrap.setLayout(h); ca_lay.addWidget(wrap); h2=QHBoxLayout(); h2.addWidget(lab); wrap2=QWidget(); wrap2.setLayout(h2); ca_lay.addWidget(wrap2); ca_rows.append((wrap,action,resp,target,status,attach,lab))
            attach_btn.clicked.connect(lambda: choose_ca_files(attach,lab))
            rem.clicked.connect(lambda: remove_ca(wrap,wrap2))
        def choose_ca_files(store,lab):
            paths,_=QFileDialog.getOpenFileNames(dialog,"Select Closeout Evidence","","Evidence Files (*)")
            if paths: store.clear(); store.extend(paths); lab.setText("; ".join(Path(x).name for x in paths))
        def remove_ca(wrap,wrap2):
            item=next((x for x in ca_rows if x[0] is wrap),None)
            if item: ca_rows.remove(item); wrap.deleteLater(); wrap2.deleteLater()
        plus_ca=QPushButton("+ Add Corrective Action"); ca_lay.addWidget(plus_ca); plus_ca.clicked.connect(lambda:add_ca())
        old_actions=[]
        if existing:
            try: old_actions=json.loads(safe(existing["investigation_details"]) or "{}").get("corrective_actions",[])
            except Exception: old_actions=[]
        for a in old_actions: add_ca(a)
        if not ca_rows: add_ca()
        form.addRow(ca_box)
        incident_capa_required = QComboBox()
        incident_capa_required.addItems(["No", "Yes"])
        if existing and safe(existing["capa_number"] if "capa_number" in existing.keys() else ""):
            incident_capa_required.setCurrentText("Yes")
        incident_capa_required.setToolTip("When Yes, link the incident corrective action plan to the CAPA Tracking Register.")
        form.addRow("Link Corrective Action Plan to CAPA Tracking?:", incident_capa_required)
        capa_link_hint = QLabel("Yes creates or updates a CAPA Tracking Register record linked to this incident report and its corrective actions.")
        capa_link_hint.setWordWrap(True)
        capa_link_hint.setStyleSheet("color:#245486;font-size:11px;")
        form.addRow("", capa_link_hint)
        # Witness statements: each statement can have its own attachment and remark.
        witness_box=QGroupBox("Witness Statements"); witness_lay=QVBoxLayout(witness_box); witness_rows=[]
        def add_witness(data=None):
            data=data or {}; n=QLineEdit(safe(data.get("name",""))); st=QTextEdit(); st.setFixedHeight(65); st.setPlainText(safe(data.get("statement",""))); rm=QLineEdit(safe(data.get("remark",""))); rm.setPlaceholderText("Remark"); files=[]; lab=QLabel(safe(data.get("attachment","")) or "No attachment"); attach=QPushButton("Attach"); remove=QPushButton("Remove")
            h=QHBoxLayout(); h.addWidget(n,2); h.addWidget(attach); h.addWidget(remove); wrap=QWidget(); wrap.setLayout(h); witness_lay.addWidget(wrap); witness_lay.addWidget(st); witness_lay.addWidget(rm); witness_lay.addWidget(lab); witness_rows.append((wrap,st,n,rm,files,lab))
            attach.clicked.connect(lambda: choose_witness(files,lab)); remove.clicked.connect(lambda: remove_witness(wrap,st,rm,lab))
        def choose_witness(store,lab):
            paths,_=QFileDialog.getOpenFileNames(dialog,"Select Witness Statement Attachment","","Files (*)")
            if paths: store.clear(); store.extend(paths); lab.setText("; ".join(Path(x).name for x in paths))
        def remove_witness(*widgets):
            for item in list(witness_rows):
                if item[0] is widgets[0]: witness_rows.remove(item); [x.deleteLater() for x in widgets]
        plus_w=QPushButton("+ Add Witness Statement"); witness_lay.addWidget(plus_w); plus_w.clicked.connect(lambda:add_witness())
        if existing:
            try:
                for x in json.loads(safe(existing["investigation_details"]) or "{}").get("witness_statements",[]): add_witness(x)
            except Exception: pass
        if not witness_rows: add_witness()
        form.addRow(witness_box)
        approval_box=QGroupBox("Prepared / Reviewed / Approved By"); approval_form=QFormLayout(approval_box)
        approval_rows={}
        approval_old={}
        try: approval_old=json.loads(safe(existing["investigation_details"]) or "{}").get("approval_signatures",{}) if existing else {}
        except Exception: approval_old={}
        for role in ["Prepared by","Reviewed by","Approved by"]:
            key=role.lower().replace(" ","_"); data=approval_old.get(key,{}) if isinstance(approval_old,dict) else {}
            n=QLineEdit(safe(data.get("name",""))); pos=QLineEdit(safe(data.get("position",""))); sig=QLineEdit(safe(data.get("signature",""))); dt=QLineEdit(safe(data.get("date","")))
            approval_form.addRow(role+" - Name:",n); approval_form.addRow(role+" - Position:",pos); approval_form.addRow(role+" - Signature:",sig); approval_form.addRow(role+" - Date:",dt)
            approval_rows[key]=(n,pos,sig,dt)
        form.addRow(approval_box)
        status=QComboBox(); status.addItems(["Draft","Open","Closed"]); status.setCurrentText(safe(existing["status"]) if existing else "Draft"); form.addRow("Investigation Status:",status)
        evidence_box=QGroupBox("Evidence Attachments"); evidence_lay=QVBoxLayout(evidence_box); evidence_rows=[]
        def add_evidence():
            h=QHBoxLayout(); file_store=[]; lab=QLabel("No file selected"); remark=QLineEdit(); remark.setPlaceholderText("Remark for this attachment"); choose=QPushButton("Add Attachment"); remove=QPushButton("Remove")
            h.addWidget(lab,2); h.addWidget(remark,3); h.addWidget(choose); h.addWidget(remove); wrap=QWidget(); wrap.setLayout(h); evidence_lay.addWidget(wrap); evidence_rows.append((wrap,file_store,remark,lab))
            def pick():
                paths,_=QFileDialog.getOpenFileNames(dialog,"Select Investigation Evidence","","Evidence Files (*)")
                if paths: file_store.clear(); file_store.extend(paths); lab.setText("; ".join(Path(x).name for x in paths))
            choose.clicked.connect(pick); remove.clicked.connect(lambda: remove_evidence(wrap))
        def remove_evidence(wrap):
            item=next((x for x in evidence_rows if x[0] is wrap),None)
            if item: evidence_rows.remove(item); wrap.deleteLater()
        plus_ev=QPushButton("+ Add Evidence Attachment"); evidence_lay.addWidget(plus_ev); plus_ev.clicked.connect(add_evidence); add_evidence(); form.addRow(evidence_box)
        # Keep action buttons permanently visible without needing to drag the form.
        action_bar=QHBoxLayout(); action_bar.addStretch()
        buttons=QDialogButtonBox(); save_draft=buttons.addButton("Save for Later",QDialogButtonBox.ButtonRole.AcceptRole); save_final=buttons.addButton("Save and Submit",QDialogButtonBox.ButtonRole.AcceptRole); save_exit=buttons.addButton("Save and Exit",QDialogButtonBox.ButtonRole.AcceptRole); cancel=buttons.addButton("Cancel",QDialogButtonBox.ButtonRole.RejectRole)
        for b in [save_draft,save_final,save_exit,cancel]: b.setMinimumHeight(38); b.setMinimumWidth(125)
        action_bar.addWidget(buttons); outer.insertLayout(1,action_bar); cancel.clicked.connect(dialog.reject)
        scroll.setWidget(content); outer.addWidget(scroll,1)
        def collect_rows(rows): return [e.text().strip() for _,e in rows if e.text().strip()]
        def save(mode):
            if not edits["location"].text().strip(): QMessageBox.warning(dialog,"Required","Location is required."); return
            try:
                number=safe(existing["number"]) if existing else next_number("HSE-INC","incidents")
                proc={}
                fishbone_rows_data=[]
                for name,w in proc_widgets:
                    if name=="fishbone_rows":
                        for _,cat,cause_cb,detail_e,evidence_e,attach_files,attach_lab in w:
                            if cat.currentText().strip() or cause_cb.currentText().strip() or detail_e.text().strip() or evidence_e.text().strip() or attach_files:
                                fishbone_rows_data.append({"category":cat.currentText().strip(),"cause":cause_cb.currentText().strip(),"detail":detail_e.text().strip(),"evidence":evidence_e.text().strip(),"attachments":list(attach_files)})
                    elif name=="fish_problem_statement":
                        proc["problem_statement"]=w.toPlainText()
                    elif name=="fish_investigation_team":
                        proc["investigation_team"]=[{"name":n.text().strip(),"position":pos.text().strip(),"department":dep.text().strip(),"role":role.text().strip()} for _,n,pos,dep,role in w if any(z.text().strip() for z in (n,pos,dep,role))]
                    elif name=="fish_timeline_events":
                        proc["timeline_events"]=[{"time":tm.text().strip(),"event":ev.text().strip()} for _,tm,ev in w if tm.text().strip() or ev.text().strip()]
                    elif name=="fish_root_cause_verification":
                        proc["root_cause_verification"]=[{"cause":cause.text().strip(),"evidence":evidence.text().strip(),"verified":verified.currentText(),"responsible":person.text().strip()} for _,cause,evidence,verified,person in w if any(z.text().strip() for z in (cause,evidence,person))]
                    elif name=="fish_effectiveness_verification":
                        proc["effectiveness_verification"]=w.toPlainText()
                    elif name=="fish_lessons_learned":
                        proc["lessons_learned"]=w.toPlainText()
                    elif name=="fish_conclusion":
                        proc["conclusion"]=w.toPlainText()
                    elif name=="icam_severity":
                        proc["icam_severity"]=w.currentText()
                    elif name=="icam_control_rows":
                        proc["icam_control_measures"]=[
                            {"factor":factor,"control_measure":cm.text().strip(),"responsible":rp.text().strip(),
                             "status":status.currentText(),"target_date":target.text().strip()}
                            for _,source_item,factor,cm,rp,status,target in w
                            if factor or cm.text().strip() or rp.text().strip() or target.text().strip()
                        ]
                    elif name.startswith("icam_text_"):
                        proc[name.replace("icam_text_","")]=w.toPlainText()
                    elif name.startswith("icam_repeat_"):
                        key=name.replace("icam_repeat_","")
                        rows_out=[]
                        for item in w:
                            vals=item[1:]
                            obj={}
                            for fld,widget in zip({
                                "icam_team":["name","position","department","role"],
                                "icam_timeline_events":["date","time","event"],
                                "icam_a":["factor","what_happened","event_condition"],
                                "icam_ita":["factor","what_happened","event_condition"],
                                "icam_tec":["factor","what_happened","event_condition"],
                                "icam_of":["factor","what_happened","event_condition"]
                            }.get("icam_repeat_"+key,[]), vals):
                                obj[fld]=(widget.currentText().strip() if isinstance(widget,QComboBox) else widget.text().strip())
                            if any(obj.values()): rows_out.append(obj)
                        proc[key]=rows_out
                    else:
                        proc[name]=(w.toPlainText() if isinstance(w,QTextEdit) else w.text())
                direct_vals=collect_rows(direct_rows); contrib_vals=collect_rows(contrib_rows); root_vals=collect_rows(root_rows)
                actions=[]
                for _,a,r,t,st,att,lab in ca_rows:
                    if a.text().strip() or r.text().strip() or t.text().strip() or att:
                        actions.append({"action":a.text().strip(),"responsible":r.text().strip(),"target_date":t.text().strip(),"status":st.currentText(),"attachments":[]})
                if incident_capa_required.currentText() == "Yes" and not any(safe(a.get("action", "")).strip() for a in actions):
                    QMessageBox.warning(dialog, "Corrective Action Required", "Add at least one corrective action before linking this incident to CAPA Tracking.")
                    return
                details=proc; details.update({"direct_causes":direct_vals,"contributing_factors_list":contrib_vals,"root_causes_list":root_vals,"corrective_actions":actions})
                details["approval_signatures"]={key:{"name":vals[0].text().strip(),"position":vals[1].text().strip(),"signature":vals[2].text().strip(),"date":vals[3].text().strip()} for key,vals in approval_rows.items()}
                if method.currentText()=="Fishbone / Ishikawa":
                    details["fishbone"]=fishbone_rows_data
                details["environmental_items"]=[{"category":widgets[0].currentText().strip(),"quantity":widgets[1].text().strip()} for _,widgets in environmental_rows if widgets[0].currentText().strip() or widgets[1].text().strip()]
                details["property_items"]=[{"name":widgets[0].text().strip(),"cost":widgets[1].text().strip()} for _,widgets in property_rows if any(w.text().strip() for w in widgets)]
                details["procedure_items"]=[{"reference":widgets[0].text().strip(),"details":widgets[1].text().strip()} for _,widgets in procedure_rows if any(w.text().strip() for w in widgets)]
                details["witness_statements"]=[]
                for _,st,n,rm,files,lab in witness_rows:
                    if n.text().strip() or st.toPlainText().strip() or files:
                        details["witness_statements"].append({"name":n.text().strip(),"statement":st.toPlainText().strip(),"remark":rm.text().strip(),"attachment":""})
                people=[{"name":n.text().strip(),"role":role.currentText(),"designation":d.text().strip(),"id":i.text().strip()} for _,n,role,d,i in people_rows if n.text().strip()]
                equipment=[{"name":n.text().strip(),"id":i.text().strip(),"specification":sp.text().strip()} for _,n,i,sp in equipment_rows if n.text().strip()]
                primary_corrective="\n".join(f"{i}. {x['action']}" for i,x in enumerate(actions,1))
                params=(today(),edits["incident_time"].text(),edits["location"].text(),edits["project"].text(),edits["company"].text(),edits["department"].text(),edits["activity"].text(),incident_type.currentText(),people[0]["name"] if people else "",people[0]["id"] if people else "",people[0]["designation"] if people else "",edits.get("supervisor",QLineEdit()).text() if "supervisor" in edits else "",edits.get("witnesses",QLineEdit()).text() if "witnesses" in edits else "",description.toPlainText(),immediate.toPlainText(),"", "", equipment[0]["name"] if equipment else "",method.currentText(),*(proc.get(f"why{i}","") for i in range(1,6)),"\n".join(direct_vals),"\n".join(contrib_vals),"\n".join(root_vals),primary_corrective,"",json.dumps(details,ensure_ascii=False),json.dumps(people,ensure_ascii=False),json.dumps(equipment,ensure_ascii=False),"", "", "", "", edits["organization"].text(), "", edits["report_reference"].text(),mode,datetime.now().isoformat())
                if existing:
                    db.execute("""UPDATE incidents SET incident_date=?,incident_time=?,location=?,project=?,company=?,department=?,activity=?,incident_type=?,person_involved=?,employee_id=?,designation=?,supervisor=?,witnesses=?,description=?,immediate_action=?,consequences=?,potential_consequences=?,equipment=?,investigation_method=?,why1=?,why2=?,why3=?,why4=?,why5=?,direct_cause=?,contributing_factors=?,root_cause=?,corrective_action=?,preventive_action=?,investigation_details=?,people_details=?,equipment_details=?,environmental_category=?,environmental_quantity=?,property_name=?,property_damage_cost=?,organization=?,procedure_reference=?,report_reference=?,status=?,created_at=? WHERE id=?""",params+(incident_id,))
                else:
                    db.execute("""INSERT INTO incidents (number,incident_date,incident_time,location,project,company,department,activity,incident_type,person_involved,employee_id,designation,supervisor,witnesses,description,immediate_action,consequences,potential_consequences,equipment,investigation_method,why1,why2,why3,why4,why5,direct_cause,contributing_factors,root_cause,corrective_action,preventive_action,investigation_details,people_details,equipment_details,environmental_category,environmental_quantity,property_name,property_damage_cost,organization,procedure_reference,report_reference,status,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(number,)+params)
                rid=incident_id or db.fetchone("SELECT id FROM incidents WHERE number=?",(number,))["id"]
                db.execute("UPDATE incidents SET incident_title=?, permit_number=? WHERE id=?",
                           (edits["incident_title"].text().strip(), edits["permit_number"].text().strip(), rid))
                # Save Fishbone cause attachments and replace temporary source paths with
                # the saved filenames so reports can display the actual evidence.
                if method.currentText()=="Fishbone / Ishikawa":
                    for frow in fishbone_rows_data:
                        srcs=frow.get("attachments",[]) or []
                        saved_names=[]
                        new_sources=[]
                        for src in srcs:
                            sp=Path(src)
                            if sp.exists():
                                new_sources.append(str(sp))
                            else:
                                # Existing Fishbone attachments are already stored in the
                                # incident attachment table; retain their filenames.
                                saved_names.append(sp.name)
                        if new_sources:
                            saved=copy_attachments(new_sources,number,"incident_attachments",rid,"Fishbone Cause")
                            saved_names.extend(Path(x).name for x in saved)
                        frow["attachments"]=saved_names
                    details["fishbone"]=fishbone_rows_data
                    db.execute("UPDATE incidents SET investigation_details=? WHERE id=?", (json.dumps(details, ensure_ascii=False), rid))
                for _,file_store,remark,lab in evidence_rows:
                    if file_store:
                        saved=copy_attachments(file_store,number,"incident_attachments",rid,"Evidence")
                        for saved_path in saved:
                            db.execute("UPDATE incident_attachments SET remark=? WHERE incident_id=? AND file_path=?",(remark.text().strip(),rid,saved_path))
                for idx, item in enumerate(witness_rows):
                    if idx >= len(details["witness_statements"]): continue
                    files=item[4]
                    if files:
                        saved=copy_attachments(files,number,"incident_attachments",rid,"Witness Statement")
                        details["witness_statements"][idx]["attachment"]="; ".join(Path(x).name for x in saved)
                        for saved_path in saved:
                            db.execute("UPDATE incident_attachments SET remark=? WHERE incident_id=? AND file_path=?",(item[3].text().strip(),rid,saved_path))
                # Save closeout evidence against the corrective-action plan items.
                for idx, item in enumerate(ca_rows):
                    if idx >= len(actions):
                        continue
                    files = item[5]
                    if files:
                        saved = copy_attachments(files, number, "incident_attachments", rid, "Closeout Evidence")
                        actions[idx]["attachments"] = [Path(x).name for x in saved]
                # Persist the final corrective-action attachment names without altering legacy fields.
                details["corrective_actions"] = actions
                db.execute("UPDATE incidents SET investigation_details=? WHERE id=?", (json.dumps(details, ensure_ascii=False), rid))

                # Keep the incident corrective action plan linked to the CAPA register,
                # following the same single-record / multi-action structure as observations.
                if incident_capa_required.currentText() == "Yes" and actions:
                    incident_row = db.fetchone("SELECT * FROM incidents WHERE id=?", (rid,))
                    linked_capa_no = safe(incident_row["capa_number"] if incident_row and "capa_number" in incident_row.keys() else "")
                    if not linked_capa_no:
                        linked_capa_no = next_number("HSE-CAPA", "capa")
                    linked_actions = [{
                        "action": safe(a.get("action", "")),
                        "evidence": [str(ATTACH_DIR / safe(n)) for n in (a.get("attachments", []) or []) if safe(n)],
                        "has_evidence": bool(a.get("attachments")),
                        "responsible": safe(a.get("responsible", "")),
                        "target_date": safe(a.get("target_date", "")),
                        "status": safe(a.get("status", "Open")),
                    } for a in actions if safe(a.get("action", "")).strip()]
                    corrective_display = format_corrective_actions_display([a["action"] for a in linked_actions])
                    incident_finding = (
                        f"Incident Investigation Report: {number}\n"
                        f"Incident Title: {safe(edits['incident_title'].text().strip())}\n"
                        f"Incident Date: {safe(incident_row['incident_date']) if incident_row else today()}\n"
                        f"Incident Type: {incident_type.currentText()}\n"
                        f"Location: {edits['location'].text().strip()}\n\n"
                        f"Incident Description: {description.toPlainText().strip()}"
                    )
                    old_capa = db.fetchone("SELECT id FROM capa WHERE number=?", (linked_capa_no,))
                    capa_values = (
                        "Incident Investigation",
                        f"{number} — Accident Corrective Action Plan",
                        incident_finding,
                        "",
                        corrective_display,
                        "",
                        "; ".join(sorted({safe(a.get("responsible", "")) for a in linked_actions if safe(a.get("responsible", "")).strip()})),
                        "High",
                        min([safe(a.get("target_date", "")) for a in linked_actions if safe(a.get("target_date", ""))] or [""]),
                        "Open",
                        "",
                        "",
                        datetime.now().isoformat(),
                        json.dumps(linked_actions, ensure_ascii=False),
                        "Pending Approval",
                        "",
                        "",
                        0,
                        "Open",
                        incident_finding,
                    )
                    if old_capa:
                        db.execute("""UPDATE capa SET source=?, reference_number=?, finding=?, root_cause=?, corrective_action=?,
                            preventive_action=?, responsible=?, priority=?, target_date=?, status=?, verification=?,
                            closeout_evidence=?, created_at=?, actions_json=?, approval_status=?, approved_by=?,
                            approved_date=?, completion_pct=?, action_status=?, observation_text=? WHERE number=?""",
                            capa_values + (linked_capa_no,))
                    else:
                        db.execute("""INSERT INTO capa (number, source, reference_number, finding, root_cause, corrective_action,
                            preventive_action, responsible, priority, target_date, status, verification, closeout_evidence,
                            created_at, actions_json, approval_status, approved_by, approved_date, completion_pct,
                            action_status, observation_text) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (linked_capa_no,) + capa_values)
                    db.execute("UPDATE incidents SET capa_required=1, capa_number=? WHERE id=?", (linked_capa_no, rid))
                else:
                    # Do not delete any previously-created CAPA history if the toggle is later changed to No.
                    db.execute("UPDATE incidents SET capa_required=? WHERE id=?", (1 if (existing and safe(existing["capa_number"] if "capa_number" in existing.keys() else "")) else 0, rid))
                dialog.accept(); refresh()
            except Exception as e: logging.exception("Incident save failed"); QMessageBox.critical(dialog,"Save Error",f"Unable to save investigation.\n\n{e}")
        save_draft.clicked.connect(lambda:save("Draft")); save_final.clicked.connect(lambda:save("Closed")); save_exit.clicked.connect(lambda:save("Draft")); dialog.exec()

    # ========================================================
    # AUDIT MODULE - self-contained implementation
    # ========================================================
    def _audit_setting_list(self, key, default):
        raw = db.setting(key, "")
        if raw:
            try:
                value = json.loads(raw)
                if isinstance(value, list) and value:
                    return [str(x) for x in value]
            except Exception:
                pass
        return list(default)

    def _audit_employees(self):
        try:
            return db.fetchall("SELECT id,name,designation,department FROM employees WHERE active=1 ORDER BY name")
        except Exception:
            return []

    def _audit_person_combo(self, include_blank=True):
        combo=QComboBox()
        if include_blank: combo.addItem("-- Select Person --", None)
        people=self._audit_employees()
        if people:
            for r in people:
                label=safe(r["name"])
                if safe(r["designation"]): label += f" — {safe(r['designation'])}"
                combo.addItem(label, r["id"])
        else:
            combo.addItem("No active employees configured", None)
        return combo

    def _audit_standards(self):
        return self._audit_setting_list("audit_standards", ["ISO 9001", "ISO 14001", "ISO 45001", "Other"])

    def _audit_types(self):
        return self._audit_setting_list("audit_types", AUDIT_TYPES)

    def _next_audit_reference(self, audit_type, audit_date, exclude_id=None):
        """Create the Audit module's company/type/year serial, e.g. AB-IA01-2026."""
        company = re.sub(r"[^A-Za-z0-9]", "", safe(company_name()).upper())
        prefix = company[:2] or "CO"
        code = "IA" if "internal" in safe(audit_type).lower() else "EA"
        try:
            year = date.fromisoformat(safe(audit_date)).year
        except Exception:
            year = date.today().year
        pattern = re.compile(rf"^{re.escape(prefix)}-{code}(\d+)-{year}$", re.I)
        nums = []
        for row in db.fetchall("SELECT id,number FROM audits"):
            if exclude_id is not None and int(row["id"]) == int(exclude_id):
                continue
            match = pattern.fullmatch(safe(row["number"]))
            if match:
                nums.append(int(match.group(1)))
        return f"{prefix}-{code}{max(nums + [0]) + 1:02d}-{year}"

    def _renumber_audits_after_delete(self, deleted_number):
        """Close serial gaps for the deleted audit's type/year; leave legacy refs untouched."""
        match = re.fullmatch(r"([A-Z0-9]{2})-(IA|EA)(\d+)-(\d{4})", safe(deleted_number).upper())
        if not match:
            return
        prefix, code, _old_seq, year = match.groups()
        rows = db.fetchall("SELECT id,number,audit_date FROM audits ORDER BY audit_date,id")
        target = []
        for row in rows:
            if re.fullmatch(rf"{re.escape(prefix)}-{code}\d+-{year}", safe(row["number"]).upper()):
                target.append(row)
        for seq, row in enumerate(target, 1):
            wanted = f"{prefix}-{code}{seq:02d}-{year}"
            if safe(row["number"]) != wanted:
                db.execute("UPDATE audits SET number=?,updated_at=? WHERE id=?", (wanted, datetime.now().isoformat(timespec="seconds"), row["id"]))

    def _audit_finding_types(self):
        return self._audit_setting_list("audit_finding_types", [
            "Major Nonconformity", "Minor Nonconformity", "Observation",
            "Opportunity for Improvement", "Positive Finding"])

    def _audit_finding_statuses(self):
        return self._audit_setting_list("audit_finding_statuses", [
            "Open", "In Progress", "Submitted for Verification", "Verified", "Closed", "Overdue"])

    def _audit_clause_map(self, standard):
        raw=db.setting("audit_clause_map", "")
        if raw:
            try:
                data=json.loads(raw)
                if isinstance(data,dict) and standard in data and isinstance(data[standard],dict):
                    return data[standard]
            except Exception:
                pass
        return ISO_CLAUSES.get(standard,{})

    def _audit_log(self, audit_id, action, details="", user=None):
        user = user or db.setting("current_user", "") or db.setting("default_observer", "") or "System User"
        db.execute("INSERT INTO audit_history(audit_id,action,user_name,details,timestamp) VALUES(?,?,?,?,?)",
                   (audit_id,action,user,details,datetime.now().isoformat(timespec="seconds")))

    def _audit_attachment_copy(self, paths, audit_number, audit_id, finding_id=None, uploaded_by=""):
        saved=[]
        for path in paths or []:
            try:
                source=Path(path)
                if not source.exists(): continue
                prefix=f"{audit_number}_{'F'+str(finding_id)+'_' if finding_id else ''}"
                destination=ATTACH_DIR / f"{prefix}{source.name}"
                n=1
                while destination.exists():
                    destination=ATTACH_DIR / f"{prefix}{n}_{source.name}"; n+=1
                shutil.copy2(source,destination)
                db.execute("INSERT INTO audit_attachments(audit_id,finding_id,file_path,attachment_type,uploaded_by,uploaded_at) VALUES(?,?,?,?,?,?)",
                           (audit_id,finding_id,str(destination),"Evidence",uploaded_by,datetime.now().isoformat(timespec="seconds")))
                saved.append(str(destination))
            except Exception:
                logging.exception("Audit attachment copy failed")
        return saved

    def _audit_attachment_rows(self, audit_id, finding_id=None):
        if finding_id is None:
            return db.fetchall("SELECT * FROM audit_attachments WHERE audit_id=? AND (finding_id IS NULL OR finding_id=0) ORDER BY id",(audit_id,))
        return db.fetchall("SELECT * FROM audit_attachments WHERE audit_id=? AND finding_id=? ORDER BY id",(audit_id,finding_id))

    def audits(self):
        """Audit home page: dashboard first; keep register actions in a separate tab."""
        w, layout = self.page("Audit Management")
        tabs = QTabWidget(); layout.addWidget(tabs, 1)
        dash = QWidget(); dl = QVBoxLayout(dash)
        title = QLabel("HSE AUDIT DASHBOARD")
        title.setStyleSheet("font-size:22px;font-weight:bold;color:#17365D;padding:8px;"); dl.addWidget(title)
        summary = QLabel(); summary.setWordWrap(True)
        summary.setStyleSheet("font-size:15px;font-weight:bold;color:#17365D;padding:12px;background:#E8F0FA;border:1px solid #C7D7EA;border-radius:5px;"); dl.addWidget(summary)
        class AuditHomeCharts(QWidget):
            def __init__(self, parent=None):
                super().__init__(parent); self.status_counts=[]; self.finding_counts=[0,0]; self.setMinimumHeight(300)
            def set_data(self, statuses, opened, closed):
                self.status_counts=list(statuses); self.finding_counts=[int(opened),int(closed)]; self.update()
            def paintEvent(self, event):
                painter=QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing); painter.fillRect(self.rect(),QColor("#FFFFFF")); painter.setPen(QColor("#17365D")); painter.setFont(QFont("Arial",11,QFont.Weight.Bold)); painter.drawText(18,24,"AUDITS BY STATUS — BAR GRAPH"); painter.drawText(max(360,self.width()//2),24,"FINDINGS — PIE CHART")
                left=QRectF(18,42,max(280,self.width()//2-35),225); painter.setPen(QPen(QColor("#D5DEE8"),1)); painter.drawRoundedRect(left,6,6)
                rows=self.status_counts or [("No audits yet",0)]; maxv=max([int(v) for _,v in rows]+[1]); bh=max(18,min(30,int((left.height()-25)/max(1,min(len(rows),6)))-5))
                for i,(label,value) in enumerate(rows[:6]):
                    y=left.top()+14+i*(bh+7); painter.setFont(QFont("Arial",9)); painter.setPen(QColor("#34495E")); painter.drawText(int(left.left()+8),int(y+bh-4),str(label)[:18]); bw=max(0,(left.width()-145)*int(value)/maxv); painter.fillRect(QRectF(left.left()+112,y,bw,bh),QColor("#4472C4")); painter.setPen(QColor("#17365D")); painter.drawText(int(left.left()+118+bw),int(y+bh-4),str(value))
                cx=max(360,self.width()//2)+(self.width()-max(360,self.width()//2))//2; cy=145; radius=78; total=sum(self.finding_counts); start_angle=0
                for i,value in enumerate(self.finding_counts):
                    span=int(360*value/total) if total else 180; painter.setBrush([QColor("#4472C4"),QColor("#70AD47")][i]); painter.setPen(QPen(QColor("#FFFFFF"),2)); painter.drawPie(QRectF(cx-radius,cy-radius,radius*2,radius*2),start_angle*16,span*16); start_angle+=span
                painter.setPen(QColor("#34495E")); painter.setFont(QFont("Arial",9)); painter.drawText(cx-90,cy+radius+24,f"Open: {self.finding_counts[0]}    Closed: {self.finding_counts[1]}")
        charts=AuditHomeCharts(); dl.addWidget(charts,1)
        actions=QHBoxLayout(); planning_btn=QPushButton("Open Audit Planning & Checklists"); planning_btn.setMinimumHeight(40); planning_btn.setStyleSheet("font-weight:bold;padding:8px 14px;background:#E8F0FA;color:#17365D;"); refresh_dash=QPushButton("Refresh Dashboard"); actions.addWidget(planning_btn); actions.addStretch(); actions.addWidget(refresh_dash); dl.addLayout(actions)
        def load_home_dashboard():
            total=db.fetchone("SELECT COUNT(*) AS n FROM audits")["n"]; scheduled=db.fetchone("SELECT COUNT(*) AS n FROM audit_schedule")["n"]
            opened=db.fetchone("SELECT COUNT(*) AS n FROM audit_findings WHERE lower(COALESCE(status,'open')) NOT IN ('closed','verified')")["n"]; closed=db.fetchone("SELECT COUNT(*) AS n FROM audit_findings WHERE lower(COALESCE(status,'')) IN ('closed','verified')")["n"]
            summary.setText(f"Total audits: {total}     |     Scheduled audits: {scheduled}     |     Open findings: {opened}     |     Closed findings: {closed}")
            rows=db.fetchall("SELECT COALESCE(NULLIF(status,''),'Unspecified') AS status,COUNT(*) AS n FROM audits GROUP BY COALESCE(NULLIF(status,''),'Unspecified') ORDER BY status")
            charts.set_data([(safe(r["status"]),int(r["n"])) for r in rows],opened,closed)
        planning_btn.clicked.connect(self.audit_planning_center); refresh_dash.clicked.connect(load_home_dashboard); load_home_dashboard(); tabs.addTab(dash,"Audit Dashboard")

        # Full register and CRUD controls are available here, not on the dashboard home page.
        reg=QWidget(); rl=QVBoxLayout(reg); top=QHBoxLayout(); reg_title=QLabel("MANAGE AUDITS"); reg_title.setStyleSheet("font-size:20px;font-weight:bold;color:#17365D;padding:6px;"); top.addWidget(reg_title); top.addStretch()
        new_btn=QPushButton("+ New Audit"); edit_btn=QPushButton("Edit Selected Audit"); delete_btn=QPushButton("Delete Selected Audit"); settings_btn=QPushButton("Audit Settings"); report_btn=QPushButton("Generate Report"); plan_btn=QPushButton("Audit Planning & Checklists")
        for b in [plan_btn,settings_btn,report_btn,edit_btn,delete_btn,new_btn]: b.setMinimumHeight(36); top.addWidget(b)
        rl.addLayout(top); filters=QHBoxLayout(); search=QLineEdit(); search.setPlaceholderText("Search reference, title, department, location or auditor..."); filters.addWidget(search,2)
        status=QComboBox(); status.addItems(["All Statuses","Draft","Submitted"]); filters.addWidget(status); typ=QComboBox(); typ.addItem("All Audit Types"); typ.addItems(self._audit_types()); filters.addWidget(typ)
        dept=QComboBox(); dept.addItem("All Departments")
        try: dept.addItems(sorted({safe(r["department"]) for r in db.fetchall("SELECT department FROM employees WHERE active=1") if safe(r["department"])}))
        except Exception: pass
        filters.addWidget(dept); rl.addLayout(filters)
        table=QTableWidget(); headers=["Audit Reference","Audit Type","Audit Title","Audit Date","Department","Location","Auditor","Findings","Status","Created By","Last Updated","Actions"]; table.setColumnCount(len(headers)); table.setHorizontalHeaderLabels(headers); table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); table.setAlternatingRowColors(True); table.setSortingEnabled(True); table.setWordWrap(True); table.setTextElideMode(Qt.TextElideMode.ElideNone); table.verticalHeader().setDefaultSectionSize(54); table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive); table.horizontalHeader().setStretchLastSection(True)
        for col,width in {0:145,1:135,2:190,3:95,4:125,5:130,6:130,7:65,8:105,9:115,10:145,11:80}.items(): table.setColumnWidth(col,width)
        rl.addWidget(table,1)
        def load():
            rows=db.fetchall("SELECT * FROM audits ORDER BY id DESC"); data=[]
            for r in rows:
                if status.currentText()!="All Statuses" and safe(r["status"])!=status.currentText(): continue
                if typ.currentText()!="All Audit Types" and safe(r["audit_type"])!=typ.currentText(): continue
                if dept.currentText()!="All Departments" and safe(r["department"])!=dept.currentText(): continue
                needle=search.text().strip().lower(); hay=" ".join(safe(r[k]) for k in ["number","audit_type","title","audit_date","department","location","auditor","lead_auditor","status","created_by"] if k in r.keys()).lower()
                if needle and needle not in hay: continue
                data.append(r)
            table.setSortingEnabled(False); table.setRowCount(len(data))
            for i,r in enumerate(data):
                count=db.fetchone("SELECT COUNT(*) AS c FROM audit_findings WHERE audit_id=?",(r["id"],))["c"]; vals=[r["number"],r["audit_type"],r["title"],r["audit_date"],r["department"],r["location"],r["auditor"] or r["lead_auditor"],count,r["status"],r["created_by"],r["updated_at"] or r["created_at"],"View" if r["status"]=="Submitted" else "Edit"]
                for c,v in enumerate(vals): item=QTableWidgetItem(safe(v)); item.setData(Qt.ItemDataRole.UserRole,r["id"]); table.setItem(i,c,item)
            table.resizeRowsToContents(); table.setSortingEnabled(True); load_home_dashboard()
        def selected_id():
            row=table.currentRow(); return table.item(row,0).data(Qt.ItemDataRole.UserRole) if row>=0 and table.item(row,0) else None
        def open_selected():
            aid=selected_id()
            if aid: self.audit_workspace(aid,load)
        def generate_selected():
            aid=selected_id()
            if aid: self._audit_register_detail(aid)
            else: QMessageBox.information(self,"Select Audit","Select an audit from the register first.")
        def edit_selected_audit():
            aid=selected_id()
            if not aid: return QMessageBox.information(self,"Select Audit","Select an audit to edit first.")
            self.audit_workspace(aid,load,allow_edit=True)
        def delete_selected_audit():
            aid=selected_id()
            if not aid: return QMessageBox.information(self,"Select Audit","Select an audit to delete first.")
            ar=db.fetchone("SELECT number,title FROM audits WHERE id=?",(aid,))
            if not ar: return load()
            answer=QMessageBox.question(self,"Delete Audit",f"Delete audit {safe(ar['number'])} — {safe(ar['title'])}?\n\nIts audit findings, evidence and audit history will be removed. Any linked CAPA record will be retained in the existing CAPA Tracking Register.",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)
            if answer!=QMessageBox.StandardButton.Yes: return
            try:
                deleted_number=safe(ar["number"])
                db.execute("UPDATE audit_schedule SET linked_audit_id=NULL,updated_at=? WHERE linked_audit_id=?",(datetime.now().isoformat(timespec="seconds"),aid)); db.execute("DELETE FROM audit_finding_history WHERE finding_id IN (SELECT id FROM audit_findings WHERE audit_id=?)",(aid,)); db.execute("DELETE FROM audit_attachments WHERE audit_id=?",(aid,)); db.execute("DELETE FROM audit_findings WHERE audit_id=?",(aid,)); db.execute("DELETE FROM audit_history WHERE audit_id=?",(aid,)); db.execute("DELETE FROM audits WHERE id=?",(aid,)); self._renumber_audits_after_delete(deleted_number); load(); QMessageBox.information(self,"Audit Deleted","Selected audit deleted. Matching company/type/year serial numbers have been resequenced. Existing CAPA records were retained.")
            except Exception as exc: logging.exception("Audit deletion failed"); QMessageBox.critical(self,"Delete Failed",f"Could not delete the selected audit.\n\n{exc}")
        new_btn.clicked.connect(lambda:self.audit_type_selector(load)); settings_btn.clicked.connect(self.audit_settings); report_btn.clicked.connect(generate_selected); edit_btn.clicked.connect(edit_selected_audit); delete_btn.clicked.connect(delete_selected_audit); plan_btn.clicked.connect(self.audit_planning_center)
        table.cellDoubleClicked.connect(lambda *_:open_selected()); table.cellClicked.connect(lambda row,col:self._audit_register_detail(table.item(row,0).data(Qt.ItemDataRole.UserRole)) if col==0 and table.item(row,0) else None)
        for ctl in [search,status,typ,dept]: ctl.textChanged.connect(load) if isinstance(ctl,QLineEdit) else ctl.currentTextChanged.connect(load)
        load(); tabs.addTab(reg,"Manage Audits")

    def _audit_procedure_library(self):
        # Built from the supplied Master Audit Prompt. Short summary checklists are
        # expanded with standard evidence/control checks so each procedure is usable
        # as a practical field checklist.
        procedures = {
            "Generic HSE Audit": ("Applicable legal, client, company and ISO requirements", [
                "Current HSE policy is approved, communicated and displayed where appropriate",
                "Roles, responsibilities, competence and management accountability are defined",
                "Hazards are identified and risks assessed for routine and non-routine activities",
                "Risk controls follow the hierarchy of controls and are implemented at the workplace",
                "Workers and contractors receive induction, task briefings and required training",
                "Permits, risk assessments and method statements are current where required",
                "Work areas, access routes, housekeeping and welfare facilities are satisfactory",
                "Plant, tools, lifting equipment and vehicles are inspected and maintained",
                "PPE is suitable, available, maintained and used correctly",
                "Emergency plans, contacts, alarms, routes and equipment are suitable and tested",
                "Incidents, near misses and unsafe conditions are reported and investigated",
                "Inspection, monitoring, exposure and statutory records are current and traceable",
                "Environmental aspects, waste, spills and emissions are controlled where applicable",
                "Previous audit findings and corrective actions are tracked and effectiveness verified",
                "Management reviews performance trends and communicates improvement actions"]),
            "General Workplace Inspection": ("Workplace condition and activity inspection", [
                "Walkways, stairs, exits and emergency access routes are clear and safe",
                "Floors are free from uncontrolled spills, trip hazards and damaged surfaces",
                "Materials, cylinders, tools and equipment are stored securely and appropriately",
                "Electrical cables, plugs, temporary supplies and distribution boards are undamaged",
                "Fire extinguishers, alarms, hose reels and emergency equipment are accessible",
                "Emergency exits, signs, lighting and assembly-point information are visible",
                "Machines and tools have suitable guards and are in safe operating condition",
                "Lifting gear, slings, shackles and accessories show valid identification and inspection",
                "Work at height areas have suitable edge protection, access and dropped-object controls",
                "Chemical containers are labelled and stored with suitable spill containment",
                "Waste is segregated, bins are suitable and hazardous waste is controlled",
                "PPE is appropriate to the task and worn correctly by workers and visitors",
                "Permits, isolations and risk controls match the activities being observed",
                "Workers can explain task hazards, stop-work expectations and emergency actions",
                "Unsafe conditions are recorded with an owner, target date and follow-up verification"]),
            "Permit to Work (PTW) System": ("8.1, 8.1.2, 6.1, 7.2", [
                "Correct type of permit selected","Scope of work clearly described and matches actual activity","JSA / Risk Assessment attached and relevant","All required isolations identified and completed","Gas testing completed within validity period","All authorized signatures present","SIMOPS conflict check completed","Permit displayed at work location","Workers understand permit conditions","Permit properly closed after completion","No expired permits still active","Supporting documents attached"]),
            "Hot Work Procedure": ("8.1, 8.1.2, 8.2", ["Hot Work Permit issued and valid","Gas test completed before start and periodically","Combustible materials removed or protected","Fire watch assigned, trained and present","Suitable fire extinguisher available","Fire watch continues after completion","Welding equipment in good condition with flashback arrestors","Area properly ventilated","Work stopped if conditions change","Permit closed and area inspected"]),
            "Confined Space Entry": ("8.1, 8.1.2, 8.2, 7.2", ["Space correctly classified","Entry Permit fully completed","Atmospheric testing done in correct sequence (O₂ → LEL → H₂S → CO)","Results within acceptable limits","Continuous monitoring during entry","All energy sources isolated and LOTO applied","Attendant present at all times","Rescue plan and equipment staged","Entrants wearing full body harness with retrieval line","Communication system tested","Entry log maintained","Forced ventilation operating if required"]),
            "Lockout / Tagout (LOTO)": ("8.1, 8.1.2", ["All energy sources identified","Written isolation procedure available","Locks and tags applied by authorized persons","Each person has individual lock","Try-start verification performed","Isolation Certificate completed","Tags legible with name, date and reason","Group lockout box used when required","De-isolation done only after all locks removed","Equipment tested safe before return to service"]),
            "Working at Height": ("8.1, 8.1.2, 7.2", ["Fall Protection Plan available","Hierarchy of controls applied","Full body harness inspected and valid","Anchor points rated and inspected","Correct lanyards / SRLs used","Scaffold / platform / MEWP inspected and tagged","Edge protection or hole covers in place","Tools tethered","Exclusion zone below work area","Rescue plan available"]),
            "Lifting & Rigging / Crane Operations": ("8.1, 8.1.2, 7.2", ["Lift categorized correctly","Lift Plan approved for complex/critical lifts","Crane and lifting gear certified","Load chart available and followed","Riggers and banksman competent","Tag lines used","Exclusion zone enforced","No personnel under suspended load","Weather conditions suitable","Pre-lift checks completed"]),
            "Excavation & Trenching": ("8.1, 8.1.2", ["Excavation Permit issued","Underground services located and marked","Soil type assessed","Protective system (shoring/sloping/benching) in place","Safe access and egress provided","Spoil placed at safe distance","Barricades and warning signs installed","Atmosphere tested if required","Daily inspection by competent person","Emergency rescue plan available"]),
            "Scaffolding": ("8.1, 8.1.2, 7.2", ["Erected by competent scaffolders","Design available for complex scaffolds","Correctly tagged","Base plates, standards, ledgers, braces correct","Guardrails, mid-rails and toe boards complete","Platforms fully boarded","Safe access provided","Inspected before use and after alteration","No overloading","Adequate ties and bracing"]),
            "Electrical Safety": ("8.1, 8.1.2, 7.2", ["Only competent electricians working","Live work controlled by permit","Temporary installations inspected","RCDs/ELCBs tested","Cables protected","LOTO applied for maintenance","Correct PPE used","Distribution boards locked and labeled","Earth leakage protection functional","No unauthorized modifications"]),
            "Hazard Communication / Chemical Safety": ("8.1, 7.3, 6.1", ["Chemical register up to date","SDS available and accessible","All containers labeled","Secondary containment provided","Incompatible chemicals separated","Spill kits available","Workers trained on hazards","Correct PPE used","Emergency eye wash/shower available","Chemical waste disposal controlled"]),
            "Personal Protective Equipment (PPE)": ("8.1, 7.1, 7.2", ["PPE selection matches identified hazards","PPE condition checked before use","Workers trained in correct use and limitations","PPE stored hygienically and protected from damage","Defective PPE quarantined and replaced","Fit testing completed where required","Issue and replacement records maintained","Supervision verifies consistent use"]),
            "Job Safety Analysis (JSA) / Risk Assessment": ("6.1.2, 8.1", ["Assessment is task-specific and current","Workers involved in hazard identification","Routine and non-routine hazards identified","Controls follow the hierarchy of controls","Residual risk assessed and acceptable","Controls implemented at worksite","Changes trigger reassessment","Approvals and review records retained"]),
            "Incident Reporting & Investigation": ("10.2, 9.1", ["Incidents and near misses reported promptly","Investigation team and responsibilities defined","Evidence and witness statements collected","Root cause analysis completed","Corrective actions assigned to responsible persons","Actions have target dates and priority","Corrective actions verified and closed","Lessons learned shared with relevant workers"]),
            "Emergency Preparedness & Response": ("8.2", ["Emergency plan current and approved","Roles and responsibilities defined","Drills conducted and recorded","Emergency equipment inspected and available","Assembly points marked and communicated","Emergency contacts displayed and current","Drill findings tracked to closure","Workers and contractors know response actions"]),
            "Fire Prevention & Protection": ("8.1, 8.2", ["Extinguishers inspected and accessible","Hot work controlled by permit","Flammable materials stored correctly","Detection and alarm systems functional","Escape routes and exits clear","Fire doors and emergency lighting serviceable","Workers know alarm and evacuation actions","Fire inspection findings tracked to closure"]),
            "H₂S Safety": ("8.1, 7.2, 8.2", ["H₂S detectors calibrated and bump-tested","Escape masks available and inspected","Training and competency current","Alarm response and evacuation practiced","Wind direction indicators available where required","Muster points and escape routes identified","Rescue arrangements do not expose rescuers","Inspection and drill records maintained"]),
            "Welding & Cutting": ("8.1", ["Cylinders secured upright and segregated","Flashback arrestors fitted correctly","Ventilation adequate for fumes","Hot work permit and gas testing completed","Correct welding PPE used","Welding leads and equipment inspected","Fire watch and extinguishers provided","Work area checked after completion"]),
            "Pressure Testing": ("8.1", ["Approved pressure-test procedure available","Exclusion zone established and controlled","Pressure gauges calibrated and suitable","Relief devices installed and verified","Competent persons assigned","Test medium and pressure limits confirmed","Connections and restraints inspected","Depressurization and emergency arrangements defined"]),
            "Dropped Objects Prevention": ("8.1", ["Tools tethered at height","Barricades established below work area","Tools and equipment checked before use","DROPS controls implemented and understood","Loose materials secured","Lifting baskets and storage arrangements inspected","Dropped-object hazards included in risk assessment","Inspections and corrective actions recorded"]),
            "SIMOPS": ("8.1, 8.1.3", ["SIMOPS matrix available and current","Coordination meetings held and recorded","Single PTW authority or clear interface defined","Interface risks assessed","Conflicting activities identified and controlled","Work groups briefed on restrictions","Changes communicated to affected teams","SIMOPS actions tracked to closure"]),
            "Contractor HSE Management": ("8.1.4, 7.2", ["Contractor prequalification completed","Induction completed before work","Competence and certificates verified","Risk assessments and method statements reviewed","Contractor performance monitored","Supervision and interfaces defined","Incidents and nonconformities reported","Corrective actions followed up"]),
            "Management of Change (MOC)": ("8.1.3", ["Change described and classified","Change risk assessment completed","Required technical and HSE approvals obtained","Affected workers informed and trained","Documents and drawings updated","Temporary changes have expiry/review date","Pre-Startup Safety Review completed where required","Post-change effectiveness verified"]),
            "Housekeeping": ("8.1", ["Walkways and access routes clear","Waste segregated correctly","Materials stored safely and secured","Spills cleaned promptly","Fire equipment and exits unobstructed","Cables and hoses routed safely","Scrap and sharp objects controlled","Housekeeping inspections and actions recorded"]),
            "Waste Management & Environmental Controls": ("8.1 + ISO 14001", ["Waste segregated and labeled","Waste storage areas suitable and controlled","Licensed disposal and transfer records kept","Spill kits available and inspected","Emissions and discharges monitored where required","Hazardous waste handled by trained personnel","Environmental incidents reported","Waste records and corrective actions reviewed"]),
            "Vehicle / Journey Management": ("8.1", ["Drivers licensed and competent","Pre-trip checks completed","Journey plans prepared for remote travel","Fatigue and driving hours controlled","Vehicle maintenance and inspection current","Seat belts and restraints used","Route and emergency communications confirmed","Journey completion and deviations monitored"]),
            "Manual Handling": ("8.1, 6.1", ["Manual handling risk assessment completed","Mechanical aids used where practicable","Team lifting arrangements defined","Workers trained in safe handling","Load weight and route assessed","Storage heights and access suitable","High-risk or repetitive tasks reviewed","Injuries and improvement actions monitored"]),
            "Noise & Hearing Conservation": ("8.1, 6.1", ["Noise exposure areas identified","Noise assessments current","Hearing protection suitable and used","Audiometry programme implemented where required","Equipment noise controls maintained","Workers informed of exposure risks","Hearing protection fit and condition checked","Records and corrective actions reviewed"]),
            "Respiratory Protection": ("8.1, 7.2", ["Respirator selected for identified hazard","Fit testing completed where required","Medical clearance obtained where applicable","Inspection, cleaning and storage controlled","Filters replaced at defined intervals","Users trained in limitations and use","Emergency respiratory equipment checked","Records are current and accessible"]),
            "First Aid & Medical Response": ("8.2, 7.1", ["Trained first aiders available for shift and location","First-aid kits stocked and inspected","Emergency numbers posted and current","Medical response and transport arrangements defined","Access routes for emergency response clear","Incident treatment records protected and maintained","First-aiders' training current","Deficiencies corrected promptly"]),
            "Stop Work Authority": ("5.1, 5.4, 8.1", ["Workers know their right and duty to stop unsafe work","No-retaliation expectations communicated","Stop-work events recorded","Supervisors respond promptly to stop-work concerns","Work restarts only after hazards controlled","Lessons learned communicated","Contractors understand stop-work process","Trends and recurring issues reviewed"]),
            "Training, Competency & Induction": ("7.2, 7.3", ["Training matrix available and current","Training and competency records current","High-risk tasks assigned only to competent persons","Induction covers site-specific hazards","Refresher training planned before expiry","Competence verified in practice","Training gaps assigned actions","Records retained and accessible"]),
            "Ladders & Platforms": ("8.1", ["Ladders and platforms inspected before use","Correct ladder type and rating selected","Ladders secured against movement","Three points of contact maintained where applicable","Defective equipment tagged and removed","Ground and access conditions suitable","Work duration and alternatives considered","Inspection records and user training maintained"]),
            "Radiography / NDT Safety": ("8.1", ["Controlled area established and secured","Barriers and warning signs displayed","Radiation monitoring and dosimetry provided","Competent and authorized operators assigned","Work permit and notification requirements met","Emergency arrangements available","Exposure equipment inspected and maintained","Survey and exposure records retained"]),
            "Abrasive Blasting / Painting": ("8.1", ["Containment used to control dust and overspray","Respiratory protection suitable and fit-tested where required","Abrasive and coating hazards assessed","Ventilation and exposure controls adequate","Waste and spent abrasive controlled","Hoses, pots and pressure equipment inspected","PPE and decontamination arrangements suitable","Environmental and inspection records maintained"]),
            "Pre-Startup Safety Review (PSSR)": ("8.1.3", ["PSSR checklist completed","Punch items categorized and closed or formally controlled","Equipment and safeguards verified against design","Operating procedures approved and available","Personnel trained before start-up","Permits and isolations status checked","Handover and authorization documented","Outstanding actions tracked to closure"]),
            "Fatigue Management": ("8.1, 6.1", ["Working hours and rest periods controlled","Fitness-for-work checks conducted","Fatigue hazards assessed for critical tasks","Rest facilities and transport arrangements adequate","Workers know fatigue reporting process","Supervisors monitor signs of fatigue","Overtime and shift patterns reviewed","Fatigue-related events and actions analysed"]),
            "Heat / Cold Stress": ("8.1", ["Work-rest regimes followed","Drinking water and hydration provided","Environmental conditions monitored","Shade, cooling or warming facilities available","Workers trained to recognize symptoms","Acclimatization and vulnerable workers considered","Emergency response for heat/cold illness defined","Records and corrective actions reviewed"]),
            "Spill Response & Environmental Incident": ("8.2 + ISO 14001", ["Response plan available and current","Spill kits suitable and ready","Drills conducted and recorded","Reporting and escalation process followed","Drain and sensitive-area protection available","Waste from cleanup safely contained and disposed","Response roles and contacts understood","Incident causes and actions tracked"]),
            "Visitor & Short Service Employee Control": ("8.1.4", ["Induction completed before access","Escort provided where required","PPE issued and suitable","Access restrictions explained","Short-service employees identified","Close supervision provided for high-risk tasks","Emergency arrangements communicated","Visitor and training records maintained"]),
            "HSE Performance Monitoring & KPI": ("9.1", ["Leading and lagging indicators tracked","Data definitions and sources consistent","Trends analysed at planned intervals","Targets and performance reviewed by management","Underperformance triggers action","Actions have owners and due dates","Data accuracy and reporting timeliness checked","Effectiveness of actions reviewed"])
                }
        governance = [
            "Document governance: approved owner, document number, revision history, effective date, review date and control of obsolete copies are defined and evidenced.",
            "Purpose and applicability: scope, exclusions, interfaces, triggering conditions and boundaries of the procedure are explicit and consistent with actual work.",
            "Roles and accountability: authority, responsibilities, competence requirements, handovers and escalation routes are assigned to roles and communicated.",
            "Requirements register: applicable legal, client, contractual, ISO and internal requirements are identified, updated and traceable to procedure controls.",
            "Risk methodology: task-specific hazards, risk assessment, hierarchy of controls and acceptance of residual risk are required before authorization.",
            "Prerequisites and approvals: required permits, method statements, drawings, isolations, resources and approvals are defined as hold points before work begins.",
            "Control design: steps, acceptance criteria, prohibited conditions, verification methods and stop-work triggers are clear, measurable and unambiguous.",
            "Interfaces: contractor responsibilities, simultaneous operations, affected departments and emergency interfaces are addressed and coordinated.",
            "Competence and communication: required training, authorization, briefing, acknowledgement and refresher requirements are defined and recorded.",
            "Records and retention: mandatory forms, registers, certificates, logs, test records, retention periods and record owners are identified and retrievable.",
            "Deviation and change control: deviations, abnormal conditions, temporary changes and revisions require authorization, risk review, communication and records.",
            "Monitoring and assurance: inspections, audits, sampling, performance indicators, review frequency and accountable reviewers are defined.",
            "Nonconformity and CAPA: containment, root-cause analysis, action ownership, due dates, effectiveness verification and closure authority are defined.",
            "Management review: incidents, audit trends, lessons learned, KPI results and operational changes trigger documented review and procedure improvement."
        ]
        expanded = {}
        for procedure, (clauses, specifics) in procedures.items():
            # Prioritize procedure-specific checks; add governance checks only to reach 15.
            criteria = [f"Procedure-specific control: verify {specific.lower()} against the current approved {procedure.lower()} procedure and retain objective evidence." for specific in specifics]
            seed = list(specifics) or ["required controls", "assigned responsibilities", "records and monitoring"]
            pad_index = 0
            while len(criteria) < 15:
                if pad_index < len(governance):
                    criteria.append(f"{procedure} — {governance[pad_index]}")
                else:
                    specific = seed[pad_index % len(seed)]
                    criteria.append(f"{procedure} — sample a recent work activity and verify that {specific.lower()} was implemented, recorded and reviewed.")
                pad_index += 1
            expanded[procedure] = (clauses, criteria[:15])
        # Audit-specific custom procedure records extend the starter library without changing
        # any other module. Custom checklists are not represented as officially verified unless reviewed.
        try:
            for record in db.fetchall("SELECT * FROM audit_procedure_records WHERE active=1 ORDER BY title"):
                try:
                    custom_items = json.loads(safe(record["checklist_json"]) or "[]")
                    if not isinstance(custom_items, list): custom_items = []
                except Exception:
                    custom_items = []
                if custom_items:
                    expanded[safe(record["title"])] = (safe(record["clause_references"]) or "Company procedure review", [str(x) for x in custom_items if str(x).strip()])
        except Exception:
            logging.exception("Unable to load custom Audit Procedure Library records")
        return expanded

    def audit_planning_center(self):
        d=QDialog(self); d.setWindowTitle("Audit Planning, Inspection Checklists & Findings Register")
        d.setWindowFlags(d.windowFlags() | Qt.WindowType.WindowMaximizeButtonHint); d.setSizeGripEnabled(True)
        screen=d.screen() or QApplication.primaryScreen()
        if screen:
            g=screen.availableGeometry(); d.setGeometry(g); d.setMinimumSize(min(700,g.width()),min(500,g.height()))
        else: d.resize(1200,800)
        root=QVBoxLayout(d)
        heading=QLabel("HSE AUDIT PLANNING & ASSURANCE"); heading.setStyleSheet("font-size:22px;font-weight:bold;color:#17365D;padding:8px;")
        root.addWidget(heading)
        intro=QLabel("Risk-based internal audit schedule, procedure compliance audits, controlled-document and record sampling, clause-based findings, evidence traceability and corrective-action follow-up.")
        intro.setWordWrap(True); root.addWidget(intro)
        tabs=QTabWidget(); root.addWidget(tabs,1)

        # ----- Schedule tab -----
        sched=QWidget(); sl=QVBoxLayout(sched)
        target_note=QLabel("Planning target: 4 procedure audits per month  |  High risk: every 6 months  |  Medium risk: annually  |  Low risk: every 18 months")
        target_note.setWordWrap(True); target_note.setStyleSheet("padding:8px;background:#E8F0FA;color:#17365D;font-weight:bold;border:1px solid #C7D7EA;border-radius:4px;")
        sl.addWidget(target_note)
        sbar=QHBoxLayout(); add=QPushButton("+ New Scheduled Audit"); edit=QPushButton("Edit Selected"); reschedule=QPushButton("Reschedule Selected Dates"); delete=QPushButton("Delete Selected"); sbar.addWidget(add); sbar.addWidget(edit); sbar.addWidget(reschedule); sbar.addWidget(delete); sbar.addStretch(); sl.addLayout(sbar)
        schedule_table=QTableWidget(0,10); schedule_table.setHorizontalHeaderLabels(["ID","Procedure / Audit Topic","Audit Area","Risk","Frequency","Planned Date","Owner","Status","Linked Audit","Notes"]); schedule_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); schedule_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection); schedule_table.setAlternatingRowColors(True); schedule_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); schedule_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents); schedule_table.horizontalHeader().setStretchLastSection(True); sl.addWidget(schedule_table,1)
        schedule_table.setColumnHidden(0,True)
        schedule_table.setColumnHidden(8,True)  # linkage is maintained internally; avoid a confusing manual-link column.
        proc_names=list(self._audit_procedure_library().keys())
        def schedule_selected_ids():
            ids=[]
            for idx in schedule_table.selectionModel().selectedRows():
                item=schedule_table.item(idx.row(),0)
                if item:
                    try: ids.append(int(item.text()))
                    except Exception: pass
            if not ids and schedule_table.currentRow()>=0 and schedule_table.item(schedule_table.currentRow(),0):
                try: ids=[int(schedule_table.item(schedule_table.currentRow(),0).text())]
                except Exception: pass
            return list(dict.fromkeys(ids))
        def schedule_selected_id():
            ids=schedule_selected_ids()
            return ids[0] if ids else None
        def load_schedule():
            rows=db.fetchall("SELECT s.*,a.number AS audit_number FROM audit_schedule s LEFT JOIN audits a ON a.id=s.linked_audit_id ORDER BY s.planned_date,s.id")
            schedule_table.setRowCount(len(rows))
            for i,r in enumerate(rows):
                vals=[r["id"],r["procedure_name"],r["area"] if "area" in r.keys() else "",r["risk_level"],r["frequency"],r["planned_date"],r["owner"],r["status"],r["audit_number"] or "",r["notes"]]
                for c,v in enumerate(vals): schedule_table.setItem(i,c,QTableWidgetItem(safe(v)))
                if safe(r["status"]).lower() in ("overdue","overdue - action required"):
                    for c in range(schedule_table.columnCount()):
                        schedule_table.item(i,c).setBackground(QColor("#FCE4D6"))
        def edit_schedule(schedule_id=None):
            row=db.fetchone("SELECT * FROM audit_schedule WHERE id=?",(schedule_id,)) if schedule_id else None
            form=QDialog(d); form.setWindowTitle("Scheduled Audit Plan"); form.resize(700,580); fl=QFormLayout(form)
            procedure=QComboBox(); procedure.addItems(proc_names)
            risk=QComboBox(); risk.addItems(["High","Medium","Low"])
            freq=QComboBox(); freq.addItems(["Every 6 months","Once per year","Every 18 months"])
            planned=QDateEdit(); planned.setCalendarPopup(True); planned.setDisplayFormat("dd-MMM-yyyy"); planned.setDate(QDate.currentDate().addMonths(1))
            area=QLineEdit(); area.setPlaceholderText("Worksite, department or audited area")
            owner=QLineEdit(); owner.setPlaceholderText("Audit owner / lead auditor")
            stat=QComboBox(); stat.addItems(["Planned","In Progress","Completed","Deferred","Cancelled"])
            notes=QTextEdit(); notes.setMaximumHeight(100)
            fl.addRow("Procedure / Audit Topic",procedure); fl.addRow("Audit Area / Department",area); fl.addRow("Risk Level",risk); fl.addRow("Frequency",freq); fl.addRow("Planned Date",planned); fl.addRow("Owner / Lead Auditor",owner); fl.addRow("Plan Status",stat); fl.addRow("Notes / Scope",notes)
            def update_frequency():
                freq.setCurrentText({"High":"Every 6 months","Medium":"Once per year","Low":"Every 18 months"}.get(risk.currentText(),"Once per year"))
            risk.currentTextChanged.connect(update_frequency)
            if row:
                procedure.setCurrentText(safe(row["procedure_name"])); area.setText(safe(row["area"] if "area" in row.keys() else "")); risk.setCurrentText(safe(row["risk_level"])); freq.setCurrentText(safe(row["frequency"])); owner.setText(safe(row["owner"])); stat.setCurrentText(safe(row["status"])); notes.setPlainText(safe(row["notes"]))
                try: planned.setDate(QDate.fromString(safe(row["planned_date"]),"yyyy-MM-dd"))
                except Exception: pass
            buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel); fl.addRow(buttons)
            def save_plan():
                now=datetime.now().isoformat(timespec="seconds")
                preserved_link = row["linked_audit_id"] if row and "linked_audit_id" in row.keys() else None
                values=(procedure.currentText(),area.text().strip(),risk.currentText(),freq.currentText(),planned.date().toString("yyyy-MM-dd"),owner.text().strip(),stat.currentText(),preserved_link,notes.toPlainText().strip(),now)
                if row:
                    db.execute("UPDATE audit_schedule SET procedure_name=?,area=?,risk_level=?,frequency=?,planned_date=?,owner=?,status=?,linked_audit_id=?,notes=?,updated_at=? WHERE id=?",values+(schedule_id,))
                else:
                    db.execute("INSERT INTO audit_schedule(procedure_name,area,risk_level,frequency,planned_date,owner,status,linked_audit_id,notes,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)",values)
                form.accept()
            buttons.accepted.connect(save_plan); buttons.rejected.connect(form.reject)
            if form.exec()==QDialog.DialogCode.Accepted:
                load_schedule(); load_checklist_schedule()
        add.clicked.connect(lambda:edit_schedule())
        edit.clicked.connect(lambda:edit_schedule(schedule_selected_id()) if schedule_selected_id() else QMessageBox.information(d,"Select Plan","Select a scheduled audit first."))
        def reschedule_plans():
            ids=schedule_selected_ids()
            if not ids:return QMessageBox.information(d,"Select Plans","Select one or more scheduled audits first.")
            form=QDialog(d); form.setWindowTitle("Reschedule Selected Audits"); fl=QFormLayout(form)
            dt=QDateEdit(); dt.setCalendarPopup(True); dt.setDisplayFormat("dd-MMM-yyyy"); dt.setDate(QDate.currentDate().addMonths(1)); fl.addRow(f"New planned date for {len(ids)} selected audit(s)",dt)
            bb=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel); fl.addRow(bb)
            bb.accepted.connect(form.accept); bb.rejected.connect(form.reject)
            if form.exec()==QDialog.DialogCode.Accepted:
                for sid in ids: db.execute("UPDATE audit_schedule SET planned_date=?,updated_at=? WHERE id=?",(dt.date().toString("yyyy-MM-dd"),datetime.now().isoformat(timespec="seconds"),sid))
                load_schedule(); load_checklist_schedule()
        def delete_plan():
            ids=schedule_selected_ids()
            if not ids:return QMessageBox.information(d,"Select Plans","Select one or more scheduled audits first.")
            if QMessageBox.question(d,"Delete Scheduled Audits",f"Delete {len(ids)} selected schedule(s) and their saved checklist responses? Existing Audit Register records and findings will not be deleted.")==QMessageBox.StandardButton.Yes:
                for sid in ids: db.execute("DELETE FROM audit_schedule WHERE id=?",(sid,))
                load_schedule(); load_checklist_schedule()
        reschedule.clicked.connect(reschedule_plans)
        delete.clicked.connect(delete_plan)
        load_schedule(); tabs.addTab(sched,"Audit Schedule")

        # ----- Internal audit checklist tab -----
        check=QWidget(); cl=QVBoxLayout(check)
        cbar=QHBoxLayout(); cbar.addWidget(QLabel("Planned audits — select multiple using the tick-box dropdown:"))
        schedule_combo=QComboBox(); schedule_combo.setEditable(True); schedule_combo.lineEdit().setReadOnly(True); schedule_combo.lineEdit().setPlaceholderText("Select planned audits…"); schedule_combo.setMinimumWidth(420); schedule_combo.setMinimumHeight(34); schedule_model=QStandardItemModel(schedule_combo); schedule_combo.setModel(schedule_model); cl.addWidget(schedule_combo)
        def schedule_combo_summary():
            count=sum(1 for i in range(schedule_model.rowCount()) if schedule_model.item(i).checkState()==Qt.CheckState.Checked)
            schedule_combo.setCurrentText(f"{count} planned audit(s) selected" if count else "Select planned audits…")
        def schedule_combo_changed(index):
            item=schedule_model.item(index)
            if item:
                item.setCheckState(Qt.CheckState.Unchecked if item.checkState()==Qt.CheckState.Checked else Qt.CheckState.Checked)
                schedule_combo_summary()
        schedule_combo.activated.connect(schedule_combo_changed)
        load_btn=QPushButton("Load Selected Checklists"); save_btn=QPushButton("Save Checklist & Update Findings"); export_checklist_btn=QPushButton("Export Excel"); export_word_btn=QPushButton("Export Word"); print_checklist_btn=QPushButton("Print Checklist")
        for b in [load_btn,save_btn,export_checklist_btn,export_word_btn,print_checklist_btn]: cbar.addWidget(b)
        cl.addLayout(cbar)
        score=QLabel("Procedure compliance: Not calculated"); score.setStyleSheet("font-weight:bold;color:#17365D;padding:8px;background:#E8F0FA;border:1px solid #C7D7EA;"); cl.addWidget(score)
        checklist_table=QTableWidget(0,14); checklist_table.setHorizontalHeaderLabels(["Audit Number","Procedure No.","Procedure / Planned Date","Procedure Compliance Audit Criteria","Primary ISO Clause","Other Applicable Clauses / Sub-clauses","Compliance Result","Objective Evidence / Gap","Evidence Attachment","Responsible Person / Action Owner","Target Date","Severity","CAPA Required","CAPA Not Required Justification"])
        checklist_table.setAlternatingRowColors(True); checklist_table.setWordWrap(True); checklist_table.setTextElideMode(Qt.TextElideMode.ElideNone); checklist_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); checklist_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection); checklist_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded); checklist_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded); checklist_table.setMinimumHeight(440); checklist_table.verticalHeader().setDefaultSectionSize(118); checklist_table.verticalHeader().setMinimumSectionSize(90)
        checklist_table.setStyleSheet("QTableWidget{gridline-color:#C9D3DF;font-size:10pt;} QTableWidget::item{padding:7px;} QHeaderView::section{background:#17365D;color:white;font-weight:bold;padding:8px;border:1px solid #D7E0EA;}")
        for col,width in {0:150,1:105,2:220,3:420,4:115,5:155,6:150,7:280,8:230,9:210,10:135,11:105,12:120,13:230}.items(): checklist_table.setColumnWidth(col,width)
        cl.addWidget(checklist_table,1)
        def load_checklist_schedule():
            selected={schedule_model.item(i).data(Qt.ItemDataRole.UserRole) for i in range(schedule_model.rowCount()) if schedule_model.item(i).checkState()==Qt.CheckState.Checked}
            schedule_model.clear()
            for r in db.fetchall("SELECT id,procedure_name,planned_date,status FROM audit_schedule ORDER BY planned_date,id"):
                item=QStandardItem(f"{r['planned_date']} | {r['procedure_name']} | {r['status']}"); item.setData(r["id"],Qt.ItemDataRole.UserRole); item.setFlags(Qt.ItemFlag.ItemIsEnabled|Qt.ItemFlag.ItemIsUserCheckable); item.setCheckable(True); item.setCheckState(Qt.CheckState.Checked if r["id"] in selected else Qt.CheckState.Unchecked); schedule_model.appendRow(item)
            schedule_combo.setCurrentIndex(-1); schedule_combo_summary()
        def selected_schedule_ids():
            return [int(schedule_model.item(i).data(Qt.ItemDataRole.UserRole)) for i in range(schedule_model.rowCount()) if schedule_model.item(i).checkState()==Qt.CheckState.Checked]
        def add_checklist_row(sid, sr, n, item_text, clause, saved=None):
            row=checklist_table.rowCount(); checklist_table.insertRow(row)
            proc=safe(sr["procedure_name"]); audit_id=ensure_internal_audit(sid,sr)
            ar=db.fetchone("SELECT number FROM audits WHERE id=?",(audit_id,)); audit_no=safe(ar["number"]) if ar else f"IA-PLAN-{sid:04d}"
            proc_no=f"PROC-{list(self._audit_procedure_library().keys()).index(proc)+1:02d}"
            for col,val in enumerate([audit_no,proc_no,f"{proc}\nPlanned: {safe(sr['planned_date'])}",item_text]):
                cell=QTableWidgetItem(safe(val)); cell.setTextAlignment(Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignVCenter); checklist_table.setItem(row,col,cell)
            checklist_table.item(row,0).setData(Qt.ItemDataRole.UserRole,sid); checklist_table.item(row,1).setData(Qt.ItemDataRole.UserRole,n)
            clauses=[x.strip() for x in safe(clause).replace("+",",").split(",") if x.strip()]
            checklist_table.setItem(row,4,QTableWidgetItem(clauses[0] if clauses else "—")); checklist_table.setItem(row,5,QTableWidgetItem("\n".join(clauses[1:]) if len(clauses)>1 else "—"))
            result=QComboBox(); result.addItems(["Not Audited","Conforming","Nonconformity","N/A","Positive Observation","Opportunity for Improvement"])
            comments=QTextEdit(); comments.setPlaceholderText("Record objective evidence: procedure/document ID and revision, sample/record reviewed, interview reference, gap and clause-based rationale."); comments.setMinimumHeight(72); comments.setMaximumHeight(112)
            evidence=QLineEdit(); evidence.setPlaceholderText("Attachment path / evidence reference"); browse=QPushButton("Browse…"); browse.setMaximumWidth(82)
            evidence_host=QWidget(); evidence_layout=QHBoxLayout(evidence_host); evidence_layout.setContentsMargins(3,3,3,3); evidence_layout.addWidget(evidence,1); evidence_layout.addWidget(browse)
            browse.clicked.connect(lambda checked=False, field=evidence: field.setText(QFileDialog.getOpenFileName(d,"Select Audit Evidence","","Evidence Files (*.png *.jpg *.jpeg *.bmp *.gif *.pdf *.doc *.docx *.xls *.xlsx *.txt *.mp4 *.avi *.mov);;All Files (*)")[0] or field.text()))
            owner=QLineEdit(); owner.setPlaceholderText("Finding owner / responsible person")
            due=QDateEdit(); due.setCalendarPopup(True); due.setDisplayFormat("dd-MMM-yyyy"); due.setDate(QDate.currentDate().addDays(30))
            severity=QComboBox(); severity.addItems(["Low","Medium","High","Critical"])
            capa_reason=QLineEdit(); capa_reason.setPlaceholderText("If CAPA = No, enter authorized justification")
            capa=QComboBox(); capa.addItems(["No","Yes"])
            if saved:
                old_result=safe(saved["result"]); result.setCurrentText({"Yes":"Conforming","No":"Nonconformity","N/A":"N/A"}.get(old_result,old_result if old_result in ["Not Audited","Conforming","Nonconformity","N/A"] else "Not Audited"))
                comments.setPlainText(safe(saved["comments"])); evidence.setText(safe(saved["evidence_path"])); owner.setText(safe(saved["responsible"])); capa.setCurrentText("Yes" if ("capa_required" in saved.keys() and saved["capa_required"]) else "No")
                if "severity" in saved.keys(): severity.setCurrentText(safe(saved["severity"]) or "Medium")
                if "capa_reason" in saved.keys(): capa_reason.setText(safe(saved["capa_reason"]))
                try: due.setDate(QDate.fromString(safe(saved["target_date"]),"yyyy-MM-dd"))
                except Exception: pass
            checklist_table.setCellWidget(row,6,result); checklist_table.setCellWidget(row,7,comments); checklist_table.setCellWidget(row,8,evidence_host); checklist_table.setCellWidget(row,9,owner); checklist_table.setCellWidget(row,10,due); checklist_table.setCellWidget(row,11,severity); checklist_table.setCellWidget(row,12,capa); checklist_table.setCellWidget(row,13,capa_reason)
            for col in range(checklist_table.columnCount()):
                it=checklist_table.item(row,col)
                if it: it.setToolTip(it.text())
            result.currentTextChanged.connect(lambda *_:update_score())
        def load_checklist():
            ids=selected_schedule_ids(); checklist_table.setRowCount(0)
            if not ids: score.setText("Select one or more planned audits first."); return
            lib=self._audit_procedure_library()
            for sid in ids:
                sr=db.fetchone("SELECT * FROM audit_schedule WHERE id=?",(sid,))
                if not sr: continue
                proc=safe(sr["procedure_name"]); clause,items=lib.get(proc,("",[]))
                saved={r["item_no"]:r for r in db.fetchall("SELECT * FROM audit_checklist_responses WHERE schedule_id=? ORDER BY item_no",(sid,))}
                for n,item_text in enumerate(items,1): add_checklist_row(sid,sr,n,item_text,clause,saved.get(n))
            update_score()
        def update_score():
            yes=no=0
            for row in range(checklist_table.rowCount()):
                w=checklist_table.cellWidget(row,6)
                if not w: continue
                if w.currentText()=="Conforming": yes+=1
                elif w.currentText()=="Nonconformity": no+=1
            total=yes+no; pct=round(yes*100/total) if total else 0
            score.setText(f"Procedure compliance: {pct}%  |  Conforming: {yes}  |  Nonconformities: {no}  |  Not Audited / N/A excluded")
        def ensure_internal_audit(sid, sr):
            linked=sr["linked_audit_id"]; audit=db.fetchone("SELECT id,audit_type FROM audits WHERE id=?",(linked,)) if linked else None
            if audit and "internal audit" in safe(audit["audit_type"]).lower(): return linked
            audit_date=safe(sr["planned_date"]) or date.today().isoformat()
            number=self._next_audit_reference("Internal Audit", audit_date)
            db.execute("INSERT INTO audits(number,audit_date,audit_type,standard,project,location,department,auditor,lead_auditor,auditee,scope,objective,criteria,status,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(number,audit_date,"Internal Audit","ISO 45001","",safe(sr["area"] if "area" in sr.keys() else ""),"HSE",safe(sr["owner"]),safe(sr["owner"]),"",f"Procedure compliance audit: {safe(sr['procedure_name'])}","Determine conformity, adequacy and effectiveness of the controlled procedure and associated records","Approved procedure, ISO 45001 and applicable legal/client/company requirements","Draft",datetime.now().isoformat(timespec="seconds")))
            aid=db.fetchone("SELECT id FROM audits WHERE number=?",(number,))["id"]; db.execute("UPDATE audit_schedule SET linked_audit_id=?,updated_at=? WHERE id=?",(aid,datetime.now().isoformat(timespec="seconds"),sid)); return aid
        def save_checklist():
            ids=selected_schedule_ids()
            if not ids:return QMessageBox.information(d,"Select Audits","Select one or more planned audits from the tick-box dropdown first.")
            updated=[]
            for row in range(checklist_table.rowCount()):
                no_item=checklist_table.item(row,0); procno_item=checklist_table.item(row,1); criterion=checklist_table.item(row,3)
                if not no_item or not procno_item or not criterion: continue
                sid=no_item.data(Qt.ItemDataRole.UserRole); n=procno_item.data(Qt.ItemDataRole.UserRole)
                if sid is None or n is None: continue
                sid=int(sid); n=int(n); sr=db.fetchone("SELECT * FROM audit_schedule WHERE id=?",(sid,))
                if not sr: continue
                proc=safe(sr["procedure_name"]); clause,items=self._audit_procedure_library().get(proc,("",[])); item_text=criterion.text()
                result_widget=checklist_table.cellWidget(row,6); comments_widget=checklist_table.cellWidget(row,7); evidence_host=checklist_table.cellWidget(row,8); owner_edit=checklist_table.cellWidget(row,9); due_widget=checklist_table.cellWidget(row,10); severity_widget=checklist_table.cellWidget(row,11); capa_widget=checklist_table.cellWidget(row,12); capa_reason_widget=checklist_table.cellWidget(row,13)
                ui_result=result_widget.currentText() if result_widget else "Not Audited"; result={"Conforming":"Yes","Nonconformity":"No","N/A":"N/A"}.get(ui_result,"Not Audited")
                comments=comments_widget.toPlainText().strip() if comments_widget else ""; evidence=evidence_host.findChild(QLineEdit).text().strip() if evidence_host else ""
                responsible=owner_edit.text().strip() if isinstance(owner_edit,QLineEdit) else ""; target_date=due_widget.date().toString("yyyy-MM-dd") if isinstance(due_widget,QDateEdit) else ""
                capa_required=1 if capa_widget and capa_widget.currentText()=="Yes" else 0
                severity_value=severity_widget.currentText() if isinstance(severity_widget,QComboBox) else "Medium"
                capa_reason_value=capa_reason_widget.text().strip() if isinstance(capa_reason_widget,QLineEdit) else ""
                if ui_result=="Nonconformity" and (not comments or not evidence):
                    QMessageBox.warning(d,"Finding Evidence Required",f"For nonconformity in {proc}, enter objective evidence/gap and an evidence reference or attachment path before saving."); return
                if ui_result=="Nonconformity" and not responsible:
                    QMessageBox.warning(d,"Finding Owner Required",f"Assign a responsible person for the finding in {proc}."); return
                if ui_result=="Nonconformity" and not capa_required and not capa_reason_value:
                    QMessageBox.warning(d,"CAPA Justification Required",f"Explain why CAPA is not required for the nonconformity in {proc}."); return
                existing_response=db.fetchone("SELECT capa_number FROM audit_checklist_responses WHERE schedule_id=? AND item_no=?",(sid,n)); capa_number=safe(existing_response["capa_number"]) if existing_response else ""
                db.execute("INSERT INTO audit_checklist_responses(schedule_id,item_no,procedure_name,iso_clause,checklist_item,result,comments,evidence_path,updated_at,responsible,target_date,capa_required,capa_number,severity,capa_reason,finding_classification) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(schedule_id,item_no) DO UPDATE SET procedure_name=excluded.procedure_name,iso_clause=excluded.iso_clause,checklist_item=excluded.checklist_item,result=excluded.result,comments=excluded.comments,evidence_path=excluded.evidence_path,updated_at=excluded.updated_at,responsible=excluded.responsible,target_date=excluded.target_date,capa_required=excluded.capa_required,capa_number=CASE WHEN excluded.capa_number='' THEN audit_checklist_responses.capa_number ELSE excluded.capa_number END,severity=excluded.severity,capa_reason=excluded.capa_reason,finding_classification=excluded.finding_classification",(sid,n,proc,clause,item_text,result,comments,evidence,datetime.now().isoformat(timespec="seconds"),responsible,target_date,capa_required,capa_number,severity_value,capa_reason_value,"Minor Nonconformity"))
                aid=ensure_internal_audit(sid,sr); existing_f=db.fetchone("SELECT id,status,capa_number FROM audit_findings WHERE source_schedule_id=? AND source_item_no=?",(sid,n))
                if result=="No":
                    detail=f"Procedure compliance nonconformity: {item_text}. Objective evidence / gap: {comments}".strip(); corrective=f"Analyse root cause; define and implement corrective action against the approved procedure; assign an owner and due date; verify effectiveness. Evidence: {evidence}".strip()
                    if existing_f:
                        fid=existing_f["id"]; db.execute("UPDATE audit_findings SET audit_id=?,procedure_name=?,clause=?,finding_type='Minor Nonconformity',finding_detail=?,observation=?,evidence=?,corrective_action=?,responsible=?,target_date=?,capa_required=?,severity=?,capa_reason=?,status=CASE WHEN status='Closed' THEN 'Open' ELSE status END,updated_at=? WHERE id=?",(aid,proc,clause,detail,detail,evidence,corrective,responsible,target_date,capa_required,severity_value,capa_reason_value,datetime.now().isoformat(timespec="seconds"),fid))
                    else:
                        next_no=db.fetchone("SELECT COALESCE(MAX(finding_number),0)+1 AS n FROM audit_findings WHERE audit_id=?",(aid,))["n"]
                        db.execute("INSERT INTO audit_findings(audit_id,finding_number,finding_standard,clause,finding_type,finding_detail,observation,evidence,corrective_action,responsible,target_date,status,created_at,updated_at,source_schedule_id,source_item_no,procedure_name,capa_required,capa_number,severity,capa_reason,finding_reference) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(aid,next_no,"ISO 45001",clause,"Minor Nonconformity",detail,detail,evidence,corrective,responsible,target_date,"Open",datetime.now().isoformat(timespec="seconds"),datetime.now().isoformat(timespec="seconds"),sid,n,proc,capa_required,"",severity_value,capa_reason_value,f"AF-{datetime.now().year}-{aid:04d}-{next_no:03d}"))
                        fid=db.fetchone("SELECT id FROM audit_findings WHERE source_schedule_id=? AND source_item_no=? ORDER BY id DESC LIMIT 1",(sid,n))["id"]
                    db.execute("INSERT INTO audit_finding_history(finding_id,action,old_status,new_status,comments,evidence_path,user_name,timestamp) VALUES(?,?,?,?,?,?,?,?)",(fid,"Checklist finding saved",safe(existing_f["status"]) if existing_f else "","Open",capa_reason_value,evidence,safe(db.setting("current_user", "") or "System User"),datetime.now().isoformat(timespec="seconds")))
                    if capa_required:
                        fr=db.fetchone("SELECT capa_number FROM audit_findings WHERE id=?",(fid,)); capano=safe(fr["capa_number"]) if fr else ""
                        if not capano:
                            capano=next_number("HSE-CAPA","capa"); db.execute("INSERT INTO capa(number,source,reference_number,finding,root_cause,corrective_action,preventive_action,responsible,priority,target_date,status,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",(capano,"Audit",f"{no_item.text()} / {proc} / Finding {n}",detail,"",corrective,"",responsible,severity_value,target_date,"Open",datetime.now().isoformat(timespec="seconds")))
                            db.execute("UPDATE audit_findings SET capa_number=? WHERE id=?",(capano,fid)); db.execute("UPDATE audit_checklist_responses SET capa_number=? WHERE schedule_id=? AND item_no=?",(capano,sid,n))
                updated.append(sid)
            for sid in set(updated):db.execute("UPDATE audit_schedule SET status='Completed',updated_at=? WHERE id=?",(datetime.now().isoformat(timespec="seconds"),sid))
            update_score(); load_findings(); load_internal_register(); load_dashboard()
            QMessageBox.information(d,"Checklist Saved","Procedure compliance checklist saved. Nonconformities are listed in the Audit Findings Register; CAPA-required findings are linked to CAPA Tracking Register.")
        def checklist_html():
            rows=[]
            for row in range(checklist_table.rowCount()):
                cells=[safe(checklist_table.item(row,c).text() if checklist_table.item(row,c) else "") for c in range(6)]
                result=checklist_table.cellWidget(row,6).currentText() if checklist_table.cellWidget(row,6) else ""
                comments=checklist_table.cellWidget(row,7).toPlainText() if checklist_table.cellWidget(row,7) else ""; evidence=checklist_table.cellWidget(row,8).findChild(QLineEdit).text() if checklist_table.cellWidget(row,8) else ""
                owner_w=checklist_table.cellWidget(row,9); due_w=checklist_table.cellWidget(row,10); sev_w=checklist_table.cellWidget(row,11); capa_w=checklist_table.cellWidget(row,12); reason_w=checklist_table.cellWidget(row,13); owner=owner_w.text() if isinstance(owner_w,QLineEdit) else ""; due=due_w.date().toString("dd-MMM-yyyy") if isinstance(due_w,QDateEdit) else ""; severity=sev_w.currentText() if isinstance(sev_w,QComboBox) else ""; capa=capa_w.currentText() if isinstance(capa_w,QComboBox) else "No"; reason=reason_w.text() if isinstance(reason_w,QLineEdit) else ""
                rows.append("<tr>"+"".join(f"<td>{safe(v)}</td>" for v in cells)+f"<td>{safe(result)}</td><td>{safe(comments)}</td><td>{safe(evidence)}</td><td>{safe(owner)}</td><td>{safe(due)}</td><td>{safe(severity)}</td><td>{safe(capa)}</td><td>{safe(reason)}</td></tr>")
            logo=report_logo_path(); logo_html=f'<p><img src="{logo}" width="120"></p>' if logo and Path(logo).exists() else ""
            headers="<th>Audit No.</th><th>Procedure No.</th><th>Procedure / Date</th><th>Procedure Compliance Criteria</th><th>ISO Main Clause</th><th>Sub-clause(s)</th><th>Result</th><th>Objective Evidence / Gap</th><th>Evidence Attachment</th><th>Responsible Person</th><th>Target Date</th><th>Severity</th><th>CAPA Required</th><th>CAPA Not Required Justification</th>"
            return f'<html><head><style>body{{font-family:Arial;font-size:9pt}}h1{{color:#17365D}}table{{border-collapse:collapse;width:100%}}th{{background:#17365D;color:white}}td,th{{border:1px solid #aab7c4;padding:6px;vertical-align:top;white-space:pre-wrap}}</style></head><body>{logo_html}<h1>{safe(company_name())} — INTERNAL AUDIT CHECKLIST</h1><p>Procedure conformity, adequacy, implementation and objective evidence.</p><table><tr>{headers}</tr>{"".join(rows)}</table></body></html>'
        def print_checklist():
            if checklist_table.rowCount()==0:return QMessageBox.information(d,"No Checklist","Load one or more audit checklists first.")
            printer=QPrinter(QPrinter.PrinterMode.HighResolution); printer.setPageOrientation(QPrinter.Orientation.Landscape); dialog=QPrintDialog(printer,d)
            if dialog.exec()==QDialog.DialogCode.Accepted:
                doc=QTextDocument(); doc.setHtml(checklist_html()); doc.print_(printer)
        def export_checklist():
            if checklist_table.rowCount()==0:return QMessageBox.information(d,"No Checklist","Load one or more audit checklists first.")
            path,_=QFileDialog.getSaveFileName(d,"Export Internal Audit Checklist","Internal_Audit_Checklist.xlsx","Excel Workbook (*.xlsx)")
            if not path:return
            try:
                wb=Workbook(); ws=wb.active; ws.title="Procedure Audit"; ws.merge_cells("A1:N1"); ws["A1"]=f"{company_name()} — INTERNAL AUDIT / PROCEDURE COMPLIANCE CHECKLIST"; ws["A1"].font=Font(bold=True,size=15,color="FFFFFF"); ws["A1"].fill=PatternFill("solid",fgColor="17365D"); ws["A1"].alignment=Alignment(horizontal="center",vertical="center"); ws.append(["Audit Number","Procedure Number","Procedure / Planned Date","Procedure Compliance Audit Criteria","Primary ISO Clause","Other Applicable Clauses / Sub-clauses","Compliance Result","Objective Evidence / Gap","Evidence Attachment","Responsible Person / Action Owner","Target Date","Severity","CAPA Required","CAPA Not Required Justification"])
                for row in range(checklist_table.rowCount()):
                    vals=[checklist_table.item(row,c).text() if checklist_table.item(row,c) else "" for c in range(6)]; result=checklist_table.cellWidget(row,6).currentText() if checklist_table.cellWidget(row,6) else ""; comments=checklist_table.cellWidget(row,7).toPlainText() if checklist_table.cellWidget(row,7) else ""; evidence=checklist_table.cellWidget(row,8).findChild(QLineEdit).text() if checklist_table.cellWidget(row,8) else ""; owner_w=checklist_table.cellWidget(row,9); due_w=checklist_table.cellWidget(row,10); sev_w=checklist_table.cellWidget(row,11); capa_w=checklist_table.cellWidget(row,12); reason_w=checklist_table.cellWidget(row,13); owner=owner_w.text() if isinstance(owner_w,QLineEdit) else ""; due=due_w.date().toString("yyyy-MM-dd") if isinstance(due_w,QDateEdit) else ""; sev=sev_w.currentText() if isinstance(sev_w,QComboBox) else ""; capa=capa_w.currentText() if isinstance(capa_w,QComboBox) else "No"; reason=reason_w.text() if isinstance(reason_w,QLineEdit) else ""; ws.append(vals+[result,comments,evidence,owner,due,sev,capa,reason])
                for cell in ws[2]:cell.font=Font(bold=True,color="FFFFFF");cell.fill=PatternFill("solid",fgColor="4472C4");cell.alignment=Alignment(wrap_text=True,vertical="center")
                for row in ws.iter_rows(min_row=3):
                    for cell in row:cell.alignment=Alignment(wrap_text=True,vertical="top")
                for col,width in {"A":19,"B":17,"C":30,"D":68,"E":18,"F":22,"G":20,"H":48,"I":35,"J":28,"K":17,"L":14,"M":16,"N":35}.items():ws.column_dimensions[col].width=width
                ws.row_dimensions[1].height=32;ws.row_dimensions[2].height=42
                for rn in range(3,ws.max_row+1):ws.row_dimensions[rn].height=86
                try:
                    logo=report_logo_path()
                    if logo and Path(logo).exists():img=XLImage(str(logo));img.width=95;img.height=38;ws.add_image(img,"A1")
                except Exception:logging.exception("Audit checklist logo embedding failed")
                ws.freeze_panes="A3";ws.auto_filter.ref=f"A2:N{ws.max_row}";wb.save(path);QMessageBox.information(d,"Export Complete","Procedure compliance audit checklist exported successfully.")
            except Exception as exc:logging.exception("Audit checklist Excel export failed");QMessageBox.critical(d,"Export Failed",f"Unable to export checklist.\n\n{exc}")
        def export_checklist_word():
            if checklist_table.rowCount()==0:return QMessageBox.information(d,"No Checklist","Load one or more audit checklists first.")
            path,_=QFileDialog.getSaveFileName(d,"Export Internal Audit Checklist","Internal_Audit_Checklist.docx","Word Document (*.docx)")
            if not path:return
            try:
                doc=Document();sec=doc.sections[0];sec.orientation=WD_ORIENT.LANDSCAPE;sec.page_width,sec.page_height=sec.page_height,sec.page_width;sec.top_margin=Inches(.35);sec.bottom_margin=Inches(.35);sec.left_margin=Inches(.35);sec.right_margin=Inches(.35);hp=sec.header.paragraphs[0];hp.alignment=WD_ALIGN_PARAGRAPH.CENTER;logo=report_logo_path()
                if logo and Path(logo).exists():
                    try:hp.add_run().add_picture(str(logo),width=Inches(.9),height=Inches(.45))
                    except Exception:logging.exception("Audit checklist Word logo failed")
                p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=p.add_run(f"{company_name()}\nINTERNAL AUDIT — PROCEDURE COMPLIANCE CHECKLIST");r.bold=True;r.font.size=Pt(15)
                table=doc.add_table(rows=1,cols=14);table.alignment=WD_TABLE_ALIGNMENT.CENTER;table.autofit=False;headers=["Audit No.","Procedure No.","Procedure / Date","Audit Criteria","Primary ISO Clause","Other Applicable Clauses / Sub-clauses","Result","Objective Evidence / Gap","Evidence Attachment","Responsible Person","Target Date","Severity","CAPA Required","CAPA Not Required Justification"]
                for i,h in enumerate(headers):table.cell(0,i).text=h
                for row in range(checklist_table.rowCount()):
                    vals=[checklist_table.item(row,c).text() if checklist_table.item(row,c) else "" for c in range(6)];result=checklist_table.cellWidget(row,6).currentText() if checklist_table.cellWidget(row,6) else "";comments=checklist_table.cellWidget(row,7).toPlainText() if checklist_table.cellWidget(row,7) else "";evidence=checklist_table.cellWidget(row,8).findChild(QLineEdit).text() if checklist_table.cellWidget(row,8) else "";owner_w=checklist_table.cellWidget(row,9);due_w=checklist_table.cellWidget(row,10);sev_w=checklist_table.cellWidget(row,11);capa_w=checklist_table.cellWidget(row,12);reason_w=checklist_table.cellWidget(row,13);owner=owner_w.text() if isinstance(owner_w,QLineEdit) else "";due=due_w.date().toString("dd-MMM-yyyy") if isinstance(due_w,QDateEdit) else "";sev=sev_w.currentText() if isinstance(sev_w,QComboBox) else "";capa=capa_w.currentText() if isinstance(capa_w,QComboBox) else "No";reason=reason_w.text() if isinstance(reason_w,QLineEdit) else "";cells=table.add_row().cells
                    for i,v in enumerate(vals+[result,comments,evidence,owner,due,sev,capa,reason]):cells[i].text=safe(v);cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
                style_docx_table(table,header=True,font_size=6);widths=[.65,.65,1.1,1.65,.55,.6,.65,1.25,1.0,.9,.65,.55,.65,1.0]
                for row in table.rows:
                    for i,c in enumerate(row.cells):c.width=Inches(widths[i])
                doc.save(path);QMessageBox.information(d,"Export Complete","Word checklist exported successfully.")
            except Exception as exc:logging.exception("Audit checklist Word export failed");QMessageBox.critical(d,"Export Failed",f"Unable to export Word checklist.\n\n{exc}")
        load_btn.clicked.connect(load_checklist); save_btn.clicked.connect(save_checklist); export_checklist_btn.clicked.connect(export_checklist); export_word_btn.clicked.connect(export_checklist_word); print_checklist_btn.clicked.connect(print_checklist)
        load_checklist_schedule(); tabs.addTab(check,"Internal Audit Checklist")

        # ----- Findings register tab -----
        findings=QWidget(); fvl=QVBoxLayout(findings)
        fbar=QHBoxLayout(); fbar.addWidget(QLabel("Internal Audit Findings Register — each saved nonconformity is traced to its procedure and checklist criterion")); finding_status_filter=QComboBox(); finding_status_filter.addItems(["All Findings","Open Findings","Closed Findings"]); fbar.addWidget(finding_status_filter); close_selected_findings=QPushButton("Close Selected Findings"); edit_finding_btn=QPushButton("Edit Selected Finding"); delete_finding_btn=QPushButton("Delete Selected Finding"); refresh_findings=QPushButton("Refresh"); fbar.addWidget(close_selected_findings); fbar.addWidget(edit_finding_btn); fbar.addWidget(delete_finding_btn); fbar.addWidget(refresh_findings); fvl.addLayout(fbar)
        findings_table=QTableWidget(0,13); findings_table.setHorizontalHeaderLabels(["Audit Ref","Procedure","Finding #","Finding Type","Clause","Finding Detail / Reference","Corrective Action","Responsible","Target Date","Status","CAPA Reference","Audit ID","Finding ID"]); findings_table.setAlternatingRowColors(True); findings_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); findings_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection); findings_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); findings_table.setWordWrap(True); findings_table.setTextElideMode(Qt.TextElideMode.ElideNone); findings_table.verticalHeader().setDefaultSectionSize(70); findings_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive); findings_table.horizontalHeader().setStretchLastSection(True)
        for col,width in {0:125,1:155,2:65,3:135,4:75,5:245,6:220,7:125,8:100,9:120,10:115}.items(): findings_table.setColumnWidth(col,width)
        findings_table.setColumnHidden(11,True); findings_table.setColumnHidden(12,True); fvl.addWidget(findings_table,1)
        def load_findings():
            rows=db.fetchall("""SELECT f.*,a.number AS audit_ref,a.title AS audit_title,a.id AS parent_audit_id
                FROM audit_findings f JOIN audits a ON a.id=f.audit_id
                WHERE f.source_schedule_id IS NOT NULL OR lower(COALESCE(a.audit_type,'')) LIKE '%internal%'
                ORDER BY CASE WHEN lower(COALESCE(f.status,''))='closed' THEN 1 ELSE 0 END,f.target_date,a.id,f.finding_number""")
            if finding_status_filter.currentText()=="Open Findings": rows=[r for r in rows if safe(r["status"]).lower() not in ("closed","verified")]
            elif finding_status_filter.currentText()=="Closed Findings": rows=[r for r in rows if safe(r["status"]).lower() in ("closed","verified")]
            findings_table.setRowCount(len(rows))
            for i,r in enumerate(rows):
                vals=[r["audit_ref"],r["procedure_name"] if "procedure_name" in r.keys() and r["procedure_name"] else (r["audit_title"] or ""),r["finding_number"],r["finding_type"],r["clause"],r["finding_detail"] or r["observation"],r["corrective_action"],r["responsible"],r["target_date"],"OVERDUE" if overdue(safe(r["target_date"]),safe(r["status"])) else r["status"],safe(r["capa_number"] if "capa_number" in r.keys() else ""),r["parent_audit_id"],r["id"]]
                for c,v in enumerate(vals):
                    item=QTableWidgetItem(safe(v)); findings_table.setItem(i,c,item)
                    if vals[9]=="OVERDUE": item.setBackground(QColor("#FCE4D6"))
            findings_table.resizeRowsToContents()
        def open_finding_audit():
            rr=findings_table.currentRow()
            if rr<0:return
            aid=findings_table.item(rr,11).text()
            if aid:
                d.accept(); self.audit_workspace(int(aid),self.audits)
        def close_selected_internal_findings():
            ids=[]
            for idx in findings_table.selectionModel().selectedRows():
                item=findings_table.item(idx.row(),12)
                if item and item.text().isdigit(): ids.append(int(item.text()))
            if not ids:return QMessageBox.information(d,"Select Findings","Select one or more findings to close.")
            if QMessageBox.question(d,"Close Findings",f"Review and close {len(ids)} selected finding(s)? CAPA-required findings can only close after the linked CAPA is completed.")==QMessageBox.StandardButton.Yes:
                now=datetime.now().isoformat(timespec="seconds")
                for fid in ids:
                    fr=db.fetchone("SELECT * FROM audit_findings WHERE id=?",(fid,))
                    if not fr: continue
                    if fr["capa_required"]:
                        cr=db.fetchone("SELECT status FROM capa WHERE number=?",(safe(fr["capa_number"]),)) if safe(fr["capa_number"]) else None
                        if not cr or safe(cr["status"]).lower() not in ("closed","verified","complete","completed"):
                            QMessageBox.warning(d,"CAPA Not Closed",f"Finding {safe(fr['finding_reference']) or fid} requires CAPA. Close it only after the linked CAPA is completed in the existing CAPA Tracking Register.")
                            return
                    elif not safe(fr["capa_reason"]):
                        QMessageBox.warning(d,"Disposition Required",f"Finding {safe(fr['finding_reference']) or fid} has CAPA Required = No but no authorized justification. Edit the checklist disposition before closure.")
                        return
                    db.execute("UPDATE audit_findings SET status='Closed',verification_status='Verified',closure_date=?,verification='Verified from Internal Audit Findings Register',updated_at=? WHERE id=?",(date.today().isoformat(),now,fid))
                    db.execute("INSERT INTO audit_finding_history(finding_id,action,old_status,new_status,comments,user_name,timestamp) VALUES(?,?,?,?,?,?,?)",(fid,"Finding closed",safe(fr["status"]),"Closed","Closed after CAPA/disposition validation",safe(db.setting("current_user", "") or "System User"),now))
                load_findings(); load_internal_register(); load_dashboard()
        def edit_selected_finding():
            rr=findings_table.currentRow()
            if rr<0:return QMessageBox.information(d,"Select Finding","Select a finding to edit first.")
            fid=int(findings_table.item(rr,12).text()); fr=db.fetchone("SELECT * FROM audit_findings WHERE id=?",(fid,))
            if not fr:return
            form=QDialog(d); form.setWindowTitle("Edit Audit Finding"); form.resize(650,520); ff=QFormLayout(form)
            detail=QTextEdit(); detail.setPlainText(safe(fr["finding_detail"] or fr["observation"]))
            corrective=QTextEdit(); corrective.setPlainText(safe(fr["corrective_action"]))
            responsible=QLineEdit(safe(fr["responsible"])); due=QDateEdit(); due.setCalendarPopup(True); due.setDisplayFormat("dd-MMM-yyyy")
            try: due.setDate(QDate.fromString(safe(fr["target_date"]),"yyyy-MM-dd"))
            except Exception: due.setDate(QDate.currentDate().addDays(30))
            severity=QComboBox(); severity.addItems(["Low","Medium","High","Critical"]); severity.setCurrentText(safe(fr["severity"]) or "Medium")
            status_edit=QComboBox(); status_edit.addItems(["Open","In Progress","Pending Verification","Closed"]); status_edit.setCurrentText(safe(fr["status"]) if safe(fr["status"]) in ["Open","In Progress","Pending Verification","Closed"] else "Open")
            ff.addRow("Finding detail / evidence gap:",detail); ff.addRow("Corrective action:",corrective); ff.addRow("Responsible person:",responsible); ff.addRow("Target date:",due); ff.addRow("Severity:",severity); ff.addRow("Status:",status_edit)
            bb=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel); ff.addRow(bb); bb.accepted.connect(form.accept); bb.rejected.connect(form.reject)
            if form.exec()!=QDialog.DialogCode.Accepted:return
            if status_edit.currentText()=="Closed":
                linked_capa=db.fetchone("SELECT status FROM capa WHERE number=?",(safe(fr["capa_number"]),)) if safe(fr["capa_number"]) else None
                if fr["capa_required"] and (not linked_capa or safe(linked_capa["status"]).lower() not in ("closed","verified","complete","completed")):
                    return QMessageBox.warning(form,"CAPA Not Closed","Complete the linked CAPA in the existing CAPA Tracking Register before closing this finding.")
                if not fr["capa_required"] and not safe(fr["capa_reason"]):
                    return QMessageBox.warning(form,"Disposition Required","Record an authorized reason for not requiring CAPA before closing this finding.")
            now=datetime.now().isoformat(timespec="seconds")
            db.execute("UPDATE audit_findings SET finding_detail=?,observation=?,corrective_action=?,responsible=?,target_date=?,severity=?,status=?,updated_at=? WHERE id=?",(detail.toPlainText().strip(),detail.toPlainText().strip(),corrective.toPlainText().strip(),responsible.text().strip(),due.date().toString("yyyy-MM-dd"),severity.currentText(),status_edit.currentText(),now,fid))
            # Keep the existing linked CAPA record synchronized without creating another CAPA.
            if safe(fr["capa_number"]):
                db.execute("UPDATE capa SET finding=?,corrective_action=?,responsible=?,priority=?,target_date=? WHERE number=?",(detail.toPlainText().strip(),corrective.toPlainText().strip(),responsible.text().strip(),severity.currentText(),due.date().toString("yyyy-MM-dd"),safe(fr["capa_number"])))
            load_findings(); load_verification(); load_dashboard()
        def delete_selected_finding():
            ids=[]
            for idx in findings_table.selectionModel().selectedRows():
                it=findings_table.item(idx.row(),12)
                if it and it.text().isdigit():ids.append(int(it.text()))
            if not ids and findings_table.currentRow()>=0 and findings_table.item(findings_table.currentRow(),12):ids=[int(findings_table.item(findings_table.currentRow(),12).text())]
            if not ids:return QMessageBox.information(d,"Select Finding","Select one or more findings to delete first.")
            if QMessageBox.question(d,"Delete Audit Finding",f"Delete {len(ids)} selected audit finding(s)? Any linked CAPA records will remain in the existing CAPA Tracking Register.",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
            for fid in ids:
                db.execute("DELETE FROM audit_finding_history WHERE finding_id=?",(fid,)); db.execute("DELETE FROM audit_attachments WHERE finding_id=?",(fid,)); db.execute("DELETE FROM audit_findings WHERE id=?",(fid,))
            load_findings(); load_verification(); load_dashboard()
        close_selected_findings.clicked.connect(close_selected_internal_findings); edit_finding_btn.clicked.connect(edit_selected_finding); delete_finding_btn.clicked.connect(delete_selected_finding)
        refresh_findings.clicked.connect(load_findings); finding_status_filter.currentTextChanged.connect(load_findings); findings_table.cellDoubleClicked.connect(lambda *_:open_finding_audit())
        load_findings(); tabs.addTab(findings,"Internal Audit Findings")

        # ----- Internal audit register / master data / dashboard -----
        dash=QWidget(); dl=QVBoxLayout(dash); dl.addWidget(QLabel("INTERNAL AUDIT PERFORMANCE DASHBOARD"))
        dashboard_summary=QLabel(); dashboard_summary.setStyleSheet("font-size:16px;font-weight:bold;color:#17365D;padding:10px;background:#E8F0FA;border:1px solid #C7D7EA;border-radius:5px;"); dashboard_summary.setWordWrap(True); dl.addWidget(dashboard_summary)
        class AuditDashboardCharts(QWidget):
            def __init__(self,parent=None):
                super().__init__(parent); self.audit_counts=[]; self.finding_counts=[0,0]; self.setMinimumHeight(250); self.setMaximumHeight(300)
            def set_data(self,audit_counts,open_findings,closed_findings):
                self.audit_counts=list(audit_counts); self.finding_counts=[int(open_findings),int(closed_findings)]; self.update()
            def paintEvent(self,event):
                painter=QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing); painter.fillRect(self.rect(),QColor("#FFFFFF")); painter.setPen(QColor("#17365D")); painter.setFont(QFont("Arial",11,QFont.Weight.Bold)); painter.drawText(18,25,"AUDIT STATUS — BAR GRAPH"); painter.drawText(max(360,self.width()//2),25,"FINDINGS — STATUS DISTRIBUTION")
                chart_left=QRectF(18,42,max(280,self.width()//2-35),190); painter.setPen(QPen(QColor("#D5DEE8"),1)); painter.drawRect(chart_left)
                data=self.audit_counts or [("No data",0)]; maxv=max([x[1] for x in data]+[1]); usable=max(1,len(data)); bar_h=min(24,max(12,int((chart_left.height()-18)/usable)-5));
                for idx,(label,value) in enumerate(data[:6]):
                    y=chart_left.top()+12+idx*(bar_h+5); width=max(0,(chart_left.width()-130)*float(value)/maxv); painter.setPen(QColor("#34495E")); painter.setFont(QFont("Arial",8)); painter.drawText(int(chart_left.left()+6),int(y+bar_h-3),str(label)[:18]); painter.fillRect(QRectF(chart_left.left()+112,y,width,bar_h),QColor("#4472C4")); painter.setPen(QColor("#17365D")); painter.drawText(int(chart_left.left()+116+width),int(y+bar_h-3),str(value))
                pie_center_x=max(360,self.width()//2)+self.width()//4; pie_center_y=132; radius=76; total=sum(self.finding_counts); colors=[QColor("#4472C4"),QColor("#70AD47")]; start=0
                for idx,value in enumerate(self.finding_counts):
                    span=int(360*value/total) if total else (180 if idx==0 else 180); painter.setBrush(colors[idx]); painter.setPen(QPen(QColor("#FFFFFF"),2)); painter.drawPie(QRectF(pie_center_x-radius,pie_center_y-radius,radius*2,radius*2),start*16,span*16); start+=span
                painter.setPen(QColor("#34495E")); painter.setFont(QFont("Arial",9)); painter.drawText(pie_center_x-90,230,f"Open: {self.finding_counts[0]}     Closed: {self.finding_counts[1]}")
        dashboard_charts=AuditDashboardCharts(); dl.addWidget(dashboard_charts)
        dashboard_table=QTableWidget(0,5); dashboard_table.setHorizontalHeaderLabels(["Procedure","Audit Area","Risk Level","Planned Audits","Open Findings"]); dashboard_table.setAlternatingRowColors(True); dashboard_table.horizontalHeader().setStretchLastSection(True); dl.addWidget(dashboard_table,1)
        def load_dashboard():
            total=db.fetchone("SELECT COUNT(*) c FROM audit_schedule") ["c"]
            open_count=db.fetchone("SELECT COUNT(*) c FROM audit_findings f JOIN audits a ON a.id=f.audit_id WHERE lower(COALESCE(f.status,'open')) NOT IN ('closed','verified') AND (f.source_schedule_id IS NOT NULL OR lower(COALESCE(a.audit_type,'')) LIKE '%internal%')")["c"]
            closed_count=db.fetchone("SELECT COUNT(*) c FROM audit_findings f JOIN audits a ON a.id=f.audit_id WHERE lower(COALESCE(f.status,'')) IN ('closed','verified') AND (f.source_schedule_id IS NOT NULL OR lower(COALESCE(a.audit_type,'')) LIKE '%internal%')")["c"]
            dashboard_summary.setText(f"Scheduled procedure audits: {total}    |    Open findings: {open_count}    |    Closed findings: {closed_count}")
            audit_status_rows=db.fetchall("SELECT status,COUNT(*) AS n FROM audit_schedule GROUP BY status ORDER BY status")
            dashboard_charts.set_data([(safe(x["status"]) or "Unspecified",int(x["n"])) for x in audit_status_rows],open_count,closed_count)
            rows=db.fetchall("SELECT procedure_name,area,risk_level,COUNT(*) AS planned FROM audit_schedule GROUP BY procedure_name,area,risk_level ORDER BY risk_level,area,procedure_name")
            dashboard_table.setRowCount(len(rows))
            for i,r in enumerate(rows):
                op=db.fetchone("SELECT COUNT(*) c FROM audit_findings WHERE procedure_name=? AND lower(COALESCE(status,'open')) NOT IN ('closed','verified')",(r["procedure_name"],))["c"]
                for c,v in enumerate([r["procedure_name"],r["area"],r["risk_level"],r["planned"],op]): dashboard_table.setItem(i,c,QTableWidgetItem(safe(v)))
        load_dashboard(); tabs.insertTab(0,dash,"Audit Dashboard")

        reg=QWidget(); rvl=QVBoxLayout(reg); rbar=QHBoxLayout(); rbar.addWidget(QLabel("Internal Audit Register — select multiple rows to review the master tracking list")); rrefresh=QPushButton("Refresh"); rbar.addStretch(); rbar.addWidget(rrefresh); rvl.addLayout(rbar)
        reg_table=QTableWidget(0,10); reg_table.setHorizontalHeaderLabels(["Audit Ref","Procedure / Audit Title","Audit Area","Audit Date","Auditor","Open Findings","Closed Findings","Audit Status","Audit ID","Linked Schedule"]); reg_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); reg_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection); reg_table.setAlternatingRowColors(True); reg_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); reg_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents); reg_table.horizontalHeader().setStretchLastSection(True); reg_table.setColumnHidden(8,True); rvl.addWidget(reg_table,1)
        def load_internal_register():
            rows=db.fetchall("SELECT a.*,s.procedure_name AS scheduled_procedure,s.area AS scheduled_area,s.id AS schedule_ref FROM audits a LEFT JOIN audit_schedule s ON s.linked_audit_id=a.id WHERE lower(COALESCE(a.audit_type,'')) LIKE '%internal%' ORDER BY a.audit_date DESC,a.id DESC")
            reg_table.setRowCount(len(rows))
            for i,r in enumerate(rows):
                aid=r["id"]; op=db.fetchone("SELECT COUNT(*) c FROM audit_findings WHERE audit_id=? AND lower(COALESCE(status,'open')) NOT IN ('closed','verified')",(aid,))["c"]; closed=db.fetchone("SELECT COUNT(*) c FROM audit_findings WHERE audit_id=? AND lower(COALESCE(status,'')) IN ('closed','verified')",(aid,))["c"]
                vals=[r["number"],r["scheduled_procedure"] or r["title"] or r["scope"],r["scheduled_area"] or r["location"],r["audit_date"],r["auditor"] or r["lead_auditor"],op,closed,r["status"],aid,r["schedule_ref"] or ""]
                for c,v in enumerate(vals): reg_table.setItem(i,c,QTableWidgetItem(safe(v)))
        rrefresh.clicked.connect(load_internal_register); reg_table.cellDoubleClicked.connect(lambda *_: (d.accept(), self.audit_workspace(int(reg_table.item(reg_table.currentRow(),8).text()),self.audits)) if reg_table.currentRow()>=0 else None)
        load_internal_register()

        master=QWidget(); mvl=QVBoxLayout(master); mvl.addWidget(QLabel("AUDIT MASTER DATA — PROCEDURE LIBRARY, ISO CLAUSES AND RISK-BASED FREQUENCY"))
        master_table=QTableWidget(0,4); master_table.setHorizontalHeaderLabels(["Procedure / Audit Topic","Applicable ISO Clauses","Checklist Criteria Count","Default Risk Frequency"]); master_table.setAlternatingRowColors(True); master_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); master_table.horizontalHeader().setSectionResizeMode(0,QHeaderView.ResizeMode.Stretch); master_table.horizontalHeader().setStretchLastSection(True); mvl.addWidget(master_table,1)
        for proc,(clause,items) in self._audit_procedure_library().items():
            rr=master_table.rowCount(); master_table.insertRow(rr)
            risk="Every 6 months (High) / Annual (Medium) / Every 18 months (Low)"
            for c,v in enumerate([proc,clause,len(items),risk]): master_table.setItem(rr,c,QTableWidgetItem(safe(v)))
        # Audit Master Data is intentionally kept out of the user-facing tab set to reduce confusion.

        # Audit notifications are generated from the existing schedule and retain revisions.
        notifications=QWidget(); nvl=QVBoxLayout(notifications)
        nvl.addWidget(QLabel("AUDIT NOTIFICATIONS — generated from planned audits; each issued revision is retained."))
        nt=QTableWidget(0,7); nt.setHorizontalHeaderLabels(["Schedule ID","Procedure / Audit","Area / Site","Planned Date","Lead Auditor","Status","Latest Notification"])
        nt.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); nt.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection); nt.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        nt.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents); nt.horizontalHeader().setStretchLastSection(True); nvl.addWidget(nt,1)
        nbar=QHBoxLayout(); nrefresh=QPushButton("Refresh"); ngen=QPushButton("Generate Word Notification"); nbar.addWidget(nrefresh); nbar.addWidget(ngen); nbar.addStretch(); nvl.addLayout(nbar)
        def load_notifications():
            rows=db.fetchall("SELECT s.*, (SELECT document_path FROM audit_notifications n WHERE n.schedule_id=s.id ORDER BY n.id DESC LIMIT 1) AS latest FROM audit_schedule s ORDER BY s.planned_date,s.id")
            nt.setRowCount(len(rows))
            for i,r in enumerate(rows):
                vals=[r["id"],r["procedure_name"],r["area"],r["planned_date"],r["owner"],r["status"],Path(safe(r["latest"])).name if safe(r["latest"]) else "Not issued"]
                for c,v in enumerate(vals): nt.setItem(i,c,QTableWidgetItem(safe(v)))
        def generate_notification():
            rr=nt.currentRow()
            if rr<0: return QMessageBox.information(d,"Select Scheduled Audit","Select a scheduled audit first.")
            sid=int(nt.item(rr,0).text()); sr=db.fetchone("SELECT * FROM audit_schedule WHERE id=?",(sid,))
            if not sr: return
            if safe(sr["status"]).lower() in ("cancelled", "deferred"):
                return QMessageBox.warning(d,"Schedule Not Active","A notification can be generated once the audit schedule is planned. Re-plan this audit before issuing a notification.")
            form=QDialog(d); form.setWindowTitle("Audit Notification Details"); form.resize(560,340); ff=QFormLayout(form)
            recipients=QLineEdit(); recipients.setPlaceholderText("To recipients; separate multiple names/emails with semicolons")
            cc=QLineEdit(); cc.setPlaceholderText("CC recipients (optional)")
            scope=QTextEdit(); scope.setPlainText(f"Review conformity, adequacy, implementation and effectiveness of the controlled procedure: {safe(sr['procedure_name'])}. Review the current approved procedure, revision history, responsibilities, risk controls, competence, required records, monitoring and corrective-action closure."); scope.setMaximumHeight(100)
            ff.addRow("To:",recipients); ff.addRow("CC:",cc); ff.addRow("Scope / documents required:",scope)
            bb=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel); ff.addRow(bb); bb.accepted.connect(form.accept); bb.rejected.connect(form.reject)
            if form.exec()!=QDialog.DialogCode.Accepted:return
            path,_=QFileDialog.getSaveFileName(d,"Save Audit Notification",f"Audit_Notification_{sid}.docx","Word Document (*.docx)")
            if not path:return
            try:
                doc=Document(); sec=doc.sections[0]; sec.top_margin=Inches(.55); sec.bottom_margin=Inches(.55); sec.left_margin=Inches(.65); sec.right_margin=Inches(.65)
                hp=sec.header.paragraphs[0]; hp.alignment=WD_ALIGN_PARAGRAPH.CENTER; logo=report_logo_path()
                # Word embeds raster images reliably; skip unsupported logo formats rather than
                # creating a notification that Word may display with a broken-image marker.
                supported_logo_exts={".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}
                if logo and Path(logo).exists() and Path(logo).suffix.lower() in supported_logo_exts:
                    try: hp.add_run().add_picture(str(logo),width=Inches(1.0),height=Inches(.5))
                    except Exception: logging.exception("Audit notification logo could not be embedded; continuing without logo")
                p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; run=p.add_run(f"{company_name()}\nINTERNAL HSE AUDIT NOTIFICATION"); run.bold=True; run.font.size=Pt(16)
                linked=db.fetchone("SELECT number,title,audit_date FROM audits WHERE id=?",(sr["linked_audit_id"],)) if sr["linked_audit_id"] else None
                data=[("Document Number",f"{document_prefix()}-AUD-NOT-{sid:04d}"),("Audit Reference",safe(linked["number"]) if linked else f"IA-PLAN-{sid:04d}"),("Audit Title",safe(linked["title"]) if linked else safe(sr["procedure_name"])),("Procedure",safe(sr["procedure_name"])),("Area / Site",safe(sr["area"])),("Planned Date",safe(sr["planned_date"])),("Lead Auditor",safe(sr["owner"])),("Audit Status",safe(sr["status"])),("To",recipients.text().strip()),("CC",cc.text().strip())]
                table=doc.add_table(rows=0,cols=2); table.style="Table Grid"
                for k,v in data:
                    cells=table.add_row().cells; cells[0].text=xml_safe(k); cells[1].text=xml_safe(v)
                doc.add_heading("Audit Scope and Documents Required",level=2); doc.add_paragraph(xml_safe(scope.toPlainText()))
                doc.add_paragraph(xml_safe("Please make the current approved procedure and revision history, applicable legal/client/ISO requirements, relevant risk assessments, competence records, implementation records, monitoring results, previous findings and corrective-action evidence available. The audit will use controlled-document sampling, interviews and objective evidence to evaluate procedure compliance."))
                doc.add_paragraph("Prepared by: ____________________     Reviewed / Approved by: ____________________     Date: ______________")
                # Save atomically and validate the DOCX package before recording it as issued.
                if not str(path).lower().endswith(".docx"):
                    path = str(path) + ".docx"
                save_docx_validated(doc, path)
                with zipfile.ZipFile(path, "r") as generated_doc:
                    bad_member = generated_doc.testzip()
                    if bad_member:
                        raise ValueError(f"Word notification validation failed at {bad_member}")
                # Re-open with python-docx as a second validation that the saved document is readable.
                Document(str(path))
                revision=db.fetchone("SELECT COALESCE(MAX(revision),0)+1 AS n FROM audit_notifications WHERE schedule_id=?",(sid,))["n"]
                db.execute("INSERT INTO audit_notifications(schedule_id,document_path,recipients,revision,issued_at,issued_by) VALUES(?,?,?,?,?,?)",(sid,path,recipients.text().strip()+("; CC: "+cc.text().strip() if cc.text().strip() else ""),revision,datetime.now().isoformat(timespec="seconds"),safe(db.setting("current_user", "") or "System User")))
                if sr["linked_audit_id"]: self._audit_log(sr["linked_audit_id"],"Audit notification generated",f"Revision {revision}: {path}")
                load_notifications()
                open_now = QMessageBox.question(d,"Notification Generated",f"Word audit notification saved and validated. Revision: {revision}\n\n{path}\n\nOpen the notification now?",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.Yes)
                if open_now == QMessageBox.StandardButton.Yes:
                    if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(path).resolve()))):
                        QMessageBox.warning(d,"Open Word Document","The file was saved and validated, but Windows could not launch the default Word application. Open the file from the saved path shown above.")
            except Exception as exc: logging.exception("Audit notification generation failed"); QMessageBox.critical(d,"Notification Failed",str(exc))
        nrefresh.clicked.connect(load_notifications); ngen.clicked.connect(generate_notification); load_notifications(); tabs.addTab(notifications,"Audit Notifications")

        # Verification is driven by the existing Audit Findings Register and reads CAPA status
        # from the existing capa table; no parallel CAPA register is created.
        verify=QWidget(); vvl=QVBoxLayout(verify); vvl.addWidget(QLabel("VERIFICATION & CLOSEOUT — CAPA status is read from the existing CAPA Tracking Register data."))
        vt=QTableWidget(0,11); vt.setHorizontalHeaderLabels(["Finding Ref","Audit Ref","Procedure","Finding / Evidence","Responsible","Target Date","CAPA Ref","CAPA Status","Finding Status","Closeout Evidence","Approved By"] )
        vt.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); vt.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection); vt.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        vt.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents); vt.horizontalHeader().setStretchLastSection(True); vvl.addWidget(vt,1)
        vbar=QHBoxLayout(); vrefresh=QPushButton("Refresh Findings"); vreview=QPushButton("Review / Verify Selected Finding"); vadd_evidence=QPushButton("Add Closeout Evidence"); vsubmit=QPushButton("Submit Closeout for Approval"); vbar.addWidget(vrefresh); vbar.addWidget(vreview); vbar.addWidget(vadd_evidence); vbar.addWidget(vsubmit); vbar.addStretch(); vvl.addLayout(vbar)
        def save_closeout_status(fid,value):
            fr=db.fetchone("SELECT * FROM audit_findings WHERE id=?",(fid,))
            if not fr:return
            if value=="Closed":
                cr=db.fetchone("SELECT status FROM capa WHERE number=?",(safe(fr["capa_number"]),)) if safe(fr["capa_number"]) else None
                if fr["capa_required"] and (not cr or safe(cr["status"]).lower() not in ("closed","verified","complete","completed")):
                    QMessageBox.warning(d,"CAPA Not Closed","This finding requires its linked CAPA to be completed in the existing CAPA Tracking Register before closure."); load_verification(); return
                if not fr["capa_required"] and not safe(fr["capa_reason"]):
                    QMessageBox.warning(d,"Disposition Required","Record an authorized reason for not requiring CAPA before closure."); load_verification(); return
            now=datetime.now().isoformat(timespec="seconds")
            mapped={"Approved":"Pending Verification","Rejected":"Action in Progress","Closed":"Closed"}.get(value,value)
            db.execute("UPDATE audit_findings SET closeout_approval_status=?,status=?,verification_status=?,closure_date=?,updated_at=? WHERE id=?",(value,mapped,"Verified" if value=="Closed" else "Pending Verification",date.today().isoformat() if value=="Closed" else "",now,fid))
            load_findings(); load_internal_register(); load_dashboard()
        def save_closeout_approver(fid,value):
            if value=="Select Role": return
            db.execute("UPDATE audit_findings SET closeout_approver=?,updated_at=? WHERE id=?",(value,datetime.now().isoformat(timespec="seconds"),fid))
        def load_verification():
            rows=db.fetchall("SELECT f.*,a.number AS audit_ref FROM audit_findings f JOIN audits a ON a.id=f.audit_id WHERE f.source_schedule_id IS NOT NULL OR lower(COALESCE(a.audit_type,'')) LIKE '%internal%' ORDER BY CASE WHEN lower(COALESCE(f.status,'')) IN ('closed','verified') THEN 1 ELSE 0 END,f.target_date")
            vt.setRowCount(len(rows))
            for i,r in enumerate(rows):
                cr=db.fetchone("SELECT status FROM capa WHERE number=?",(safe(r["capa_number"]),)) if safe(r["capa_number"]) else None
                ref=safe(r["finding_reference"]) or f"AF-{int(r['id']):05d}"
                vals=[ref,r["audit_ref"],r["procedure_name"],r["finding_detail"] or r["observation"],r["responsible"],r["target_date"],r["capa_number"],safe(cr["status"]) if cr else ("Not Required" if not r["capa_required"] else "CAPA link missing")]
                for c,v in enumerate(vals): vt.setItem(i,c,QTableWidgetItem(safe(v)))
                status_cell=QComboBox(); status_cell.addItems(["Pending Verification","Action in Progress","Pending Approval","Approved","Rejected","Closed"]); current_status=safe(r["closeout_approval_status"]) if "closeout_approval_status" in r.keys() else "Pending Approval"
                if current_status not in [status_cell.itemText(j) for j in range(status_cell.count())]: current_status="Pending Approval"
                status_cell.setCurrentText(current_status); status_cell.currentTextChanged.connect(lambda value,fid=int(r["id"]): save_closeout_status(fid,value)); vt.setCellWidget(i,8,status_cell)
                evidence_btn=QPushButton("Add / View Evidence"); evidence_btn.setToolTip(safe(r["closeout_evidence"] if "closeout_evidence" in r.keys() else "")); evidence_btn.clicked.connect(lambda checked=False,fid=int(r["id"]): add_closeout_evidence_for_id(fid)); vt.setCellWidget(i,9,evidence_btn)
                approver=QComboBox(); approver.addItems(["Select Role","Lead Auditor","Supervisor","Manager"]); approver.setCurrentText(safe(r["closeout_approver"]) if "closeout_approver" in r.keys() and safe(r["closeout_approver"]) else "Select Role"); approver.currentTextChanged.connect(lambda value,fid=int(r["id"]): save_closeout_approver(fid,value)); vt.setCellWidget(i,10,approver)
                vt.item(i,0).setData(Qt.ItemDataRole.UserRole,int(r["id"]))
        def verify_selected_finding():
            rr=vt.currentRow()
            if rr<0:return QMessageBox.information(d,"Select Finding","Select a finding first.")
            fid=vt.item(rr,0).data(Qt.ItemDataRole.UserRole); fr=db.fetchone("SELECT * FROM audit_findings WHERE id=?",(fid,))
            if not fr:return
            form=QDialog(d); form.setWindowTitle("Finding Verification / Closeout"); form.resize(600,360); ff=QFormLayout(form)
            action=QComboBox(); action.addItems(["Pending Verification","Action in Progress","Pending Response","Closed","Return for Action"])
            action.setCurrentText(safe(fr["status"]) if safe(fr["status"]) in ["Pending Verification","Action in Progress","Pending Response","Closed","Return for Action"] else "Pending Verification")
            comments=QTextEdit(); comments.setPlaceholderText("Record verification comments, evidence reviewed, and reasons for rejection or return.")
            evidence=QLineEdit(); evidence.setPlaceholderText("Evidence file path / reference (optional)"); browse=QPushButton("Browse…")
            browse.clicked.connect(lambda: evidence.setText(QFileDialog.getOpenFileName(form,"Select Verification Evidence","","Evidence Files (*.*)")[0] or evidence.text()))
            evhost=QWidget(); eh=QHBoxLayout(evhost); eh.setContentsMargins(0,0,0,0); eh.addWidget(evidence); eh.addWidget(browse)
            ff.addRow("New finding status:",action); ff.addRow("Verification comments:",comments); ff.addRow("Evidence:",evhost)
            bb=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel); ff.addRow(bb); bb.accepted.connect(form.accept); bb.rejected.connect(form.reject)
            if form.exec()!=QDialog.DialogCode.Accepted:return
            new_status=action.currentText(); cr=db.fetchone("SELECT status FROM capa WHERE number=?",(safe(fr["capa_number"]),)) if safe(fr["capa_number"]) else None
            if new_status=="Closed" and fr["capa_required"] and (not cr or safe(cr["status"]).lower() not in ("closed","verified","complete","completed")):
                return QMessageBox.warning(d,"CAPA Not Closed","This finding requires CAPA. The linked record in the existing CAPA Tracking Register must reach its approved completed/closed status before the audit finding can be closed.")
            if new_status=="Closed" and not fr["capa_required"] and not safe(fr["capa_reason"]):
                return QMessageBox.warning(d,"Disposition Required","An authorized reason for not requiring CAPA must be recorded before this finding can be closed.")
            now=datetime.now().isoformat(timespec="seconds"); close_date=date.today().isoformat() if new_status=="Closed" else ""
            db.execute("UPDATE audit_findings SET status=?,verification=?,verification_status=?,closure_date=?,updated_at=? WHERE id=?",(new_status,comments.toPlainText().strip(),"Verified" if new_status=="Closed" else "Pending Verification",close_date,now,fid))
            db.execute("INSERT INTO audit_finding_history(finding_id,action,old_status,new_status,comments,evidence_path,user_name,timestamp) VALUES(?,?,?,?,?,?,?,?)",(fid,"Finding verification updated",safe(fr["status"]),new_status,comments.toPlainText().strip(),evidence.text().strip(),safe(db.setting("current_user", "") or "System User"),now))
            load_verification(); load_findings(); load_internal_register(); load_dashboard(); QMessageBox.information(d,"Verification Saved","Finding verification history was recorded.")
        def add_closeout_evidence_for_id(fid):
            path,_=QFileDialog.getOpenFileName(d,"Select Closeout Evidence","","Evidence Files (*.*)")
            if not path:return
            fr=db.fetchone("SELECT * FROM audit_findings WHERE id=?",(fid,))
            if not fr:return
            saved=self._audit_attachment_copy([path],safe(fr["finding_reference"]) or f"F{fid}",fr["audit_id"],fid,safe(db.setting("current_user","") or "System User"))
            db.execute("UPDATE audit_findings SET closeout_evidence=?,updated_at=? WHERE id=?",((safe(fr["closeout_evidence"]) + "\n" if "closeout_evidence" in fr.keys() and safe(fr["closeout_evidence"]) else "") + (saved[0] if saved else path),datetime.now().isoformat(timespec="seconds"),fid))
            load_verification()
        def add_closeout_evidence():
            rr=vt.currentRow()
            if rr<0:return QMessageBox.information(d,"Select Finding","Select a finding row first.")
            fid=vt.item(rr,0).data(Qt.ItemDataRole.UserRole); add_closeout_evidence_for_id(fid)
        def submit_closeout_for_approval():
            rr=vt.currentRow()
            if rr<0:return QMessageBox.information(d,"Select Finding","Select a finding row first.")
            fid=vt.item(rr,0).data(Qt.ItemDataRole.UserRole); fr=db.fetchone("SELECT * FROM audit_findings WHERE id=?",(fid,))
            if not fr:return
            if not safe(fr["closeout_approver"] if "closeout_approver" in fr.keys() else ""):
                return QMessageBox.warning(d,"Approver Required","Select Lead Auditor, Supervisor, or Manager in the Approved By column before submitting.")
            if not safe(fr["closeout_evidence"] if "closeout_evidence" in fr.keys() else ""):
                return QMessageBox.warning(d,"Evidence Required","Add closeout evidence before submitting this finding for approval.")
            if fr["capa_required"]:
                cr=db.fetchone("SELECT status FROM capa WHERE number=?",(safe(fr["capa_number"]),)) if safe(fr["capa_number"]) else None
                if not cr:return QMessageBox.warning(d,"CAPA Link Missing","This finding requires CAPA but has no linked record in the existing CAPA Tracking Register.")
                if safe(cr["status"]).lower() not in ("closed","verified","complete","completed"):
                    return QMessageBox.warning(d,"CAPA Not Closed","Complete the linked CAPA in the existing CAPA Tracking Register before submitting closeout.")
            elif not safe(fr["capa_reason"]):
                return QMessageBox.warning(d,"CAPA Justification Required","Enter an authorized reason for not requiring CAPA before submitting closeout.")
            now=datetime.now().isoformat(timespec="seconds")
            db.execute("UPDATE audit_findings SET closeout_approval_status='Pending Approval',status='Pending Approval',updated_at=? WHERE id=?",(now,fid))
            db.execute("INSERT INTO audit_finding_history(finding_id,action,old_status,new_status,comments,evidence_path,user_name,timestamp) VALUES(?,?,?,?,?,?,?,?)",(fid,"Closeout submitted for approval",safe(fr["status"]),"Pending Approval","Closeout submitted from Verification & Closeout.",safe(fr["closeout_evidence"] if "closeout_evidence" in fr.keys() else ""),safe(db.setting("current_user","") or "System User"),now))
            load_verification(); load_findings(); load_dashboard(); QMessageBox.information(d,"Closeout Submitted","The finding closeout has been submitted for approval.")
        vrefresh.clicked.connect(load_verification); vreview.clicked.connect(verify_selected_finding); vadd_evidence.clicked.connect(add_closeout_evidence); vsubmit.clicked.connect(submit_closeout_for_approval); load_verification(); tabs.addTab(verify,"Verification & Closeout")

        # Editable custom procedure records are Audit-only. New checklists are labelled by their
        # declared source/approval status and do not claim to be verified external requirements.
        proc_admin=QWidget(); pvl=QVBoxLayout(proc_admin); pvl.addWidget(QLabel("PROCEDURE LIBRARY — verify company/client source documents before approving checklist criteria."))
        pt=QTableWidget(0,8); pt.setHorizontalHeaderLabels(["Procedure ID","Title","Document No.","Revision","Clauses","Owner","Source / Approval Status","Active"])
        pt.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); pt.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection); pt.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        pt.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents); pt.horizontalHeader().setStretchLastSection(True); pvl.addWidget(pt,1)
        pbar=QHBoxLayout(); padd=QPushButton("Add / Edit Procedure"); pdel=QPushButton("Deactivate Selected"); pref=QPushButton("Refresh"); pbar.addWidget(padd); pbar.addWidget(pdel); pbar.addWidget(pref); pbar.addStretch(); pvl.addLayout(pbar)
        def load_proc_admin():
            rows=db.fetchall("SELECT * FROM audit_procedure_records ORDER BY title"); pt.setRowCount(len(rows))
            for i,r in enumerate(rows):
                vals=[r["procedure_id"],r["title"],r["document_number"],r["revision"],r["clause_references"],r["owner"],r["approval_status"],"Yes" if r["active"] else "No"]
                for c,v in enumerate(vals):pt.setItem(i,c,QTableWidgetItem(safe(v)))
                pt.item(i,0).setData(Qt.ItemDataRole.UserRole,int(r["id"]))
        def edit_proc_record():
            rr=pt.currentRow(); pid=pt.item(rr,0).data(Qt.ItemDataRole.UserRole) if rr>=0 and pt.item(rr,0) else None
            rec=db.fetchone("SELECT * FROM audit_procedure_records WHERE id=?",(pid,)) if pid else None
            form=QDialog(d); form.setWindowTitle("Audit Procedure Library Record"); form.resize(700,650); ff=QFormLayout(form)
            fields={k:QLineEdit() for k in ["procedure_id","title","document_number","issuing_organization","site_contract","revision","effective_date","review_date","clause_references","owner","applicability","source_file"]}
            for k,w in fields.items(): ff.addRow(k.replace("_"," ").title()+":",w)
            approval=QComboBox(); approval.addItems(["Draft checklist pending review","Company-controlled procedure","Official verified requirement","Approved"]); ff.addRow("Source / approval status:",approval)
            checklist=QTextEdit(); checklist.setPlaceholderText("One procedure-compliance audit criterion per line; include document control, responsibilities, controls, records, monitoring and corrective-action requirements."); ff.addRow("Checklist criteria (one per line):",checklist)
            if rec:
                for k,w in fields.items():w.setText(safe(rec[k]))
                approval.setCurrentText(safe(rec["approval_status"]))
                try:checklist.setPlainText("\n".join(json.loads(safe(rec["checklist_json"]) or "[]")))
                except Exception:pass
            bb=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel); ff.addRow(bb); bb.accepted.connect(form.accept); bb.rejected.connect(form.reject)
            if form.exec()!=QDialog.DialogCode.Accepted:return
            title=fields["title"].text().strip(); items=[x.strip() for x in checklist.toPlainText().splitlines() if x.strip()]
            if not title or not items:return QMessageBox.warning(d,"Required Fields","Enter a procedure title and at least one checklist criterion.")
            vals=[fields[k].text().strip() for k in fields]+[approval.currentText(),json.dumps(items,ensure_ascii=False),datetime.now().isoformat(timespec="seconds")]
            if rec:db.execute("UPDATE audit_procedure_records SET procedure_id=?,title=?,document_number=?,issuing_organization=?,site_contract=?,revision=?,effective_date=?,review_date=?,clause_references=?,owner=?,applicability=?,source_file=?,approval_status=?,checklist_json=?,updated_at=? WHERE id=?",tuple(vals)+(pid,))
            else:db.execute("INSERT INTO audit_procedure_records(procedure_id,title,document_number,issuing_organization,site_contract,revision,effective_date,review_date,clause_references,owner,applicability,source_file,approval_status,checklist_json,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",tuple(vals))
            load_proc_admin(); load_dashboard(); load_checklist_schedule()
        def deactivate_proc_record():
            rr=pt.currentRow()
            if rr<0:return QMessageBox.information(d,"Select Procedure","Select a custom procedure first.")
            pid=pt.item(rr,0).data(Qt.ItemDataRole.UserRole)
            if QMessageBox.question(d,"Deactivate Procedure","Deactivate this custom procedure? Existing audits and saved findings will be retained.")==QMessageBox.StandardButton.Yes:
                db.execute("UPDATE audit_procedure_records SET active=0,updated_at=? WHERE id=?",(datetime.now().isoformat(timespec="seconds"),pid));load_proc_admin()
        padd.clicked.connect(edit_proc_record); pdel.clicked.connect(deactivate_proc_record); pref.clicked.connect(load_proc_admin); load_proc_admin(); tabs.addTab(proc_admin,"Procedure Library")

        close=QPushButton("Close"); close.clicked.connect(d.accept); root.addWidget(close)
        d.exec()

    def _audit_register_detail(self, audit_id):
        a, fs = self._audit_report_data(audit_id)
        if not a:
            return
        d=QDialog(self); d.setWindowTitle(f"Audit Register — {safe(a['number'])}"); d.resize(1180,820)
        lay=QVBoxLayout(d)
        head=QLabel(f"AUDIT REGISTER — {safe(a['number'])}"); head.setStyleSheet("font-size:21px;font-weight:bold;color:#17365D;padding:6px;"); lay.addWidget(head)
        info=QTableWidget(0,2); info.setHorizontalHeaderLabels(["Audit Field","Value"]); info.horizontalHeader().setStretchLastSection(True); info.setAlternatingRowColors(True)
        for k,v in [("Audit Reference",a["number"]),("Audit Type",a["audit_type"]),("Audit Title",a["title"]),("Audit Date",a["audit_date"]),("Department",a["department"]),("Location",a["location"]),("Audit Standard",a["standard"]),("Audit Scope",a["scope"]),("Audit Criteria",a["criteria"]),("Status",a["status"]),("Auditor",a["auditor"]),("Reviewer",a["reviewer"]),("Approver",a["approver"])]:
            r=info.rowCount(); info.insertRow(r); info.setItem(r,0,QTableWidgetItem(safe(k))); info.setItem(r,1,QTableWidgetItem(safe(v)))
        info.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); lay.addWidget(info,1)
        if fs:
            lay.addWidget(QLabel("FINDINGS"))
            ft=QTableWidget(0,9); ft.setHorizontalHeaderLabels(["No.","Standard","Clause","Sub-Clause","Type","Detail","Responsible","Location","Status / Target"]); ft.setAlternatingRowColors(True); ft.horizontalHeader().setStretchLastSection(True)
            for f in fs:
                r=ft.rowCount(); ft.insertRow(r); vals=[f["finding_number"],f["finding_standard"] if "finding_standard" in f.keys() else a["standard"],f["clause"],f["sub_clause"],f["finding_type"],f["finding_detail"] or f["observation"],f["responsible"],f["location"] if "location" in f.keys() else "",("OVERDUE" if overdue(f["target_date"],f["status"]) else f["status"]) + (f" / {f['target_date']}" if f["target_date"] else "")]
                for c,v in enumerate(vals): ft.setItem(r,c,QTableWidgetItem(safe(v)))
            ft.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); lay.addWidget(ft,2)
        buttons=QHBoxLayout(); word=QPushButton("Generate Word"); excel=QPushButton("Generate Excel"); close=QPushButton("Close")
        buttons.addWidget(word); buttons.addWidget(excel); buttons.addStretch(); buttons.addWidget(close); lay.addLayout(buttons)
        word.clicked.connect(lambda:self._audit_export_word(audit_id)); excel.clicked.connect(lambda:self._audit_export_excel(audit_id)); close.clicked.connect(d.accept)
        d.exec()

    def audit_type_selector(self, refresh):
        d=QDialog(self); d.setWindowTitle("Select Audit Type"); d.resize(560,320)
        lay=QVBoxLayout(d); h=QLabel("SELECT AUDIT TYPE"); h.setStyleSheet("font-size:20px;font-weight:bold;color:#17365D;padding:10px;"); lay.addWidget(h)
        info=QLabel("Select the audit type before opening the Audit Workspace."); info.setWordWrap(True); lay.addWidget(info)
        combo=QComboBox(); combo.addItems(self._audit_types()); lay.addWidget(combo)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel); lay.addWidget(buttons)
        def continue_with_type():
            chosen=combo.currentText()
            d.accept()
            if "internal audit" in chosen.lower():
                self.audit_planning_center()
            else:
                self.audit_workspace(None,refresh,chosen)
        buttons.accepted.connect(continue_with_type)
        buttons.rejected.connect(d.reject); d.exec()

    def audit_workspace(self, audit_id=None, refresh=None, selected_type=None, allow_edit=False):
        existing=db.fetchone("SELECT * FROM audits WHERE id=?",(audit_id,)) if audit_id else None
        if existing and existing["status"]=="Submitted" and not allow_edit: readonly=True
        else: readonly=False
        d=QDialog(self); d.setWindowTitle("Audit Workspace")
        d.setWindowFlags(d.windowFlags() | Qt.WindowType.WindowMaximizeButtonHint)
        d.setSizeGripEnabled(True)
        # Keep the dialog within the usable desktop area. A large minimum size
        # can force Qt to extend the window beyond the screen on smaller displays.
        screen=d.screen() or QApplication.primaryScreen()
        if screen:
            available=screen.availableGeometry()
            d.setMinimumSize(min(640, available.width()), min(440, available.height()))
            d.setGeometry(available)
        else:
            d.setMinimumSize(640,440)
            d.resize(1100,760)
        outer=QVBoxLayout(d)
        outer.setContentsMargins(12,12,12,10)
        outer.setSpacing(8)
        banner=QHBoxLayout(); head=QLabel("AUDIT WORKSPACE"); head.setStyleSheet("font-size:22px;font-weight:bold;color:#17365D;"); banner.addWidget(head); banner.addStretch(); outer.addLayout(banner)
        tabs=QTabWidget(); tabs.setDocumentMode(True); tabs.setUsesScrollButtons(True); tabs.tabBar().setExpanding(False); outer.addWidget(tabs,1)
        tabs.setMinimumHeight(180)
        overview=QWidget(); of=QFormLayout(overview); of.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        audit_type=QComboBox(); audit_type.addItems(self._audit_types());
        title=QLineEdit(); audit_date=QDateEdit(); audit_date.setCalendarPopup(True); audit_date.setDisplayFormat("dd-MMM-yyyy"); audit_date.setDate(datetime.now().date())
        reference=QLineEdit(); reference.setPlaceholderText("Automatically generated from company, audit type and year"); reference.setReadOnly(True); department=QLineEdit(); department.setPlaceholderText("Enter department..."); location=QLineEdit(); location.setPlaceholderText("Enter audit location..."); standard=QComboBox(); standard.addItems(self._audit_standards()); scope=QTextEdit(); criteria=QTextEdit()
        of.addRow("Audit Type:",audit_type); of.addRow("Audit Reference:",reference); of.addRow("Audit Title:",title); of.addRow("Audit Date:",audit_date); of.addRow("Department:",department); of.addRow("Location:",location); of.addRow("Audit Standard:",standard); of.addRow("Audit Scope:",scope); of.addRow("Audit Criteria:",criteria)
        overview_scroll=QScrollArea(); overview_scroll.setWidgetResizable(True); overview_scroll.setFrameShape(QScrollArea.Shape.NoFrame); overview_scroll.setWidget(overview)
        tabs.addTab(overview_scroll,"Overview")

        findings_tab=QWidget(); fl=QVBoxLayout(findings_tab); finding_scroll=QScrollArea(); finding_scroll.setWidgetResizable(True); finding_host=QWidget(); finding_layout=QVBoxLayout(finding_host); finding_layout.setAlignment(Qt.AlignmentFlag.AlignTop); finding_scroll.setWidget(finding_host); fl.addWidget(finding_scroll,1); add_finding=QPushButton("+ ADD FINDING"); add_finding.setMinimumHeight(40); fl.addWidget(add_finding); tabs.addTab(findings_tab,"Findings")
        review=QWidget(); rf=QFormLayout(review); rf.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        def person_row(label):
            row=QHBoxLayout(); c=QLineEdit(); c.setPlaceholderText("Enter name..."); pos=QLineEdit(); pos.setPlaceholderText("Position..."); dt=QDateEdit(); dt.setCalendarPopup(True); dt.setDisplayFormat("dd-MMM-yyyy"); row.addWidget(c,2); row.addWidget(pos,1); row.addWidget(dt); box=QWidget(); box.setLayout(row); rf.addRow(label,box); return c,pos,dt
        auditor, auditor_pos, auditor_date=person_row("Auditor (Name / Position / Date):")
        reviewer, reviewer_pos, reviewer_date=person_row("Reviewed By:")
        approver, approver_pos, approver_date=person_row("Approved By:")
        review_scroll=QScrollArea(); review_scroll.setWidgetResizable(True); review_scroll.setFrameShape(QScrollArea.Shape.NoFrame); review_scroll.setWidget(review)
        tabs.addTab(review_scroll,"Review & Approval")
        att_tab=QWidget(); al=QVBoxLayout(att_tab); attach_list=QListWidget(); al.addWidget(QLabel("AUDIT-LEVEL ATTACHMENTS")); al.addWidget(attach_list,1); att_btns=QHBoxLayout(); add_att=QPushButton("+ ADD ATTACHMENT"); open_att=QPushButton("View / Open"); del_att=QPushButton("Delete Selected"); att_btns.addWidget(add_att); att_btns.addWidget(open_att); att_btns.addWidget(del_att); al.addLayout(att_btns); att_scroll=QScrollArea(); att_scroll.setWidgetResizable(True); att_scroll.setFrameShape(QScrollArea.Shape.NoFrame); att_scroll.setWidget(att_tab); tabs.addTab(att_scroll,"Attachments")
        hist_tab=QWidget(); hl=QVBoxLayout(hist_tab); history=QTableWidget(); history.setColumnCount(4); history.setHorizontalHeaderLabels(["Date","User","Action","Details"]); history.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents); history.horizontalHeader().setStretchLastSection(True); hl.addWidget(history); hist_scroll=QScrollArea(); hist_scroll.setWidgetResizable(True); hist_scroll.setFrameShape(QScrollArea.Shape.NoFrame); hist_scroll.setWidget(hist_tab); tabs.addTab(hist_scroll,"History")

        finding_widgets=[]
        def clause_map(): return self._audit_clause_map(standard.currentText())
        def make_finding(data=None):
            idx=len(finding_widgets)+1; box=QGroupBox(f"FINDING {idx}"); box.setStyleSheet("QGroupBox{font-size:15px;font-weight:bold;border:1px solid #B8C2CC;margin-top:10px;padding:10px;}QGroupBox::title{subcontrol-origin:margin;left:10px;padding:0 5px;}")
            form=QFormLayout(box); form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
            ft=QComboBox(); ft.addItems(self._audit_finding_types())
            fs=QComboBox(); fs.addItems(self._audit_standards())
            cl=QComboBox(); sc=QComboBox()
            detail=QTextEdit(); detail.setMinimumHeight(100); ca=QTextEdit(); ca.setMinimumHeight(90)
            rp=QLineEdit(); rp.setPlaceholderText("Enter responsible person...")
            loc=QLineEdit(); loc.setPlaceholderText("Enter finding location...")
            status=QComboBox(); status.addItems(self._audit_finding_statuses())
            td=QDateEdit(); td.setCalendarPopup(True); td.setDisplayFormat("dd-MMM-yyyy"); td.setDate(datetime.now().date())
            closeout=QLineEdit(); closeout.setPlaceholderText("Enter close-out evidence / verification...")
            form.addRow("Finding Type:",ft); form.addRow("Standard:",fs); form.addRow("Clause:",cl); form.addRow("Sub-Clause:",sc); form.addRow("Finding Detail:",detail); form.addRow("Corrective Action:",ca); form.addRow("Responsible Person:",rp); form.addRow("Location:",loc); form.addRow("Status:",status); form.addRow("Target Date:",td); form.addRow("Close Out Evidence:",closeout)
            evidence=QListWidget(); evidence.setMaximumHeight(110); eb=QHBoxLayout(); add=QPushButton("+ ADD ATTACHMENT"); open_a=QPushButton("View / Open"); rem=QPushButton("Delete Selected"); eb.addWidget(add); eb.addWidget(open_a); eb.addWidget(rem); ew=QWidget(); ee=QVBoxLayout(ew); ee.setContentsMargins(0,0,0,0); ee.addWidget(evidence); ee.addLayout(eb); form.addRow("Attachments:",ew)
            remove=QPushButton("Remove Finding"); form.addRow("",remove)
            def load_cl():
                cl.blockSignals(True); cl.clear(); cl.addItems(list(self._audit_clause_map(fs.currentText()).keys())); cl.blockSignals(False); load_sub()
            def load_sub():
                sc.clear(); sc.addItems(self._audit_clause_map(fs.currentText()).get(cl.currentText(),[]))
            fs.currentTextChanged.connect(load_cl); cl.currentTextChanged.connect(load_sub); load_cl()
            if data:
                fs.setCurrentText(safe(data["finding_standard"]) if "finding_standard" in data.keys() and safe(data["finding_standard"]) else safe(existing["standard"]) if existing else self._audit_standards()[0]); load_cl()
                ft.setCurrentText(safe(data["finding_type"])); cl.setCurrentText(safe(data["clause"])); load_sub(); sc.setCurrentText(safe(data["sub_clause"])); detail.setPlainText(safe(data["finding_detail"] or data["observation"])); ca.setPlainText(safe(data["corrective_action"])); status.setCurrentText("Overdue" if overdue(safe(data["target_date"]),safe(data["status"])) else safe(data["status"])); loc.setText(safe(data["location"]) if "location" in data.keys() else ""); closeout.setText(safe(data["close_out_evidence"]) if "close_out_evidence" in data.keys() else "")
                if data["target_date"]:
                    try: td.setDate(date.fromisoformat(data["target_date"]))
                    except Exception: pass
                rp.setText(safe(data.get("responsible", "")))
                if existing:
                    for ar in self._audit_attachment_rows(existing["id"], data["id"]):
                        fp=safe(ar["file_path"])
                        if fp: evidence.addItem(fp)
            finding_widgets.append({"box":box,"type":ft,"standard":fs,"clause":cl,"sub":sc,"detail":detail,"ca":ca,"person":rp,"location":loc,"status":status,"target":td,"closeout":closeout,"attachments":evidence,"add":add,"remove":remove,"data":data})
            closeout.textChanged.connect(lambda text, st=status: st.setCurrentText("Closed") if text.strip() else None)
            def choose():
                paths,_=QFileDialog.getOpenFileNames(d,"Select Finding Evidence","","Evidence Files (*.png *.jpg *.jpeg *.bmp *.gif *.pdf *.doc *.docx *.xls *.xlsx *.txt *.mp4 *.avi *.mov);;All Files (*)")
                for p in paths: evidence.addItem(p)
            add.clicked.connect(choose)
            open_a.clicked.connect(lambda: self._audit_open_attachment(evidence))
            rem.clicked.connect(lambda: evidence.takeItem(evidence.currentRow()) if evidence.currentRow()>=0 else None)
            remove.clicked.connect(lambda: remove_finding(box))
            finding_layout.addWidget(box)
            box.setEnabled(not readonly)
        def remove_finding(box):
            if readonly:return
            for i,x in enumerate(finding_widgets):
                if x["box"] is box:
                    box.setParent(None); box.deleteLater(); finding_widgets.pop(i); break
            for n,x in enumerate(finding_widgets,1): x["box"].setTitle(f"FINDING {n}")
        add_finding.clicked.connect(lambda:make_finding())
        if selected_type:
            idx = audit_type.findText(selected_type)
            if idx >= 0:
                audit_type.setCurrentIndex(idx)
        def update_reference_preview(*_):
            if not existing:
                reference.setText(self._next_audit_reference(audit_type.currentText(), audit_date.date().toString("yyyy-MM-dd")))
        audit_type.currentTextChanged.connect(update_reference_preview)
        audit_date.dateChanged.connect(update_reference_preview)

        def load_audit():
            nonlocal existing
            if not existing:return
            audit_type.setCurrentText(safe(existing["audit_type"])); reference.setText(safe(existing["number"])); title.setText(safe(existing["title"])); department.setText(safe(existing["department"])); location.setText(safe(existing["location"])); standard.setCurrentText(safe(existing["standard"])); scope.setPlainText(safe(existing["scope"])); criteria.setPlainText(safe(existing["criteria"]))
            try:audit_date.setDate(date.fromisoformat(safe(existing["audit_date"])))
            except Exception:pass
            auditor.setText(safe(existing["auditor"])); reviewer.setText(safe(existing["reviewer"])); approver.setText(safe(existing["approver"]))
            auditor_pos.setText(safe(existing["auditor_position"])); reviewer_pos.setText(safe(existing["reviewer_position"])); approver_pos.setText(safe(existing["approver_position"]))
            for widget,key in [(auditor_date,"auditor_date"),(reviewer_date,"reviewer_date"),(approver_date,"approver_date")]:
                try: widget.setDate(date.fromisoformat(safe(existing[key])))
                except Exception: pass
            for r in db.fetchall("SELECT * FROM audit_findings WHERE audit_id=? ORDER BY finding_number,id",(existing["id"],)): make_finding(r)
            if not finding_widgets: make_finding()
            load_attachments(); load_history()
        def load_attachments():
            attach_list.clear()
            if not existing:return
            for a in self._audit_attachment_rows(existing["id"]):
                item=QListWidgetItem(Path(safe(a["file_path"])).name); item.setData(Qt.ItemDataRole.UserRole,a["id"]); attach_list.addItem(item)
        def load_history():
            if not existing:return
            rows=db.fetchall("SELECT timestamp,user_name,action,details FROM audit_history WHERE audit_id=? ORDER BY id",(existing["id"],)); history.setRowCount(len(rows))
            for i,r in enumerate(rows):
                for c,v in enumerate([r["timestamp"],r["user_name"],r["action"],r["details"]]): history.setItem(i,c,QTableWidgetItem(safe(v)))

        add_att.clicked.connect(lambda:self._audit_add_general_attachment(existing,attach_list) if existing and not readonly else None)
        open_att.clicked.connect(lambda:self._audit_open_attachment(attach_list, audit_id=existing["id"] if existing else None))
        del_att.clicked.connect(lambda:self._audit_delete_general_attachment(existing,attach_list) if existing and not readonly else None)
        if existing: load_audit()
        else: make_finding()
        if existing and readonly:
            overview.setEnabled(False); review.setEnabled(False); add_finding.setEnabled(False); add_att.setEnabled(False); open_att.setEnabled(True); del_att.setEnabled(False)
            for x in finding_widgets:
                for key in ["type","standard","clause","sub","detail","ca","person","location","target","attachments","add","remove"]:
                    x[key].setEnabled(False)
                x["status"].setEnabled(True); x["closeout"].setEnabled(True)
            d.setWindowTitle(f"Audit View — {existing['number']} — SUBMITTED / STATUS & CLOSE-OUT ONLY")
        else:
            if not existing:
                update_reference_preview()
            if refresh is None: refresh=lambda:None

        buttons=QHBoxLayout(); buttons.setSpacing(8)
        save=QPushButton("SAVE & EDIT LATER"); submit=QPushButton("SAVE & SUBMIT"); close=QPushButton("Close")
        for b in [save,submit,close]:
            b.setMinimumHeight(42)
            buttons.addWidget(b)
        buttons.addStretch()
        outer.addLayout(buttons)
        if readonly: save.setEnabled(True); submit.setEnabled(False)
        def validate_submission():
            if not title.text().strip(): return "Audit Title is required."
            if not audit_type.currentText(): return "Audit Type is required."
            if not standard.currentText(): return "Audit Standard is required."
            if not finding_widgets: return "At least one finding is required."
            for i,x in enumerate(finding_widgets,1):
                required=[(x["type"].currentText(),"Finding Type"),(x["standard"].currentText(),"Standard"),(x["clause"].currentText(),"Clause"),(x["sub"].currentText(),"Sub-Clause"),(x["detail"].toPlainText().strip(),"Finding Detail"),(x["ca"].toPlainText().strip(),"Corrective Action"),(x["person"].text().strip(),"Responsible Person"),(x["location"].text().strip(),"Location")]
                for val,label in required:
                    if not val or val.startswith("--") or val.startswith("No active"): return f"Finding {i}: {label} is required."
                if not x["target"].date(): return f"Finding {i}: Target Date is required."
            if not auditor.text().strip() or not reviewer.text().strip() or not approver.text().strip(): return "Auditor, Reviewer and Approver are required."
            return ""
        def save_submitted_status_closeout():
            nonlocal existing
            if not existing:
                return
            try:
                now=datetime.now().isoformat(timespec="seconds")
                for x in finding_widgets:
                    closeout_value=x["closeout"].text().strip()
                    status_value=x["status"].currentText().strip()
                    if closeout_value:
                        status_value="Closed"
                        x["status"].setCurrentText("Closed")
                    if status_value=="Closed" and not closeout_value:
                        QMessageBox.warning(d,"Close-out Evidence Required","A finding can only be Closed after Close Out Evidence is entered.")
                        return
                    fid=x["data"]["id"] if x.get("data") and "id" in x["data"].keys() else None
                    if fid:
                        db.execute("UPDATE audit_findings SET status=?, close_out_evidence=?, updated_at=? WHERE id=? AND audit_id=?",(status_value,closeout_value,now,fid,existing["id"]))
                        self._audit_log(existing["id"],"Submitted audit finding updated",f"Finding {x['data']['finding_number']}: Status={status_value}; Close Out Evidence updated")
                db.execute("UPDATE audits SET updated_at=? WHERE id=?",(now,existing["id"]))
                if refresh: refresh()
                QMessageBox.information(d,"Saved","Status and close-out evidence were saved.")
            except Exception as e:
                logging.exception("Audit submitted update failed"); QMessageBox.critical(d,"Save Error",f"Unable to save status/close-out evidence.\n\n{e}")

        def save_audit(final=False):
            nonlocal existing
            if readonly:
                save_submitted_status_closeout()
                return
            if final:
                msg=validate_submission()
                if msg: QMessageBox.warning(d,"Validation",msg); return
                if QMessageBox.question(d,"Confirm Submission","Once this audit is submitted, only finding Status and Close Out Evidence can be changed.")!=QMessageBox.StandardButton.Yes: return
            try:
                now=datetime.now().isoformat(timespec="seconds")
                ref_value = safe(existing["number"]) if existing else self._next_audit_reference(audit_type.currentText(), audit_date.date().toString("yyyy-MM-dd"))
                reference.setText(ref_value)
                if not ref_value:
                    QMessageBox.warning(d,"Reference Error","Unable to generate an Audit Reference. Check the company name in Settings."); return
                duplicate=db.fetchone("SELECT id FROM audits WHERE number=? AND id<>?",(ref_value, existing["id"] if existing else -1))
                if duplicate:
                    QMessageBox.warning(d,"Duplicate Audit Reference","This Audit Reference already exists. Please enter a unique reference."); return
                vals=(audit_date.date().toString("yyyy-MM-dd"),audit_type.currentText(),title.text().strip(),standard.currentText(),department.text().strip(),location.text().strip(),scope.toPlainText(),criteria.toPlainText(),auditor.text().strip(),auditor_pos.text(),auditor_date.date().toString("yyyy-MM-dd"),reviewer.text().strip(),reviewer_pos.text(),reviewer_date.date().toString("yyyy-MM-dd"),approver.text().strip(),approver_pos.text(),approver_date.date().toString("yyyy-MM-dd"),"Submitted" if final else "Draft",now,now if final else "")
                if not existing:
                    db.execute("INSERT INTO audits(number,audit_date,audit_type,title,standard,department,location,scope,criteria,auditor,auditor_position,auditor_date,reviewer,reviewer_position,reviewer_date,approver,approver_position,approver_date,status,created_by,created_at,updated_at,submitted_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                               (reference.text(),)+vals[:18]+(db.setting("current_user","") or "System User",now,now,vals[-1]))
                    existing=db.fetchone("SELECT * FROM audits WHERE number=?",(reference.text(),)); self._audit_log(existing["id"],"Audit created", "Draft" if not final else "Submitted")
                else:
                    db.execute("UPDATE audits SET number=?,audit_date=?,audit_type=?,title=?,standard=?,department=?,location=?,scope=?,criteria=?,auditor=?,auditor_position=?,auditor_date=?,reviewer=?,reviewer_position=?,reviewer_date=?,approver=?,approver_position=?,approver_date=?,status=?,updated_at=?,submitted_at=? WHERE id=?", (ref_value,)+vals+(existing["id"],))
                    self._audit_log(existing["id"],"Audit submitted" if final else "Audit saved", "Status: Submitted" if final else "Status: Draft")
                    # Preserve existing findings. They are updated individually below so linked CAPA
                    # records, checklist-originated findings, evidence and audit history are not orphaned.
                # If newly created, there are no finding rows yet; write current finding state.
                # Preserve the currently listed finding attachments while rebuilding the
                # finding rows. Removing an attachment from the UI therefore removes it
                # from the saved audit on the next draft save.
                preserved_attachments=[]
                if existing:
                    for n,x in enumerate(finding_widgets,1):
                        kept=[]
                        for i in range(x["attachments"].count()):
                            candidate=x["attachments"].item(i).text()
                            if Path(candidate).exists(): kept.append(candidate)
                        preserved_attachments.append((n,kept))
                prior_manual_findings = db.fetchall("SELECT id,status,capa_number FROM audit_findings WHERE audit_id=? AND source_schedule_id IS NULL",(existing["id"],)) if existing else []
                visible_existing_ids = {int(x["data"]["id"]) for x in finding_widgets if x.get("data") is not None and "id" in x["data"].keys()}
                for n,x in enumerate(finding_widgets,1):
                    pid=None
                    prior=x.get("data")
                    fid=int(prior["id"]) if prior is not None and "id" in prior.keys() else None
                    prior_row=db.fetchone("SELECT source_schedule_id FROM audit_findings WHERE id=? AND audit_id=?",(fid,existing["id"])) if fid else None
                    # Checklist-originated findings are managed by the checklist workflow. Do not
                    # overwrite or recreate them from the general Audit Workspace.
                    if prior_row and prior_row["source_schedule_id"] is not None:
                        continue
                    values=(n,x["standard"].currentText(),x["clause"].currentText(),x["sub"].currentText(),x["type"].currentText(),x["detail"].toPlainText(),x["detail"].toPlainText(),x["ca"].toPlainText(),x["person"].text().strip(),pid,x["location"].text().strip(),x["target"].date().toString("yyyy-MM-dd"),x["status"].currentText(),x["closeout"].text().strip(),now)
                    if fid:
                        db.execute("UPDATE audit_findings SET finding_number=?,finding_standard=?,clause=?,sub_clause=?,finding_type=?,finding_detail=?,observation=?,corrective_action=?,responsible=?,responsible_person_id=?,location=?,target_date=?,status=?,close_out_evidence=?,updated_at=? WHERE id=? AND audit_id=?",values+(fid,existing["id"]))
                        db.execute("DELETE FROM audit_attachments WHERE audit_id=? AND finding_id=?",(existing["id"],fid))
                    else:
                        db.execute("INSERT INTO audit_findings(audit_id,finding_number,finding_standard,clause,sub_clause,finding_type,finding_detail,observation,corrective_action,responsible,responsible_person_id,location,target_date,status,close_out_evidence,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(existing["id"],)+values)
                        fid=db.fetchone("SELECT id FROM audit_findings WHERE audit_id=? ORDER BY id DESC LIMIT 1",(existing["id"],))["id"]
                    paths=[x["attachments"].item(i).text() for i in range(x["attachments"].count()) if Path(x["attachments"].item(i).text()).exists()]
                    for fp in paths:
                        try:
                            # Existing files in the application attachment store are retained
                            # as-is; newly selected external files are copied into the store.
                            if str(Path(fp).resolve()).startswith(str(ATTACH_DIR.resolve())):
                                db.execute("INSERT INTO audit_attachments(audit_id,finding_id,file_path,attachment_type,uploaded_by,uploaded_at) VALUES(?,?,?,?,?,?)",
                                           (existing["id"],fid,fp,"Evidence",db.setting("current_user","") or "System User",now))
                            else:
                                self._audit_attachment_copy([fp],reference.text(),existing["id"],fid,db.setting("current_user","") or "System User")
                        except Exception:
                            logging.exception("Unable to preserve audit finding attachment")
                # A finding removed from the editor is never physically deleted. Keep its history
                # and CAPA relationship, and flag it for controlled review/cancellation instead.
                for prior_finding in prior_manual_findings:
                    if int(prior_finding["id"]) not in visible_existing_ids:
                        old_status=safe(prior_finding["status"])
                        db.execute("UPDATE audit_findings SET status='Cancellation Pending Review',updated_at=? WHERE id=?",(now,prior_finding["id"]))
                        db.execute("INSERT INTO audit_finding_history(finding_id,action,old_status,new_status,comments,user_name,timestamp) VALUES(?,?,?,?,?,?,?)",(prior_finding["id"],"Cancellation requested from Audit Workspace",old_status,"Cancellation Pending Review","Finding retained to preserve audit history and any CAPA relationship; authorized review is required.",safe(db.setting("current_user", "") or "System User"),now))
                # General attachments already saved independently.
                existing=db.fetchone("SELECT * FROM audits WHERE id=?",(existing["id"],))
                if refresh: refresh()
                if final: QMessageBox.information(d,"Audit Submitted","Audit submitted successfully and is now read-only."); d.accept()
                else: QMessageBox.information(d,"Saved","Audit saved as Draft. You can edit it later.")
            except Exception as e:
                logging.exception("Audit save failed"); QMessageBox.critical(d,"Save Error",f"Unable to save audit.\n\n{e}")
        save.clicked.connect(lambda:save_audit(False)); submit.clicked.connect(lambda:save_audit(True)); close.clicked.connect(d.reject)
        d.exec()

    def _audit_add_general_attachment(self, audit, widget):
        paths,_=QFileDialog.getOpenFileNames(self,"Select Audit Attachment","","Documents and Evidence (*.pdf *.doc *.docx *.xls *.xlsx *.png *.jpg *.jpeg *.bmp *.txt);;All Files (*)")
        if paths:
            self._audit_attachment_copy(paths,audit["number"],audit["id"],None,db.setting("current_user","") or "System User"); self._audit_log(audit["id"],"Attachment added", "; ".join(Path(x).name for x in paths));
            for p in paths: widget.addItem(Path(p).name)

    def _audit_open_attachment(self, widget, audit_id=None):
        item=widget.currentItem()
        if not item:return
        path=item.text()
        if audit_id is not None:
            row=db.fetchone("SELECT file_path FROM audit_attachments WHERE id=? AND audit_id=?",(item.data(Qt.ItemDataRole.UserRole),audit_id))
            if row:path=safe(row["file_path"])
        fp=Path(path)
        if not fp.exists():
            QMessageBox.warning(self,"Attachment","The attachment file could not be found.")
            return
        try:
            if hasattr(os,"startfile"): os.startfile(str(fp))
            else:
                import subprocess
                subprocess.Popen(["xdg-open",str(fp)])
        except Exception as e:
            QMessageBox.warning(self,"Attachment","Unable to open the attachment.\n\n"+str(e))

    def _audit_delete_general_attachment(self,audit,widget):
        item=widget.currentItem()
        if not item:return
        aid=item.data(Qt.ItemDataRole.UserRole)
        row=db.fetchone("SELECT file_path FROM audit_attachments WHERE id=? AND audit_id=? AND (finding_id IS NULL OR finding_id=0)",(aid,audit["id"]))
        if row:
            try: Path(safe(row["file_path"])).unlink(missing_ok=True)
            except Exception: pass
            db.execute("DELETE FROM audit_attachments WHERE id=?",(aid,)); self._audit_log(audit["id"],"Attachment removed",item.text()); widget.takeItem(widget.row(item))

    def audit_settings(self):
        d=QDialog(self); d.setWindowTitle("Audit Settings"); d.resize(900,700); lay=QVBoxLayout(d)
        tabs=QTabWidget(); lay.addWidget(tabs)
        def list_editor(key, default, title):
            w=QWidget(); l=QVBoxLayout(w); l.addWidget(QLabel(title)); lw=QListWidget(); lw.addItems(self._audit_setting_list(key,default)); l.addWidget(lw)
            row=QHBoxLayout(); add=QLineEdit(); add.setPlaceholderText("Add value..."); ba=QPushButton("Add"); br=QPushButton("Remove Selected"); row.addWidget(add); row.addWidget(ba); row.addWidget(br); l.addLayout(row)
            ba.clicked.connect(lambda: (lw.addItem(add.text().strip()),add.clear()) if add.text().strip() else None); br.clicked.connect(lambda: lw.takeItem(lw.currentRow()) if lw.currentRow()>=0 else None); tabs.addTab(w,title); return lw
        types=list_editor("audit_types",AUDIT_TYPES,"Audit Types"); ftypes=list_editor("audit_finding_types",["Major Nonconformity","Minor Nonconformity","Observation","Opportunity for Improvement","Positive Finding"],"Finding Types"); statuses=list_editor("audit_finding_statuses",["Open","In Progress","Submitted for Verification","Verified","Closed","Overdue"],"Finding Statuses"); standards=list_editor("audit_standards",["ISO 9001","ISO 14001","ISO 45001","Other"],"Standards")
        clause_tab=QWidget(); cl=QVBoxLayout(clause_tab); cl.addWidget(QLabel("Configure Standards → Clauses → Sub-Clauses using JSON. Example: {\"ISO 45001\": {\"7\": [\"7.2 Competence\"]}}")); clause_edit=QTextEdit(); clause_edit.setPlainText(json.dumps(json.loads(db.setting("audit_clause_map","")) if db.setting("audit_clause_map","") else ISO_CLAUSES,ensure_ascii=False,indent=2)); cl.addWidget(clause_edit); tabs.addTab(clause_tab,"Clauses / Sub-Clauses")
        info=QLabel("Clause and Sub-Clause selections in each finding are dependent on the selected Audit Standard. Changes here affect only the Audit module."); info.setWordWrap(True); lay.addWidget(info)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel); lay.addWidget(buttons)
        def save():
            for key,lw in [("audit_types",types),("audit_finding_types",ftypes),("audit_finding_statuses",statuses),("audit_standards",standards)]: db.set_setting(key,json.dumps([lw.item(i).text() for i in range(lw.count())],ensure_ascii=False))
            try:
                clause_data=json.loads(clause_edit.toPlainText())
                if not isinstance(clause_data,dict): raise ValueError("Clause configuration must be a JSON object.")
                db.set_setting("audit_clause_map",json.dumps(clause_data,ensure_ascii=False))
            except Exception as e:
                QMessageBox.warning(d,"Invalid Clause Configuration",str(e)); return
            d.accept()
        buttons.accepted.connect(save); buttons.rejected.connect(d.reject); d.exec()

    def _audit_report_data(self,audit_id):
        a=db.fetchone("SELECT * FROM audits WHERE id=?",(audit_id,)); fs=db.fetchall("SELECT * FROM audit_findings WHERE audit_id=? ORDER BY finding_number,id",(audit_id,)); return a,fs

    def _audit_export_word(self,audit_id):
        a,fs=self._audit_report_data(audit_id)
        if not a:return
        path=self._export_path(f"{a['number']}_Audit_Report.docx","Save Audit Word Report","Word Documents (*.docx)")
        if not path:return
        try:
            doc=Document()
            sec=doc.sections[0]
            sec.orientation=WD_ORIENT.LANDSCAPE
            sec.page_width,sec.page_height=sec.page_height,sec.page_width
            sec.top_margin=Inches(0.35); sec.bottom_margin=Inches(0.35); sec.left_margin=Inches(0.35); sec.right_margin=Inches(0.35)
            # Professional header matching the existing report style.
            hp=sec.header.paragraphs[0]; hp.alignment=WD_ALIGN_PARAGRAPH.CENTER
            if report_logo_path() and Path(report_logo_path()).exists():
                try: hp.add_run().add_picture(report_logo_path(),width=Inches(1.05),height=Inches(0.58))
                except Exception: logging.exception("Unable to embed audit Word logo")
            p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            r=p.add_run(f"{company_name()}\nAUDIT REPORT"); r.bold=True; r.font.size=Pt(14)
            hdr=doc.add_table(rows=2,cols=6); hdr.alignment=WD_TABLE_ALIGNMENT.CENTER; hdr.autofit=False
            hvals=["Audit Reference",safe(a["number"]),"Audit Date",safe(a["audit_date"]),"Document No.",f"{document_prefix()}-AUDIT"]
            hvals2=["Audit Type",safe(a["audit_type"]),"Department",safe(a["department"]),"Location",safe(a["location"])]
            for i,v in enumerate(hvals): hdr.cell(0,i).text=v
            for i,v in enumerate(hvals2): hdr.cell(1,i).text=v
            for row in hdr.rows:
                for c in row.cells: c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            style_docx_table(hdr,header=True,font_size=8)
            for row in hdr.rows:
                for i,c in enumerate(row.cells): c.width=Inches(1.18 if i%2==0 else 2.05)
            doc.add_paragraph()
            info=doc.add_table(rows=2,cols=4); info.alignment=WD_TABLE_ALIGNMENT.CENTER; info.autofit=False
            vals=[("Audit Title",a["title"],"Standard",a["standard"]),("Scope",a["scope"],"Audit Criteria",a["criteria"])]
            for ri,rowvals in enumerate(vals):
                for ci,v in enumerate(rowvals): info.cell(ri,ci).text=safe(v)
            for row in info.rows:
                for c in row.cells: c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
            style_docx_table(info,header=False,font_size=8)
            for row in info.rows:
                row.cells[0].width=Inches(1.15); row.cells[1].width=Inches(4.55); row.cells[2].width=Inches(1.15); row.cells[3].width=Inches(4.55)

            p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            r=p.add_run("AUDIT FINDINGS"); r.bold=True; r.font.size=Pt(12)
            # One finding = one landscape row. Evidence is embedded in the same row/cell.
            cols=["Standard","Clause","Sub-Clause","Finding Type","Location","Finding Details","Corrective Action","Responsible Person","Evidence","Target Date","Close Out Evidence"]
            ft=doc.add_table(rows=1,cols=len(cols)); ft.alignment=WD_TABLE_ALIGNMENT.CENTER; ft.autofit=False
            for i,h in enumerate(cols): ft.cell(0,i).text=h
            style_docx_table(ft,header=True,font_size=7)
            widths=[0.70,0.61,0.98,0.80,0.80,1.54,1.54,0.98,1.08,0.80,1.16]
            for row in ft.rows:
                for i,c in enumerate(row.cells): c.width=Inches(widths[i]); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for f in fs:
                cells=ft.add_row().cells
                values=[safe(f["finding_standard"] if "finding_standard" in f.keys() else a["standard"]),safe(f["clause"]),safe(f["sub_clause"]),safe(f["finding_type"]),safe(f["location"] if "location" in f.keys() else ""),safe(f["finding_detail"] or f["observation"]),safe(f["corrective_action"]),safe(f["responsible"]),"",safe(f["target_date"]),safe(f["close_out_evidence"] if "close_out_evidence" in f.keys() else "")]
                for i,v in enumerate(values): cells[i].text=v
                # Evidence photo(s) stay inside the Evidence cell on this same finding row.
                evidence_cell=cells[8]
                evidence_rows=self._audit_attachment_rows(audit_id,f["id"])
                if not evidence_rows:
                    evidence_cell.text="No evidence attached"
                for ar in evidence_rows:
                    fp=Path(safe(ar["file_path"]))
                    if evidence_cell.paragraphs and not evidence_cell.paragraphs[0].text and not evidence_cell.paragraphs[0].runs:
                        pp=evidence_cell.paragraphs[0]
                    else:
                        pp=evidence_cell.add_paragraph()
                    pp.add_run(fp.name if fp.name else "Evidence attachment")
                    if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
                        try:
                            ip=evidence_cell.add_paragraph(); ip.alignment=WD_ALIGN_PARAGRAPH.CENTER
                            ip.add_run().add_picture(str(fp),width=Inches(0.85),height=Inches(0.75))
                        except Exception: logging.exception("Unable to embed audit finding evidence photo")
                for c in cells: c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
                trPr=cells[0]._tc.getparent().get_or_add_trPr()
                trHeight=OxmlElement("w:trHeight"); trHeight.set(qn("w:val"),str(2880)); trHeight.set(qn("w:hRule"),"atLeast"); trPr.append(trHeight)
                for i,c in enumerate(cells): c.width=Inches(widths[i])
            style_docx_table(ft,header=True,font_size=7)

            # Include audit-level evidence files as well as finding-level evidence above.
            general_evidence=self._audit_attachment_rows(audit_id)
            p=doc.add_paragraph(); r=p.add_run("AUDIT-LEVEL ATTACHMENTS / EVIDENCE"); r.bold=True; r.font.size=11
            if general_evidence:
                at=doc.add_table(rows=1,cols=4)
                for ci,h in enumerate(["Attachment / Evidence","Type","Uploaded By","Upload Date"]): at.cell(0,ci).text=h
                for ar in general_evidence:
                    fp=Path(safe(ar["file_path"]))
                    cells=at.add_row().cells
                    cells[0].text=fp.name or "Evidence attachment"
                    cells[1].text=safe(ar["attachment_type"] if "attachment_type" in ar.keys() else "Evidence")
                    cells[2].text=safe(ar["uploaded_by"] if "uploaded_by" in ar.keys() else "")
                    cells[3].text=safe(ar["uploaded_at"] if "uploaded_at" in ar.keys() else "")
                    if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
                        try:
                            pp=cells[0].add_paragraph(); pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
                            pp.add_run().add_picture(str(fp),width=Inches(0.8),height=Inches(0.65))
                        except Exception: logging.exception("Unable to embed audit-level evidence photo")
                style_docx_table(at,header=True,font_size=7)
            else:
                doc.add_paragraph("No audit-level attachments recorded.")

            doc.add_paragraph()
            p=doc.add_paragraph(); r=p.add_run("REVIEW / APPROVAL"); r.bold=True; r.font.size=11
            st=doc.add_table(rows=3,cols=4); st.alignment=WD_TABLE_ALIGNMENT.CENTER; st.autofit=False
            sign_rows=[("AUDITOR",a["auditor"],a["auditor_position"],a["auditor_date"]),("REVIEWED BY",a["reviewer"],a["reviewer_position"],a["reviewer_date"]),("APPROVED BY",a["approver"],a["approver_position"],a["approver_date"])]
            for ri,(role,name,pos,dt) in enumerate(sign_rows):
                vals=[role,f"Name: {safe(name)}",f"Position: {safe(pos)}",f"Date: {safe(dt)}\nSignature: __________________"]
                for ci,v in enumerate(vals): st.cell(ri,ci).text=v
            for row in st.rows:
                for i,c in enumerate(row.cells): c.width=Inches([1.35,3.3,3.0,3.5][i]); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            style_docx_table(st,header=False,font_size=8)
            footer=doc.add_paragraph(safe(db.setting("report_footer",""))); footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
            doc.save(path); self.show_export_success(path)
        except Exception as e:
            logging.exception("Audit Word export failed"); QMessageBox.critical(self,"Export Error",str(e))

    def _audit_export_excel(self,audit_id):
        a,fs=self._audit_report_data(audit_id)
        if not a:return
        path=self._export_path(f"{a['number']}_Audit_Report.xlsx","Save Audit Excel Report","Excel Files (*.xlsx)")
        if not path:return
        try:
            from openpyxl.worksheet.table import Table, TableStyleInfo
            from openpyxl.worksheet.page import PageMargins

            wb=Workbook()
            wb.properties.title=f"Audit Report - {safe(a['number'])}"
            wb.properties.subject="HSE Audit Report"
            wb.properties.creator=company_name()

            # ------------------------------------------------------------
            # Professional report header helper used on each worksheet.
            # ------------------------------------------------------------
            def report_header(ws,title,columns_end):
                # Keep the report header visually consistent even for sheets with
                # only 5-7 data columns; the actual table can still use fewer columns.
                layout_cols=max(columns_end,11)
                ws.sheet_view.showGridLines=False
                ws.merge_cells(start_row=1,start_column=1,end_row=3,end_column=3)
                ws.merge_cells(start_row=1,start_column=4,end_row=3,end_column=8)
                right_start=9
                ws.merge_cells(start_row=1,start_column=right_start,end_row=1,end_column=layout_cols)
                ws.merge_cells(start_row=2,start_column=right_start,end_row=2,end_column=layout_cols)
                ws.merge_cells(start_row=3,start_column=right_start,end_row=3,end_column=layout_cols)
                ws.cell(1,4).value=f"{title}\n{company_name()}"
                ws.cell(1,right_start).value=f"Audit No.: {safe(a['number'])}"
                ws.cell(2,right_start).value=f"Document No.: {document_prefix()}-AUDIT"
                ws.cell(3,right_start).value=f"Audit Date: {safe(a['audit_date'])}"
                logo_path=report_logo_path()
                if logo_path and Path(logo_path).exists():
                    try:
                        logo=XLImage(str(logo_path)); logo.width=125; logo.height=62; logo.anchor="A1"; ws.add_image(logo)
                    except Exception: logging.exception("Unable to embed audit Excel logo")
                thin=Side(style="thin",color="A6A6A6")
                header_fill=PatternFill("solid",fgColor="17365D")
                light_fill=PatternFill("solid",fgColor="D9E2F3")
                for r in range(1,4): ws.row_dimensions[r].height=24
                ws.row_dimensions[1].height=30
                for row in ws.iter_rows(min_row=1,max_row=3,min_col=1,max_col=layout_cols):
                    for c in row:
                        if isinstance(c,MergedCell): continue
                        c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
                        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
                ws.cell(1,4).font=copy(ws.cell(1,4).font); ws.cell(1,4).font=ws.cell(1,4).font.copy(bold=True,size=14)
                for ref in [f"{chr(64+right_start)}1",f"{chr(64+right_start)}2",f"{chr(64+right_start)}3"]:
                    if ref in ws:
                        ws[ref].fill=light_fill; ws[ref].font=ws[ref].font.copy(bold=True,size=10)
                ws.page_setup.orientation="landscape"
                ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
                ws.sheet_properties.pageSetUpPr.fitToPage=True
                ws.page_margins=PageMargins(left=0.25,right=0.25,top=0.45,bottom=0.45,header=0.2,footer=0.2)
                ws.oddFooter.center.text="Page &P of &N"
                ws.oddFooter.right.text=company_name()

            def style_data_sheet(ws,header_row,table_ref,name):
                ws.freeze_panes=f"A{header_row+1}"
                ws.auto_filter.ref=table_ref
                thin=Side(style="thin",color="D0D0D0")
                for c in ws[header_row]:
                    if isinstance(c,MergedCell): continue
                    c.font=c.font.copy(bold=True,color="FFFFFF")
                    c.fill=PatternFill("solid",fgColor="17365D")
                    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
                    c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
                for row in ws.iter_rows(min_row=header_row+1,max_row=ws.max_row,min_col=1,max_col=ws.max_column):
                    for c in row:
                        if isinstance(c,MergedCell): continue
                        c.alignment=Alignment(vertical="top",horizontal="left",wrap_text=True)
                        c.border=Border(bottom=thin)
                if ws.max_row >= header_row+1:
                    tab=Table(displayName=name,ref=table_ref)
                    tab.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2",showFirstColumn=False,showLastColumn=False,showRowStripes=True,showColumnStripes=False)
                    ws.add_table(tab)

            # ------------------------------------------------------------
            # 1. Executive / audit summary
            # ------------------------------------------------------------
            ws=wb.active; ws.title="Audit Summary"
            report_header(ws,"AUDIT REPORT",8)
            ws["A5"]="AUDIT SUMMARY"; ws.merge_cells("A5:H5")
            ws["A5"].font=ws["A5"].font.copy(bold=True,size=12,color="FFFFFF"); ws["A5"].fill=PatternFill("solid",fgColor="17365D"); ws["A5"].alignment=Alignment(horizontal="left",vertical="center")
            ws.append(["Field","Value","Field","Value","Field","Value","Field","Value"])
            summary_pairs=[
                ("Audit Reference",a["number"]),("Audit Type",a["audit_type"]),("Audit Title",a["title"]),("Audit Date",a["audit_date"]),
                ("Department",a["department"]),("Location",a["location"]),("Standard",a["standard"]),("Scope",a["scope"]),
                ("Audit Criteria",a["criteria"]),("Auditor",a["auditor"]),("Reviewer",a["reviewer"]),("Approver",a["approver"]),
                ("Status",a["status"]),("Total Findings",len(fs))]
            # Keep the summary compact but readable: two field/value pairs per row.
            for i in range(0,len(summary_pairs),2):
                p1=summary_pairs[i]; p2=summary_pairs[i+1] if i+1<len(summary_pairs) else ("","")
                ws.append([safe(p1[0]),safe(p1[1]),safe(p2[0]),safe(p2[1]),"","","",""])
            for r in range(6,ws.max_row+1):
                for c in [1,3]:
                    ws.cell(r,c).font=ws.cell(r,c).font.copy(bold=True)
                    ws.cell(r,c).fill=PatternFill("solid",fgColor="EAF0F7")
                for c in range(1,9): ws.cell(r,c).alignment=Alignment(vertical="top",wrap_text=True)
            for col,w in {"A":22,"B":38,"C":22,"D":38,"E":4,"F":4,"G":4,"H":4}.items(): ws.column_dimensions[col].width=w
            ws.row_dimensions[6].height=28

            # ------------------------------------------------------------
            # 2. Findings register as a real Excel Table
            # ------------------------------------------------------------
            fws=wb.create_sheet("Findings")
            headers=["Finding No.","Standard","Clause","Sub-Clause","Finding Type","Location","Finding Details","Corrective Action","Responsible Person","Evidence Photo","Target Date","Status","Close Out Evidence"]
            report_header(fws,"AUDIT FINDINGS",len(headers))
            fws.append([]); fws.append(headers)
            for f in fs:
                fws.append([
                    safe(f["finding_number"]),safe(f["finding_standard"] if "finding_standard" in f.keys() else a["standard"]),safe(f["clause"]),safe(f["sub_clause"]),
                    safe(f["finding_type"]),safe(f["location"] if "location" in f.keys() else ""),safe(f["finding_detail"] or f["observation"]),
                    safe(f["corrective_action"]),safe(f["responsible"]),"",safe(f["target_date"]),
                    "OVERDUE" if overdue(f["target_date"],f["status"]) else safe(f["status"]),safe(f["close_out_evidence"] if "close_out_evidence" in f.keys() else "")
                ])
                rr=fws.max_row; fws.row_dimensions[rr].height=100
                photos=[]; other_files=[]
                for ar in self._audit_attachment_rows(audit_id,f["id"]):
                    fp=Path(safe(ar["file_path"]))
                    if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}: photos.append(fp)
                    elif fp.exists(): other_files.append(fp.name)
                if photos:
                    try:
                        img=XLImage(str(photos[0])); ratio=min(135/max(img.width,1),85/max(img.height,1),1); img.width=max(1,int(img.width*ratio)); img.height=max(1,int(img.height*ratio)); img.anchor=f"J{rr}"; fws.add_image(img)
                    except Exception: logging.exception("Unable to embed audit finding photo")
                if other_files: fws.cell(rr,10).value="; ".join(other_files)
                elif photos: fws.cell(rr,10).value="Photo evidence"
                else: fws.cell(rr,10).value="No photo"
            style_data_sheet(fws,5,f"A5:M{fws.max_row}","AuditFindings")
            for col,w in {"A":13,"B":16,"C":12,"D":18,"E":20,"F":20,"G":42,"H":42,"I":24,"J":24,"K":16,"L":20,"M":32}.items(): fws.column_dimensions[col].width=w

            # ------------------------------------------------------------
            # 3. Evidence / attachment register
            # ------------------------------------------------------------
            aws=wb.create_sheet("Attachments")
            headers=["Finding Number","Attachment / Evidence","File Type","Uploaded By","Upload Date"]
            report_header(aws,"AUDIT EVIDENCE REGISTER",len(headers))
            aws.append([]); aws.append(headers)
            attachment_rows=[("Audit Level",ar) for ar in self._audit_attachment_rows(audit_id)]
            for f in fs: attachment_rows += [(f["finding_number"],ar) for ar in self._audit_attachment_rows(audit_id,f["id"])]
            for label,ar in attachment_rows:
                fp=Path(safe(ar["file_path"]))
                aws.append([safe(label),"","Photo" if fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"} else "File",safe(ar["uploaded_by"]),safe(ar["uploaded_at"])])
                rr=aws.max_row; aws.row_dimensions[rr].height=90
                if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
                    try:
                        img=XLImage(str(fp)); ratio=min(150/max(img.width,1),75/max(img.height,1),1); img.width=max(1,int(img.width*ratio)); img.height=max(1,int(img.height*ratio)); img.anchor=f"B{rr}"; aws.add_image(img); aws.cell(rr,2).value="Photo evidence"
                    except Exception: aws.cell(rr,2).value=fp.name
                else: aws.cell(rr,2).value=fp.name
            if aws.max_row>=5: style_data_sheet(aws,5,f"A5:E{aws.max_row}","AuditAttachments")
            for col,w in {"A":18,"B":42,"C":15,"D":24,"E":22}.items(): aws.column_dimensions[col].width=w

            # ------------------------------------------------------------
            # 4. Corrective action tracker as a real Excel Table
            # ------------------------------------------------------------
            cws=wb.create_sheet("Corrective Action Tracker")
            headers=["Finding No.","Finding Detail","Corrective Action","Responsible Person","Target Date","Status","Close Out Evidence"]
            report_header(cws,"CORRECTIVE ACTION TRACKER",len(headers))
            cws.append([]); cws.append(headers)
            for f in fs:
                cws.append([safe(f["finding_number"]),safe(f["finding_detail"] or f["observation"]),safe(f["corrective_action"]),safe(f["responsible"]),safe(f["target_date"]),"OVERDUE" if overdue(f["target_date"],f["status"]) else safe(f["status"]),safe(f["close_out_evidence"] if "close_out_evidence" in f.keys() else "")])
                cws.row_dimensions[cws.max_row].height=60
            if cws.max_row>=5: style_data_sheet(cws,5,f"A5:G{cws.max_row}","AuditCAPA")
            for col,w in {"A":14,"B":42,"C":42,"D":24,"E":16,"F":20,"G":32}.items(): cws.column_dimensions[col].width=w

            # Consistent professional formatting / print setup.
            for sh in wb.worksheets:
                sh.sheet_properties.pageSetUpPr.fitToPage=True
                sh.print_title_rows="1:5"
                for row in sh.iter_rows():
                    for c in row:
                        if isinstance(c,MergedCell): continue
                        c.alignment=Alignment(vertical="top",horizontal="left",wrap_text=True)
                if sh.max_row>=5:
                    for c in sh[5]:
                        if isinstance(c,MergedCell): continue
                        c.font=c.font.copy(bold=True,color="FFFFFF")
                        c.fill=PatternFill("solid",fgColor="17365D")
                        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
                        c.border=Border(left=Side(style="thin",color="A6A6A6"),right=Side(style="thin",color="A6A6A6"),top=Side(style="thin",color="A6A6A6"),bottom=Side(style="thin",color="A6A6A6"))
                sh.auto_filter.ref=sh.auto_filter.ref or (f"A5:{chr(64+min(sh.max_column,26))}{sh.max_row}" if sh.max_row>=5 else None)

            target=Path(path); temp=target.with_name(target.name+".tmp.xlsx")
            try:
                wb.save(temp)
                with zipfile.ZipFile(temp,"r") as z:
                    bad=z.testzip()
                    if bad: raise ValueError(f"Generated Excel report is corrupt: {bad}")
                os.replace(temp,target)
            finally:
                temp.unlink(missing_ok=True)
            self.show_export_success(path)
        except Exception as e:
            logging.exception("Audit Excel export failed"); QMessageBox.critical(self,"Export Error",f"Unable to generate the Audit Excel report.\n\n{e}")

    def _audit_export_pdf(self,audit_id):
        a,fs=self._audit_report_data(audit_id)
        if not a:return
        path=self._export_path(f"{a['number']}_Audit_Report.pdf","Save Audit Professional Report","PDF Files (*.pdf)");
        if not path:return
        try:
            styles=getSampleStyleSheet(); doc=SimpleDocTemplate(path,pagesize=A4,rightMargin=28,leftMargin=28,topMargin=30,bottomMargin=28); story=[]
            logo=report_logo_path()
            if logo and Path(logo).exists(): story.append(RLImage(str(logo),width=90,height=55,preserveAspectRatio=True))
            story += [Paragraph(f"<b>AUDIT REPORT</b>",styles["Title"]),Paragraph(f"<b>{xml_safe(company_name())}</b>",styles["Heading2"]),Spacer(1,8)]
            meta=[["Audit Reference",a["number"]],["Audit Type",a["audit_type"]],["Audit Title",a["title"]],["Audit Date",a["audit_date"]],["Department",a["department"]],["Location",a["location"]],["Standard",a["standard"]],["Scope",a["scope"]],["Audit Criteria",a["criteria"]]]
            mt=Table(meta,colWidths=[120,400]); mt.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.grey),("BACKGROUND",(0,0),(0,-1),colors.HexColor("#EAF0F6")),("VALIGN",(0,0),(-1,-1),"TOP") ])); story += [mt,Spacer(1,12),Paragraph("<b>FINDINGS</b>",styles["Heading2"])]
            data=[["No.","Type","Clause","Sub-Clause","Finding Detail","Corrective Action","Responsible","Status","Target"]]
            for f in fs:data.append([str(f["finding_number"]),safe(f["finding_type"]),safe(f["clause"]),safe(f["sub_clause"]),safe(f["finding_detail"] or f["observation"]),safe(f["corrective_action"]),safe(f["responsible"]),"OVERDUE" if overdue(f["target_date"],f["status"]) else safe(f["status"]),safe(f["target_date"])])
            ps=styles["Normal"]; ps.fontSize=6.2; ps.leading=7.2
            pdata=[[Paragraph(xml_safe(v),ps) for v in row] for row in data]
            ft=Table(pdata,repeatRows=1,colWidths=[25,55,55,60,115,100,65,55,55]); ft.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.4,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP") ])); story.append(ft)
            # Add a consolidated evidence register so audit-level and finding-level
            # attachments are visible in the professional PDF report.
            story += [Spacer(1,12),Paragraph("<b>ATTACHMENTS / EVIDENCE REGISTER</b>",styles["Heading2"])]
            evidence_data=[["Finding / Level","Attachment / Evidence","Type","Uploaded By","Upload Date"]]
            evidence_items=[("Audit Level",ar) for ar in self._audit_attachment_rows(audit_id)]
            for finding in fs:
                evidence_items += [(safe(finding["finding_number"]),ar) for ar in self._audit_attachment_rows(audit_id,finding["id"])]
            for label,ar in evidence_items:
                fp=Path(safe(ar["file_path"]))
                evidence_data.append([label,fp.name or "Evidence attachment",safe(ar["attachment_type"] if "attachment_type" in ar.keys() else "Evidence"),safe(ar["uploaded_by"] if "uploaded_by" in ar.keys() else ""),safe(ar["uploaded_at"] if "uploaded_at" in ar.keys() else "")])
                if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
                    try:
                        evidence_data.append(["",Paragraph("Photo evidence",ps),"", "", ""])
                        evidence_data[-1][1]=RLImage(str(fp),width=1.0*72,height=0.75*72,preserveAspectRatio=True)
                    except Exception:
                        logging.exception("Unable to embed audit evidence image in PDF")
            if len(evidence_data)>1:
                ep=[[Paragraph(xml_safe(str(v)) if not isinstance(v,RLImage) else "",ps) if not isinstance(v,RLImage) else v for v in row] for row in evidence_data]
                et=Table(ep,repeatRows=1,colWidths=[65,175,55,85,80])
                et.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.4,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP")])); story.append(et)
            else:
                story.append(Paragraph("No attachments or evidence files are recorded for this audit.",ps))
            story += [Spacer(1,12),Paragraph("<b>REVIEW & APPROVAL</b>",styles["Heading2"]),Paragraph(xml_safe(f"Auditor: {a['auditor']} | {a['auditor_position']} | {a['auditor_date']}"),ps),Paragraph(xml_safe(f"Reviewer: {a['reviewer']} | {a['reviewer_position']} | {a['reviewer_date']}"),ps),Paragraph(xml_safe(f"Approver: {a['approver']} | {a['approver_position']} | {a['approver_date']}"),ps)]
            doc.build(story); self.show_export_success(path)
        except Exception as e: logging.exception("Audit PDF export failed"); QMessageBox.critical(self,"Export Error",str(e))

    def capa(self):
        w, layout = self.page("CAPA Tracking Register")
        info = QLabel(
            "Tracks formal corrective actions linked from HSE observations and incident investigations. "
            "Upload evidence against each corrective action. Completion % = actions with evidence ÷ total actions. "
            "100% = CAPA Action Complete. Approval is separate (Approved / Not Approved)."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color:#245486;padding:4px;")
        layout.addWidget(info)

        toolbar = QHBoxLayout()
        btns = {}
        for text_label in ["+ New CAPA", "Manage Action Evidence", "Set Approval", "Delete Selected",
                           "Export CSV", "Export Excel", "Export PDF", "Export Word"]:
            b = QPushButton(text_label)
            toolbar.addWidget(b)
            btns[text_label] = b
        layout.addLayout(toolbar)

        table = QTableWidget()
        headers = [
            "ID", "S.No.", "CAPA No.", "Reference (Observation / Incident)", "Finding / Incident Details", "Corrective Action(s)",
            "Evidence Progress", "Completion %", "Action Status", "Approval", "Approved By",
            "Priority", "Responsible", "Target Date", "Overdue"
        ]
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setColumnHidden(0, True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setWordWrap(True)
        table.setAlternatingRowColors(True)
        layout.addWidget(table)

        def selected_capa_id():
            r = table.currentRow()
            if r < 0 or not table.item(r, 0):
                QMessageBox.warning(self, "CAPA", "Select a CAPA record first.")
                return None
            return int(table.item(r, 0).text())

        def load():
            rows = db.fetchall("SELECT * FROM capa ORDER BY id DESC")
            table.setRowCount(len(rows))
            for r, row in enumerate(rows):
                actions = parse_capa_actions(row["actions_json"] if "actions_json" in row.keys() else row["corrective_action"])
                done, total, pct, action_status = capa_actions_completion(actions)
                try:
                    approval = safe(row["approval_status"]) or "Pending Approval"
                except Exception:
                    approval = "Pending Approval"
                try:
                    approved_by = safe(row["approved_by"])
                except Exception:
                    approved_by = ""
                try:
                    observed = safe(row["observation_text"]) or safe(row["finding"])
                except Exception:
                    observed = safe(row["finding"])
                progress = f"{done}/{total} with evidence" if total else "No actions"
                actions_display = format_corrective_actions_display([a["action"] for a in actions])
                vals = [
                    row["id"], r + 1, row["number"], row["reference_number"], observed, actions_display,
                    progress, f"{pct}%", action_status, approval, approved_by or "—",
                    row["priority"], row["responsible"], row["target_date"],
                    "OVERDUE" if overdue(row["target_date"], row["status"]) else ""
                ]
                for c, v in enumerate(vals):
                    item = QTableWidgetItem(safe(v))
                    item.setToolTip(safe(v))
                    table.setItem(r, c, item)
                status_item = table.item(r, 8)
                if action_status == "CAPA Action Complete":
                    status_item.setForeground(Qt.GlobalColor.darkGreen)
                elif action_status == "Pending":
                    status_item.setForeground(Qt.GlobalColor.darkYellow)
                else:
                    status_item.setForeground(Qt.GlobalColor.darkRed)
                if safe(row["source"]) == "HSE Inspection":
                    for c in range(table.columnCount()):
                        if table.item(r, c):
                            table.item(r, c).setBackground(QBrush(QColor("#E8F1FB")))
                table.setRowHeight(r, max(52, 18 + 16 * max(1, len(actions))))
            for row in rows:
                actions = parse_capa_actions(row["actions_json"] if "actions_json" in row.keys() else "[]")
                done, total, pct, action_status = capa_actions_completion(actions)
                db.execute(
                    "UPDATE capa SET completion_pct=?, action_status=? WHERE id=?",
                    (pct, action_status, row["id"])
                )

        def manage_evidence():
            capa_id = selected_capa_id()
            if capa_id is None:
                return
            row = db.fetchone("SELECT * FROM capa WHERE id=?", (capa_id,))
            if not row:
                return
            actions = parse_capa_actions(row["actions_json"] if "actions_json" in row.keys() else "[]")
            if not actions:
                actions = [{"action": a, "evidence": [], "has_evidence": False}
                           for a in parse_corrective_actions(row["corrective_action"])]
            if not actions:
                QMessageBox.information(self, "CAPA", "No corrective actions are listed on this CAPA.")
                return

            dialog = QDialog(self)
            dialog.setWindowTitle(f"CAPA Action Evidence — {row['number']}")
            dialog.resize(900, 560)
            v = QVBoxLayout(dialog)
            v.addWidget(QLabel(
                f"<b>{safe(row['number'])}</b> | Ref: {safe(row['reference_number'])}<br>"
                f"Upload evidence for each corrective action. Completion is based on actions with at least one attachment."
            ))
            grid = QTableWidget()
            grid.setColumnCount(4)
            grid.setHorizontalHeaderLabels(["S.No.", "Corrective Action", "Evidence Files", "Attach"])
            grid.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            grid.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            grid.setRowCount(len(actions))
            grid.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

            def refresh_grid():
                for i, action in enumerate(actions):
                    grid.setItem(i, 0, QTableWidgetItem(str(i + 1)))
                    grid.setItem(i, 1, QTableWidgetItem(action["action"]))
                    names = ", ".join(Path(p).name for p in action.get("evidence") or []) or "No evidence"
                    grid.setItem(i, 2, QTableWidgetItem(names))
                    btn = QPushButton("Add / Edit Evidence")
                    def make_handler(idx):
                        def handler():
                            existing = actions[idx].get("evidence") or []
                            if existing:
                                choice = QMessageBox.question(
                                    dialog, "Manage Evidence",
                                    "Choose Yes to replace the current evidence files, or No to add more files.",
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
                                    QMessageBox.StandardButton.Yes
                                )
                                if choice == QMessageBox.StandardButton.Cancel:
                                    return
                                replace_existing = choice == QMessageBox.StandardButton.Yes
                            else:
                                replace_existing = False
                            paths, _ = QFileDialog.getOpenFileNames(
                                dialog, "Select Evidence for Corrective Action",
                                "",
                                "Evidence Files (*.png *.jpg *.jpeg *.bmp *.gif *.pdf *.doc *.docx *.xls *.xlsx *.txt);;All Files (*)"
                            )
                            if not paths:
                                return
                            saved = []
                            for path in paths:
                                try:
                                    source = Path(path)
                                    if not source.exists():
                                        continue
                                    dest = ATTACH_DIR / f"{row['number']}_A{idx+1}_{source.name}"
                                    counter = 1
                                    while dest.exists():
                                        dest = ATTACH_DIR / f"{row['number']}_A{idx+1}_{counter}_{source.name}"
                                        counter += 1
                                    shutil.copy2(source, dest)
                                    saved.append(str(dest))
                                    db.execute(
                                        "INSERT INTO capa_attachments (capa_id,file_path,attachment_type) VALUES(?,?,?)",
                                        (capa_id, str(dest), f"Action {idx+1} Evidence")
                                    )
                                except Exception:
                                    logging.exception("CAPA action evidence save failed")
                            if saved:
                                if replace_existing:
                                    old_paths = set(existing)
                                    actions[idx]["evidence"] = saved
                                    for old_path in old_paths:
                                        try:
                                            db.execute("DELETE FROM capa_attachments WHERE capa_id=? AND file_path=?", (capa_id, old_path))
                                        except Exception:
                                            logging.exception("Unable to update replaced CAPA evidence record")
                                else:
                                    actions[idx].setdefault("evidence", []).extend(saved)
                                actions[idx]["has_evidence"] = True
                                refresh_grid()
                        return handler
                    btn.clicked.connect(make_handler(i))
                    grid.setCellWidget(i, 3, btn)
                    grid.setRowHeight(i, 48)
            refresh_grid()
            v.addWidget(grid, 1)

            def save_and_close():
                done, total, pct, action_status = capa_actions_completion(actions)
                db.execute(
                    """UPDATE capa SET actions_json=?, completion_pct=?, action_status=?,
                       corrective_action=?, status=? WHERE id=?""",
                    (
                        json.dumps(actions, ensure_ascii=False),
                        pct,
                        action_status,
                        format_corrective_actions_display([a["action"] for a in actions]),
                        "Closed" if action_status == "CAPA Action Complete" else ("In Progress" if pct > 0 else "Open"),
                        capa_id,
                    )
                )
                updated = db.fetchone("SELECT * FROM capa WHERE id=?", (capa_id,))
                if updated:
                    sync_observation_from_capa(updated)
                dialog.accept()
                load()
                QMessageBox.information(
                    self, "CAPA Tracking",
                    f"Evidence saved.\nCompletion: {pct}% ({done}/{total})\nStatus: {action_status}"
                )

            buttons = QDialogButtonBox()
            save_b = buttons.addButton("Save Evidence / Update Status", QDialogButtonBox.ButtonRole.AcceptRole)
            cancel_b = buttons.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
            save_b.clicked.connect(save_and_close)
            cancel_b.clicked.connect(dialog.reject)
            v.addWidget(buttons)
            dialog.exec()

        def set_approval():
            capa_id = selected_capa_id()
            if capa_id is None:
                return
            row = db.fetchone("SELECT * FROM capa WHERE id=?", (capa_id,))
            if not row:
                return
            dialog = QDialog(self)
            dialog.setWindowTitle(f"CAPA Approval — {row['number']}")
            dialog.resize(420, 220)
            form = QFormLayout(dialog)
            approval = QComboBox()
            approval.addItems(["Pending Approval", "Approved", "Not Approved", "Rejected"])
            try:
                approval.setCurrentText(safe(row["approval_status"]) or "Pending Approval")
            except Exception:
                pass
            approved_by = QLineEdit(safe(row["approved_by"]) if "approved_by" in row.keys() else "")
            form.addRow("CAPA Approval:", approval)
            form.addRow("Approved By:", approved_by)
            buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
            form.addRow(buttons)

            def save_approval():
                if approval.currentText() in ("Approved", "Not Approved", "Rejected") and not approved_by.text().strip():
                    QMessageBox.warning(dialog, "Required", "Enter the name of the person approving / rejecting.")
                    return
                adate = today() if approval.currentText() != "Pending Approval" else ""
                db.execute(
                    "UPDATE capa SET approval_status=?, approved_by=?, approved_date=? WHERE id=?",
                    (approval.currentText(), approved_by.text().strip(), adate, capa_id)
                )
                updated = db.fetchone("SELECT * FROM capa WHERE id=?", (capa_id,))
                if updated:
                    sync_observation_from_capa(updated)
                dialog.accept()
                load()

            buttons.accepted.connect(save_approval)
            buttons.rejected.connect(dialog.reject)
            dialog.exec()

        btns["+ New CAPA"].clicked.connect(lambda: self.capa_form(load))
        btns["Manage Action Evidence"].clicked.connect(manage_evidence)
        btns["Set Approval"].clicked.connect(set_approval)
        btns["Delete Selected"].clicked.connect(lambda: self.delete_selected_record(table, "capa", "capa_attachments", "CAPA", load))
        btns["Export CSV"].clicked.connect(lambda: self.export_csv("capa"))
        btns["Export Excel"].clicked.connect(lambda: self.export_excel("capa"))
        btns["Export PDF"].clicked.connect(lambda: self.export_pdf("capa"))
        btns["Export Word"].clicked.connect(lambda: self.export_docx("capa"))
        load()

    def capa_form(self, refresh):
        dialog = QDialog(self)
        dialog.setWindowTitle("New CAPA")
        dialog.resize(950, 850)
        dialog.setMinimumSize(900, 720)
        outer = QVBoxLayout(dialog)

        title = QLabel("CORRECTIVE AND PREVENTIVE ACTION (CAPA) TRACKING")
        title.setStyleSheet("font-size:20px;font-weight:bold;color:#17365D;padding:6px;")
        outer.addWidget(title)

        scroll_area = QScrollArea(); scroll_area.setWidgetResizable(True)
        content = QWidget(); form = QFormLayout(content)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)

        source = QComboBox(); source.addItems(CAPA_SOURCES); form.addRow("Source:", source)
        reference = QLineEdit(); form.addRow("Reference Number:", reference)
        finding = QTextEdit(); finding.setMinimumHeight(90); form.addRow("Finding / Observed:", finding)
        root = QTextEdit(); root.setMinimumHeight(70); form.addRow("Root Cause:", root)

        actions_box = QGroupBox("Corrective Action(s)")
        actions_layout = QVBoxLayout(actions_box)
        action_rows = []

        def add_action_row(text=""):
            row_w = QWidget()
            row_l = QHBoxLayout(row_w)
            row_l.setContentsMargins(0, 0, 0, 0)
            te = QTextEdit(); te.setFixedHeight(50); te.setPlainText(text)
            rem = QPushButton("Remove"); rem.setFixedWidth(80)
            row_l.addWidget(te, 1); row_l.addWidget(rem)
            actions_layout.insertWidget(max(0, actions_layout.count() - 1), row_w)
            entry = {"widget": row_w, "edit": te}
            action_rows.append(entry)
            def remove():
                if len(action_rows) <= 1:
                    te.clear(); return
                action_rows.remove(entry); row_w.deleteLater()
            rem.clicked.connect(remove)

        add_btn = QPushButton("+ Add Corrective Action")
        add_btn.clicked.connect(lambda: add_action_row())
        actions_layout.addWidget(add_btn)
        add_action_row()
        form.addRow(actions_box)

        preventive = QTextEdit(); preventive.setMinimumHeight(70); form.addRow("Preventive Action:", preventive)
        responsible = QLineEdit(); form.addRow("Responsible Person:", responsible)
        priority = QComboBox(); priority.addItems(PRIORITIES); form.addRow("Priority:", priority)
        target = QLineEdit(); form.addRow("Target Date (YYYY-MM-DD):", target)
        status = QComboBox(); status.addItems(STATUSES); form.addRow("Status:", status)
        verification = QTextEdit(); verification.setMinimumHeight(60); form.addRow("Verification:", verification)
        evidence = QTextEdit(); evidence.setMinimumHeight(60); form.addRow("Closeout Evidence:", evidence)

        scroll_area.setWidget(content); outer.addWidget(scroll_area, 1)
        buttons = QDialogButtonBox()
        save_button = buttons.addButton("Save CAPA", QDialogButtonBox.ButtonRole.AcceptRole)
        cancel_button = buttons.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
        save_button.setMinimumHeight(40); cancel_button.setMinimumHeight(40)
        outer.addWidget(buttons); cancel_button.clicked.connect(dialog.reject)

        def save():
            try:
                actions = [e["edit"].toPlainText().strip() for e in action_rows if e["edit"].toPlainText().strip()]
                if not actions:
                    QMessageBox.warning(dialog, "Required", "Add at least one corrective action.")
                    return
                number = next_number("HSE-CAPA", "capa")
                corrective_text = format_corrective_actions_display(actions)
                capa_actions = [{"action": a, "evidence": [], "has_evidence": False} for a in actions]
                db.execute("""
                    INSERT INTO capa (
                        number,source,reference_number,finding,root_cause,corrective_action,
                        preventive_action,responsible,priority,target_date,status,verification,closeout_evidence,created_at,
                        actions_json,approval_status,approved_by,approved_date,completion_pct,action_status,observation_text
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    number, source.currentText(), reference.text(), finding.toPlainText(), root.toPlainText(),
                    corrective_text, preventive.toPlainText(), responsible.text(), priority.currentText(),
                    target.text(), status.currentText(), verification.toPlainText(), evidence.toPlainText(),
                    datetime.now().isoformat(),
                    json.dumps(capa_actions, ensure_ascii=False),
                    "Pending Approval", "", "", 0, "Open", finding.toPlainText()
                ))
                dialog.accept(); refresh()
            except Exception as e:
                logging.exception("CAPA save failed")
                QMessageBox.critical(dialog, "Save Error", f"Unable to save CAPA.\n\n{e}")

        save_button.clicked.connect(save)
        dialog.exec()


    # ========================================================
    # TRAINING MANAGEMENT (new additive module)
    # ========================================================

    def training(self):
        w, layout = self.page("Training Management")
        intro = QLabel(
            "Training Matrix, employee Training Register and Master Training Record. "
            "Trade-based requirements are fetched automatically from the matrix."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        tabs = QTabWidget()
        tabs.setDocumentMode(True)
        layout.addWidget(tabs, 1)

        trade_options = [
            "Welder", "Rigger", "Fabricator", "Electrician", "Mechanical Technician",
            "Instrument Technician", "Scaffolder", "Scaffold Inspector", "Forklift Operator",
            "Crane Operator", "Banksman / Signaler", "Helper", "General Worker", "Driver",
            "HSE Officer", "HSE Supervisor", "Supervisor", "Engineer", "Storekeeper", "Other"
        ]
        department_options = [
            "HSE", "Operations", "Construction", "Maintenance", "Engineering", "Electrical",
            "Mechanical", "Welding", "Fabrication", "Rigging", "Scaffolding", "Logistics",
            "Warehouse", "Transport", "Administration", "HR", "IT", "Quality", "Environment",
            "Security", "Facilities", "Other"
        ]

        def distinct_values(column):
            try:
                rows = db.fetchall(
                    f"SELECT DISTINCT {column} AS value FROM employees "
                    f"WHERE COALESCE(active,1)=1 AND COALESCE({column},'')<>'' ORDER BY {column} COLLATE NOCASE"
                )
                return [safe(row["value"]).strip() for row in rows if safe(row["value"]).strip()]
            except Exception:
                return []

        def fill_combo(combo, options, first_label=None, first_data=""):
            combo.clear()
            if first_label is not None:
                combo.addItem(first_label, first_data)
            for value in options:
                if value and value not in [combo.itemText(i) for i in range(combo.count())]:
                    combo.addItem(value, value)

        # ---------------- Training Matrix ----------------
        matrix_page = QWidget()
        matrix_layout = QVBoxLayout(matrix_page)
        matrix_layout.setContentsMargins(8, 8, 8, 8)
        matrix_layout.setSpacing(8)
        form_box = QGroupBox("Training Matrix — applicable course requirements")
        form_layout = QGridLayout(form_box)
        form_layout.setHorizontalSpacing(8)
        form_layout.setVerticalSpacing(7)

        course_name = QLineEdit()
        course_name.setPlaceholderText("e.g. HSE Induction, First Aid, LOTO")
        trade_combo = QComboBox()
        trade_combo.setEditable(True)
        trade_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        trade_combo.setToolTip("Open the dropdown and check every trade this course applies to. Select All Trades for universal applicability.")
        trade_model = QStandardItemModel(trade_combo)
        trade_combo.setModel(trade_model)
        for trade_label in ["All Trades"] + trade_options:
            trade_item = QStandardItem(trade_label)
            trade_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            trade_item.setData(Qt.CheckState.Unchecked, Qt.ItemDataRole.CheckStateRole)
            trade_model.appendRow(trade_item)
        trade_combo.lineEdit().setReadOnly(True)
        trade_combo.lineEdit().setPlaceholderText("Select one or more trades…")
        trade_other_label = QLabel("Other trade")
        trade_other_input = QLineEdit()
        trade_other_input.setPlaceholderText("Enter other trade name")
        trade_other_label.setVisible(False)
        trade_other_input.setVisible(False)
        department = QComboBox()
        fill_combo(department, department_options + distinct_values("department"), "All Departments", "")
        mandatory = QComboBox()
        mandatory.addItems(["Mandatory", "Recommended"])
        training_type = QComboBox()
        training_type.addItems(["Internal", "External"])
        validity = QSpinBox()
        validity.setRange(0, 600)
        validity.setSuffix(" months")
        validity.setSpecialValueText("No expiry / not set")
        source = QComboBox()
        source.addItems([
            "Qatar law / regulation", "Regulatory / authority requirement", "Client / project requirement",
            "Risk assessment", "Company procedure / policy", "Contractual requirement", "Other"
        ])
        certificate_required = QCheckBox("Certificate required")
        active = QCheckBox("Active requirement")
        active.setChecked(True)

        form_layout.addWidget(QLabel("Course name"), 0, 0)
        form_layout.addWidget(course_name, 0, 1, 1, 2)
        form_layout.addWidget(QLabel("Applicable trade(s)"), 0, 3)
        form_layout.addWidget(trade_combo, 0, 4, 1, 2)
        form_layout.addWidget(QLabel("Department"), 1, 0)
        form_layout.addWidget(department, 1, 1)
        form_layout.addWidget(QLabel("Requirement"), 1, 2)
        form_layout.addWidget(mandatory, 1, 3)
        form_layout.addWidget(QLabel("Training type"), 1, 4)
        form_layout.addWidget(training_type, 1, 5)
        form_layout.addWidget(QLabel("Validity"), 2, 0)
        form_layout.addWidget(validity, 2, 1)
        form_layout.addWidget(QLabel("Requirement source"), 2, 2)
        form_layout.addWidget(source, 2, 3, 1, 2)
        form_layout.addWidget(certificate_required, 2, 5)
        form_layout.addWidget(active, 3, 0, 1, 2)
        form_layout.addWidget(trade_other_label, 3, 2)
        form_layout.addWidget(trade_other_input, 3, 3, 1, 3)
        trade_other_label.setVisible(False)
        form_layout.setColumnStretch(1, 2)
        form_layout.setColumnStretch(4, 2)
        matrix_layout.addWidget(form_box)

        matrix_buttons = QHBoxLayout()
        save_course_btn = QPushButton("Save Course")
        clear_course_btn = QPushButton("Clear Form")
        deactivate_course_btn = QPushButton("Delete Selected")
        export_matrix_btn = QPushButton("Export Matrix to Excel")
        refresh_matrix_btn = QPushButton("Refresh")
        for btn in (save_course_btn, clear_course_btn, deactivate_course_btn, export_matrix_btn, refresh_matrix_btn):
            btn.setMinimumHeight(34)
            btn.setMinimumWidth(100)
            matrix_buttons.addWidget(btn)
        matrix_buttons.addStretch(1)
        matrix_layout.addLayout(matrix_buttons)

        course_table = QTableWidget()
        course_table.setColumnCount(9)
        course_table.setHorizontalHeaderLabels([
            "ID", "Training Course", "Applicable Trade(s)", "Department", "Requirement",
            "Training Type", "Validity (months)", "Requirement Source", "Status"
        ])
        course_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        course_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        course_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        course_table.setAlternatingRowColors(True)
        course_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        course_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        course_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        course_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeMode.ResizeToContents)
        course_table.setColumnHidden(0, True)
        matrix_layout.addWidget(course_table, 1)
        tabs.addTab(matrix_page, "Training Matrix")

        # ---------------- Training Register ----------------
        register_page = QWidget()
        register_layout = QVBoxLayout(register_page)
        register_layout.setContentsMargins(8, 8, 8, 8)
        register_layout.setSpacing(8)
        register_help = QLabel(
            "One employee per row and one course per column. Select an applicable course cell to add a training date, "
            "expiry date and certificate. The employee's trade and department determine applicable courses."
        )
        register_help.setWordWrap(True)
        register_layout.addWidget(register_help)

        employee_box = QGroupBox("Add employee to the training register")
        employee_form = QGridLayout(employee_box)
        employee_id_input = QLineEdit()
        employee_id_input.setPlaceholderText("Employee ID (optional)")
        employee_name_input = QLineEdit()
        employee_name_input.setPlaceholderText("Employee full name")
        employee_department_input = QComboBox()
        employee_department_input.setEditable(True)
        fill_combo(employee_department_input, department_options + distinct_values("department"), "Select department", "")
        employee_trade_input = QComboBox()
        employee_trade_input.setEditable(True)
        fill_combo(employee_trade_input, trade_options + distinct_values("designation"), "Select trade", "")
        joining_date = QDateEdit()
        joining_date.setCalendarPopup(True)
        joining_date.setDisplayFormat("yyyy-MM-dd")
        joining_date.setDate(QDate.currentDate())
        safety_induction_date = QDateEdit()
        safety_induction_date.setCalendarPopup(True)
        safety_induction_date.setDisplayFormat("yyyy-MM-dd")
        safety_induction_date.setSpecialValueText("Not completed")
        safety_induction_date.setDate(QDate(2000, 1, 1))
        safety_induction_date.setMinimumDate(QDate(2000, 1, 1))
        safety_induction_date.setMaximumDate(QDate.currentDate().addYears(20))
        safety_induction_date.setToolTip("Set to 2000-01-01 if safety induction has not yet been completed.")
        add_employee_btn = QPushButton("Save Employee")
        edit_employee_btn = QPushButton("Edit Selected Person")
        delete_employee_btn = QPushButton("Delete Selected Person")
        selected_employee_id = {"value": None}
        employee_form.addWidget(QLabel("Employee ID"), 0, 0)
        employee_form.addWidget(employee_id_input, 0, 1)
        employee_form.addWidget(QLabel("Employee name *"), 0, 2)
        employee_form.addWidget(employee_name_input, 0, 3)
        employee_form.addWidget(QLabel("Department"), 0, 4)
        employee_form.addWidget(employee_department_input, 0, 5)
        employee_form.addWidget(QLabel("Trade"), 1, 0)
        employee_form.addWidget(employee_trade_input, 1, 1)
        employee_form.addWidget(QLabel("Date joined"), 1, 2)
        employee_form.addWidget(joining_date, 1, 3)
        employee_form.addWidget(QLabel("Safety induction date"), 1, 4)
        employee_form.addWidget(safety_induction_date, 1, 5)
        employee_form.addWidget(add_employee_btn, 2, 3)
        employee_form.addWidget(edit_employee_btn, 2, 4)
        employee_form.addWidget(delete_employee_btn, 2, 5)
        register_layout.addWidget(employee_box)

        register_toolbar = QHBoxLayout()
        register_search = QLineEdit()
        register_search.setPlaceholderText("Search employee, ID, trade or department…")
        register_search.setClearButtonEnabled(True)
        refresh_register_btn = QPushButton("Refresh Register")
        export_register_btn = QPushButton("Export Register to Excel")
        register_toolbar.addWidget(register_search, 1)
        register_toolbar.addWidget(refresh_register_btn)
        register_toolbar.addWidget(export_register_btn)
        register_layout.addLayout(register_toolbar)
        status_summary = QLabel("")
        status_summary.setStyleSheet("font-weight: bold; padding: 4px;")
        register_layout.addWidget(status_summary)
        register_table = QTableWidget()
        register_table.setAlternatingRowColors(True)
        register_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        register_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        register_table.setWordWrap(True)
        register_table.horizontalHeader().setMinimumSectionSize(100)
        register_table.horizontalHeader().setDefaultSectionSize(170)
        register_table.verticalHeader().setDefaultSectionSize(58)
        register_layout.addWidget(register_table, 1)
        tabs.addTab(register_page, "Training Register")

        # ---------------- Master Training Record ----------------
        master_page = QWidget()
        master_layout = QVBoxLayout(master_page)
        master_layout.setContentsMargins(8, 8, 8, 8)
        master_layout.setSpacing(8)
        master_info = QLabel("Complete employee-by-course history, including joining date, safety induction, trade, training dates, expiry, type and compliance status.")
        master_info.setWordWrap(True)
        master_layout.addWidget(master_info)
        master_toolbar = QHBoxLayout()
        master_search = QLineEdit()
        master_search.setPlaceholderText("Search employee, course, trade or department…")
        master_search.setClearButtonEnabled(True)
        refresh_master_btn = QPushButton("Refresh Master Record")
        export_master_excel_btn = QPushButton("Export Master Record to Excel")
        export_master_word_btn = QPushButton("Export Master Record to Word")
        delete_master_record_btn = QPushButton("Delete Selected Training Record")
        for btn in (refresh_master_btn, export_master_excel_btn, export_master_word_btn, delete_master_record_btn):
            btn.setMinimumHeight(34)
        master_toolbar.addWidget(master_search, 1)
        master_toolbar.addWidget(refresh_master_btn)
        master_toolbar.addWidget(export_master_excel_btn)
        master_toolbar.addWidget(export_master_word_btn)
        master_toolbar.addWidget(delete_master_record_btn)
        master_layout.addLayout(master_toolbar)
        master_table = QTableWidget()
        master_table.setColumnCount(14)
        master_table.setHorizontalHeaderLabels([
            "Employee ID", "Employee Name", "Department", "Trade", "Date Joined", "Safety Induction Date",
            "Training Course", "Training Type", "Requirement Source", "Training Date", "Expiry Date",
            "Status", "Certificate No.", "Attachment"
        ])
        master_table.setAlternatingRowColors(True)
        master_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        master_table.setWordWrap(True)
        master_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        master_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        master_layout.addWidget(master_table, 1)
        tabs.addTab(master_page, "Master Training Record")

        selected_course_id = {"value": None}

        def checked_trades():
            selected = []
            all_item = trade_model.item(0)
            if all_item and all_item.checkState() == Qt.CheckState.Checked:
                return []
            for idx in range(1, trade_model.rowCount()):
                item = trade_model.item(idx)
                if item and item.checkState() == Qt.CheckState.Checked:
                    value = item.text().strip()
                    if value == "Other":
                        other_value = trade_other_input.text().strip()
                        if other_value:
                            selected.append(other_value)
                    else:
                        selected.append(value)
            return list(dict.fromkeys(selected))

        def update_trade_summary(*args):
            all_item = trade_model.item(0)
            selected = checked_trades()
            if all_item and all_item.checkState() == Qt.CheckState.Checked:
                summary = "All Trades"
            else:
                labels = []
                for idx in range(1, trade_model.rowCount()):
                    item = trade_model.item(idx)
                    if item and item.checkState() == Qt.CheckState.Checked:
                        labels.append(item.text())
                summary = ", ".join(labels) if labels else "Select trade(s)…"
            trade_combo.lineEdit().setText(summary)
            show_other = any(trade_model.item(i).text() == "Other" and trade_model.item(i).checkState() == Qt.CheckState.Checked for i in range(1, trade_model.rowCount()))
            trade_other_label.setVisible(show_other)
            trade_other_input.setVisible(show_other)

        def toggle_trade_item(index):
            if not index.isValid():
                return
            item = trade_model.itemFromIndex(index)
            if not item:
                return
            new_state = Qt.CheckState.Unchecked if item.checkState() == Qt.CheckState.Checked else Qt.CheckState.Checked
            if item.row() == 0 and new_state == Qt.CheckState.Checked:
                for idx in range(1, trade_model.rowCount()):
                    trade_model.item(idx).setCheckState(Qt.CheckState.Unchecked)
            elif item.row() > 0 and new_state == Qt.CheckState.Checked:
                trade_model.item(0).setCheckState(Qt.CheckState.Unchecked)
            item.setCheckState(new_state)
            update_trade_summary()

        trade_combo.view().pressed.connect(toggle_trade_item)
        trade_other_input.textChanged.connect(update_trade_summary)

        def set_trade_selection(saved_trade):
            saved = [x.strip().lower() for x in safe(saved_trade).split(",") if x.strip()]
            if not saved:
                trade_model.item(0).setCheckState(Qt.CheckState.Checked)
                for idx in range(1, trade_model.rowCount()):
                    trade_model.item(idx).setCheckState(Qt.CheckState.Unchecked)
                trade_other_input.clear()
            else:
                trade_model.item(0).setCheckState(Qt.CheckState.Unchecked)
                known = {trade_model.item(i).text().lower(): i for i in range(1, trade_model.rowCount())}
                other_saved = []
                for idx in range(1, trade_model.rowCount()):
                    trade_model.item(idx).setCheckState(Qt.CheckState.Unchecked)
                for trade in saved:
                    if trade in known and trade != "other":
                        trade_model.item(known[trade]).setCheckState(Qt.CheckState.Checked)
                    elif trade == "other":
                        trade_model.item(known[trade]).setCheckState(Qt.CheckState.Checked)
                    elif trade not in known:
                        # Backward compatibility for custom trade values already saved in older matrix records.
                        trade_other_input.setText(trade)
                        trade_model.item(known["other"]).setCheckState(Qt.CheckState.Checked)
                if "other" not in saved and not any(x not in known for x in saved):
                    trade_other_input.clear()
            update_trade_summary()

        def reset_course_form():
            selected_course_id["value"] = None
            course_table.clearSelection()
            course_name.clear()
            set_trade_selection("")
            trade_other_input.clear()
            department.setCurrentIndex(0)
            mandatory.setCurrentIndex(0)
            training_type.setCurrentIndex(0)
            validity.setValue(0)
            source.setCurrentText("Company procedure / policy")
            certificate_required.setChecked(False)
            active.setChecked(True)
            save_course_btn.setText("Save Course")

        def load_courses():
            rows = db.fetchall("SELECT * FROM training_courses ORDER BY active DESC, course_name COLLATE NOCASE")
            course_table.setRowCount(len(rows))
            for r, row in enumerate(rows):
                saved_trades = [x.strip() for x in safe(row["designation"]).split(",") if x.strip()]
                trade_display = "All Trades" if not saved_trades else "\n• " + "\n• ".join(saved_trades)
                values = [
                    row["id"], row["course_name"], trade_display,
                    row["department"] or "All Departments", "Mandatory" if row["mandatory"] else "Recommended",
                    row["provider_type"] or "Internal", row["validity_months"] or 0,
                    row["requirement_source"] or "Company procedure / policy", "Active" if row["active"] else "Inactive"
                ]
                for c, value in enumerate(values):
                    item = QTableWidgetItem(str(value))
                    if c == 0:
                        item.setData(Qt.ItemDataRole.UserRole, int(row["id"]))
                    course_table.setItem(r, c, item)
            course_table.resizeRowsToContents()
            load_register()
            load_master_record()

        def select_course_row(row, column=0):
            item = course_table.item(row, 0)
            if not item:
                return
            course_id = item.data(Qt.ItemDataRole.UserRole)
            rec = db.fetchone("SELECT * FROM training_courses WHERE id=?", (course_id,))
            if not rec:
                return
            selected_course_id["value"] = int(course_id)
            course_name.setText(safe(rec["course_name"]))
            saved_trade = safe(rec["designation"]).strip()
            set_trade_selection(saved_trade)
            saved_dept = safe(rec["department"]).strip()
            if department.findText(saved_dept) < 0 and saved_dept:
                department.addItem(saved_dept, saved_dept)
            department.setCurrentText(saved_dept if saved_dept else "All Departments")
            mandatory.setCurrentIndex(0 if rec["mandatory"] else 1)
            training_type.setCurrentText(safe(rec["provider_type"]) if safe(rec["provider_type"]) in ("Internal", "External") else "Internal")
            validity.setValue(int(rec["validity_months"] or 0))
            req_source = safe(rec["requirement_source"])
            if source.findText(req_source) < 0 and req_source:
                source.addItem(req_source)
            source.setCurrentText(req_source or "Company procedure / policy")
            certificate_required.setChecked(bool(rec["certificate_required"]))
            active.setChecked(bool(rec["active"]))
            save_course_btn.setText("Update Course")

        def save_course():
            name = course_name.text().strip()
            if not name:
                QMessageBox.warning(self, "Training Matrix", "Enter a training course name.")
                course_name.setFocus()
                return
            trades_selected = checked_trades()
            all_checked = bool(trade_model.item(0) and trade_model.item(0).checkState() == Qt.CheckState.Checked)
            if not all_checked and not trades_selected:
                QMessageBox.warning(self, "Training Matrix", "Select All Trades or at least one applicable trade.")
                return
            if any(trade_model.item(i).text() == "Other" and trade_model.item(i).checkState() == Qt.CheckState.Checked for i in range(1, trade_model.rowCount())) and not trade_other_input.text().strip():
                QMessageBox.warning(self, "Training Matrix", "Enter the trade name in the Other trade box.")
                trade_other_input.setFocus()
                return
            trades = "" if all_checked else ", ".join(dict.fromkeys(trades_selected))
            dept_text = department.currentText().strip()
            dept = "" if dept_text.lower() in ("all departments", "") else dept_text
            values = (
                name, trades, dept, 1 if mandatory.currentIndex() == 0 else 0,
                training_type.currentText(), int(validity.value()), source.currentText(),
                1 if certificate_required.isChecked() else 0, 1 if active.isChecked() else 0
            )
            existing = db.fetchone(
                "SELECT id FROM training_courses WHERE lower(course_name)=lower(?) "
                "AND COALESCE(designation,'')=? AND COALESCE(department,'')=? AND active=1 AND id<>COALESCE(?, -1)",
                (name, trades, dept, selected_course_id["value"])
            )
            if existing:
                QMessageBox.warning(self, "Training Matrix", "An active course with the same course name, trade(s) and department already exists.")
                return
            if selected_course_id["value"] is None:
                db.execute(
                    """INSERT INTO training_courses
                    (course_name,designation,department,mandatory,provider_type,validity_months,requirement_source,certificate_required,active)
                    VALUES (?,?,?,?,?,?,?,?,?)""", values
                )
            else:
                db.execute(
                    """UPDATE training_courses SET course_name=?,designation=?,department=?,mandatory=?,provider_type=?,
                    validity_months=?,requirement_source=?,certificate_required=?,active=? WHERE id=?""",
                    values + (selected_course_id["value"],)
                )
            reset_course_form()
            load_courses()

        def deactivate_course():
            row = course_table.currentRow()
            if row < 0:
                QMessageBox.information(self, "Training Matrix", "Select a course first.")
                return
            course_id = course_table.item(row, 0).data(Qt.ItemDataRole.UserRole)
            history = db.fetchone("SELECT COUNT(*) AS n FROM training_records WHERE course_id=?", (course_id,))
            has_history = bool(history and int(history["n"] or 0) > 0)
            message = ("Training records exist for this course. Delete is not destructive: the course will be marked inactive and historical records retained. Continue?"
                       if has_history else "Delete this training course from the matrix?")
            answer = QMessageBox.question(
                self, "Delete Training Course", message,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if answer == QMessageBox.StandardButton.Yes:
                if has_history:
                    db.execute("UPDATE training_courses SET active=0 WHERE id=?", (course_id,))
                else:
                    db.execute("DELETE FROM training_courses WHERE id=?", (course_id,))
                reset_course_form()
                load_courses()

        def add_months(date_text, months):
            try:
                base = datetime.strptime(date_text, "%Y-%m-%d").date()
                month_index = base.month - 1 + int(months)
                year = base.year + month_index // 12
                month = month_index % 12 + 1
                import calendar
                day = min(base.day, calendar.monthrange(year, month)[1])
                return date(year, month, day).isoformat()
            except Exception:
                return ""

        def latest_training(employee_db_id, course_id):
            return db.fetchone(
                "SELECT * FROM training_records WHERE employee_db_id=? AND course_id=? ORDER BY training_date DESC, id DESC LIMIT 1",
                (employee_db_id, course_id)
            )

        def course_applies(course, employee):
            raw_trades = safe(course["designation"]).strip()
            emp_trade = safe(employee["designation"]).strip().lower()
            allowed_trades = [t.strip().lower() for t in raw_trades.split(",") if t.strip()]
            trade_match = not allowed_trades or emp_trade in allowed_trades or "all trades" in allowed_trades or "all designations" in allowed_trades
            course_dept = safe(course["department"]).strip().lower()
            emp_dept = safe(employee["department"]).strip().lower()
            return trade_match and (not course_dept or course_dept == emp_dept or course_dept == "all departments")

        def training_status(record, course=None):
            if not record:
                return "Training Required"
            expiry = safe(record["expiry_date"]).strip()
            if not expiry:
                return "Completed"
            try:
                expiry_day = datetime.strptime(expiry, "%Y-%m-%d").date()
                days_left = (expiry_day - date.today()).days
                if days_left < 0:
                    return "Expired"
                if course and safe(course["provider_type"]).lower() == "external" and days_left <= 45:
                    return "Due Soon"
                if course and safe(course["provider_type"]).lower() == "internal" and days_left <= 30:
                    return "Due Soon"
                return "Valid"
            except Exception:
                return "Completed"

        def set_status_style(item, status, course=None):
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setToolTip(item.toolTip() or status)
            bg = None
            if status == "Expired" or status == "Training Required":
                bg = QColor("#FDE2E2")
                fg = QColor("#B42318")
            elif status == "Due Soon" and course and safe(course["provider_type"]).lower() == "external":
                bg = QColor("#FDE2E2")
                fg = QColor("#B42318")
            elif status == "Due Soon" and course and safe(course["provider_type"]).lower() == "internal":
                bg = QColor("#DDF5E5")
                fg = QColor("#176B3A")
            elif status in ("Valid", "Completed"):
                fg = QColor("#176B3A")
            elif status == "N/A":
                fg = QColor("#777777")
            else:
                fg = QColor("#9A6700")
            item.setForeground(QBrush(fg))
            if bg:
                item.setBackground(QBrush(bg))
            font = item.font()
            font.setBold(True)
            item.setFont(font)

        def get_employees():
            return db.fetchall(
                "SELECT id,employee_id,name,designation,department,active,date_joining,safety_induction_date "
                "FROM employees WHERE COALESCE(active,1)=1 ORDER BY name COLLATE NOCASE"
            )

        def expanded_course_entries(courses):
            entries = []
            for course in courses:
                saved_trades = [x.strip() for x in safe(course["designation"]).split(",") if x.strip()]
                if not saved_trades:
                    saved_trades = [""]
                for trade_label in saved_trades:
                    entry = dict(course)
                    entry["designation"] = trade_label
                    entry["display_course_name"] = safe(course["course_name"]) if not trade_label else f"{safe(course['course_name'])} ({trade_label})"
                    entries.append(entry)
            return entries

        def load_register():
            employees = get_employees()
            courses = expanded_course_entries(db.fetchall("SELECT * FROM training_courses WHERE active=1 OR EXISTS (SELECT 1 FROM training_records tr WHERE tr.course_id=training_courses.id) ORDER BY active DESC, course_name COLLATE NOCASE"))
            base_headers = ["Employee ID", "Employee Name", "Department", "Trade", "Date Joined", "Safety Induction Date"]
            headers = base_headers + [safe(c["display_course_name"]) for c in courses]
            register_table.clear()
            register_table.setColumnCount(len(headers))
            register_table.setHorizontalHeaderLabels(headers)
            register_table.setRowCount(len(employees))
            for r, employee in enumerate(employees):
                fixed_values = [
                    safe(employee["employee_id"]), safe(employee["name"]), safe(employee["department"]),
                    safe(employee["designation"]), safe(employee["date_joining"]), safe(employee["safety_induction_date"])
                ]
                for c, value in enumerate(fixed_values):
                    item = QTableWidgetItem(value)
                    if c == 0:
                        item.setData(Qt.ItemDataRole.UserRole, int(employee["id"]))
                    register_table.setItem(r, c, item)
                for offset, course in enumerate(courses, start=6):
                    item = QTableWidgetItem()
                    item.setData(Qt.ItemDataRole.UserRole, int(course["id"]))
                    record = latest_training(employee["id"], course["id"])
                    applicable = course_applies(course, employee)
                    if not applicable and not record:
                        item.setText("N/A")
                        item.setToolTip("Not applicable to this employee's trade/department. You can still click to record optional/event training.")
                        set_status_style(item, "N/A", course)
                    else:
                        status = training_status(record, course)
                        if record:
                            training_date_text = safe(record["training_date"])
                            expiry_text = safe(record["expiry_date"])
                            display = status + "\nTrained: " + training_date_text
                            if expiry_text:
                                display += "\nExpiry: " + expiry_text
                            item.setText(display)
                            item.setToolTip(
                                "Status: " + status + "\nTraining date: " + (training_date_text or "—") +
                                "\nExpiry date: " + (expiry_text or "No expiry recorded") +
                                "\nTraining type: " + safe(course["provider_type"]) +
                                "\nClick to edit this training record; previous renewals are retained."
                            )
                        else:
                            item.setText("Training Required")
                            item.setToolTip("No training record exists. Click to enter a completion date, including for recommended or optional training.")
                        set_status_style(item, status, course)
                    register_table.setItem(r, offset, item)
            register_table.resizeRowsToContents()
            for col in range(min(6, register_table.columnCount())):
                register_table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
            for col in range(6, register_table.columnCount()):
                register_table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
            counts = {"Training Required": 0, "Expired": 0, "Due Soon": 0, "Valid": 0, "Completed": 0}
            for r in range(register_table.rowCount()):
                for c in range(6, register_table.columnCount()):
                    cell = register_table.item(r, c)
                    if not cell:
                        continue
                    status = cell.toolTip().split("\n")[0].replace("Status: ", "") if cell.toolTip().startswith("Status: ") else cell.text().split("\n")[0]
                    if status in counts:
                        counts[status] += 1
            status_summary.setText(
                f"Employees: {len(employees)} | Active courses: {len(courses)} | "
                f"Training required: {counts['Training Required']} | Expired: {counts['Expired']} | "
                f"Due soon: {counts['Due Soon']} | Valid: {counts['Valid']} | No expiry: {counts['Completed']}"
            )
            apply_register_filter()

        def apply_register_filter(*args):
            query = register_search.text().strip().lower()
            for r in range(register_table.rowCount()):
                haystack = " ".join(register_table.item(r, c).text() for c in range(min(6, register_table.columnCount())) if register_table.item(r, c)).lower()
                register_table.setRowHidden(r, bool(query) and query not in haystack)

        def reset_employee_form():
            selected_employee_id["value"] = None
            employee_id_input.clear()
            employee_name_input.clear()
            employee_department_input.setCurrentIndex(0)
            employee_trade_input.setCurrentIndex(0)
            joining_date.setDate(QDate.currentDate())
            safety_induction_date.setDate(QDate(2000, 1, 1))
            add_employee_btn.setText("Save Employee")

        def select_employee_from_register(row, column=0):
            if column >= 6:
                return
            item = register_table.item(row, 0)
            if not item:
                return
            emp_db_id = item.data(Qt.ItemDataRole.UserRole)
            employee = db.fetchone("SELECT * FROM employees WHERE id=?", (emp_db_id,))
            if not employee:
                return
            selected_employee_id["value"] = int(emp_db_id)
            employee_id_input.setText(safe(employee["employee_id"]))
            employee_name_input.setText(safe(employee["name"]))
            dept = safe(employee["department"])
            trade = safe(employee["designation"])
            if dept and employee_department_input.findText(dept) < 0:
                employee_department_input.addItem(dept, dept)
            if trade and employee_trade_input.findText(trade) < 0:
                employee_trade_input.addItem(trade, trade)
            employee_department_input.setCurrentText(dept or "Select department")
            employee_trade_input.setCurrentText(trade or "Select trade")
            joined = QDate.fromString(safe(employee["date_joining"]), "yyyy-MM-dd")
            joining_date.setDate(joined if joined.isValid() else QDate.currentDate())
            induction = QDate.fromString(safe(employee["safety_induction_date"]), "yyyy-MM-dd")
            safety_induction_date.setDate(induction if induction.isValid() else QDate(2000, 1, 1))
            add_employee_btn.setText("Update Employee")

        def add_employee():
            name = employee_name_input.text().strip()
            emp_id = employee_id_input.text().strip()
            dept = employee_department_input.currentText().strip()
            trade = employee_trade_input.currentText().strip()
            if dept.lower() == "select department": dept = ""
            if trade.lower() == "select trade": trade = ""
            if not name:
                QMessageBox.warning(self, "Training Register", "Enter the employee name.")
                employee_name_input.setFocus()
                return
            duplicate = db.fetchone("SELECT id FROM employees WHERE lower(employee_id)=lower(?) AND id<>COALESCE(?, -1)", (emp_id, selected_employee_id["value"])) if emp_id else None
            if duplicate:
                QMessageBox.warning(self, "Training Register", "That employee ID already exists. Enter a different ID.")
                employee_id_input.setFocus()
                return
            induction = safety_induction_date.date().toString("yyyy-MM-dd")
            if induction == "2000-01-01": induction = ""
            values = (emp_id, name, trade, dept, joining_date.date().toString("yyyy-MM-dd"), induction)
            if selected_employee_id["value"] is None:
                db.execute("INSERT INTO employees (employee_id,name,designation,department,date_joining,safety_induction_date,active) VALUES (?,?,?,?,?,?,1)", values)
                message = "Employee saved and training applicability refreshed."
            else:
                db.execute("UPDATE employees SET employee_id=?,name=?,designation=?,department=?,date_joining=?,safety_induction_date=? WHERE id=?", values + (selected_employee_id["value"],))
                message = "Employee updated and training applicability refreshed."
            reset_employee_form()
            load_register()
            load_master_record()
            QMessageBox.information(self, "Training Register", message)

        def delete_selected_employee():
            row = register_table.currentRow()
            item = register_table.item(row, 0) if row >= 0 else None
            emp_db_id = item.data(Qt.ItemDataRole.UserRole) if item else selected_employee_id["value"]
            if emp_db_id is None:
                QMessageBox.information(self, "Training Register", "Select a person in the register first.")
                return
            history = db.fetchone("SELECT COUNT(*) AS n FROM training_records WHERE employee_db_id=?", (emp_db_id,))
            has_history = bool(history and int(history["n"] or 0))
            message = ("This person has training history. The person will be marked inactive and all training records will be retained. Continue?" if has_history else "Delete this person from the training register?")
            answer = QMessageBox.question(self, "Delete Person", message, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes: return
            if has_history:
                db.execute("UPDATE employees SET active=0 WHERE id=?", (emp_db_id,))
            else:
                db.execute("DELETE FROM employees WHERE id=?", (emp_db_id,))
            reset_employee_form()
            load_register()
            load_master_record()

        def edit_training_cell(row, column):
            if column < 6:
                return
            employee_item = register_table.item(row, 0)
            course_item = register_table.item(row, column)
            if not employee_item or not course_item:
                return
            employee_db_id = employee_item.data(Qt.ItemDataRole.UserRole)
            course_id = course_item.data(Qt.ItemDataRole.UserRole)
            employee = db.fetchone("SELECT * FROM employees WHERE id=?", (employee_db_id,))
            course = db.fetchone("SELECT * FROM training_courses WHERE id=?", (course_id,))
            if not employee or not course:
                return
            record = latest_training(employee_db_id, course_id)
            dialog = QDialog(self)
            dialog.setWindowTitle("Record Training — " + safe(course["course_name"]))
            dialog.setMinimumSize(540, 460)
            dialog.resize(620, 520)
            dlg_layout = QVBoxLayout(dialog)
            form = QFormLayout()
            employee_label = QLabel(f"{safe(employee['name'])} | {safe(employee['designation'])} | {safe(employee['department'])}")
            employee_label.setWordWrap(True)
            form.addRow("Employee", employee_label)
            training_date = QDateEdit()
            training_date.setCalendarPopup(True)
            training_date.setDisplayFormat("yyyy-MM-dd")
            training_date.setDate(QDate.currentDate())
            expiry_date = QDateEdit()
            expiry_date.setCalendarPopup(True)
            expiry_date.setDisplayFormat("yyyy-MM-dd")
            expiry_date.setDate(QDate.currentDate().addMonths(max(1, int(course["validity_months"] or 1))))
            no_expiry = QCheckBox("No expiry date")
            provider = QLineEdit()
            provider.setPlaceholderText("Internal trainer / external training organisation")
            provider_type_label = QLabel(safe(course["provider_type"]))
            cert_no = QLineEdit()
            cert_no.setPlaceholderText("Certificate or record reference")
            attachment = QLineEdit()
            attachment.setReadOnly(True)
            attachment.setPlaceholderText("Optional certificate file")
            browse_attachment = QPushButton("Browse…")
            attachment_row = QHBoxLayout()
            attachment_row.addWidget(attachment, 1)
            attachment_row.addWidget(browse_attachment)
            notes = QTextEdit()
            notes.setPlaceholderText("Assessment result, remarks or competency notes")
            notes.setMaximumHeight(85)
            form.addRow("Training type", provider_type_label)
            form.addRow("Training date", training_date)
            form.addRow("Expiry date", expiry_date)
            form.addRow("", no_expiry)
            form.addRow("Trainer / organisation", provider)
            form.addRow("Certificate number", cert_no)
            form.addRow("Attachment", attachment_row)
            form.addRow("Notes", notes)
            dlg_layout.addLayout(form)
            validity_months = int(course["validity_months"] or 0)

            def update_expiry_from_training(qdate=None):
                if validity_months <= 0 or no_expiry.isChecked():
                    return
                calculated = add_months(training_date.date().toString("yyyy-MM-dd"), validity_months)
                if calculated:
                    parsed = QDate.fromString(calculated, "yyyy-MM-dd")
                    if parsed.isValid():
                        expiry_date.setDate(parsed)

            training_date.dateChanged.connect(update_expiry_from_training)
            no_expiry.toggled.connect(lambda checked: expiry_date.setEnabled(not checked))
            no_expiry.setChecked(validity_months <= 0)
            if validity_months > 0:
                update_expiry_from_training()
            if record:
                parsed = QDate.fromString(safe(record["training_date"]), "yyyy-MM-dd")
                if parsed.isValid():
                    training_date.setDate(parsed)
                if safe(record["expiry_date"]):
                    parsed = QDate.fromString(safe(record["expiry_date"]), "yyyy-MM-dd")
                    if parsed.isValid():
                        expiry_date.setDate(parsed)
                    no_expiry.setChecked(False)
                else:
                    no_expiry.setChecked(True)
                provider.setText(safe(record["provider"]))
                cert_no.setText(safe(record["certificate_no"]))
                attachment.setText(safe(record["attachment_path"]))
                notes.setPlainText(safe(record["notes"]))
            browse_attachment.clicked.connect(lambda: attachment.setText(
                QFileDialog.getOpenFileName(dialog, "Select Training Certificate", "",
                "Documents and Images (*.pdf *.png *.jpg *.jpeg *.doc *.docx);;All Files (*)")[0] or attachment.text()))
            button_row = QHBoxLayout()
            save_btn = QPushButton("Save Training Record")
            delete_training_btn = QPushButton("Delete Training Record")
            delete_training_btn.setEnabled(bool(record))
            cancel_btn = QPushButton("Cancel")
            button_row.addStretch(1)
            button_row.addWidget(save_btn)
            button_row.addWidget(delete_training_btn)
            button_row.addWidget(cancel_btn)
            dlg_layout.addLayout(button_row)

            def save_training_record():
                date_text = training_date.date().toString("yyyy-MM-dd")
                expiry_text = "" if no_expiry.isChecked() else expiry_date.date().toString("yyyy-MM-dd")
                if expiry_text and expiry_text < date_text:
                    QMessageBox.warning(dialog, "Training Register", "Expiry date cannot be before the training date.")
                    return
                record_values = (date_text, expiry_text, provider.text().strip(), cert_no.text().strip(), attachment.text().strip(), notes.toPlainText().strip())
                if record:
                    db.execute("UPDATE training_records SET training_date=?,expiry_date=?,provider=?,certificate_no=?,attachment_path=?,notes=? WHERE id=?", record_values + (record["id"],))
                else:
                    db.execute("INSERT INTO training_records (employee_db_id,course_id,training_date,expiry_date,provider,certificate_no,attachment_path,notes) VALUES (?,?,?,?,?,?,?,?)", (employee_db_id, course_id) + record_values)
                dialog.accept()
                load_register()
                load_master_record()

            def delete_training_record():
                if not record:
                    return
                answer = QMessageBox.question(dialog, "Delete Training Record", "Delete this employee's training record for this course? This cannot be undone.", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if answer == QMessageBox.StandardButton.Yes:
                    db.execute("DELETE FROM training_records WHERE id=?", (record["id"],))
                    dialog.accept()
                    load_register()
                    load_master_record()

            save_btn.clicked.connect(save_training_record)
            delete_training_btn.clicked.connect(delete_training_record)
            cancel_btn.clicked.connect(dialog.reject)
            dialog.exec()

        def master_rows():
            result = []
            employees = get_employees()
            courses = expanded_course_entries(db.fetchall("SELECT * FROM training_courses ORDER BY course_name COLLATE NOCASE"))
            for employee in employees:
                applicable_found = False
                for course in courses:
                    record = latest_training(employee["id"], course["id"])
                    if not course_applies(course, employee) and not record:
                        continue
                    applicable_found = True
                    status = training_status(record, course)
                    result.append([
                        safe(employee["employee_id"]), safe(employee["name"]), safe(employee["department"]),
                        safe(employee["designation"]), safe(employee["date_joining"]), safe(employee["safety_induction_date"]),
                        safe(course["display_course_name"]), safe(course["provider_type"]), safe(course["requirement_source"]),
                        safe(record["training_date"]) if record else "", safe(record["expiry_date"]) if record else "",
                        status, safe(record["certificate_no"]) if record else "", safe(record["attachment_path"]) if record else ""
                    ])
                if not applicable_found:
                    result.append([
                        safe(employee["employee_id"]), safe(employee["name"]), safe(employee["department"]),
                        safe(employee["designation"]), safe(employee["date_joining"]), safe(employee["safety_induction_date"]),
                        "No applicable course configured", "", "", "", "", "N/A", "", ""
                    ])
            return result

        def load_master_record():
            rows = master_rows()
            master_table.setRowCount(len(rows))
            for r, values in enumerate(rows):
                for c, value in enumerate(values):
                    item = QTableWidgetItem(str(value or ""))
                    if c == 9 and values[6] != "No applicable course configured":
                        employee_match = db.fetchone("SELECT id FROM employees WHERE COALESCE(employee_id,'')=? AND name=? LIMIT 1", (values[0], values[1]))
                        display_course = safe(values[6])
                        course_name_base = display_course.split(" (", 1)[0]
                        trade_suffix = display_course[len(course_name_base):].strip()
                        trade_suffix = trade_suffix[1:-1].strip() if trade_suffix.startswith("(") and trade_suffix.endswith(")") else ""
                        course_match = db.fetchone("SELECT id FROM training_courses WHERE course_name=? AND (COALESCE(designation,'')=? OR COALESCE(designation,'')='' OR designation LIKE ?) ORDER BY active DESC LIMIT 1", (course_name_base, trade_suffix, '%' + trade_suffix + '%'))
                        if employee_match and course_match:
                            rec_match = latest_training(employee_match["id"], course_match["id"])
                            if rec_match:
                                item.setData(Qt.ItemDataRole.UserRole, int(rec_match["id"]))
                    if c == 11:
                        course = {"provider_type": values[7]} if values[6] != "No applicable course configured" else None
                        set_status_style(item, values[11], course)
                    master_table.setItem(r, c, item)
            master_table.resizeRowsToContents()
            apply_master_filter()

        def delete_selected_master_training_record():
            row = master_table.currentRow()
            item = master_table.item(row, 9) if row >= 0 else None
            record_id = item.data(Qt.ItemDataRole.UserRole) if item else None
            if record_id is None:
                QMessageBox.information(self, "Master Training Record", "Select a row with an existing training date first.")
                return
            answer = QMessageBox.question(self, "Delete Training Record", "Delete this training record? The employee and training course will remain.", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if answer == QMessageBox.StandardButton.Yes:
                db.execute("DELETE FROM training_records WHERE id=?", (record_id,))
                load_register()
                load_master_record()

        def apply_master_filter(*args):
            query = master_search.text().strip().lower()
            for r in range(master_table.rowCount()):
                haystack = " ".join(master_table.item(r, c).text() for c in range(master_table.columnCount()) if master_table.item(r, c)).lower()
                master_table.setRowHidden(r, bool(query) and query not in haystack)

        def style_excel(ws, title, headers):
            ws.append([title])
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, len(headers)))
            ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
            ws["A1"].fill = PatternFill("solid", fgColor="17365D")
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 28
            ws.append(headers)
            for cell in ws[2]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="245486")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin = Side(style="thin", color="D5DADF")
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
            ws.freeze_panes = "A3"
            ws.auto_filter.ref = f"A2:{ws.cell(row=max(2, ws.max_row), column=len(headers)).coordinate}"
            ws.sheet_view.showGridLines = False
            for col in range(1, len(headers) + 1):
                ws.column_dimensions[ws.cell(row=2, column=col).column_letter].width = 22

        def finish_excel(ws, headers):
            thin = Side(style="thin", color="D5DADF")
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
            ws.freeze_panes = "A3"
            ws.auto_filter.ref = f"A2:{ws.cell(row=max(2, ws.max_row), column=len(headers)).coordinate}"
            ws.sheet_view.showGridLines = False

        def export_training_register():
            path, _ = QFileDialog.getSaveFileName(self, "Export Training Register", "Training_Register.xlsx", "Excel Workbook (*.xlsx)")
            if not path:
                return
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            employees = get_employees()
            courses = expanded_course_entries(db.fetchall("SELECT * FROM training_courses WHERE active=1 OR EXISTS (SELECT 1 FROM training_records tr WHERE tr.course_id=training_courses.id) ORDER BY active DESC, course_name COLLATE NOCASE"))
            wb = Workbook()
            ws = wb.active
            ws.title = "Training Register"
            headers = ["Employee ID", "Employee Name", "Department", "Trade", "Date Joined", "Safety Induction Date"] + [safe(c["display_course_name"]) for c in courses]
            style_excel(ws, "HSE TRAINING REGISTER", headers)
            register_attachment_links = []
            for employee in employees:
                values = [safe(employee["employee_id"]), safe(employee["name"]), safe(employee["department"]), safe(employee["designation"]), safe(employee["date_joining"]), safe(employee["safety_induction_date"])]
                employee_links = []
                for course in courses:
                    if not course_applies(course, employee):
                        values.append("N/A")
                        employee_links.append(None)
                        continue
                    record = latest_training(employee["id"], course["id"])
                    status = training_status(record, course)
                    if record:
                        text_value = status + "\nTraining: " + safe(record["training_date"])
                        if safe(record["expiry_date"]):
                            text_value += "\nExpiry: " + safe(record["expiry_date"])
                        attachment_path = safe(record["attachment_path"]).strip()
                        if attachment_path:
                            text_value += "\nAttachment: " + Path(attachment_path).name
                        values.append(text_value)
                        employee_links.append(attachment_path or None)
                    else:
                        values.append("Training Required")
                        employee_links.append(None)
                ws.append(values)
                register_attachment_links.append(employee_links)
            for col in range(1, len(headers) + 1):
                ws.column_dimensions[ws.cell(row=2, column=col).column_letter].width = 20 if col <= 6 else 24
            ws.freeze_panes = "G3"
            logo_path = report_logo_path()
            if logo_path and Path(logo_path).exists():
                try:
                    logo = XLImage(str(logo_path)); logo.width = 115; logo.height = 58; logo.anchor = "A1"; ws.add_image(logo)
                    ws.row_dimensions[1].height = max(ws.row_dimensions[1].height or 15, 48)
                except Exception:
                    logging.exception("Unable to embed Training Register Excel logo")
            # Preserve each original certificate attachment as a clickable link in its employee/course cell.
            for row_offset, employee_links in enumerate(register_attachment_links, start=3):
                for course_offset, attachment_path in enumerate(employee_links, start=7):
                    if attachment_path:
                        cell = ws.cell(row=row_offset, column=course_offset)
                        cell.hyperlink = Path(attachment_path).resolve().as_uri() if Path(attachment_path).exists() else attachment_path
                        cell.style = "Hyperlink"
            finish_excel(ws, headers)
            try:
                wb.save(path)
                QMessageBox.information(self, "Training Register", f"Register exported successfully:\n{path}")
            except Exception as exc:
                QMessageBox.critical(self, "Training Register", f"Could not export Excel file:\n{exc}")

        def export_training_matrix():
            path, _ = QFileDialog.getSaveFileName(self, "Export Training Matrix", "Training_Matrix.xlsx", "Excel Workbook (*.xlsx)")
            if not path:
                return
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            rows = db.fetchall("SELECT * FROM training_courses ORDER BY course_name COLLATE NOCASE")
            headers = ["Course", "Applicable Trade(s)", "Department", "Mandatory", "Training Type", "Validity (months)", "Requirement Source", "Certificate Required", "Status"]
            wb = Workbook()
            ws = wb.active
            ws.title = "Training Matrix"
            style_excel(ws, "HSE TRAINING MATRIX", headers)
            for row in rows:
                trades_for_export = [x.strip() for x in safe(row["designation"]).split(",") if x.strip()]
                trade_cell = "All Trades" if not trades_for_export else "\n• " + "\n• ".join(trades_for_export)
                ws.append([safe(row["course_name"]), trade_cell, safe(row["department"]) or "All Departments", "Mandatory" if row["mandatory"] else "Recommended", safe(row["provider_type"]), int(row["validity_months"] or 0), safe(row["requirement_source"]), "Yes" if row["certificate_required"] else "No", "Active" if row["active"] else "Inactive"])
            logo_path = report_logo_path()
            if logo_path and Path(logo_path).exists():
                try:
                    logo = XLImage(str(logo_path)); logo.width = 115; logo.height = 58; logo.anchor = "A1"; ws.add_image(logo)
                    ws.row_dimensions[1].height = max(ws.row_dimensions[1].height or 15, 48)
                except Exception:
                    logging.exception("Unable to embed Training Matrix Excel logo")
            finish_excel(ws, headers)
            try:
                wb.save(path)
                QMessageBox.information(self, "Training Matrix", f"Matrix exported successfully:\n{path}")
            except Exception as exc:
                QMessageBox.critical(self, "Training Matrix", f"Could not export Excel file:\n{exc}")

        def export_master_excel():
            path, _ = QFileDialog.getSaveFileName(self, "Export Master Training Record", "Master_Training_Record.xlsx", "Excel Workbook (*.xlsx)")
            if not path:
                return
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            headers = [master_table.horizontalHeaderItem(i).text() for i in range(master_table.columnCount())]
            rows = master_rows()
            wb = Workbook()
            ws = wb.active
            ws.title = "Master Training Record"
            style_excel(ws, "MASTER TRAINING RECORD", headers)
            for row in rows:
                ws.append(row)
            logo_path = report_logo_path()
            if logo_path and Path(logo_path).exists():
                try:
                    logo = XLImage(str(logo_path)); logo.width = 115; logo.height = 58; logo.anchor = "A1"; ws.add_image(logo)
                    ws.row_dimensions[1].height = max(ws.row_dimensions[1].height or 15, 48)
                except Exception:
                    logging.exception("Unable to embed Master Training Record Excel logo")
            attachment_col = headers.index("Attachment") + 1 if "Attachment" in headers else None
            if attachment_col:
                for row_idx in range(3, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=attachment_col)
                    attachment_path = str(cell.value or "").strip()
                    if attachment_path:
                        cell.value = Path(attachment_path).name
                        cell.hyperlink = Path(attachment_path).resolve().as_uri() if Path(attachment_path).exists() else attachment_path
                        cell.style = "Hyperlink"
            finish_excel(ws, headers)
            try:
                wb.save(path)
                QMessageBox.information(self, "Master Training Record", f"Excel export completed:\n{path}")
            except Exception as exc:
                QMessageBox.critical(self, "Master Training Record", f"Could not export Excel file:\n{exc}")

        def export_master_word():
            path, _ = QFileDialog.getSaveFileName(self, "Export Master Training Record", "Master_Training_Record.docx", "Word Document (*.docx)")
            if not path:
                return
            if not path.lower().endswith(".docx"):
                path += ".docx"
            headers = [master_table.horizontalHeaderItem(i).text() for i in range(master_table.columnCount())]
            rows = master_rows()
            doc = Document()
            section = doc.sections[0]
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width, section.page_height = section.page_height, section.page_width
            doc.add_heading("Master Training Record", 0)
            doc.add_paragraph(f"Organisation: {safe(company_name())} | Generated: {date.today().isoformat()}")
            table = doc.add_table(rows=1, cols=len(headers))
            table.style = "Table Grid"
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            for i, header in enumerate(headers):
                table.rows[0].cells[i].text = header
            for row in rows:
                cells = table.add_row().cells
                for i, value in enumerate(row):
                    cells[i].text = str(value or "")
                    cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for cell in table.rows[0].cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
                        run.font.size = Pt(8)
            for row in table.rows[1:]:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(7)
            try:
                doc.save(path)
                QMessageBox.information(self, "Master Training Record", f"Word export completed:\n{path}")
            except Exception as exc:
                QMessageBox.critical(self, "Master Training Record", f"Could not export Word file:\n{exc}")

        def filter_course_rows(*args):
            return

        save_course_btn.clicked.connect(save_course)
        clear_course_btn.clicked.connect(reset_course_form)
        deactivate_course_btn.clicked.connect(deactivate_course)
        export_matrix_btn.clicked.connect(export_training_matrix)
        refresh_matrix_btn.clicked.connect(load_courses)
        course_table.cellClicked.connect(select_course_row)
        add_employee_btn.clicked.connect(add_employee)
        edit_employee_btn.clicked.connect(lambda: select_employee_from_register(register_table.currentRow(), 0) if register_table.currentRow() >= 0 else QMessageBox.information(self, "Training Register", "Select a person first."))
        delete_employee_btn.clicked.connect(delete_selected_employee)
        refresh_register_btn.clicked.connect(load_register)
        export_register_btn.clicked.connect(export_training_register)
        register_search.textChanged.connect(apply_register_filter)
        register_table.cellClicked.connect(lambda row, column: select_employee_from_register(row, column) if column < 6 else edit_training_cell(row, column))
        refresh_master_btn.clicked.connect(load_master_record)
        delete_master_record_btn.clicked.connect(delete_selected_master_training_record)
        export_master_excel_btn.clicked.connect(export_master_excel)
        export_master_word_btn.clicked.connect(export_master_word)
        master_search.textChanged.connect(apply_master_filter)

        load_courses()
        tabs.setCurrentIndex(0)

    def reports(self):
        w, layout = self.page("Reports & Export")
        banner = QHBoxLayout()
        logo_path = db.setting("company_logo", "")
        logo = QLabel()
        logo.setFixedSize(90, 70)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        if logo_path and Path(logo_path).exists():
            pix = QPixmap(logo_path)
            logo.setPixmap(pix.scaled(80, 65, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        banner.addWidget(logo)
        company_label = QLabel(f"<b>{safe(company_name())}</b><br>Reports & Export<br>Project: {safe(db.setting('project_name',''))}")
        company_label.setStyleSheet("font-size:16px;padding:6px;")
        banner.addWidget(company_label, 1)
        open_folder = QPushButton("Open Reports Folder")
        open_folder.clicked.connect(self.open_reports_folder)
        banner.addWidget(open_folder)
        layout.addLayout(banner)

        info = QLabel("Select a register and format. The report is extracted from the saved HSE database and can be saved to any folder on the computer.")
        info.setWordWrap(True)
        layout.addWidget(info)

        for title, table_name in [
            ("HSE Inspection / Observation Report", "observations"),
            ("Incident Investigation Register Report", "incidents"),
            ("Audit Register Report", "audits"),
            ("CAPA Register Report", "capa")
        ]:
            box = QGroupBox(title)
            row = QHBoxLayout(box)
            for label, handler in [
                ("Download CSV", lambda t=table_name: self.export_csv(t)),
                ("Download Excel", lambda t=table_name: self.export_excel(t)),
                ("Download PDF", lambda t=table_name: self.export_pdf(t)),
                ("Download Word", lambda t=table_name: self.export_docx(t))
            ]:
                b = QPushButton(label)
                b.setMinimumHeight(38)
                row.addWidget(b)
                b.clicked.connect(handler)
            layout.addWidget(box)

        incident_box = QGroupBox("Professional Incident Investigation Word Report")
        ir = QHBoxLayout(incident_box)
        incident_combo = QComboBox()
        incidents = db.fetchall("SELECT id,number FROM incidents ORDER BY id DESC")
        for row in incidents:
            incident_combo.addItem(safe(row["number"]), row["id"])
        ir.addWidget(QLabel("Investigation:"))
        ir.addWidget(incident_combo, 1)
        b = QPushButton("Download Investigation Report")
        b.setMinimumHeight(38)
        ir.addWidget(b)

        def generate():
            if incident_combo.count() == 0:
                QMessageBox.information(self, "No Incidents", "No incident investigations are available.")
                return
            incident_id = incident_combo.currentData()
            number = incident_combo.currentText()
            default_path = str(REPORT_DIR / f"{number}_Investigation_Report.docx")
            path, _ = QFileDialog.getSaveFileName(
                self, "Save Investigation Report", default_path, "Word Document (*.docx)"
            )
            if not path:
                return
            try:
                generate_incident_docx(int(incident_id), path)
                self.show_export_success(path)
            except Exception as e:
                logging.exception("Incident report failed")
                QMessageBox.critical(self, "Report Error", str(e))

        b.clicked.connect(generate)
        layout.addWidget(incident_box)
        layout.addStretch(1)

    def show_export_success(self, path):
        msg = QMessageBox(self)
        msg.setWindowTitle("Export Complete")
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText("Report created successfully.")
        msg.setInformativeText(str(path))
        open_button = msg.addButton("Open Folder", QMessageBox.ButtonRole.ActionRole)
        msg.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        msg.exec()
        if msg.clickedButton() == open_button:
            self.open_path(Path(path).parent)

    def open_path(self, path):
        try:
            path = Path(path)
            if sys.platform.startswith("win"):
                os.startfile(str(path))
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception as e:
            logging.exception("Unable to open path")
            QMessageBox.warning(self, "Open Folder", f"Unable to open folder.\n\n{e}")

    def open_reports_folder(self):
        REPORT_DIR.mkdir(parents=True, exist_ok=True)
        self.open_path(REPORT_DIR)

    def _export_path(self, filename, title, file_filter):
        REPORT_DIR.mkdir(parents=True, exist_ok=True)
        return QFileDialog.getSaveFileName(self, title, str(REPORT_DIR / filename), file_filter)[0]

    @staticmethod
    def table_data_static(table_name, record_ids=None):
        allowed = {"observations", "incidents", "audits", "capa"}
        if table_name not in allowed:
            raise ValueError("Invalid report table.")
        if table_name == "observations":
            columns, data, _ = observation_export_rows(record_ids)
            return columns, data
        params = []
        where = ""
        if record_ids:
            placeholders = ",".join("?" for _ in record_ids)
            where = f" WHERE id IN ({placeholders})"
            params = list(record_ids)
        rows = db.fetchall(f"SELECT * FROM {table_name}{where} ORDER BY id DESC", params)
        if not rows:
            return [], []
        columns = list(rows[0].keys())
        attach_table = attachment_export_map(table_name)
        if attach_table and "Attachments" not in columns:
            columns.append("Attachments")
        data=[]
        id_key="id"
        for row in rows:
            values=[xml_safe(row[c]) for c in columns if c != "Attachments"]
            if attach_table:
                values.append(xml_safe(attachment_names(attach_table, row[id_key])))
            data.append(values)
        return columns, data

    def table_data(self, table_name, record_ids=None):
        return self.table_data_static(table_name, record_ids=record_ids)

    def selected_id(self, table, title):
        r=table.currentRow()
        if r < 0 or not table.item(r,0):
            QMessageBox.warning(self, title, "Select a record first.")
            return None
        return int(table.item(r,0).text())

    def delete_selected_record(self, table, table_name, attachment_table, title, refresh):
        record_id=self.selected_id(table, title)
        if record_id is None:
            return
        record=db.fetchone(f"SELECT number FROM {table_name} WHERE id=?", (record_id,))
        if not record:
            return
        if QMessageBox.question(self, "Confirm Delete", f"Delete {title.lower()} {record['number']} and its evidence?\n\nThis cannot be undone.") != QMessageBox.Yes:
            return
        for a in attachment_rows(attachment_table, record_id):
            try: Path(safe(a["file_path"])).unlink(missing_ok=True)
            except Exception: logging.exception("Unable to remove attachment")
        db.execute(f"DELETE FROM {table_name} WHERE id=?", (record_id,))
        refresh()

    def export_csv(self, table_name, record_ids=None):
        try:
            columns, data = self.table_data(table_name, record_ids)
            if not columns:
                QMessageBox.information(self, "No Data", "There is no data to export.")
                return
            path = self._export_path(f"{table_name}_report.csv", "Download CSV Report", "CSV Files (*.csv)")
            if not path: return
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer=csv.writer(f); writer.writerow(columns); writer.writerows(data)
            self.show_export_success(path)
        except Exception as e:
            logging.exception("CSV export failed"); QMessageBox.critical(self,"Export Error",str(e))

    def export_excel(self, table_name, record_ids=None):
        try:
            columns, data = self.table_data(table_name, record_ids)
            if not columns:
                QMessageBox.information(self, "No Data", "There is no data to export."); return
            path=self._export_path(f"{table_name}_report.xlsx","Download Excel Report","Excel Files (*.xlsx)")
            if not path: return
            wb=Workbook(); ws=wb.active; ws.title=table_name[:31]
            if table_name == "observations":
                preferred = ["S.No.", "Observation No.", "Observation Date", "Location", "Observation Type", "Category", "Observation Detail", "Action Taken", "Responsible", "Evidence", "Status"]
                # Professional three-row report header.  The layout is fixed to A1:K3:
                # A1:C3 = logo, D1:H3 = report heading, I1:K1 = report number,
                # I2:K2 = document number, I3:K3 = project name.
                ws.merge_cells("A1:C3")
                ws.merge_cells("D1:H3")
                ws.merge_cells("I1:K1")
                ws.merge_cells("I2:K2")
                ws.merge_cells("I3:K3")
                ws["D1"] = f"HSE OBSERVATION REPORT\n{company_name()}"
                report_no = f"{document_prefix()}-OBS-REPORT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                ws["I1"] = f"Report No.: {report_no}"
                ws["I2"] = f"Document No.: {document_prefix()}-OBS"
                ws["I3"] = f"Project: {db.setting('project_name', '')}"
                ws.append(preferred)
                obs_rows = db.fetchall("SELECT * FROM observations ORDER BY id DESC")
                if record_ids:
                    obs_rows = [r for r in obs_rows if r["id"] in record_ids]
                for serial, r in enumerate(obs_rows, 1):
                    vals=[serial, r["number"], r["obs_date"], r["location"], r["obs_type"], r["category"], r["observation"], format_corrective_actions_display(r), r["responsible"], "", r["status"]]
                    ws.append([xml_safe(v) for v in vals])
                    excel_row=ws.max_row
                    evcell=ws.cell(excel_row,10)
                    evcell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
                    image_files=[]; other_files=[]
                    for a in attachment_rows("observation_attachments", r["id"]):
                        fp=Path(safe(a["file_path"]))
                        if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp"}:
                            image_files.append(fp)
                        elif fp.exists():
                            other_files.append(fp.name)
                    # Show the first real image directly in the Evidence cell; all other evidence names remain visible.
                    if image_files:
                        try:
                            img=XLImage(str(image_files[0]))
                            max_w=145; max_h=62
                            ratio=min(max_w/max(img.width,1), max_h/max(img.height,1), 1)
                            img.width=max(1,int(img.width*ratio)); img.height=max(1,int(img.height*ratio))
                            img.anchor=f"J{excel_row}"
                            ws.add_image(img)
                            extra=[p.name for p in image_files[1:]] + other_files
                            evcell.value="; ".join(extra) if extra else "Photo evidence"
                        except Exception:
                            logging.exception("Unable to embed observation image in Excel")
                            evcell.value="; ".join([p.name for p in image_files] + other_files)
                    elif other_files:
                        evcell.value="; ".join(other_files)
                    else:
                        evcell.value="None"
                    ws.row_dimensions[excel_row].height=70
                # Professional register heading. The logo is fixed inside the merged A1:C3 area.
                logo_path=db.setting("company_logo", "")
                if logo_path and Path(logo_path).exists():
                    try:
                        logo=XLImage(logo_path)
                        logo.width=125; logo.height=62
                        logo.anchor="A1"
                        ws.add_image(logo)
                    except Exception:
                        logging.exception("Unable to embed company logo in Excel heading")
                # Header area is deliberately fixed to three rows; the register header begins on row 4.
                for r in range(1,4):
                    ws.row_dimensions[r].height = 24
                ws.row_dimensions[4].height = 30
                header_fill=PatternFill("solid", fgColor="D9E2F3")
                for cell_ref in ["A1","D1","I1","I2","I3"]:
                    c=ws[cell_ref]
                    c.font=copy(c.font); c.font=c.font.copy(bold=True, size=13 if cell_ref=="D1" else 10)
                    c.fill=header_fill
                    c.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
                    c.border=Border(left=Side(style="thin", color="808080"), right=Side(style="thin", color="808080"), top=Side(style="thin", color="808080"), bottom=Side(style="thin", color="808080"))
                # Give every cell in the merged header areas a border so the visual box surrounds the logo/title.
                thin=Side(style="thin", color="808080")
                for row in ws.iter_rows(min_row=1,max_row=3,min_col=1,max_col=11):
                    for c in row:
                        c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
                for cell in ws[4]:
                    cell.font=copy(cell.font); cell.font=cell.font.copy(bold=True, color="FFFFFF")
                    cell.fill=PatternFill("solid", fgColor="17365D")
                    cell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
                # All data cells are centered/middle and wrapped; long narrative fields get extra width.
                for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=1, max_col=len(preferred)):
                    for cell in row:
                        cell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
                widths=[8,20,20,20,20,20,50,50,20,20,20]
                for i,w in enumerate(widths,1):
                    ws.column_dimensions[chr(64+i)].width=w
                ws.freeze_panes="A5"
                ws.auto_filter.ref=f"A4:K{ws.max_row}"
                # Real Excel table for the register rows.
                from openpyxl.worksheet.table import Table, TableStyleInfo
                if ws.max_row >= 4:
                    tab=Table(displayName="HSEObservationRegister", ref=f"A4:K{ws.max_row}")
                    style=TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                    tab.tableStyleInfo=style
                    ws.add_table(tab)
                target=Path(path); temp=target.with_name(target.name + ".tmp.xlsx")
                try:
                    wb.save(temp)
                    with zipfile.ZipFile(temp,"r") as z:
                        bad=z.testzip()
                        if bad: raise ValueError(f"Generated Excel report is corrupt: {bad}")
                    os.replace(temp,target)
                finally: temp.unlink(missing_ok=True)
                self.show_export_success(path); return
            wb=Workbook(); ws=wb.active; ws.title=table_name[:31]; ws.append(columns)
            for row in data: ws.append([xml_safe(v) for v in row])
            # Use the company logo selected in Settings without inserting rows or changing
            # the report's column structure. The image is placed at the far-right header area.
            logo_path=report_logo_path()
            if logo_path and Path(logo_path).exists():
                try:
                    logo=XLImage(str(logo_path))
                    scale=min(90/max(logo.width,1), 42/max(logo.height,1), 1)
                    logo.width=max(1,int(logo.width*scale)); logo.height=max(1,int(logo.height*scale))
                    logo.anchor=f"{ws.cell(1, max(1,len(columns)+1)).column_letter}1"
                    ws.add_image(logo)
                    ws.row_dimensions[1].height=max(ws.row_dimensions[1].height or 15, 34)
                except Exception:
                    logging.exception("Unable to embed company logo in Excel report")
            for cell in ws[1]:
                cell.font=copy(cell.font); cell.font=cell.font.copy(bold=True)
            ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
            if "Attachments" in columns:
                col=columns.index("Attachments")+1
                for r in range(2, ws.max_row+1):
                    ws.cell(r,col).alignment=Alignment(wrap_text=True, vertical="top")
                    # Make the attachment cell readable; the dedicated Attachments sheet below contains file-level hyperlinks.
                    ws.cell(r,col).hyperlink = f"#'Attachments'!A1"
            for column_cells in ws.columns:
                letter=column_cells[0].column_letter
                max_len=min(55,max(len(xml_safe(c.value)) if c.value is not None else 0 for c in column_cells)+2)
                ws.column_dimensions[letter].width=max(12,max_len)

            # Include a dedicated attachment register in the Excel report so evidence is
            # not lost when the main register has many columns.
            attach_table=attachment_export_map(table_name)
            if attach_table:
                ids=[r["id"] for r in db.fetchall(f"SELECT id FROM {table_name} ORDER BY id DESC") if not record_ids or r["id"] in record_ids]
                aws=wb.create_sheet("Attachments")
                aws.append(["Record ID","Record Number","File Name","Attachment Type","Stored File Path","Embedded Preview"])
                for rid in ids:
                    rec=db.fetchone(f"SELECT number FROM {table_name} WHERE id=?",(rid,))
                    for a in attachment_rows(attach_table,rid):
                        fpath=safe(a["file_path"])
                        aws.append([rid, safe(rec["number"]) if rec else "", Path(fpath).name, xml_safe(a["attachment_type"]), fpath, ""])
                        current_row=aws.max_row
                        cell=aws.cell(current_row,5)
                        if fpath and Path(fpath).exists():
                            cell.hyperlink=Path(fpath).as_uri()
                            cell.style="Hyperlink"
                            fp=Path(fpath)
                            if fp.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp", ".gif"}:
                                try:
                                    embedded=XLImage(str(fp))
                                    scale=min(100/max(embedded.width,1), 65/max(embedded.height,1), 1)
                                    embedded.width=max(1,int(embedded.width*scale)); embedded.height=max(1,int(embedded.height*scale))
                                    embedded.anchor=f"F{current_row}"
                                    aws.add_image(embedded)
                                    aws.row_dimensions[current_row].height=52
                                except Exception:
                                    logging.exception("Unable to embed attachment preview in Excel report")
                aws.freeze_panes="A2"; aws.auto_filter.ref=aws.dimensions
                for column_cells in aws.columns:
                    letter=column_cells[0].column_letter
                    max_len=min(80,max(len(xml_safe(c.value)) if c.value is not None else 0 for c in column_cells)+2)
                    aws.column_dimensions[letter].width=max(12,max_len)
                aws.column_dimensions["F"].width=18
            target=Path(path)
            temp=target.with_name(target.name + ".tmp.xlsx")
            try:
                wb.save(temp)
                with zipfile.ZipFile(temp, "r") as z:
                    bad=z.testzip()
                    if bad: raise ValueError(f"Generated Excel report is corrupt: {bad}")
                os.replace(temp,target)
            finally:
                temp.unlink(missing_ok=True)
            self.show_export_success(path)
        except Exception as e:
            logging.exception("Excel export failed"); QMessageBox.critical(self,"Export Error",str(e))

    def export_pdf(self, table_name, record_ids=None):
        try:
            columns,data=self.table_data(table_name, record_ids)
            if not columns:
                QMessageBox.information(self,"No Data","There is no data to export."); return
            path=self._export_path(f"{table_name}_report.pdf","Download PDF Report","PDF Files (*.pdf)")
            if not path:return
            styles=getSampleStyleSheet()
            project=db.setting("project_name","")
            if table_name == "observations":
                # Observation PDF header mirrors the fixed Excel report header.
                def draw_page_border(canvas, doc):
                    canvas.saveState()
                    canvas.setStrokeColor(colors.HexColor("#808080"))
                    canvas.setLineWidth(0.8)
                    page_w,page_h=landscape(A4)
                    canvas.rect(12,12,page_w-24,page_h-24,stroke=1,fill=0)
                    canvas.restoreState()
                doc=SimpleDocTemplate(path,pagesize=landscape(A4),rightMargin=22,leftMargin=22,topMargin=92,bottomMargin=24,
                                      onFirstPage=draw_page_border,onLaterPages=draw_page_border)
                story=[]
                report_no=f"{document_prefix()}-OBS-REPORT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                logo_path=report_logo_path()
                logo_flow=Paragraph(xml_safe(company_name()),styles["Normal"])
                if logo_path and Path(logo_path).exists():
                    try:
                        logo_flow=RLImage(str(logo_path),width=85,height=55,preserveAspectRatio=True)
                    except Exception:
                        logging.exception("Unable to embed company logo in observation PDF header")
                title_style=styles["Title"]
                title_style.fontSize=15; title_style.leading=17; title_style.alignment=1
                meta_style=styles["Normal"]
                meta_style.fontSize=7.5; meta_style.leading=9
                header_right=Table([
                    [Paragraph(f"<b>Report No.:</b> {xml_safe(report_no)}",meta_style)],
                    [Paragraph(f"<b>Document No.:</b> {xml_safe(document_prefix())}-OBS",meta_style)],
                    [Paragraph(f"<b>Project:</b> {xml_safe(project)}",meta_style)],
                ], colWidths=[185], rowHeights=[20,20,20])
                header_right.setStyle(TableStyle([
                    ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#808080")),
                    ("INNERGRID",(0,0),(-1,-1),0.3,colors.HexColor("#BFBFBF")),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                    ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),
                ]))
                header=Table([[logo_flow,Paragraph(f"<b>HSE OBSERVATION REPORT</b><br/>{xml_safe(company_name())}",title_style),header_right]],
                             colWidths=[130,455,185], rowHeights=[60])
                header.setStyle(TableStyle([
                    ("BOX",(0,0),(-1,-1),0.7,colors.HexColor("#808080")),
                    ("INNERGRID",(0,0),(-1,-1),0.4,colors.HexColor("#BFBFBF")),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                    ("ALIGN",(0,0),(-1,-1),"CENTER"),
                    ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),
                ]))
                story.append(header)
                story.append(Spacer(1,8))
            else:
                def draw_page_border(canvas, doc):
                    canvas.saveState(); canvas.setStrokeColor(colors.HexColor("#808080")); canvas.setLineWidth(0.8); canvas.rect(12,12,landscape(A4)[0]-24,landscape(A4)[1]-24,stroke=1,fill=0); canvas.restoreState()
                doc=SimpleDocTemplate(path,pagesize=landscape(A4),rightMargin=22,leftMargin=22,topMargin=90,bottomMargin=24,onFirstPage=draw_page_border,onLaterPages=draw_page_border)
                story=[]
                report_no=f"{document_prefix()}-{table_name[:3].upper()}-REPORT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                logo_path=report_logo_path()
                logo_flow=Paragraph(xml_safe(company_name()),styles["Normal"])
                if logo_path and Path(logo_path).exists():
                    try: logo_flow=RLImage(str(logo_path),width=85,height=55,preserveAspectRatio=True)
                    except Exception: pass
                ms=styles["Normal"]; ms.fontSize=7.5; ms.leading=9
                meta=Table([[Paragraph(f"<b>Report No.:</b> {xml_safe(report_no)}",ms)],[Paragraph(f"<b>Document No.:</b> {xml_safe(document_prefix())}-{table_name[:3].upper()}",ms)],[Paragraph(f"<b>Project:</b> {xml_safe(project)}",ms)]],colWidths=[185],rowHeights=[20,20,20])
                meta.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#808080")),("INNERGRID",(0,0),(-1,-1),0.3,colors.HexColor("#BFBFBF")),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
                head=Table([[logo_flow,Paragraph(f"<b>{xml_safe(table_name.title())} REPORT</b><br/>{xml_safe(company_name())}",styles["Title"]),meta]],colWidths=[130,455,185],rowHeights=[60])
                head.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.7,colors.HexColor("#808080")),("INNERGRID",(0,0),(-1,-1),0.4,colors.HexColor("#BFBFBF")),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(-1,-1),"CENTER")]))
                story += [head,Spacer(1,8)]
            if table_name == "observations":
                preferred = ["S.No.", "Observation Date", "Location", "Observation Type", "Category", "Observation Detail", "Action Taken", "Responsible", "Evidence", "Status"]
                # Build a visible evidence cell using real images, not hyperlinks.
                obs_rows=db.fetchall("SELECT * FROM observations ORDER BY id DESC")
                if record_ids:
                    obs_rows=[r for r in obs_rows if r["id"] in record_ids]
                pdf_data=[preferred]
                for serial,r in enumerate(obs_rows,1):
                    ps=styles["Normal"]
                    ps.fontSize=5.2; ps.leading=6
                    evidence=[]
                    for a in attachment_rows("observation_attachments",r["id"]):
                        fp=Path(safe(a["file_path"]))
                        if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp"}:
                            try:
                                im=RLImage(str(fp),width=38,height=30,preserveAspectRatio=True)
                                evidence.append(im)
                            except Exception:
                                logging.exception("Unable to embed observation image in PDF")
                        elif fp.exists():
                            evidence.append(Paragraph(xml_safe(fp.name), ps))
                    if not evidence:
                        evidence=[Paragraph("None",ps)]
                    wrap=lambda v: Paragraph(xml_safe(v).replace("\n","<br/>"), ps)
                    pdf_data.append([wrap(serial),wrap(r["obs_date"]),wrap(r["location"]),wrap(r["obs_type"]),wrap(r["category"]),wrap(r["observation"]),wrap(format_corrective_actions_display(r)),wrap(r["responsible"]),evidence,wrap(r["status"])])
                header_style=styles["Normal"]
                header_style.fontSize=5.5; header_style.leading=6
                pdf_data[0]=[Paragraph(x,header_style) for x in preferred]
                col_widths=[0.32*72,0.66*72,0.88*72,0.78*72,0.78*72,1.55*72,1.35*72,0.78*72,1.62*72,0.58*72]
                t=Table(pdf_data,repeatRows=1,colWidths=col_widths)
            else:
                if table_name == "incidents" and "Attachments" in columns:
                    display_indices=list(range(9))+[columns.index("Attachments")]
                    pdf_data=[[xml_safe(columns[i]) for i in display_indices]]
                    incident_rows=db.fetchall("SELECT * FROM incidents ORDER BY id DESC")
                    if record_ids:
                        wanted=set(record_ids); incident_rows=[r for r in incident_rows if r["id"] in wanted]
                    cell_style=styles["Normal"]; cell_style.fontSize=5.5; cell_style.leading=6
                    for r in incident_rows[:500]:
                        vals=[xml_safe(r[columns[i]] if columns[i] in r.keys() else "") for i in display_indices[:-1]]
                        evidence=[]
                        for a in attachment_rows("incident_attachments",r["id"]):
                            fp=Path(safe(a["file_path"]))
                            remark=safe(a["remark"]) if "remark" in a.keys() else ""
                            if fp.exists() and fp.suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif"}:
                                try:
                                    evidence.append(RLImage(str(fp),width=144,height=144,preserveAspectRatio=True))
                                    evidence.append(Paragraph(xml_safe(fp.name + (f" | {remark}" if remark else "")),cell_style))
                                except Exception:
                                    evidence.append(Paragraph(xml_safe(fp.name + (f" | {remark}" if remark else "")),cell_style))
                            else:
                                evidence.append(Paragraph(xml_safe(fp.name + (f" | {remark}" if remark else "")),cell_style))
                        if not evidence: evidence=[Paragraph("No attachment",cell_style)]
                        pdf_data.append([Paragraph(v.replace("\n","<br/>"),cell_style) for v in vals]+[evidence])
                    t=Table(pdf_data,repeatRows=1,colWidths=[55,55,65,65,65,90,75,65,75,125])
                elif len(columns) > 10 and "Attachments" in columns:
                    display_indices=list(range(9))+[columns.index("Attachments")]
                    pdf_data=[[xml_safe(columns[i]) for i in display_indices]]+[[xml_safe(r[i]) for i in display_indices] for r in data[:500]]
                    t=Table(pdf_data,repeatRows=1)
                else:
                    display_indices=list(range(min(len(columns),10)))
                    pdf_data=[[xml_safe(columns[i]) for i in display_indices]]+[[xml_safe(r[i]) for i in display_indices] for r in data[:500]]
                    t=Table(pdf_data,repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("GRID",(0,0),(-1,-1),0.4,colors.grey),("FONTSIZE",(0,0),(-1,-1),6),("VALIGN",(0,0),(-1,-1),"TOP"),
                ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
                ("TOPPADDING",(0,0),(-1,-1),1.5),("BOTTOMPADDING",(0,0),(-1,-1),1.5)]))
            story.append(t)
            footer=db.setting("report_footer","")
            if footer: story.extend([Spacer(1,10),Paragraph(xml_safe(footer),styles["Normal"])])
            doc.build(story); self.show_export_success(path)
        except Exception as e:
            logging.exception("PDF export failed"); QMessageBox.critical(self,"Export Error",str(e))

    def export_docx(self, table_name, record_ids=None):
        try:
            columns,data=self.table_data(table_name, record_ids)
            if not columns:
                QMessageBox.information(self,"No Data","There is no data to export."); return
            path=self._export_path(f"{table_name}_report.docx","Download Word Report","Word Documents (*.docx)")
            if not path:return
            generate_table_docx(table_name,path,record_ids=record_ids)
            with zipfile.ZipFile(path,"r") as z: z.testzip()
            self.show_export_success(path)
        except Exception as e:
            logging.exception("Word report failed"); QMessageBox.critical(self,"Export Error",str(e))

    def export_table(self, table, filename):
        self._export_table_widget_csv(table, filename)

    def _export_table_widget_csv(self, table, filename):
        path,_=QFileDialog.getSaveFileName(self,"Export CSV",f"{filename}.csv","CSV (*.csv)")
        if not path:return
        with open(path,"w",newline="",encoding="utf-8-sig") as f:
            writer=csv.writer(f)
            writer.writerow([table.horizontalHeaderItem(c).text() for c in range(table.columnCount())])
            for r in range(table.rowCount()):
                writer.writerow([table.item(r,c).text() if table.item(r,c) else "" for c in range(table.columnCount())])
        self.show_export_success(path)

    # ========================================================
    # MASTER DATA
    # ========================================================


    def master_data(self):
        w,layout=self.page("Master Data")
        layout.addWidget(QLabel("Manage employees, companies, projects, categories and company profile."))
        profile=QPushButton("Company Profile / Logo / Document Number")
        profile.clicked.connect(lambda:self.settings_page())
        layout.addWidget(profile)
        for title,table_name in [
            ("Employees","employees"),("Companies","companies"),
            ("Projects / Locations","projects"),("Observation Categories","categories")
        ]:
            button=QPushButton(title)
            button.clicked.connect(lambda checked=False,t=table_name,n=title:self.master_editor(t,n))
            layout.addWidget(button)


    def master_editor(
        self,
        table_name,
        title
    ):

        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(650, 550)

        layout = QVBoxLayout(dialog)

        table = QTableWidget()

        layout.addWidget(table)

        form = QFormLayout()

        if table_name == "employees":

            fields = [
                ("Employee ID", "employee_id"),
                ("Name", "name"),
                ("Designation", "designation"),
                ("Company", "company"),
                ("Department", "department"),
                ("Contact", "contact")
            ]

        elif table_name == "companies":

            fields = [
                ("Company Name", "name"),
                ("Company Type", "company_type"),
                ("Contact Person", "contact_person"),
                ("Contact Number", "contact_number")
            ]

        elif table_name == "projects":

            fields = [
                ("Project", "project"),
                ("Area", "area"),
                ("Location", "location"),
                ("Description", "description")
            ]

        else:

            fields = [
                ("Category", "category"),
                ("Subcategory", "subcategory")
            ]

        inputs = {}

        for label, key in fields:

            e = QLineEdit()

            inputs[key] = e

            form.addRow(
                label,
                e
            )

        layout.addLayout(form)

        add = QPushButton("Add")
        layout.addWidget(add)

        close = QPushButton("Close")
        layout.addWidget(close)

        close.clicked.connect(
            dialog.accept
        )

        def load():

            rows = db.fetchall(
                f"SELECT * FROM {table_name}"
            )

            if not rows:
                table.setRowCount(0)
                return

            columns = list(rows[0].keys())

            table.setColumnCount(
                len(columns)
            )

            table.setHorizontalHeaderLabels(
                columns
            )

            table.setRowCount(
                len(rows)
            )

            for r, row in enumerate(rows):

                for c, col in enumerate(columns):

                    table.setItem(
                        r,
                        c,
                        QTableWidgetItem(
                            safe(row[col])
                        )
                    )

            table.horizontalHeader().setSectionResizeMode(
                QHeaderView.Stretch
            )

        def save():

            values = [
                inputs[key].text()
                for _, key in fields
            ]

            if not values[0].strip():
                return

            placeholders = ",".join(
                ["?"] * len(values)
            )

            columns = ",".join(
                [key for _, key in fields]
            )

            db.execute(
                f"""
                INSERT INTO {table_name}
                ({columns})
                VALUES ({placeholders})
                """,
                values
            )

            for e in inputs.values():
                e.clear()

            load()

        add.clicked.connect(save)

        load()

        dialog.exec()

    # ========================================================
    # SETTINGS
    # ========================================================


    def settings_page(self):
        w,layout=self.page("Settings")
        form=QFormLayout()
        company=QLineEdit(db.setting("company_name",""))
        project=QLineEdit(db.setting("project_name",""))
        location=QLineEdit(db.setting("default_location",""))
        observer=QLineEdit(db.setting("default_observer",""))
        footer=QLineEdit(db.setting("report_footer",""))
        prefix=QLineEdit(db.setting("document_prefix","HSE"))
        logo_path=QLineEdit(db.setting("company_logo",""))
        logo_path.setReadOnly(True)
        logo_button=QPushButton("Choose Company Logo")
        logo_preview=QLabel()
        logo_preview.setMinimumHeight(80)
        logo_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)

        form.addRow("Company Name",company)
        form.addRow("Project Name",project)
        form.addRow("Default Location",location)
        form.addRow("Default Observer",observer)
        form.addRow("Document Number / Prefix",prefix)
        form.addRow("Report Footer",footer)
        logo_row=QHBoxLayout(); logo_row.addWidget(logo_path); logo_row.addWidget(logo_button)
        form.addRow("Company Logo",logo_row)
        layout.addLayout(form); layout.addWidget(logo_preview)

        def choose_logo():
            path,_=QFileDialog.getOpenFileName(self,"Select Company Logo","","Images (*.png *.jpg *.jpeg *.bmp)")
            if not path:return
            try:
                src=Path(path); dest=ATTACH_DIR / f"company_logo{src.suffix.lower()}"
            except Exception:
                dest=ATTACH_DIR / f"company_logo{Path(path).suffix.lower()}"
            try:
                for old in ATTACH_DIR.glob("company_logo.*"):
                    if old != dest: old.unlink(missing_ok=True)
                shutil.copy2(path,dest)
                logo_path.setText(str(dest))
                pix=QPixmap(str(dest))
                logo_preview.setPixmap(pix.scaled(180,80,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation))
            except Exception as e:
                QMessageBox.critical(self,"Logo Error",str(e))

        logo_button.clicked.connect(choose_logo)
        existing_logo=db.setting("company_logo","")
        if existing_logo and Path(existing_logo).exists():
            pix=QPixmap(existing_logo)
            logo_preview.setPixmap(pix.scaled(180,80,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation))

        save=QPushButton("Save Settings")
        save.setMinimumHeight(44)
        save.setToolTip("Save company, project and logo settings. New values will be used by all reports.")
        layout.addWidget(save)
        save.clicked.connect(lambda:self.save_settings(company,project,location,observer,footer,prefix,logo_path))
        backup=QPushButton("Backup Complete Application Data"); layout.addWidget(backup); backup.clicked.connect(self.backup)
        restore=QPushButton("Restore Application Backup"); layout.addWidget(restore); restore.clicked.connect(self.restore)
        about=QPushButton("About"); layout.addWidget(about)
        about.clicked.connect(lambda:QMessageBox.information(self,"About",f"{APP_NAME}\nVersion {APP_VERSION}\n\nOffline HSE Management System"))

    def save_settings(self, company, project, location, observer, footer, prefix=None, logo_path=None):
        # Persist all report branding settings in the application database.
        # The logo is copied into the application's attachment directory so the
        # saved report branding does not depend on the user's original file path.
        db.set_setting("company_name", company.text().strip())
        db.set_setting("project_name", project.text().strip())
        db.set_setting("default_location", location.text().strip())
        db.set_setting("default_observer", observer.text().strip())
        db.set_setting("report_footer", footer.text().strip())
        if prefix is not None:
            db.set_setting("document_prefix", prefix.text().strip() or "HSE")

        if logo_path is not None:
            selected = logo_path.text().strip()
            if selected:
                source = Path(selected)
                if source.exists() and source.is_file():
                    try:
                        ATTACH_DIR.mkdir(parents=True, exist_ok=True)
                        suffix = source.suffix.lower() or ".png"
                        destination = ATTACH_DIR / f"company_logo{suffix}"
                        # Remove older saved logo formats so reports always use the
                        # current logo selected in Settings.
                        for old in ATTACH_DIR.glob("company_logo.*"):
                            if old != destination:
                                old.unlink(missing_ok=True)
                        if source.resolve() != destination.resolve():
                            shutil.copy2(source, destination)
                        selected = str(destination)
                    except Exception as e:
                        logging.exception("Unable to persist company logo")
                        QMessageBox.critical(self, "Logo Error", f"Unable to save the company logo.\n\n{e}")
                        return
                elif not source.exists():
                    QMessageBox.warning(self, "Logo Not Found", "The selected company logo could not be found. The previous saved logo will be kept.")
                    selected = db.setting("company_logo", "")
            db.set_setting("company_logo", str(Path(selected).resolve()) if selected and Path(selected).exists() else selected)

        # All report generators read these settings from the database at export
        # time, so the next Word/Excel/PDF/CSV report immediately uses the new
        # company name and logo without changing any other module.
        QMessageBox.information(self, "Saved", "Settings saved successfully. New company name and logo will be used in all new reports.")
        self.dashboard()


    def backup(self):

        destination, _ = QFileDialog.getSaveFileName(
            self,
            "Save HSE Backup",
            f"HSE_Backup_{today()}.zip",
            "ZIP (*.zip)"
        )

        if not destination:
            return

        try:

            with zipfile.ZipFile(
                destination,
                "w",
                zipfile.ZIP_DEFLATED
            ) as z:

                if DB_FILE.exists():

                    z.write(
                        DB_FILE,
                        "database/hse.db"
                    )

                for file in ATTACH_DIR.rglob("*"):

                    if file.is_file():

                        z.write(
                            file,
                            f"attachments/{file.name}"
                        )

            QMessageBox.information(
                self,
                "Backup Complete",
                "HSE backup created successfully."
            )

        except Exception as e:

            logging.exception(
                "Backup failed"
            )

            QMessageBox.critical(
                self,
                "Backup Error",
                str(e)
            )

    # --------------------------------------------------------

    def restore(self):

        source, _ = QFileDialog.getOpenFileName(
            self,
            "Select HSE Backup",
            "",
            "ZIP (*.zip)"
        )

        if not source:
            return

        answer = QMessageBox.question(
            self,
            "Restore Backup",
            "Restoring will replace the current database.\n\n"
            "Do you want to continue?"
        )

        if answer != QMessageBox.Yes:
            return

        try:

            # Close current database connection.
            db.conn.close()

            with zipfile.ZipFile(
                source,
                "r"
            ) as z:

                temp = APP_DIR / "restore_temp"

                if temp.exists():
                    shutil.rmtree(temp)

                temp.mkdir()

                z.extractall(temp)

                restored_db = (
                    temp /
                    "database" /
                    "hse.db"
                )

                if not restored_db.exists():
                    raise Exception(
                        "Backup database not found."
                    )

                shutil.copy2(
                    restored_db,
                    DB_FILE
                )

                restored_attach = (
                    temp /
                    "attachments"
                )

                if restored_attach.exists():

                    for file in restored_attach.iterdir():

                        if file.is_file():

                            shutil.copy2(
                                file,
                                ATTACH_DIR /
                                file.name
                            )

                shutil.rmtree(temp)

            QMessageBox.information(
                self,
                "Restore Complete",
                "Backup restored successfully.\n"
                "Please restart the application."
            )

        except Exception as e:

            logging.exception(
                "Restore failed"
            )

            QMessageBox.critical(
                self,
                "Restore Error",
                str(e)
            )

        finally:

            # Reopen database.
            db.conn = sqlite3.connect(
                DB_FILE
            )

            db.conn.row_factory = sqlite3.Row
            db.conn.execute(
                "PRAGMA foreign_keys = ON"
            )

    # ========================================================
    # CLOSE
    # ========================================================

    def closeEvent(self, event):

        try:
            db.conn.commit()
            db.conn.close()
        except Exception:
            pass

        event.accept()




# ============================================================
# EMBEDDED PTW MODULE (self-contained in main.py)
# ============================================================

PERMIT_TYPES = [
    "Hot Work Permit", "Cold Work Permit", "General Work Permit",
    "Confined Space Entry Permit", "Excavation / Ground Disturbance Permit",
    "Electrical Work Permit (LV)", "High Voltage Work Permit",
    "Mechanical Work Permit", "Energy Isolation / LOTO Permit",
    "Work at Height / Roof Work Permit", "Lifting / Critical Lift Permit",
    "Radiography / Radiation Work Permit", "Breaking Containment / Line Breaking Permit",
    "Pressure Testing Permit", "Diving / Underwater Work Permit",
    "Demolition Permit", "Asbestos Removal / Disturbance Permit",
    "Road Closure / Traffic Management Permit", "Work Near Water Permit",
    "Blasting Permit", "Other"
]
WORKFLOWS = ["PTW", "CPW", "e-PTW / EPTW"]
PTW_STATUSES = ["Draft", "Submitted", "Under Review", "Returned for Correction", "Approved", "Issued / Active", "Suspended", "On Hold", "Awaiting Revalidation", "Work Completed", "Pending Closure", "Closed", "Cancelled", "Expired"]
CERT_TYPES = ["ECC / Excavation Clearance", "Confined Space Entry Certificate", "Electrical Work / Isolation", "Mechanical Isolation / LOTO", "Preparation / Reinstatement", "Sanction for Test", "Gas Test Record", "Lifting Plan / Equipment Certificate", "Work at Height / Rescue", "Radiography Supporting Document", "JSA / JHA", "Approved Method Statement", "Toolbox Talk Attendance", "Drawing / Utility Clearance", "Other"]


def _s(value):
    return "" if value is None else str(value)


def _now():
    return datetime.now().isoformat(timespec="seconds")


def _stamp():
    return datetime.now().strftime("%Y%m%d-%H%M%S")


class PTWModule(QWidget):
    """Standalone PTW workspace mounted inside the existing application stack."""
    def __init__(self, parent_window, db):
        super().__init__(parent_window)
        self.host = parent_window
        self.db = db
        self.selected_permit_id = None
        self.editing_certificate_id = None
        self.editing_renewal_id = None
        self.editing_document_id = None
        self.editing_closure_id = None
        self.editing_suspension_event_id = None
        self.form_fields = {}
        self._ensure_schema()
        self._build_ui()
        self.refresh_all()

    def _ensure_schema(self):
        # Additive-only migration: never alters or drops existing application tables.
        self.db.conn.executescript("""
        CREATE TABLE IF NOT EXISTS ptw_permit_types (
          id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE,
          active INTEGER NOT NULL DEFAULT 1, workflow_profile TEXT DEFAULT 'PTW',
          validity_hours INTEGER DEFAULT 12, procedure_name TEXT DEFAULT '',
          procedure_number TEXT DEFAULT '', procedure_revision TEXT DEFAULT '',
          effective_date TEXT DEFAULT '', mandatory_fields TEXT DEFAULT '', updated_at TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS ptw_permits (
          id INTEGER PRIMARY KEY AUTOINCREMENT, permit_number TEXT NOT NULL UNIQUE,
          workflow_type TEXT NOT NULL DEFAULT 'PTW', permit_type TEXT NOT NULL,
          revision INTEGER NOT NULL DEFAULT 1, original_permit_ref TEXT DEFAULT '',
          asset_owner TEXT DEFAULT '', client_authority TEXT DEFAULT '', performing_company TEXT DEFAULT '',
          subcontractor TEXT DEFAULT '', project TEXT DEFAULT '', site TEXT DEFAULT '', department TEXT DEFAULT '',
          contract_number TEXT DEFAULT '', work_order TEXT DEFAULT '', job_reference TEXT DEFAULT '',
          procedure_number TEXT DEFAULT '', procedure_revision TEXT DEFAULT '', client_reference TEXT DEFAULT '',
          location TEXT NOT NULL DEFAULT '', plant_area TEXT DEFAULT '', equipment_tag TEXT DEFAULT '',
          work_scope TEXT NOT NULL DEFAULT '', work_description TEXT DEFAULT '', discipline TEXT DEFAULT '',
          tools_equipment TEXT DEFAULT '', special_equipment TEXT DEFAULT '', worker_count INTEGER DEFAULT 0,
          work_crew TEXT DEFAULT '', planned_start TEXT DEFAULT '', planned_end TEXT DEFAULT '', shift TEXT DEFAULT '',
          supervisor TEXT DEFAULT '', performing_authority TEXT DEFAULT '', applicant_name TEXT DEFAULT '',
          applicant_employee_id TEXT DEFAULT '', applicant_company TEXT DEFAULT '', applicant_contact TEXT DEFAULT '',
          permit_holder TEXT DEFAULT '', permit_holder_employee_id TEXT DEFAULT '', permit_holder_company TEXT DEFAULT '',
          permit_holder_designation TEXT DEFAULT '', competency_ref TEXT DEFAULT '', competency_expiry TEXT DEFAULT '',
          permit_issuer TEXT DEFAULT '', area_authority TEXT DEFAULT '', hse_representative TEXT DEFAULT '',
          gas_tester TEXT DEFAULT '', isolation_authority TEXT DEFAULT '', emergency_arrangements TEXT DEFAULT '',
          simultaneous_operations TEXT DEFAULT '', hazards TEXT DEFAULT '', risk_controls TEXT DEFAULT '',
          safety_checks TEXT DEFAULT '', status TEXT NOT NULL DEFAULT 'Draft', created_by TEXT DEFAULT '',
          created_at TEXT DEFAULT '', updated_at TEXT DEFAULT '', closure_date TEXT DEFAULT '', notes TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS ptw_certificates (
          id INTEGER PRIMARY KEY AUTOINCREMENT, certificate_ref TEXT NOT NULL UNIQUE,
          permit_id INTEGER NOT NULL, certificate_type TEXT NOT NULL, status TEXT DEFAULT 'Draft',
          responsible_person TEXT DEFAULT '', issue_date TEXT DEFAULT '', valid_from TEXT DEFAULT '', valid_until TEXT DEFAULT '',
          revision TEXT DEFAULT '1', verification_status TEXT DEFAULT 'Pending', closure_status TEXT DEFAULT 'Open',
          attachment_path TEXT DEFAULT '', comments TEXT DEFAULT '', created_at TEXT DEFAULT '', updated_at TEXT DEFAULT '',
          FOREIGN KEY(permit_id) REFERENCES ptw_permits(id)
        );
        CREATE TABLE IF NOT EXISTS ptw_approvals (
          id INTEGER PRIMARY KEY AUTOINCREMENT, permit_id INTEGER NOT NULL, stage TEXT NOT NULL,
          user_id TEXT DEFAULT '', user_name TEXT DEFAULT '', role_name TEXT DEFAULT '', decision TEXT NOT NULL,
          comments TEXT DEFAULT '', document_revision INTEGER DEFAULT 1, approval_ref TEXT DEFAULT '', created_at TEXT DEFAULT '',
          FOREIGN KEY(permit_id) REFERENCES ptw_permits(id)
        );
        CREATE TABLE IF NOT EXISTS ptw_events (
          id INTEGER PRIMARY KEY AUTOINCREMENT, permit_id INTEGER NOT NULL, event_type TEXT NOT NULL,
          previous_status TEXT DEFAULT '', new_status TEXT DEFAULT '', details TEXT DEFAULT '',
          actor TEXT DEFAULT '', attachment_path TEXT DEFAULT '', created_at TEXT DEFAULT '',
          FOREIGN KEY(permit_id) REFERENCES ptw_permits(id)
        );
        CREATE TABLE IF NOT EXISTS ptw_renewals (
          id INTEGER PRIMARY KEY AUTOINCREMENT, permit_id INTEGER NOT NULL, reason TEXT DEFAULT '',
          previous_start TEXT DEFAULT '', previous_end TEXT DEFAULT '', proposed_start TEXT DEFAULT '', proposed_end TEXT DEFAULT '',
          status TEXT DEFAULT 'Requested', responsible_person TEXT DEFAULT '', reinspection TEXT DEFAULT '',
          gas_retest TEXT DEFAULT '', approval_ref TEXT DEFAULT '', created_at TEXT DEFAULT '',
          FOREIGN KEY(permit_id) REFERENCES ptw_permits(id)
        );
        CREATE TABLE IF NOT EXISTS ptw_handovers (
          id INTEGER PRIMARY KEY AUTOINCREMENT, permit_id INTEGER NOT NULL, outgoing_person TEXT DEFAULT '',
          incoming_person TEXT DEFAULT '', outgoing_company TEXT DEFAULT '', incoming_company TEXT DEFAULT '',
          shift_name TEXT DEFAULT '', work_completed TEXT DEFAULT '', work_remaining TEXT DEFAULT '',
          site_condition TEXT DEFAULT '', outstanding_hazards TEXT DEFAULT '', gas_test_status TEXT DEFAULT '',
          isolation_status TEXT DEFAULT '', outstanding_precautions TEXT DEFAULT '', outgoing_ack TEXT DEFAULT '',
          incoming_ack TEXT DEFAULT '', authority_confirmation TEXT DEFAULT '', status TEXT DEFAULT 'Pending',
          created_at TEXT DEFAULT '', FOREIGN KEY(permit_id) REFERENCES ptw_permits(id)
        );
        CREATE TABLE IF NOT EXISTS ptw_closures (
          id INTEGER PRIMARY KEY AUTOINCREMENT, permit_id INTEGER NOT NULL, work_complete TEXT DEFAULT '',
          site_inspection TEXT DEFAULT '', tools_removed TEXT DEFAULT '', personnel_accounted TEXT DEFAULT '',
          plant_handover TEXT DEFAULT '', remaining_hazards TEXT DEFAULT '', isolation_disposition TEXT DEFAULT '',
          reinstatement TEXT DEFAULT '', certificates_checked TEXT DEFAULT '', outstanding_actions TEXT DEFAULT '',
          performing_ack TEXT DEFAULT '', issuer_approval TEXT DEFAULT '', closure_comments TEXT DEFAULT '',
          evidence_path TEXT DEFAULT '', status TEXT DEFAULT 'Pending', created_at TEXT DEFAULT '',
          FOREIGN KEY(permit_id) REFERENCES ptw_permits(id)
        );
        CREATE TABLE IF NOT EXISTS ptw_documents (
          id INTEGER PRIMARY KEY AUTOINCREMENT, permit_id INTEGER NOT NULL, document_type TEXT DEFAULT '',
          file_path TEXT NOT NULL, description TEXT DEFAULT '', uploaded_by TEXT DEFAULT '', uploaded_at TEXT DEFAULT '',
          FOREIGN KEY(permit_id) REFERENCES ptw_permits(id)
        );
        CREATE TABLE IF NOT EXISTS ptw_settings (
          key TEXT PRIMARY KEY, value TEXT DEFAULT '', updated_at TEXT DEFAULT '', updated_by TEXT DEFAULT ''
        );
        CREATE INDEX IF NOT EXISTS idx_ptw_permits_status ON ptw_permits(status);
        CREATE INDEX IF NOT EXISTS idx_ptw_permits_number ON ptw_permits(permit_number);
        CREATE INDEX IF NOT EXISTS idx_ptw_permits_company ON ptw_permits(performing_company);
        CREATE INDEX IF NOT EXISTS idx_ptw_permits_expiry ON ptw_permits(planned_end);
        CREATE INDEX IF NOT EXISTS idx_ptw_certificates_permit ON ptw_certificates(permit_id);
        CREATE INDEX IF NOT EXISTS idx_ptw_events_permit ON ptw_events(permit_id);
        """)
        # Additive PTW-only extensions. Existing application and PTW records are preserved.
        existing_cols = {r[1] for r in self.db.conn.execute("PRAGMA table_info(ptw_permits)").fetchall()}
        ptw_extra_cols = {
            "project_number": "TEXT DEFAULT ''",
            "jsa_number": "TEXT DEFAULT ''",
            "method_statement_number": "TEXT DEFAULT ''",
            "supporting_document_types": "TEXT DEFAULT ''",
            "supporting_document_reference": "TEXT DEFAULT ''",
            "loto_equipment": "TEXT DEFAULT ''",
            "loto_number": "TEXT DEFAULT ''",
            "renewal_of_permit_id": "INTEGER",
            "approved_by": "TEXT DEFAULT ''",
            "approval_date": "TEXT DEFAULT ''",
        }
        for col, definition in ptw_extra_cols.items():
            if col not in existing_cols:
                self.db.conn.execute(f"ALTER TABLE ptw_permits ADD COLUMN {col} {definition}")
        self.db.conn.executescript("""
        CREATE TABLE IF NOT EXISTS ptw_daily_revalidation (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          permit_id INTEGER NOT NULL,
          work_date TEXT NOT NULL,
          day_status TEXT NOT NULL DEFAULT 'Pending',
          comments TEXT DEFAULT '',
          recorded_by TEXT DEFAULT '',
          recorded_at TEXT DEFAULT '',
          UNIQUE(permit_id, work_date),
          FOREIGN KEY(permit_id) REFERENCES ptw_permits(id)
        );
        CREATE INDEX IF NOT EXISTS idx_ptw_daily_permit_date ON ptw_daily_revalidation(permit_id, work_date);
        """)
        daily_cols = {r[1] for r in self.db.conn.execute("PRAGMA table_info(ptw_daily_revalidation)").fetchall()}
        for col in ("permit_holder", "permit_authority"):
            if col not in daily_cols:
                self.db.conn.execute(f"ALTER TABLE ptw_daily_revalidation ADD COLUMN {col} TEXT DEFAULT ''")
        self.db.conn.commit()
        for name in PERMIT_TYPES:
            self.db.conn.execute("INSERT OR IGNORE INTO ptw_permit_types(name,active,workflow_profile,validity_hours,updated_at) VALUES(?,1,'PTW',12,?)", (name, _now()))
        defaults = {
            "expiry_warning_hours": "24",
            "site_procedure_notice": "Configure this module against the approved site/client PTW procedure before operational use.",
            "approval_roles": "Permit Issuer;Area Authority;Performing Authority;HSE Representative",
            "required_certificate_types": "",
        }
        for k, v in defaults.items():
            self.db.conn.execute("INSERT OR IGNORE INTO ptw_settings(key,value,updated_at) VALUES(?,?,?)", (k, v, _now()))
        self.db.conn.commit()

    def _build_ui(self):
        root = QVBoxLayout(self)
        heading = QLabel("PERMIT TO WORK (PTW) MANAGEMENT")
        heading.setStyleSheet("font-size:22px;font-weight:bold;color:#17365D;padding:7px;")
        root.addWidget(heading)
        notice = QLabel("Complete each PTW record against the approved client/site procedure. A saved or approved application does not by itself authorize work to start.")
        notice.setWordWrap(True)
        notice.setStyleSheet("padding:8px;background:#FFF4D6;border:1px solid #E6D18A;border-radius:4px;color:#5B4B16;")
        root.addWidget(notice)
        self.tabs = QTabWidget()
        root.addWidget(self.tabs, 1)
        self.application_tab = QWidget(); self.tabs.addTab(self.application_tab, "New Permit Application"); self._build_application()
        self.documents_tab = QWidget(); self.tabs.addTab(self.documents_tab, "Attachments / Safety Documents"); self._build_documents()
        self.cert_tab = QWidget(); self.tabs.addTab(self.cert_tab, "Supporting Certificates"); self._build_certificates()
        self.review_tab = QWidget(); self.tabs.addTab(self.review_tab, "Permit Review & Approval"); self._build_review_approval()
        self.issue_tab = QWidget(); self.tabs.addTab(self.issue_tab, "Permit Issue Authority"); self._build_issue_authority()
        self.register_tab = QWidget(); self.tabs.addTab(self.register_tab, "Permit Register"); self._build_register()
        self.revalidation_tab = QWidget(); self.tabs.addTab(self.revalidation_tab, "Daily Revalidation"); self._build_revalidation()
        self.renewal_tab = QWidget(); self.tabs.addTab(self.renewal_tab, "Renewal / Extension"); self._build_renewals()
        self.suspension_tab = QWidget(); self.tabs.addTab(self.suspension_tab, "Suspension / Restart"); self._build_suspension()
        self.closure_tab = QWidget(); self.tabs.addTab(self.closure_tab, "Permit Closure"); self._build_closure()

    def _table(self, headers, parent_layout):
        t = QTableWidget(); t.setColumnCount(len(headers)); t.setHorizontalHeaderLabels(headers)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); t.setAlternatingRowColors(True)
        t.setWordWrap(True); t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        t.horizontalHeader().setStretchLastSection(True); t.setSortingEnabled(True)
        parent_layout.addWidget(t, 1); return t

    def _build_dashboard(self):
        lay = QVBoxLayout(self.dashboard_tab)
        self.dashboard_summary = QLabel(); self.dashboard_summary.setWordWrap(True)
        self.dashboard_summary.setStyleSheet("font-size:15px;font-weight:bold;padding:12px;background:#E8F0FA;color:#17365D;border-radius:5px;")
        lay.addWidget(self.dashboard_summary)
        self.dashboard_table = self._table(["Permit", "Type", "Company", "Location", "Expiry", "Status", "Required Action"], lay)
        row = QHBoxLayout(); refresh = QPushButton("Refresh Dashboard"); new = QPushButton("New Permit Application")
        refresh.clicked.connect(self.refresh_all); new.clicked.connect(lambda: self.tabs.setCurrentWidget(self.application_tab))
        row.addWidget(refresh); row.addWidget(new); row.addStretch(); lay.addLayout(row)

    def _build_application(self):
        outer = QVBoxLayout(self.application_tab)
        scroll = QScrollArea(); scroll.setWidgetResizable(True); outer.addWidget(scroll, 1)
        body = QWidget(); scroll.setWidget(body); root = QVBoxLayout(body)
        company_box = QGroupBox("1. Permit Details")
        form = QFormLayout(company_box); form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        def line(key, label, required=False):
            w=QLineEdit(); w.setClearButtonEnabled(True)
            if required: w.setPlaceholderText("Required")
            self.form_fields[key]=w; form.addRow(label + (" *" if required else ""),w); return w
        self.form_fields["performing_company"] = QLineEdit(); self.form_fields["performing_company"].setClearButtonEnabled(True); form.addRow("Company / Contractor *",self.form_fields["performing_company"])
        self.form_fields["client_authority"] = QLineEdit(); self.form_fields["client_authority"].setClearButtonEnabled(True); form.addRow("Client Name",self.form_fields["client_authority"])
        self.form_fields["subcontractor"] = QLineEdit(); self.form_fields["subcontractor"].setClearButtonEnabled(True); form.addRow("Subcontractor",self.form_fields["subcontractor"])
        self.form_fields["planned_start"] = QDateTimeEdit(QDateTime.currentDateTime()); self.form_fields["planned_start"].setCalendarPopup(True); self.form_fields["planned_start"].setDisplayFormat("yyyy-MM-dd HH:mm"); form.addRow("Planned Start *",self.form_fields["planned_start"])
        self.form_fields["planned_end"] = QDateTimeEdit(QDateTime.currentDateTime().addSecs(12*3600)); self.form_fields["planned_end"].setCalendarPopup(True); self.form_fields["planned_end"].setDisplayFormat("yyyy-MM-dd HH:mm"); form.addRow("Planned End *",self.form_fields["planned_end"])
        self.form_fields["permit_type"] = QComboBox(); self.form_fields["permit_type"].addItems(PERMIT_TYPES); form.addRow("Type of Permit *",self.form_fields["permit_type"])
        line("location","Facility / Site Location",True)
        line("project_number","Project Number")
        line("jsa_number","JSA Number")
        line("method_statement_number","Method Statement Number")
        root.addWidget(company_box)

        docs_box=QGroupBox("2. Supporting Documents")
        docs_layout=QVBoxLayout(docs_box)
        docs_hint=QLabel("Add each applicable certificate/document separately. Select Other to enter its document name.")
        docs_hint.setWordWrap(True); docs_layout.addWidget(docs_hint)
        self.support_doc_rows=[]
        self.support_doc_container=QWidget(); self.support_doc_layout=QVBoxLayout(self.support_doc_container); self.support_doc_layout.setContentsMargins(0,0,0,0)
        docs_layout.addWidget(self.support_doc_container)
        self.add_support_doc_button=QPushButton("+ Add Supporting Document")
        self.add_support_doc_button.setMinimumHeight(34)
        self.add_support_doc_button.setAutoDefault(False)
        self.add_support_doc_button.clicked.connect(lambda _checked=False: self._add_support_doc_row())
        docs_layout.addWidget(self.add_support_doc_button)
        self.form_fields["supporting_document_types"] = self.support_doc_container
        self.form_fields["supporting_document_reference"] = self.support_doc_container
        self._add_support_doc_row()
        root.addWidget(docs_box)

        loto_box=QGroupBox("3. LOTO Details")
        loto=QVBoxLayout(loto_box)
        loto.addWidget(QLabel("Add a separate row for every isolation / LOTO point."))
        self.loto_rows=[]
        self.loto_container=QWidget(); self.loto_layout=QVBoxLayout(self.loto_container); self.loto_layout.setContentsMargins(0,0,0,0); loto.addWidget(self.loto_container)
        self.add_loto_button=QPushButton("+ Add LOTO / Isolation Point")
        self.add_loto_button.setMinimumHeight(34)
        self.add_loto_button.setAutoDefault(False)
        self.add_loto_button.clicked.connect(lambda _checked=False: self._add_loto_row())
        loto.addWidget(self.add_loto_button)
        self.form_fields["loto_equipment"] = self.loto_container
        self.form_fields["loto_number"] = self.loto_container
        self._add_loto_row()
        root.addWidget(loto_box)

        work_box=QGroupBox("4. Work Details and Permit Holder")
        work=QFormLayout(work_box); work.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        self.form_fields["work_scope"]=QTextEdit(); self.form_fields["work_scope"].setMaximumHeight(90); self.form_fields["work_scope"].setPlaceholderText("Describe the work activity and scope clearly")
        work.addRow("Work Details / Scope *",self.form_fields["work_scope"])
        self.form_fields["permit_holder"]=QLineEdit(); self.form_fields["permit_holder"].setReadOnly(True); work.addRow("Permit Holder Name *",self.form_fields["permit_holder"])
        self.form_fields["permit_holder_designation"]=QLineEdit(); self.form_fields["permit_holder_designation"].setReadOnly(True); work.addRow("Permit Holder Position",self.form_fields["permit_holder_designation"])
        self.form_fields["permit_holder_company"]=QLineEdit(); work.addRow("Permit Holder Company",self.form_fields["permit_holder_company"])
        self.form_fields["permit_holder_employee_id"]=QLineEdit(); self.form_fields["permit_holder_employee_id"].setPlaceholderText("Enter employee ID number")
        holder_id_row=QHBoxLayout(); holder_id_row.addWidget(self.form_fields["permit_holder_employee_id"],1)
        self.search_ptw_holder_btn=QPushButton("Search")
        self.search_ptw_holder_btn.setAutoDefault(False)
        self.search_ptw_holder_btn.clicked.connect(self._search_ptw_holder_training)
        holder_id_row.addWidget(self.search_ptw_holder_btn)
        holder_id_wrap=QWidget(); holder_id_wrap.setLayout(holder_id_row); work.addRow("Permit Holder ID *",holder_id_wrap)
        self.form_fields["permit_holder_employee_id"].textChanged.connect(self._clear_ptw_holder_details_on_id_change)
        self.form_fields["work_description"]=QTextEdit(); self.form_fields["work_description"].setMaximumHeight(70); self.form_fields["work_description"].setPlaceholderText("Additional work description / precautions")
        work.addRow("Additional Work Description",self.form_fields["work_description"])
        root.addWidget(work_box)

        renewal_box=QGroupBox("5. Renewal (if applicable)")
        renewal=QFormLayout(renewal_box)
        self.renewal_check=QCheckBox("This permit is a renewal of an earlier permit")
        renewal.addRow(self.renewal_check)
        self.renewal_previous=QComboBox(); self.renewal_previous.addItem("Select previous permit",None)
        renewal.addRow("Previous Permit Number",self.renewal_previous)
        self.renewal_check.toggled.connect(self.renewal_previous.setEnabled); self.renewal_previous.setEnabled(False)
        root.addWidget(renewal_box)

        btns=QHBoxLayout(); self.save_draft_btn=QPushButton("Save Draft"); submit=QPushButton("Submit for Review"); clear=QPushButton("Clear Form")
        self.save_draft_btn.clicked.connect(lambda:self._save_permit("Draft")); submit.clicked.connect(self._submit_permit); clear.clicked.connect(self._clear_form)
        btns.addWidget(self.save_draft_btn); btns.addWidget(submit); btns.addWidget(clear); btns.addStretch(); root.addLayout(btns)
        self.application_permit_summary=QLabel("Save a draft or submit the application for review.")
        self.application_permit_summary.setWordWrap(True)
        self.application_permit_summary.setStyleSheet("padding:10px;background:#F2F6FA;border:1px solid #C8D5E3;border-radius:4px;")
        root.addWidget(self.application_permit_summary)
        self._update_application_permit_summary()

    def _build_revalidation(self):
        lay=QVBoxLayout(self.revalidation_tab)
        top=QFormLayout(); lay.addLayout(top)
        self.reval_permit=QComboBox(); self.reval_permit.currentIndexChanged.connect(self._load_revalidation_days)
        self.reval_holder=QLineEdit(); self.reval_authority=QLineEdit()
        top.addRow("Select Permit from Register",self.reval_permit)
        self.reval_info=QLabel("Select a permit to view its planned dates and daily revalidation record."); self.reval_info.setWordWrap(True); lay.addWidget(self.reval_info)
        self.reval_table=QTableWidget(); self.reval_table.setColumnCount(6); self.reval_table.setHorizontalHeaderLabels(["Date","Day","Daily Status","Daily PTW Holder Name","Permit Issuing Authority Name","Comments"])
        self.reval_table.setAlternatingRowColors(True); self.reval_table.setWordWrap(True); self.reval_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        for col in (0,1,2): self.reval_table.horizontalHeader().setSectionResizeMode(col,QHeaderView.ResizeMode.ResizeToContents)
        self.reval_table.horizontalHeader().setSectionResizeMode(3,QHeaderView.ResizeMode.Interactive)
        self.reval_table.horizontalHeader().setSectionResizeMode(4,QHeaderView.ResizeMode.Interactive)
        self.reval_table.horizontalHeader().setSectionResizeMode(5,QHeaderView.ResizeMode.Stretch)
        lay.addWidget(self.reval_table,1)
        row=QHBoxLayout(); save=QPushButton("Save Daily Revalidation"); delete=QPushButton("Delete Selected Day Record"); save.clicked.connect(self._save_revalidation); delete.clicked.connect(self._delete_revalidation_day); row.addWidget(save); row.addWidget(delete); row.addStretch(); lay.addLayout(row)
        self.reval_summary=QLabel("Days issued: 0 | Days worked: 0"); lay.addWidget(self.reval_summary)

    def _build_review_approval(self):
        lay=QVBoxLayout(self.review_tab)
        intro=QLabel("Independent reviewer / approver workspace. Only submitted applications awaiting review are listed here.")
        intro.setWordWrap(True); lay.addWidget(intro)
        form=QFormLayout(); lay.addLayout(form)
        self.review_permit=QComboBox(); self.review_permit.currentIndexChanged.connect(self._refresh_review_summary)
        self.review_summary=QLabel("Select a submitted permit to review its details."); self.review_summary.setWordWrap(True)
        self.review_summary.setStyleSheet("padding:10px;background:#F2F6FA;border:1px solid #C8D5E3;border-radius:4px;")
        self.reviewed_by=QLineEdit(); self.review_comments=QTextEdit(); self.review_comments.setMaximumHeight(75)
        form.addRow("Submitted Permit",self.review_permit); lay.addWidget(self.review_summary)
        form.addRow("Reviewed / Approved By",self.reviewed_by); form.addRow("Review Comments",self.review_comments)
        row=QHBoxLayout(); approve=QPushButton("Approve Application"); return_btn=QPushButton("Return for Correction")
        approve.clicked.connect(lambda _checked=False:self._review_decision(True)); return_btn.clicked.connect(lambda _checked=False:self._review_decision(False))
        row.addWidget(approve); row.addWidget(return_btn); row.addStretch(); lay.addLayout(row)
        self.review_table=self._table(["Permit","Company","Permit Type","Location","Applicant","Submitted","Status"],lay)
        lay.addWidget(QLabel("Submitted Safety Documents / Attachments"))
        self.review_documents_table=self._table(["Document Type","Filename","Description","Uploaded By","Uploaded At"],lay)
        lay.addWidget(QLabel("Supporting Certificates"))
        self.review_certificates_table=self._table(["Certificate Ref","Certificate Type","Responsible Person","Valid Until","Status","Verification","Attachment"],lay)

    def _build_issue_authority(self):
        lay=QVBoxLayout(self.issue_tab)
        intro=QLabel("Permit issuing authority workspace. Only approved permits are listed. Review the permit summary before issuing authorization to work.")
        intro.setWordWrap(True); lay.addWidget(intro)
        form=QFormLayout(); lay.addLayout(form)
        self.issue_permit=QComboBox(); self.issue_permit.currentIndexChanged.connect(self._refresh_issue_summary)
        self.issue_summary=QLabel("Select an approved permit to review its details."); self.issue_summary.setWordWrap(True)
        self.issue_summary.setStyleSheet("padding:10px;background:#F2F6FA;border:1px solid #C8D5E3;border-radius:4px;")
        self.issued_by=QLineEdit(); self.issued_by.setPlaceholderText("Authorized permit issuer name")
        form.addRow("Approved Permit",self.issue_permit); lay.addWidget(self.issue_summary); form.addRow("Issued By / Permit Authority",self.issued_by)
        issue=QPushButton("Issue Permit / Authorize Work"); issue.setMinimumHeight(40); issue.clicked.connect(lambda _checked=False:self._issue_from_authority_tab()); lay.addWidget(issue)
        self.issue_table=self._table(["Permit","Company","Permit Type","Location","Permit Holder","Approved By","Approval Date","Status"],lay)
        lay.addWidget(QLabel("Safety Documents / Attachments for Selected Permit"))
        self.issue_documents_table=self._table(["Document Type","Filename","Description","Uploaded By","Uploaded At"],lay)
        lay.addWidget(QLabel("Supporting Certificates for Selected Permit"))
        self.issue_certificates_table=self._table(["Certificate Ref","Certificate Type","Responsible Person","Valid Until","Status","Verification","Attachment"],lay)

    def _build_register(self):
        lay=QVBoxLayout(self.register_tab); filters=QHBoxLayout()
        self.reg_search=QLineEdit(); self.reg_search.setPlaceholderText("Search permit, company, client, project, location or permit holder")
        self.reg_status=QComboBox(); self.reg_status.addItems(["All Statuses"]+PTW_STATUSES)
        self.reg_type=QComboBox(); self.reg_type.addItems(["All Types"]+PERMIT_TYPES)
        filters.addWidget(self.reg_search,2); filters.addWidget(self.reg_status); filters.addWidget(self.reg_type); lay.addLayout(filters)
        headers=["Permit No.","Previous Permit","Company","Client","Subcontractor","Permit Type","Facility / Location","Project No.","JSA No.","Method Statement No.","Supporting Docs","Document Ref.","LOTO Equipment","LOTO No.","Work Details","Holder Name","Holder Company","Holder ID","Planned Start","Planned End","Days Issued","Days Worked","Approved By","Approval Date","Status"]
        self.permit_table=self._table(headers,lay)
        self.permit_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.permit_table.horizontalHeader().setSectionResizeMode(14,QHeaderView.ResizeMode.Stretch)
        self.permit_table.setWordWrap(True); self.permit_table.setMinimumHeight(260)
        row=QHBoxLayout()
        for label,fn in [("View / Load",self._load_selected),("Edit Draft",self._edit_selected),("Delete Permit",self._delete_selected_permit),("Issue Selected",self._issue_selected_from_register),("Export Excel Register",self._export_excel_register),("Export Selected Permit",self._export_selected),("Daily Revalidation",self._open_revalidation_selected),("History",self._show_history)]:
            b=QPushButton(label); b.clicked.connect(fn); row.addWidget(b)
        row.addStretch(); lay.addLayout(row)
        self.reg_search.textChanged.connect(self.refresh_register); self.reg_status.currentTextChanged.connect(self.refresh_register); self.reg_type.currentTextChanged.connect(self.refresh_register)

    def _build_certificates(self):
        lay=QVBoxLayout(self.cert_tab); f=QFormLayout(); lay.addLayout(f)
        self.cert_permit=QComboBox(); self.cert_permit.currentIndexChanged.connect(self._refresh_certificates_for_selected_permit); self.cert_type=QComboBox(); self.cert_type.addItems(CERT_TYPES); self.cert_responsible=QLineEdit(); self.cert_valid_until=QDateTimeEdit(QDateTime.currentDateTime()); self.cert_valid_until.setCalendarPopup(True); self.cert_valid_until.setDisplayFormat("yyyy-MM-dd HH:mm")
        self.cert_status=QComboBox(); self.cert_status.addItems(["Draft","Submitted","Approved","Valid","Expired","Suspended","Closed","Rejected"])
        self.cert_verification=QComboBox(); self.cert_verification.addItems(["Pending","Verified","Rejected"])
        self.cert_comments=QLineEdit(); f.addRow("Parent Permit",self.cert_permit); f.addRow("Certificate Type",self.cert_type); f.addRow("Responsible Person",self.cert_responsible); f.addRow("Valid Until",self.cert_valid_until); f.addRow("Status",self.cert_status); f.addRow("Verification",self.cert_verification); f.addRow("Comments",self.cert_comments)
        btn=QHBoxLayout(); add=QPushButton("Add Certificate"); edit=QPushButton("Update Selected"); delete=QPushButton("Delete Selected"); attach=QPushButton("Attach Certificate File"); add.clicked.connect(self._add_certificate); edit.clicked.connect(self._update_certificate); delete.clicked.connect(self._delete_certificate); attach.clicked.connect(self._attach_certificate); load=QPushButton("Load Selected for Edit"); load.clicked.connect(self._load_certificate_for_edit); btn.addWidget(add); btn.addWidget(load); btn.addWidget(edit); btn.addWidget(delete); btn.addWidget(attach); btn.addStretch(); lay.addLayout(btn)
        self.cert_table=self._table(["Reference","Permit","Certificate Type","Responsible Person","Valid Until","Status","Verification","File","Comments"],lay)

    def _build_renewals(self):
        lay=QVBoxLayout(self.renewal_tab); form=QFormLayout(); lay.addLayout(form)
        self.renew_permit=QComboBox(); self.renew_reason=QLineEdit(); self.renew_start=QDateTimeEdit(QDateTime.currentDateTime()); self.renew_start.setCalendarPopup(True); self.renew_start.setDisplayFormat("yyyy-MM-dd HH:mm"); self.renew_end=QDateTimeEdit(QDateTime.currentDateTime().addSecs(12*3600)); self.renew_end.setCalendarPopup(True); self.renew_end.setDisplayFormat("yyyy-MM-dd HH:mm"); self.renew_reinspection=QComboBox(); self.renew_reinspection.addItems(["Not required by configured procedure","Required - pending","Completed"]); self.renew_gas=QComboBox(); self.renew_gas.addItems(["Not required by configured procedure","Required - pending","Completed"]); self.renew_person=QLineEdit()
        for label,w in [("Permit",self.renew_permit),("Reason",self.renew_reason),("Proposed Start (YYYY-MM-DD HH:MM)",self.renew_start),("Proposed End (YYYY-MM-DD HH:MM)",self.renew_end),("Reinspection",self.renew_reinspection),("Gas Retest",self.renew_gas),("Responsible Person",self.renew_person)]: form.addRow(label,w)
        row=QHBoxLayout(); btn=QPushButton("Record Renewal Request"); load=QPushButton("Load Selected for Edit"); update=QPushButton("Update Selected Renewal"); delete=QPushButton("Delete Selected Renewal"); btn.clicked.connect(self._add_renewal); load.clicked.connect(self._load_selected_renewal); update.clicked.connect(self._update_selected_renewal); delete.clicked.connect(self._delete_selected_renewal); row.addWidget(btn); row.addWidget(load); row.addWidget(update); row.addWidget(delete); lay.addLayout(row)
        self.renewal_table=self._table(["Permit","Previous Start","Previous End","Proposed Start","Proposed End","Reason","Status","Reinspection","Gas Retest","Created"],lay)

    def _build_handovers(self):
        lay=QVBoxLayout(self.handover_tab); form=QFormLayout(); lay.addLayout(form)
        self.hand_permit=QComboBox(); self.hand_out=QLineEdit(); self.hand_in=QLineEdit(); self.hand_shift=QLineEdit(); self.hand_done=QTextEdit(); self.hand_done.setMaximumHeight(55); self.hand_remaining=QTextEdit(); self.hand_remaining.setMaximumHeight(55); self.hand_hazards=QTextEdit(); self.hand_hazards.setMaximumHeight(55); self.hand_gas=QComboBox(); self.hand_gas.addItems(["Not applicable per procedure","Verified - record evidence in notes","Pending verification"]); self.hand_iso=QComboBox(); self.hand_iso.addItems(["Not applicable per procedure","Verified - record evidence in notes","Pending verification"]); self.hand_out_ack=QCheckBox("Outgoing person acknowledged"); self.hand_in_ack=QCheckBox("Incoming person accepted"); self.hand_auth=QCheckBox("Authority confirmation completed (if required)")
        for label,w in [("Permit",self.hand_permit),("Outgoing Responsible Person",self.hand_out),("Incoming Responsible Person",self.hand_in),("Shift",self.hand_shift),("Work Completed",self.hand_done),("Work Remaining",self.hand_remaining),("Outstanding Hazards",self.hand_hazards),("Gas Test Status",self.hand_gas),("Isolation / LOTO Status",self.hand_iso),("Acknowledgement",self.hand_out_ack),("Acceptance",self.hand_in_ack),("Authority Confirmation",self.hand_auth)]: form.addRow(label,w)
        btn=QPushButton("Record Shift Handover"); btn.clicked.connect(self._add_handover); lay.addWidget(btn)
        self.handover_table=self._table(["Permit","Outgoing","Incoming","Shift","Work Remaining","Hazards","Gas","Isolation","Status","Created"],lay)

    def _build_suspension(self):
        lay=QVBoxLayout(self.suspension_tab)
        info=QLabel("Select any permit, record the reason, attach supporting evidence if available, choose Suspend, On Hold or Restart, then submit.")
        info.setWordWrap(True); lay.addWidget(info)
        form=QFormLayout(); lay.addLayout(form)
        self.susp_permit=QComboBox(); self.susp_action=QComboBox(); self.susp_action.addItems(["Suspend", "On Hold", "Restart"])
        self.susp_reason=QTextEdit(); self.susp_reason.setMaximumHeight(85)
        self.susp_actor=QLineEdit(); self.susp_actor.setPlaceholderText("Responsible person / user")
        self.susp_evidence=QLineEdit(); browse=QPushButton("Browse Evidence"); browse.clicked.connect(lambda:self._choose_into(self.susp_evidence))
        evidence_row=QHBoxLayout(); evidence_row.addWidget(self.susp_evidence); evidence_row.addWidget(browse)
        form.addRow("Select Permit",self.susp_permit); form.addRow("Reason",self.susp_reason); form.addRow("Evidence Attachment",evidence_row); form.addRow("Action",self.susp_action); form.addRow("Recorded By",self.susp_actor)
        submit=QPushButton("Submit Suspension / Hold / Restart"); submit.setMinimumHeight(38); submit.clicked.connect(lambda _checked=False:self._suspension_action()); lay.addWidget(submit)
        self.susp_table=self._table(["Permit","Event","Previous Status","New Status","Reason","Recorded By","Evidence","Timestamp"],lay)

    def _build_closure(self):
        lay=QVBoxLayout(self.closure_tab)
        info=QLabel("Select the permit, confirm the work is complete, then close it. This action is available regardless of the permit's current status.")
        info.setWordWrap(True)
        info.setStyleSheet("padding:10px;background:#F2F6FA;border:1px solid #C8D5E3;border-radius:4px;")
        lay.addWidget(info)
        form=QFormLayout()
        self.close_permit=QComboBox()
        self.close_permit.setMinimumWidth(360)
        self.close_confirm_check=QCheckBox("I confirm the work is complete and the permit can be closed.")
        form.addRow("Select Permit",self.close_permit)
        form.addRow("Confirmation",self.close_confirm_check)
        lay.addLayout(form)
        close_btn=QPushButton("Close Selected Permit")
        close_btn.setMinimumHeight(40)
        close_btn.clicked.connect(lambda _checked=False: self._close_permit_simple())
        lay.addWidget(close_btn)
        self.closure_table=self._table(["Permit","Work Complete","Inspection","Tools","Personnel","Hand-back","Isolation","Reinstatement","Status","Closure Date"],lay)

    def _build_documents(self):
        lay=QVBoxLayout(self.documents_tab); form=QFormLayout(); lay.addLayout(form)
        self.doc_permit=QComboBox(); self.doc_permit.currentIndexChanged.connect(self._refresh_documents_for_selected_permit); self.doc_type=QComboBox(); self.doc_type.addItems(["JSA / JHA","Method Statement","Drawing","Risk Assessment","Gas Test Record","Isolation / LOTO","Lifting Plan","Rescue Plan","Toolbox Talk","Other"]); self.doc_desc=QLineEdit(); self.doc_path=QLineEdit(); browse=QPushButton("Choose File"); browse.clicked.connect(lambda:self._choose_into(self.doc_path)); form.addRow("Permit",self.doc_permit); form.addRow("Document Type",self.doc_type); form.addRow("Description",self.doc_desc); row=QHBoxLayout(); row.addWidget(self.doc_path); row.addWidget(browse); form.addRow("File",row)
        btn=QPushButton("Attach Document"); btn.clicked.connect(self._add_document); lay.addWidget(btn)
        self.doc_table=self._table(["Permit","Document Type","Filename","Description","Uploaded By","Uploaded At","Path"],lay)
        openrow=QHBoxLayout(); openb=QPushButton("Open Selected Document"); load_doc=QPushButton("Load Selected for Edit"); update_doc=QPushButton("Update Selected Document"); deldoc=QPushButton("Delete Selected Document"); openb.clicked.connect(self._open_document); load_doc.clicked.connect(self._load_selected_document); update_doc.clicked.connect(self._update_selected_document); deldoc.clicked.connect(self._delete_document); openrow.addWidget(openb); openrow.addWidget(load_doc); openrow.addWidget(update_doc); openrow.addWidget(deldoc); lay.addLayout(openrow)

    def _build_reports(self):
        lay=QVBoxLayout(self.reports_tab); row=QHBoxLayout(); refresh=QPushButton("Refresh Audit Trail"); refresh.clicked.connect(self.refresh_all); row.addWidget(refresh); row.addStretch(); lay.addLayout(row)
        self.event_table=self._table(["Permit","Event","Previous Status","New Status","Details","Actor","Timestamp"],lay)

    def _build_config(self):
        lay=QVBoxLayout(self.config_tab); notice=QLabel("Only authorized administrators should configure these values. No site-specific QatarEnergy/client requirements are assumed by default."); notice.setWordWrap(True); lay.addWidget(notice)
        form=QFormLayout(); lay.addLayout(form)
        self.cfg_type=QLineEdit(); self.cfg_profile=QComboBox(); self.cfg_profile.addItems(WORKFLOWS); self.cfg_validity=QSpinBox(); self.cfg_validity.setRange(1,720); self.cfg_proc=QLineEdit(); self.cfg_procno=QLineEdit(); self.cfg_procrev=QLineEdit(); self.cfg_effective=QLineEdit(); self.cfg_expiry=QSpinBox(); self.cfg_expiry.setRange(1,720); self.cfg_expiry.setValue(24); self.cfg_roles=QLineEdit(); self.cfg_required_certs=QLineEdit()
        for label,w in [("Permit Type Name",self.cfg_type),("Workflow Profile",self.cfg_profile),("Default Validity Hours (configure to site procedure)",self.cfg_validity),("Procedure Name",self.cfg_proc),("Procedure Number",self.cfg_procno),("Procedure Revision",self.cfg_procrev),("Effective Date",self.cfg_effective),("Expiry Warning Hours",self.cfg_expiry),("Approval Roles (semicolon separated)",self.cfg_roles),("Default Required Certificate Types (semicolon separated)",self.cfg_required_certs)]: form.addRow(label,w)
        btn=QPushButton("Save / Update Permit Type and Configuration"); btn.clicked.connect(self._save_config); lay.addWidget(btn)
        self.config_table=self._table(["Permit Type","Active","Workflow","Validity Hours","Procedure","Number","Revision","Effective Date"],lay)

    def _add_support_doc_row(self, selected="", other_name="", reference=""):
        row_widget = QWidget(self.support_doc_container)
        row = QHBoxLayout(row_widget)
        row.setContentsMargins(0, 2, 0, 2)
        typ = QComboBox(row_widget)
        typ.setMinimumWidth(220)
        typ.addItems(["Select document", "ECC / Excavation Clearance", "Confined Space Entry",
                      "Electrical / Isolation Certificate", "Mechanical Isolation / LOTO",
                      "Gas Test Record", "Lifting Plan", "Work at Height / Rescue",
                      "Radiography Supporting Document", "Other"])
        ix = typ.findText(selected)
        typ.setCurrentIndex(ix if ix >= 0 else 0)
        other = QLineEdit(row_widget)
        other.setPlaceholderText("Enter document name (Other)")
        other.setText(other_name)
        other.setVisible(typ.currentText() == "Other")
        ref = QLineEdit(row_widget)
        ref.setPlaceholderText("Document number / reference")
        ref.setText(reference)
        typ.currentTextChanged.connect(lambda value, w=other: w.setVisible(value == "Other"))
        remove = QPushButton("Remove", row_widget)
        remove.setAutoDefault(False)
        remove.clicked.connect(lambda _checked=False, rw=row_widget: self._remove_dynamic_row(rw, self.support_doc_rows, self.support_doc_layout, self._add_support_doc_row))
        row.addWidget(typ, 2)
        row.addWidget(other, 2)
        row.addWidget(ref, 2)
        row.addWidget(remove)
        self.support_doc_layout.addWidget(row_widget)
        row_widget.setVisible(True)
        self.support_doc_rows.append((typ, other, ref, row_widget))

    def _add_loto_row(self, equipment="", number=""):
        row_widget = QWidget(self.loto_container)
        row = QHBoxLayout(row_widget)
        row.setContentsMargins(0, 2, 0, 2)
        eq = QLineEdit(row_widget)
        eq.setPlaceholderText("Plant / equipment to isolate")
        eq.setText(equipment)
        no = QLineEdit(row_widget)
        no.setPlaceholderText("LOTO / isolation number")
        no.setText(number)
        remove = QPushButton("Remove", row_widget)
        remove.setAutoDefault(False)
        remove.clicked.connect(lambda _checked=False, rw=row_widget: self._remove_dynamic_row(rw, self.loto_rows, self.loto_layout, self._add_loto_row))
        row.addWidget(eq, 2)
        row.addWidget(no, 1)
        row.addWidget(remove)
        self.loto_layout.addWidget(row_widget)
        row_widget.setVisible(True)
        self.loto_rows.append((eq, no, row_widget))

    def _remove_dynamic_row(self, row_widget, collection, layout, add_empty_callback):
        collection[:] = [item for item in collection if item[-1] is not row_widget]
        layout.removeWidget(row_widget)
        row_widget.setParent(None)
        row_widget.deleteLater()
        if not collection:
            add_empty_callback()

    def _set_dynamic_values(self, key, value, refs=""):
        if key=="supporting_document_types":
            while self.support_doc_layout.count():
                item=self.support_doc_layout.takeAt(0)
                child=item.layout()
                if child:
                    while child.count():
                        sub=child.takeAt(0)
                        if sub.widget(): sub.widget().deleteLater()
                elif item.widget(): item.widget().deleteLater()
            self.support_doc_rows=[]
            names=[x.strip() for x in _s(value).split(";") if x.strip()]
            refs_list=[x.strip() for x in _s(refs).split(";")]
            known=["ECC / Excavation Clearance","Confined Space Entry","Electrical / Isolation Certificate","Mechanical Isolation / LOTO","Gas Test Record","Lifting Plan","Work at Height / Rescue","Radiography Supporting Document","Other"]
            for i,name in enumerate(names):
                is_other=name not in known
                self._add_support_doc_row("Other" if is_other else name,name if is_other else "",refs_list[i] if i<len(refs_list) else "")
            if not self.support_doc_rows:self._add_support_doc_row()
        elif key=="loto_equipment":
            while self.loto_layout.count():
                item=self.loto_layout.takeAt(0); child=item.layout()
                if child:
                    while child.count():
                        sub=child.takeAt(0)
                        if sub.widget(): sub.widget().deleteLater()
                elif item.widget(): item.widget().deleteLater()
            self.loto_rows=[]
            eqs=[x.strip() for x in _s(value).split(";") if x.strip()]
            nums=[x.strip() for x in _s(refs).split(";")]
            for i,eq in enumerate(eqs): self._add_loto_row(eq,nums[i] if i<len(nums) else "")
            if not self.loto_rows:self._add_loto_row()

    def _update_application_permit_summary(self):
        if not hasattr(self, "application_permit_summary"):
            return
        permit = self._permit_row(self.selected_permit_id)
        if not permit:
            self.application_permit_summary.setText("Save a draft or submit the application for review.")
            return
        self.application_permit_summary.setText(
            f"Permit: {permit['permit_number']} | Status: {permit['status']}\n"
            f"Company: {permit['performing_company']} | Location: {permit['location']}\n"
            f"Permit Holder: {permit['permit_holder']} | Planned: {permit['planned_start']} to {permit['planned_end']}\n"
            f"Supporting Documents: {permit['supporting_document_types'] or 'None'}")

    def _refresh_authority_evidence(self, permit_id, stage):
        document_table=getattr(self,f"{stage}_documents_table",None)
        certificate_table=getattr(self,f"{stage}_certificates_table",None)
        if not permit_id:
            if document_table is not None: document_table.setRowCount(0)
            if certificate_table is not None: certificate_table.setRowCount(0)
            return
        docs=self.db.fetchall("SELECT document_type,file_path,description,uploaded_by,uploaded_at,id FROM ptw_documents WHERE permit_id=? ORDER BY id DESC",(permit_id,))
        certs=self.db.fetchall("SELECT certificate_ref,certificate_type,responsible_person,valid_until,status,verification_status,attachment_path,id FROM ptw_certificates WHERE permit_id=? ORDER BY id DESC",(permit_id,))
        if document_table is not None:self._fill_table(document_table,docs,["document_type","file_path","description","uploaded_by","uploaded_at"])
        if certificate_table is not None:self._fill_table(certificate_table,certs,["certificate_ref","certificate_type","responsible_person","valid_until","status","verification_status","attachment_path"])

    def _refresh_review_summary(self, *_args):
        if not hasattr(self, "review_permit"):
            return
        pid=self.review_permit.currentData(); r=self._permit_row(pid)
        if not r:
            self.review_summary.setText("Select a submitted permit to review its details."); self._refresh_authority_evidence(None, "review"); return
        self.review_summary.setText(f"Permit: {r['permit_number']} | {r['status']}\nType: {r['permit_type']} | Company: {r['performing_company']}\nLocation: {r['location']} | Work: {r['work_scope']}\nApplicant: {r['applicant_name']} | Holder: {r['permit_holder']}\nPlanned: {r['planned_start']} to {r['planned_end']}\nSupporting Documents: {r['supporting_document_types'] or 'None'} | LOTO: {r['loto_equipment'] or 'None'}")
        self._refresh_authority_evidence(pid, "review")

    def _refresh_issue_summary(self, *_args):
        if not hasattr(self, "issue_permit"):
            return
        pid=self.issue_permit.currentData(); r=self._permit_row(pid)
        if not r:
            self.issue_summary.setText("Select an approved permit to review its details."); self._refresh_authority_evidence(None, "issue"); return
        self.issue_summary.setText(f"Permit: {r['permit_number']} | Status: {r['status']}\nType: {r['permit_type']} | Company: {r['performing_company']}\nLocation: {r['location']} | Work: {r['work_scope']}\nPermit Holder: {r['permit_holder']} | Approved By: {r['approved_by']}\nApproval Date: {r['approval_date']} | Validity: {r['planned_start']} to {r['planned_end']}\nSupporting Documents: {r['supporting_document_types'] or 'None'} | LOTO: {r['loto_equipment'] or 'None'}")
        self._refresh_authority_evidence(pid, "issue")

    def _review_decision(self, approve):
        pid=self.review_permit.currentData(); r=self._permit_row(pid)
        if not r:
            return QMessageBox.warning(self,"Permit Required","Select a submitted application.")
        if r["status"] not in ("Submitted", "Under Review"):
            return QMessageBox.warning(self,"Review Not Available","Only submitted applications can be reviewed here.")
        reviewer=self.reviewed_by.text().strip()
        if not reviewer:
            return QMessageBox.warning(self,"Reviewer Required","Enter the authorized reviewer / approver name.")
        comments=self.review_comments.toPlainText().strip()
        if approve:
            new_status="Approved"
            self.db.execute("UPDATE ptw_permits SET status='Approved',approved_by=?,approval_date=?,updated_at=? WHERE id=?",(reviewer,_now(),_now(),pid))
            decision="Approved"
        else:
            new_status="Returned for Correction"
            self.db.execute("UPDATE ptw_permits SET status=?,updated_at=? WHERE id=?",(new_status,_now(),pid))
            decision="Returned for Correction"
        self.db.execute("INSERT INTO ptw_approvals(permit_id,stage,user_name,role_name,decision,comments,document_revision,created_at) VALUES(?,?,?,?,?,?,?,?)",(pid,"Application Review",reviewer,"Permit Reviewer / Approver",decision,comments,int(r["revision"] or 1),_now()))
        self._event(pid,"Permit application review",r["status"],new_status,f"{decision} by {reviewer}. {comments}")
        self.db.conn.commit(); self.refresh_all()
        QMessageBox.information(self,"Review Recorded",f"Permit {r['permit_number']} is now {new_status}.")

    def _refresh_issue_authority(self):
        if not hasattr(self,"issue_permit"):
            return
        current=self.issue_permit.currentData(); self.issue_permit.blockSignals(True); self.issue_permit.clear(); self.issue_permit.addItem("Select approved permit",None)
        rows=self.db.fetchall("SELECT id,permit_number,permit_type,status FROM ptw_permits WHERE status='Approved' ORDER BY id DESC")
        for r in rows:self.issue_permit.addItem(f"{r['permit_number']} — {r['permit_type']} ({r['status']})",r["id"])
        idx=self.issue_permit.findData(current); self.issue_permit.setCurrentIndex(idx if idx>=0 else 0); self.issue_permit.blockSignals(False); self._refresh_issue_summary()

    def _issue_from_authority_tab(self):
        pid=self.issue_permit.currentData(); issuer=self.issued_by.text().strip()
        if not pid:
            return QMessageBox.warning(self,"Permit Required","Select an approved permit.")
        if not issuer:
            return QMessageBox.warning(self,"Issuer Required","Enter the authorized permit issuer / authority name.")
        self._issue_permit_id(pid,issuer)
        self.refresh_all()

    def _issue_selected_from_register(self):
        pid=self._selected_id()
        if not pid:
            return QMessageBox.information(self,"Select Permit","Select a permit row first.")
        r=self._permit_row(pid)
        if not r or r["status"]!="Approved":
            return QMessageBox.warning(self,"Approval Required","Only approved permits can be issued. Use the Permit Issue Authority tab.")
        self.tabs.setCurrentWidget(self.issue_tab)
        idx=self.issue_permit.findData(pid)
        if idx<0:self._refresh_issue_authority(); idx=self.issue_permit.findData(pid)
        if idx>=0:self.issue_permit.setCurrentIndex(idx)
        self._refresh_issue_summary()

    def _issue_permit_id(self, pid, issuer=""):
        permit = self._permit_row(pid)
        if not permit:
            return QMessageBox.information(self, "Select Permit", "The selected permit could not be found.")
        if permit["status"] != "Approved":
            return QMessageBox.warning(self, "Approval Required", "Only an approved permit can be issued.")
        required = [x.strip() for x in self._setting("required_certificate_types", "").split(";") if x.strip()]
        certs = {_s(c["certificate_type"]) for c in self.db.fetchall(
            "SELECT certificate_type FROM ptw_certificates WHERE permit_id=? AND status IN ('Approved','Valid') AND verification_status='Verified'", (pid,))}
        missing = [x for x in required if x not in certs]
        if missing:
            return QMessageBox.warning(self, "Certificates Required", "Complete these configured certificates before issue:\n- " + "\n- ".join(missing))
        if not permit["planned_start"] or not permit["planned_end"]:
            return QMessageBox.warning(self, "Validity Required", "Planned start and end dates are required.")
        try:
            if datetime.fromisoformat(permit["planned_end"]) <= datetime.now():
                return QMessageBox.warning(self, "Expired", "The planned end date/time has passed. Update the permit validity before issue.")
        except Exception:
            return QMessageBox.warning(self, "Invalid Date", "The permit validity dates are invalid.")
        issuer = issuer.strip() or self._actor()
        self.db.execute("UPDATE ptw_permits SET permit_issuer=?,updated_at=? WHERE id=?",(issuer,_now(),pid))
        self._change_status(pid, "Issued / Active", f"Permit issued by {issuer}. Permit is now active.")
        QMessageBox.information(self, "Permit Issued", f"Permit {permit['permit_number']} is now Issued / Active and available in Suspension / Restart.")

    def _approve_application(self):
        self.tabs.setCurrentWidget(self.review_tab)
        self._refresh_authority_workflows()
        QMessageBox.information(self,"Separate Approval","Permit applications are submitted by the applicant and approved separately in the Permit Review & Approval tab.")

    def _approve_selected(self):
        pid=self._selected_id(); r=self._permit_row(pid)
        if not r:return QMessageBox.information(self,"Select Permit","Select a permit row first.")
        name,ok=QInputDialog.getText(self,"Approve Permit","Approved by:",text=_s(r["approved_by"]))
        if not ok:return
        if not name.strip():return QMessageBox.warning(self,"Approver Required","Enter the approver name.")
        self.db.execute("UPDATE ptw_permits SET status='Approved',approved_by=?,approval_date=?,updated_at=? WHERE id=?",(name.strip(),_now(),_now(),pid)); self._event(pid,"Permit approved",r["status"],"Approved",f"Approved by {name.strip()}"); self.refresh_all()
        if self.selected_permit_id == pid:
            self._update_application_permit_summary()

    def _delete_selected_permit(self):
        pid=self._selected_id(); r=self._permit_row(pid)
        if not r:return QMessageBox.information(self,"Select Permit","Select a permit row first.")
        ans=QMessageBox.question(self,"Delete Permit",f"Delete {r['permit_number']} and its PTW-only linked records? This cannot be undone.",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)
        if ans!=QMessageBox.StandardButton.Yes:return
        for table in ["ptw_daily_revalidation","ptw_certificates","ptw_approvals","ptw_events","ptw_renewals","ptw_handovers","ptw_closures","ptw_documents"]: self.db.execute(f"DELETE FROM {table} WHERE permit_id=?",(pid,))
        self.db.execute("DELETE FROM ptw_permits WHERE id=?",(pid,)); self.selected_permit_id=None; self.refresh_all()

    def _delete_certificate(self):
        row=self.cert_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Certificate","Select a certificate row first.")
        cid=self.cert_table.item(row,0).data(Qt.ItemDataRole.UserRole); c=self.db.fetchone("SELECT * FROM ptw_certificates WHERE id=?",(cid,))
        if not c:return
        if QMessageBox.question(self,"Delete Certificate",f"Delete certificate {c['certificate_ref']}?",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
        self.db.execute("DELETE FROM ptw_certificates WHERE id=?",(cid,)); self._event(c["permit_id"],"Supporting certificate deleted",details=c["certificate_ref"]); self.refresh_all()

    def _load_selected_closure(self):
        row=self.closure_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Closure","Select a closure row first.")
        cid=self.closure_table.item(row,0).data(Qt.ItemDataRole.UserRole); c=self.db.fetchone("SELECT * FROM ptw_closures WHERE id=?",(cid,))
        if not c:return
        self.editing_closure_id=cid; self.close_permit.setCurrentIndex(self.close_permit.findData(c["permit_id"]))
        for w,key in [(self.close_work,"work_complete"),(self.close_inspection,"site_inspection"),(self.close_tools,"tools_removed"),(self.close_people,"personnel_accounted"),(self.close_plant,"plant_handover"),(self.close_isolation,"isolation_disposition"),(self.close_reinstatement,"reinstatement")]:
            ix=w.findText(_s(c[key])); w.setCurrentIndex(ix if ix>=0 else 0)
        self.close_performing.setText(_s(c["performing_ack"])); self.close_issuer.setText(_s(c["issuer_approval"])); self.close_comments.setPlainText(_s(c["closure_comments"])); self.close_evidence.setText(_s(c["evidence_path"]))
        self.closure_edit_loaded=True

    def _update_selected_closure(self):
        if not self.editing_closure_id:return QMessageBox.information(self,"Load Closure","Load a closure record first.")
        c=self.db.fetchone("SELECT * FROM ptw_closures WHERE id=?",(self.editing_closure_id,))
        if not c:return
        self.db.execute("UPDATE ptw_closures SET work_complete=?,site_inspection=?,tools_removed=?,personnel_accounted=?,plant_handover=?,isolation_disposition=?,reinstatement=?,performing_ack=?,issuer_approval=?,closure_comments=?,evidence_path=? WHERE id=?",(self.close_work.currentText(),self.close_inspection.currentText(),self.close_tools.currentText(),self.close_people.currentText(),self.close_plant.currentText(),self.close_isolation.currentText(),self.close_reinstatement.currentText(),self.close_performing.text().strip(),self.close_issuer.text().strip(),self.close_comments.toPlainText(),self.close_evidence.text().strip(),self.editing_closure_id))
        self._event(c["permit_id"],"Closure record updated"); self.editing_closure_id=None; self.refresh_all()

    def _delete_selected_closure(self):
        row=self.closure_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Closure","Select a closure row first.")
        cid=self.closure_table.item(row,0).data(Qt.ItemDataRole.UserRole); c=self.db.fetchone("SELECT * FROM ptw_closures WHERE id=?",(cid,))
        if not c:return
        if QMessageBox.question(self,"Delete Closure","Delete the selected closure record?",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
        self.db.execute("DELETE FROM ptw_closures WHERE id=?",(cid,)); self._event(c["permit_id"],"Closure record deleted"); self.refresh_all()

    def _actor(self):
        return _s(self.db.setting("current_user", "") or self.db.setting("username", "") or "Application User")

    def _next_number(self):
        year=datetime.now().strftime("%Y")
        prefix=_s(self.db.setting("ptw_document_prefix", "PTW") or "PTW").strip().upper()
        row=self.db.fetchone("SELECT COUNT(*) AS n FROM ptw_permits WHERE permit_number LIKE ?",(f"{prefix}-{year}-%",))
        seq=int(row["n"] or 0)+1
        while self.db.fetchone("SELECT id FROM ptw_permits WHERE permit_number=?",(f"{prefix}-{year}-{seq:05d}",)):
            seq+=1
        return f"{prefix}-{year}-{seq:05d}"

    def _form_value(self, key):
        w=self.form_fields[key]
        if isinstance(w,QLineEdit): return w.text().strip()
        if isinstance(w,QTextEdit): return w.toPlainText().strip()
        if isinstance(w,QComboBox): return w.currentText().strip()
        if key in ("supporting_document_types", "supporting_document_reference"):
            docs=[]
            for typ,other,ref,_row_widget in self.support_doc_rows:
                name=other.text().strip() if typ.currentText()=="Other" else typ.currentText()
                if name and name!="Select document": docs.append((name,ref.text().strip()))
            return "; ".join(x[0] for x in docs) if key=="supporting_document_types" else "; ".join(x[1] for x in docs)
        if key in ("loto_equipment", "loto_number"):
            vals=[(eq.text().strip(),no.text().strip()) for eq,no,_row_widget in self.loto_rows]
            return "; ".join(x[0] if key=="loto_equipment" else x[1] for x in vals if x[0] or x[1])
        if isinstance(w,QSpinBox): return w.value()
        if isinstance(w,QDateTimeEdit): return w.dateTime().toString("yyyy-MM-dd HH:mm")
        return ""

    def _clear_form(self):
        self.selected_permit_id=None
        for k,w in self.form_fields.items():
            if isinstance(w,QLineEdit): w.clear()
            elif isinstance(w,QTextEdit): w.clear()
            elif isinstance(w,QComboBox): w.setCurrentIndex(0)
            elif isinstance(w,QSpinBox): w.setValue(0)
        self._set_dynamic_values("supporting_document_types", "", "")
        self._set_dynamic_values("loto_equipment", "", "")
        self.form_fields["planned_start"].setDateTime(QDateTime.currentDateTime())
        self.form_fields["planned_end"].setDateTime(QDateTime.currentDateTime().addSecs(12*3600))
        if hasattr(self,"renewal_check"):
            self.renewal_check.setChecked(False); self.renewal_previous.setCurrentIndex(0)
        self.save_draft_btn.setText("Save Permit")
        self._update_application_permit_summary()

    def _sync_application_supporting_certificates(self, permit_id, types_text, refs_text):
        names=[x.strip() for x in _s(types_text).split(";") if x.strip()]
        refs=[x.strip() for x in _s(refs_text).split(";")]
        permit=self._permit_row(permit_id)
        if not permit:return
        now=_now()
        for i,name in enumerate(names):
            ref_text=refs[i] if i<len(refs) else ""
            exists=self.db.fetchone("SELECT id FROM ptw_certificates WHERE permit_id=? AND certificate_type=? AND comments=? LIMIT 1",(permit_id,name,ref_text))
            if exists:continue
            n=self.db.fetchone("SELECT COUNT(*) AS n FROM ptw_certificates WHERE permit_id=?",(permit_id,))["n"]+1
            cert_ref=f"{permit['permit_number']}-CERT-{n:02d}"
            self.db.execute("INSERT OR IGNORE INTO ptw_certificates(certificate_ref,permit_id,certificate_type,status,responsible_person,valid_until,verification_status,comments,created_at,updated_at) VALUES(?,?,?,'Draft','','','Pending',?,?,?)",(cert_ref,permit_id,name,ref_text,now,now))

    def _clear_ptw_holder_details_on_id_change(self, _text=""):
        # Permit holder identity is always sourced from the Training module search.
        for key in ("permit_holder", "permit_holder_designation"):
            widget=self.form_fields.get(key)
            if isinstance(widget,QLineEdit): widget.clear()

    def _search_ptw_holder_training(self):
        employee_id=self.form_fields["permit_holder_employee_id"].text().strip()
        if not employee_id:
            return QMessageBox.warning(self,"Employee ID Required","Enter the permit holder's ID number, then click Search.")
        record=self.db.fetchone(
            """SELECT e.id,e.employee_id,e.name,e.designation,e.company,
                      tr.training_date,tr.expiry_date,c.course_name
               FROM employees e
               JOIN training_records tr ON tr.employee_db_id=e.id
               JOIN training_courses c ON c.id=tr.course_id
               WHERE lower(trim(COALESCE(e.employee_id,'')))=lower(trim(?))
                 AND (lower(COALESCE(c.course_name,'')) LIKE '%ptw%'
                      OR lower(COALESCE(c.course_name,'')) LIKE '%permit to work%')
                 AND (COALESCE(trim(tr.expiry_date),'')='' OR trim(tr.expiry_date)>=?)
               ORDER BY tr.training_date DESC, tr.id DESC LIMIT 1""",
            (employee_id,today())
        )
        if not record:
            self.form_fields["permit_holder"].clear()
            self.form_fields["permit_holder_designation"].clear()
            return QMessageBox.warning(
                self,"PTW Training Record Not Found",
                f"PTW record for the user with ID {employee_id} not found.\n\n"
                "A valid PTW / Permit to Work training record is required to submit this permit application."
            )
        self.form_fields["permit_holder"].setText(safe(record["name"]))
        self.form_fields["permit_holder_designation"].setText(safe(record["designation"]))
        if safe(record["company"]).strip() and not self.form_fields["permit_holder_company"].text().strip():
            self.form_fields["permit_holder_company"].setText(safe(record["company"]))
        QMessageBox.information(
            self,"PTW Training Verified",
            f"Permit holder verified from Training Register.\nName: {safe(record['name'])}\n"
            f"Position: {safe(record['designation'])}\nTraining: {safe(record['course_name'])}"
        )

    def _save_permit(self, status="Draft"):
        values={k:self._form_value(k) for k in self.form_fields}
        if any(t.currentText()=="Other" and not other.text().strip() for t,other,ref,_row_widget in self.support_doc_rows):
            QMessageBox.warning(self,"Document Name Required","When Other is selected, enter the supporting document name."); return
        if not values.get("performing_company") or not values.get("location") or not values.get("work_scope"):
            QMessageBox.warning(self,"Required Information","Enter the company / contractor, facility location and work details / scope."); return
        if values["planned_end"] <= values["planned_start"]:
            QMessageBox.warning(self,"Invalid Validity","Planned end must be later than planned start."); return
        if status != "Draft":
            if not values.get("permit_holder_employee_id"):
                QMessageBox.warning(self,"Permit Holder ID Required","Enter the permit holder's ID number and search the Training Register before submitting."); return
            training_record=self.db.fetchone(
                """SELECT e.name,e.designation
                   FROM employees e
                   JOIN training_records tr ON tr.employee_db_id=e.id
                   JOIN training_courses c ON c.id=tr.course_id
                   WHERE lower(trim(COALESCE(e.employee_id,'')))=lower(trim(?))
                     AND (lower(COALESCE(c.course_name,'')) LIKE '%ptw%'
                          OR lower(COALESCE(c.course_name,'')) LIKE '%permit to work%')
                     AND (COALESCE(trim(tr.expiry_date),'')='' OR trim(tr.expiry_date)>=?)
                   ORDER BY tr.training_date DESC, tr.id DESC LIMIT 1""",
                (values["permit_holder_employee_id"],today())
            )
            if not training_record:
                QMessageBox.warning(self,"PTW Training Record Not Found",
                    f"PTW record for the user with ID {values['permit_holder_employee_id']} not found.\n\n"
                    "The permit application cannot be submitted without a valid PTW / Permit to Work training record.")
                return
            values["permit_holder"]=safe(training_record["name"])
            values["permit_holder_designation"]=safe(training_record["designation"])
            self.form_fields["permit_holder"].setText(values["permit_holder"])
            self.form_fields["permit_holder_designation"].setText(values["permit_holder_designation"])
        values["renewal_of_permit_id"] = self.renewal_previous.currentData() if self.renewal_check.isChecked() else None
        if self.renewal_check.isChecked() and not values["renewal_of_permit_id"]:
            QMessageBox.warning(self,"Previous Permit Required","Select the previous permit number for this renewal."); return
        now=_now(); values.update(status=status,updated_at=now)
        if self.selected_permit_id:
            old=self.db.fetchone("SELECT status FROM ptw_permits WHERE id=?",(self.selected_permit_id,))
            if not old: self.selected_permit_id=None
            elif old["status"] not in ("Draft","Returned for Correction"):
                QMessageBox.warning(self,"Editing Restricted","Only draft or returned permits can be edited. Use the revision/reissue process configured by your site procedure."); return
        if self.selected_permit_id:
            cols=list(values.keys()); sql="UPDATE ptw_permits SET "+", ".join(f"{c}=?" for c in cols)+" WHERE id=?"
            self.db.execute(sql,[values[c] for c in cols]+[self.selected_permit_id])
            pid=self.selected_permit_id; number=self.db.fetchone("SELECT permit_number FROM ptw_permits WHERE id=?",(pid,))["permit_number"]
            self._event(pid,"Permit updated",status,status,"Draft fields updated.")
        else:
            number=self._next_number(); values.update(permit_number=number,revision=1,created_by=self._actor(),created_at=now)
            cols=list(values.keys()); cur=self.db.execute("INSERT INTO ptw_permits("+", ".join(cols)+") VALUES("+", ".join("?" for _ in cols)+")",[values[c] for c in cols]); pid=cur.lastrowid
            self.selected_permit_id=pid; self._event(pid,"Permit created","",status,"New permit application saved.")
        self._sync_application_supporting_certificates(pid, values.get("supporting_document_types", ""), values.get("supporting_document_reference", ""))
        self.db.conn.commit()
        self.refresh_all(); self.tabs.setCurrentWidget(self.application_tab)
        self._update_application_permit_summary()
        QMessageBox.information(self,"Permit Saved",f"Permit {number} saved with status {status}. Work is not authorized until the permit is approved and issued.")

    def _submit_permit(self):
        self._save_permit("Submitted")

    def _selected_id(self, table=None):
        t=table or self.permit_table; row=t.currentRow()
        if row<0 or not t.item(row,0): return None
        return t.item(row,0).data(Qt.ItemDataRole.UserRole)

    def _permit_row(self, pid): return self.db.fetchone("SELECT * FROM ptw_permits WHERE id=?",(pid,)) if pid else None

    def _set_combo_permits(self, combo, include_all=False):
        current = combo.currentData()
        combo.blockSignals(True)
        combo.clear()
        if include_all:
            combo.addItem("All Permits", None)
        if combo is getattr(self, "reval_permit", None):
            rows = self.db.fetchall("SELECT id,permit_number,permit_type,status FROM ptw_permits WHERE status='Issued / Active' ORDER BY id DESC")
        elif combo is getattr(self, "susp_permit", None):
            rows = self.db.fetchall("SELECT id,permit_number,permit_type,status FROM ptw_permits ORDER BY id DESC")
        else:
            rows = self.db.fetchall("SELECT id,permit_number,permit_type,status FROM ptw_permits ORDER BY id DESC")
        combo.addItem("Select permit", None)
        for r in rows:
            combo.addItem(f"{r['permit_number']} — {r['permit_type']} ({r['status']})", r["id"])
        idx = combo.findData(current)
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        combo.blockSignals(False)

    def _refresh_suspension_permits(self, *_args):
        if hasattr(self, "susp_permit"):
            self._set_combo_permits(self.susp_permit)
            self.susp_permit.setToolTip("All permits are listed. The selected action is validated against the permit's current status when submitted.")

    def _refresh_certificates_for_selected_permit(self, *_args):
        if not hasattr(self, "cert_table"):
            return
        permit_id = self.cert_permit.currentData() if hasattr(self, "cert_permit") else None
        if not permit_id:
            self.cert_table.setRowCount(0)
            return
        rows = self.db.fetchall(
            "SELECT c.*,p.permit_number FROM ptw_certificates c JOIN ptw_permits p ON p.id=c.permit_id WHERE c.permit_id=? ORDER BY c.id DESC",
            (permit_id,)
        )
        self._fill_table(self.cert_table, rows, ["certificate_ref","permit_number","certificate_type","responsible_person","valid_until","status","verification_status","attachment_path","comments"])

    def _refresh_documents_for_selected_permit(self, *_args):
        if not hasattr(self, "doc_table"):
            return
        permit_id = self.doc_permit.currentData() if hasattr(self, "doc_permit") else None
        if not permit_id:
            self.doc_table.setRowCount(0)
            return
        rows = self.db.fetchall(
            "SELECT d.*,p.permit_number FROM ptw_documents d JOIN ptw_permits p ON p.id=d.permit_id WHERE d.permit_id=? ORDER BY d.id DESC",
            (permit_id,)
        )
        self._fill_table(self.doc_table, rows, ["permit_number", "document_type", "file_path", "description", "uploaded_by", "uploaded_at", "file_path"])

    def _refresh_combo_lists(self):
        for name in ["cert_permit","renew_permit","susp_permit","close_permit","doc_permit","reval_permit"]:
            c=getattr(self,name,None)
            if c is not None:self._set_combo_permits(c)
        if hasattr(self,"renewal_previous"):
            current=self.renewal_previous.currentData(); self.renewal_previous.clear(); self.renewal_previous.addItem("Select previous permit",None)
            for r in self.db.fetchall("SELECT id,permit_number,permit_type,status FROM ptw_permits ORDER BY id DESC"):
                self.renewal_previous.addItem(f"{r['permit_number']} — {r['permit_type']} ({r['status']})",r["id"])
            ix=self.renewal_previous.findData(current)
            if ix>=0:self.renewal_previous.setCurrentIndex(ix)

    def _event(self,pid,action,old_status="",new_status="",details="",attachment=""):
        self.db.execute("INSERT INTO ptw_events(permit_id,event_type,previous_status,new_status,details,actor,attachment_path,created_at) VALUES(?,?,?,?,?,?,?,?)",(pid,action,old_status,new_status,details,self._actor(),attachment,_now()))

    def _fill_table(self, table, rows, fields, id_field="id"):
        table.setSortingEnabled(False); table.setRowCount(len(rows))
        for i,r in enumerate(rows):
            for j,field in enumerate(fields):
                val=r[field] if field in r.keys() else ""
                item=QTableWidgetItem(_s(val)); item.setData(Qt.ItemDataRole.UserRole,r[id_field]); item.setToolTip(_s(val)); table.setItem(i,j,item)
        table.resizeRowsToContents(); table.setSortingEnabled(True)

    def refresh_register(self):
        if not hasattr(self,"permit_table"): return
        rows=self.db.fetchall("SELECT * FROM ptw_permits ORDER BY id DESC")
        needle=self.reg_search.text().strip().lower(); status=self.reg_status.currentText(); typ=self.reg_type.currentText()
        filtered=[]
        for r in rows:
            if status!="All Statuses" and r["status"]!=status: continue
            if typ!="All Types" and r["permit_type"]!=typ: continue
            searchable=[_s(r[k]).lower() for k in ["permit_number","performing_company","client_authority","subcontractor","project_number","location","permit_holder","work_scope","permit_type"]]
            if needle and needle not in " ".join(searchable): continue
            filtered.append(r)
        data=[]
        for r in filtered:
            previous=self.db.fetchone("SELECT permit_number FROM ptw_permits WHERE id=?",(r["renewal_of_permit_id"],)) if r["renewal_of_permit_id"] else None
            issued=0; worked=0
            try:
                start=datetime.fromisoformat(r["planned_start"]).date(); end=datetime.fromisoformat(r["planned_end"]).date()
                issued=max(0,(end-start).days+1)
            except Exception: pass
            dayrows=self.db.fetchall("SELECT day_status FROM ptw_daily_revalidation WHERE permit_id=?",(r["id"],))
            worked=sum(1 for d in dayrows if d["day_status"]=="Revalidated")
            data.append(dict(r,previous_permit=previous["permit_number"] if previous else "",days_issued=issued,days_worked=worked))
        fields=["permit_number","previous_permit","performing_company","client_authority","subcontractor","permit_type","location","project_number","jsa_number","method_statement_number","supporting_document_types","supporting_document_reference","loto_equipment","loto_number","work_scope","permit_holder","permit_holder_company","permit_holder_employee_id","planned_start","planned_end","days_issued","days_worked","approved_by","approval_date","status"]
        self._fill_table(self.permit_table,data,fields)

    def refresh_all(self):
        self.refresh_register()
        rows=self.db.fetchall("SELECT * FROM ptw_permits ORDER BY id DESC")
        counts={s:sum(1 for r in rows if r["status"]==s) for s in PTW_STATUSES}
        warning=int(self._setting("expiry_warning_hours","24")); expiring=0; expired=0; missing=0; pending_handover=0; pending_closure=0
        now=datetime.now()
        for r in rows:
            try:
                end=datetime.fromisoformat(r["planned_end"])
                if end<now and r["status"] not in ("Closed","Cancelled"): expired+=1
                elif now<=end<=now+timedelta(hours=warning) and r["status"] not in ("Closed","Cancelled"): expiring+=1
            except Exception: pass
            required=[x.strip() for x in self._setting("required_certificate_types","").split(";") if x.strip()]
            existing={_s(c["certificate_type"]) for c in self.db.fetchall("SELECT certificate_type FROM ptw_certificates WHERE permit_id=? AND status IN ('Approved','Valid')",(r["id"],))}
            if any(x not in existing for x in required): missing+=1
            h=self.db.fetchone("SELECT status FROM ptw_handovers WHERE permit_id=? ORDER BY id DESC LIMIT 1",(r["id"],))
            if h and h["status"]=="Pending": pending_handover+=1
            c=self.db.fetchone("SELECT status FROM ptw_closures WHERE permit_id=? ORDER BY id DESC LIMIT 1",(r["id"],))
            if r["status"]=="Pending Closure" or (c and c["status"]=="Pending"): pending_closure+=1
        if hasattr(self, "dashboard_summary"):
            self.dashboard_summary.setText("   |   ".join([f"Total permits: {len(rows)}",f"Active: {counts['Issued / Active']}",f"Draft: {counts['Draft']}",f"Pending approval: {counts['Submitted']+counts['Under Review']}",f"Expiring soon: {expiring}",f"Expired: {expired}",f"Suspended: {counts['Suspended']}",f"Pending handovers: {pending_handover}",f"Pending closures: {pending_closure}",f"Missing configured certificates: {missing}"]))
        drows=[]
        for r in rows:
            action="Review application" if r["status"] in ("Submitted","Under Review") else "Revalidate expired permit" if r["status"]=="Expired" else "Complete required checks" if r["status"] in ("Draft","Suspended","Pending Closure") else "Monitor expiry"
            if r["status"] not in ("Closed","Cancelled") and r["planned_end"]:
                try:
                    end=datetime.fromisoformat(r["planned_end"])
                    if end<now and r["status"]!="Expired": action="Expired — do not continue work"
                    elif end<=now+timedelta(hours=warning): action="Expiring soon — review"
                except Exception: pass
            drows.append(dict(r, action=action))
        if hasattr(self, "dashboard_table"):
            self._fill_table(self.dashboard_table,drows,["permit_number","permit_type","performing_company","location","planned_end","status","action"])
        self._refresh_combo_lists(); self._refresh_other_tables(); self._refresh_authority_workflows()
        if hasattr(self,"reval_permit"): self._load_revalidation_days()

    def _open_revalidation_selected(self):
        pid=self._selected_id()
        if not pid:
            return QMessageBox.information(self,"Select Permit","Select a permit row in the Permit Register first.")
        idx=self.reval_permit.findData(pid)
        if idx<0:
            self._refresh_combo_lists(); idx=self.reval_permit.findData(pid)
        if idx>=0:self.reval_permit.setCurrentIndex(idx)
        self.tabs.setCurrentWidget(self.revalidation_tab)
        self._load_revalidation_days()

    def _load_revalidation_days(self):
        if not hasattr(self,"reval_table"): return
        pid=self.reval_permit.currentData(); self.reval_table.setRowCount(0)
        r=self._permit_row(pid)
        if not r:
            self.reval_info.setText("Select an approved permit to view its planned dates and daily revalidation record."); self.reval_summary.setText("Days issued: 0 | Days worked: 0"); return
        if r["status"] != "Issued / Active":
            self.reval_info.setText("Only issued / active permits are available for Daily Revalidation."); self.reval_summary.setText("Days issued: 0 | Days worked: 0"); return
        self.reval_holder.setText(_s(r["permit_holder"]))
        self.reval_authority.setText(_s(r["permit_issuer"] or r["performing_authority"]))
        try:
            start=datetime.fromisoformat(r["planned_start"]).date(); end=datetime.fromisoformat(r["planned_end"]).date()
        except Exception:
            self.reval_info.setText("The selected permit does not have valid planned start/end dates."); return
        if end<start:
            self.reval_info.setText("Permit end date is earlier than start date. Correct the permit dates first."); return
        issued=(end-start).days+1
        existing={x["work_date"]:x for x in self.db.fetchall("SELECT * FROM ptw_daily_revalidation WHERE permit_id=?",(pid,))}
        if existing:
            latest=existing[sorted(existing.keys())[-1]]
            self.reval_holder.setText(_s(latest["permit_holder"]) or _s(r["permit_holder"]))
            self.reval_authority.setText(_s(latest["permit_authority"]) or _s(r["permit_issuer"] or r["performing_authority"]))
        days=(end-start).days+1
        self.reval_table.setRowCount(days)
        statuses=["Pending","Revalidated","Suspended","No Work","On Hold","Rejected"]
        worked=0
        for i in range(days):
            d=start+timedelta(days=i); key=d.isoformat(); saved=existing.get(key)
            date_item=QTableWidgetItem(key); date_item.setData(Qt.ItemDataRole.UserRole,pid); self.reval_table.setItem(i,0,date_item)
            self.reval_table.setItem(i,1,QTableWidgetItem(d.strftime("%A")))
            combo=QComboBox(); combo.addItems(statuses)
            if saved:
                ix=combo.findText(saved["day_status"]); combo.setCurrentIndex(ix if ix>=0 else 0)
                holder_text=_s(saved["permit_holder"]) or _s(r["permit_holder"])
                authority_text=_s(saved["permit_authority"]) or _s(r["permit_issuer"] or r["performing_authority"])
                comments=QTableWidgetItem(_s(saved["comments"]))
            else:
                holder_text=_s(r["permit_holder"])
                authority_text=_s(r["permit_issuer"] or r["performing_authority"])
                comments=QTableWidgetItem("")
            if combo.currentText()=="Revalidated": worked+=1
            self.reval_table.setCellWidget(i,2,combo)
            self.reval_table.setItem(i,3,QTableWidgetItem(holder_text))
            self.reval_table.setItem(i,4,QTableWidgetItem(authority_text))
            self.reval_table.setItem(i,5,comments)
        self.reval_info.setText(f"Permit: {r['permit_number']} | Planned period: {start.isoformat()} to {end.isoformat()} | Select one status for every day.")
        self.reval_summary.setText(f"Days issued: {issued} | Days worked (revalidated): {worked}")

    def _save_revalidation(self):
        pid=self.reval_permit.currentData(); r=self._permit_row(pid)
        if not r:return QMessageBox.warning(self,"Permit Required","Select a permit from the Permit Register.")
        now=_now(); worked=0
        for i in range(self.reval_table.rowCount()):
            date_item=self.reval_table.item(i,0); status_widget=self.reval_table.cellWidget(i,2)
            holder_item=self.reval_table.item(i,3); authority_item=self.reval_table.item(i,4); comment_item=self.reval_table.item(i,5)
            if not date_item or not status_widget: continue
            day=date_item.text(); status=status_widget.currentText()
            holder=holder_item.text().strip() if holder_item else ""
            authority=authority_item.text().strip() if authority_item else ""
            comments=comment_item.text().strip() if comment_item else ""
            if status=="Revalidated" and (not holder or not authority):
                return QMessageBox.warning(self,"Daily Names Required",f"Enter the PTW holder and permit issuing authority for {day} before marking it Revalidated.")
            self.db.execute("INSERT INTO ptw_daily_revalidation(permit_id,work_date,day_status,comments,recorded_by,recorded_at,permit_holder,permit_authority) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(permit_id,work_date) DO UPDATE SET day_status=excluded.day_status,comments=excluded.comments,recorded_by=excluded.recorded_by,recorded_at=excluded.recorded_at,permit_holder=excluded.permit_holder,permit_authority=excluded.permit_authority",(pid,day,status,comments,self._actor(),now,holder,authority))
            if status=="Revalidated":worked+=1
            self._event(pid,"Daily permit revalidation",r["status"],r["status"],f"{day}: {status}. PTW Holder: {holder}. Permit Issuing Authority: {authority}. {comments}")
        self.db.conn.commit(); self.refresh_register(); self.reval_summary.setText(f"Days issued: {self.reval_table.rowCount()} | Days worked (revalidated): {worked}")
        QMessageBox.information(self,"Daily Revalidation Saved",f"Daily revalidation entries saved for {r['permit_number']}.")

    def _delete_revalidation_day(self):
        row=self.reval_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Day","Select a daily revalidation row first.")
        pid=self.reval_permit.currentData(); item=self.reval_table.item(row,0)
        if not pid or not item:return
        day=item.text()
        if QMessageBox.question(self,"Delete Daily Record",f"Delete the revalidation record for {day}?",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
        self.db.execute("DELETE FROM ptw_daily_revalidation WHERE permit_id=? AND work_date=?",(pid,day)); self._event(pid,"Daily revalidation record deleted",details=day); self.refresh_all()

    def _load_selected_renewal(self):
        row=self.renewal_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Renewal","Select a renewal row first.")
        rid=self.renewal_table.item(row,0).data(Qt.ItemDataRole.UserRole); r=self.db.fetchone("SELECT * FROM ptw_renewals WHERE id=?",(rid,))
        if not r:return
        self.editing_renewal_id=rid; self.renew_permit.setCurrentIndex(self.renew_permit.findData(r["permit_id"]))
        self.renew_reason.setText(_s(r["reason"]))
        for w,key in [(self.renew_start,"proposed_start"),(self.renew_end,"proposed_end")]:
            dt=QDateTime.fromString(_s(r[key]),"yyyy-MM-dd HH:mm")
            if dt.isValid():w.setDateTime(dt)
        for w,key in [(self.renew_reinspection,"reinspection"),(self.renew_gas,"gas_retest")]:
            ix=w.findText(_s(r[key])); w.setCurrentIndex(ix if ix>=0 else 0)
        self.renew_person.setText(_s(r["responsible_person"]))

    def _update_selected_renewal(self):
        if not self.editing_renewal_id:return QMessageBox.information(self,"Load Renewal","Load a renewal row for editing first.")
        r=self.db.fetchone("SELECT * FROM ptw_renewals WHERE id=?",(self.editing_renewal_id,))
        if not r:return
        start=self.renew_start.dateTime().toString("yyyy-MM-dd HH:mm"); end=self.renew_end.dateTime().toString("yyyy-MM-dd HH:mm")
        if datetime.fromisoformat(end)<=datetime.fromisoformat(start):return QMessageBox.warning(self,"Invalid Date","Proposed end must be later than proposed start.")
        self.db.execute("UPDATE ptw_renewals SET permit_id=?,reason=?,proposed_start=?,proposed_end=?,reinspection=?,gas_retest=?,responsible_person=? WHERE id=?",(self.renew_permit.currentData(),self.renew_reason.text().strip(),start,end,self.renew_reinspection.currentText(),self.renew_gas.currentText(),self.renew_person.text().strip(),self.editing_renewal_id))
        self._event(r["permit_id"],"Renewal request updated"); self.editing_renewal_id=None; self.refresh_all()

    def _delete_selected_renewal(self):
        row=self.renewal_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Renewal","Select a renewal row first.")
        rid=self.renewal_table.item(row,0).data(Qt.ItemDataRole.UserRole); r=self.db.fetchone("SELECT * FROM ptw_renewals WHERE id=?",(rid,))
        if not r:return
        if QMessageBox.question(self,"Delete Renewal","Delete the selected renewal request?",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
        self.db.execute("DELETE FROM ptw_renewals WHERE id=?",(rid,)); self._event(r["permit_id"],"Renewal request deleted"); self.refresh_all()

    def _load_selected_suspension(self):
        row=self.susp_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Action","Select a suspension/restart record first.")
        eid=self.susp_table.item(row,0).data(Qt.ItemDataRole.UserRole); e=self.db.fetchone("SELECT * FROM ptw_events WHERE id=?",(eid,))
        if not e:return
        self.editing_suspension_event_id=eid
        ix=self.susp_permit.findData(e["permit_id"]); self.susp_permit.setCurrentIndex(ix if ix>=0 else 0)
        action_map={"Permit suspended":"Suspend","Permit placed on hold":"On Hold","Restart requested":"Restart","Permit restarted":"Restart"}
        ix=self.susp_action.findText(action_map.get(e["event_type"],"Suspend")); self.susp_action.setCurrentIndex(ix if ix>=0 else 0)
        self.susp_reason.setPlainText(_s(e["details"])); self.susp_actor.setText(_s(e["actor"])); self.susp_evidence.setText(_s(e["attachment_path"]))

    def _update_selected_suspension(self):
        if not self.editing_suspension_event_id:return QMessageBox.information(self,"Load Action","Load a suspension/restart record first.")
        e=self.db.fetchone("SELECT * FROM ptw_events WHERE id=?",(self.editing_suspension_event_id,))
        if not e:return
        details=self.susp_reason.toPlainText().strip()
        if not details:return QMessageBox.warning(self,"Reason Required","Enter the reason / site conditions.")
        self.db.execute("UPDATE ptw_events SET details=?,actor=?,attachment_path=? WHERE id=?",(details,self.susp_actor.text().strip() or self._actor(),self.susp_evidence.text().strip(),self.editing_suspension_event_id))
        self.editing_suspension_event_id=None; self.refresh_all()

    def _delete_selected_suspension(self):
        row=self.susp_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Action","Select a suspension/restart record first.")
        eid=self.susp_table.item(row,0).data(Qt.ItemDataRole.UserRole); e=self.db.fetchone("SELECT * FROM ptw_events WHERE id=?",(eid,))
        if not e:return
        if QMessageBox.question(self,"Delete Action Record","Delete the selected PTW suspension/restart event from its workspace?",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
        self.db.execute("DELETE FROM ptw_events WHERE id=?",(eid,)); self.refresh_all()

    def _refresh_authority_workflows(self):
        if hasattr(self,"review_permit"):
            current=self.review_permit.currentData(); self.review_permit.blockSignals(True); self.review_permit.clear(); self.review_permit.addItem("Select submitted permit",None)
            rows=self.db.fetchall("SELECT id,permit_number,permit_type,status FROM ptw_permits WHERE status IN ('Submitted','Under Review') ORDER BY id DESC")
            for r in rows:self.review_permit.addItem(f"{r['permit_number']} — {r['permit_type']} ({r['status']})",r["id"])
            idx=self.review_permit.findData(current); self.review_permit.setCurrentIndex(idx if idx>=0 else 0); self.review_permit.blockSignals(False); self._refresh_review_summary()
            self._fill_table(self.review_table,self.db.fetchall("SELECT permit_number,performing_company,permit_type,location,applicant_name,created_at,status,id FROM ptw_permits WHERE status IN ('Submitted','Under Review','Returned for Correction') ORDER BY id DESC"),["permit_number","performing_company","permit_type","location","applicant_name","created_at","status"])
        self._refresh_issue_authority()
        if hasattr(self,"issue_table"):
            self._fill_table(self.issue_table,self.db.fetchall("SELECT permit_number,performing_company,permit_type,location,permit_holder,approved_by,approval_date,status,id FROM ptw_permits WHERE status IN ('Approved','Issued / Active') ORDER BY id DESC"),["permit_number","performing_company","permit_type","location","permit_holder","approved_by","approval_date","status"])

    def _refresh_other_tables(self):
        if hasattr(self,"cert_table"):
            self._refresh_certificates_for_selected_permit()
        if hasattr(self,"renewal_table"):
            self._fill_table(self.renewal_table,self.db.fetchall("SELECT r.*,p.permit_number FROM ptw_renewals r JOIN ptw_permits p ON p.id=r.permit_id ORDER BY r.id DESC"),["permit_number","previous_start","previous_end","proposed_start","proposed_end","reason","status","reinspection","gas_retest","created_at"])
        if hasattr(self,"susp_table"):
            self._fill_table(self.susp_table,self.db.fetchall("SELECT e.*,p.permit_number FROM ptw_events e JOIN ptw_permits p ON p.id=e.permit_id WHERE e.event_type IN ('Permit suspended','Permit placed on hold','Restart requested','Permit restarted') ORDER BY e.id DESC"),["permit_number","event_type","previous_status","new_status","details","actor","attachment_path","created_at"])
        if hasattr(self,"closure_table"):
            self._fill_table(self.closure_table,self.db.fetchall("SELECT c.*,p.permit_number FROM ptw_closures c JOIN ptw_permits p ON p.id=c.permit_id ORDER BY c.id DESC"),["permit_number","work_complete","site_inspection","tools_removed","personnel_accounted","plant_handover","isolation_disposition","reinstatement","status","created_at"])
        if hasattr(self,"doc_table"):
            self._refresh_documents_for_selected_permit()
    def _setting(self,key,default=""):
        r=self.db.fetchone("SELECT value FROM ptw_settings WHERE key=?",(key,)); return r["value"] if r else default

    def _load_selected(self):
        pid=self._selected_id(); r=self._permit_row(pid)
        if not r: return QMessageBox.information(self,"Select Permit","Select a permit row first.")
        self.selected_permit_id=pid
        holder_id_widget=self.form_fields.get("permit_holder_employee_id")
        if holder_id_widget is not None: holder_id_widget.blockSignals(True)
        for k,w in self.form_fields.items():
            val=r[k] if k in r.keys() else ""
            if isinstance(w,QLineEdit): w.setText(_s(val))
            elif isinstance(w,QTextEdit): w.setPlainText(_s(val))
            elif isinstance(w,QComboBox):
                i=w.findText(_s(val)); w.setCurrentIndex(i if i>=0 else 0)
            elif isinstance(w,QSpinBox): w.setValue(int(val or 0))
            elif isinstance(w,QDateTimeEdit):
                dt=QDateTime.fromString(_s(val),"yyyy-MM-dd HH:mm")
                if dt.isValid(): w.setDateTime(dt)
        if holder_id_widget is not None: holder_id_widget.blockSignals(False)
        self._set_dynamic_values("supporting_document_types", r["supporting_document_types"], r["supporting_document_reference"])
        self._set_dynamic_values("loto_equipment", r["loto_equipment"], r["loto_number"])
        linked_id = r["renewal_of_permit_id"] if "renewal_of_permit_id" in r.keys() else None
        self.renewal_check.setChecked(bool(linked_id))
        idx=self.renewal_previous.findData(linked_id)
        self.renewal_previous.setCurrentIndex(idx if idx>=0 else 0)
        self.save_draft_btn.setText("Update Permit"); self.tabs.setCurrentWidget(self.application_tab)

    def _edit_selected(self):
        pid=self._selected_id(); r=self._permit_row(pid)
        if not r: return QMessageBox.information(self,"Select Permit","Select a permit row first.")
        if r["status"] not in ("Draft","Returned for Correction"):
            return QMessageBox.warning(self,"Editing Restricted","Only draft or returned permits can be edited. Issued history is retained.")
        self._load_selected()

    def _submit_selected(self):
        pid=self._selected_id(); r=self._permit_row(pid)
        if not r: return QMessageBox.information(self,"Select Permit","Select a permit row first.")
        if r["status"] not in ("Draft","Returned for Correction"): return QMessageBox.warning(self,"Invalid State","Only draft or returned permits can be submitted.")
        if not r["permit_holder_employee_id"]:
            return QMessageBox.warning(self,"Permit Holder ID Required","A permit holder ID number linked to a valid PTW / Permit to Work training record is required before submission.")
        training_record=self.db.fetchone(
            """SELECT e.name,e.designation FROM employees e
               JOIN training_records tr ON tr.employee_db_id=e.id
               JOIN training_courses c ON c.id=tr.course_id
               WHERE lower(trim(COALESCE(e.employee_id,'')))=lower(trim(?))
                 AND (lower(COALESCE(c.course_name,'')) LIKE '%ptw%'
                      OR lower(COALESCE(c.course_name,'')) LIKE '%permit to work%')
                 AND (COALESCE(trim(tr.expiry_date),'')='' OR trim(tr.expiry_date)>=?)
               ORDER BY tr.training_date DESC, tr.id DESC LIMIT 1""",
            (r["permit_holder_employee_id"],today())
        )
        if not training_record:
            return QMessageBox.warning(self,"PTW Training Record Not Found",
                f"PTW record for the user with ID {r['permit_holder_employee_id']} not found.\n\n"
                "The permit application cannot be submitted without a valid PTW / Permit to Work training record.")
        if not r["location"] or not r["work_scope"] or not r["work_description"]: return QMessageBox.warning(self,"Incomplete Application","Permit holder, exact location, scope and description are required.")
        self.db.execute("UPDATE ptw_permits SET permit_holder=?,permit_holder_designation=? WHERE id=?",
                        (safe(training_record["name"]),safe(training_record["designation"]),pid))
        self._change_status(pid,"Submitted","Permit submitted for review. Permit holder verified against the Training Register PTW record. Work is not authorized.")

    def _approval_dialog(self):
        pid=self._selected_id(); r=self._permit_row(pid)
        if not r: return QMessageBox.information(self,"Select Permit","Select a permit row first.")
        if r["status"] not in ("Submitted","Under Review","Approved","Returned for Correction"): return QMessageBox.warning(self,"Invalid State","This permit is not in an approval state.")
        decision,ok=QInputDialog.getItem(self,"Permit Approval","Decision",["Approve stage","Return for correction","Reject"],0,False)
        if not ok:return
        role,ok=QInputDialog.getText(self,"Approval Role","Enter your configured role/stage:")
        if not ok or not role.strip():return
        comments,ok=QInputDialog.getText(self,"Approval Comments","Comments / reason (required for return/reject):")
        if not ok:return
        if decision!="Approve stage" and not comments.strip(): return QMessageBox.warning(self,"Reason Required","Enter a return/rejection reason.")
        status="Approved" if decision=="Approve stage" else ("Returned for Correction" if decision=="Return for correction" else "Cancelled")
        # An approval decision alone does not issue a permit.
        self.db.execute("INSERT INTO ptw_approvals(permit_id,stage,user_id,user_name,role_name,decision,comments,document_revision,approval_ref,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",(pid,role.strip(),self._actor(),self._actor(),role.strip(),decision,comments.strip(),r["revision"],f"PTW-APR-{pid}-{_stamp()}",_now()))
        if status=="Approved": self.db.execute("UPDATE ptw_permits SET approved_by=?,approval_date=? WHERE id=?",(self._actor(),_now(),pid))
        self._change_status(pid,status,f"Approval stage: {role}. Decision: {decision}. {comments}")

    def _issue_selected(self):
        pid=self._selected_id()
        if not pid:
            return QMessageBox.information(self,"Select Permit","Select a permit row first.")
        self._issue_permit_id(pid)
        self.refresh_all()
        if self.selected_permit_id == pid:
            self._update_application_permit_summary()

    def _change_status(self,pid,status,details=""):
        r=self._permit_row(pid)
        if not r:return
        # Do not silently reactivate expired permits.
        if r["status"] in ("Closed","Cancelled") and status not in (r["status"],): return QMessageBox.warning(self,"Final Status","Closed/cancelled permits cannot be reactivated. Create a linked new permit if required by procedure.")
        old=r["status"]
        self.db.execute("UPDATE ptw_permits SET status=?,updated_at=? WHERE id=?",(status,_now(),pid))
        self._event(pid,details.split(".")[0] or "Status changed",old,status,details)
        self.refresh_all()

    def _add_certificate(self):
        pid=self.cert_permit.currentData()
        if pid is None:return QMessageBox.warning(self,"Permit Required","Create a permit before adding a supporting certificate.")
        number=self._permit_row(pid)["permit_number"]
        ref=f"{number}-CERT-{self.db.fetchone('SELECT COUNT(*) n FROM ptw_certificates WHERE permit_id=?',(pid,))['n']+1:02d}"
        now=_now(); self.db.execute("INSERT INTO ptw_certificates(certificate_ref,permit_id,certificate_type,status,responsible_person,valid_until,verification_status,comments,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)",(ref,pid,self.cert_type.currentText(),self.cert_status.currentText(),self.cert_responsible.text().strip(),self.cert_valid_until.dateTime().toString("yyyy-MM-dd HH:mm"),self.cert_verification.currentText(),self.cert_comments.text().strip(),now,now))
        self._event(pid,"Supporting certificate added",details=f"{ref}: {self.cert_type.currentText()}"); self.refresh_all()

    def _update_certificate(self):
        row=self.cert_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Certificate","Select a certificate row first.")
        cid=self.editing_certificate_id or self.cert_table.item(row,0).data(Qt.ItemDataRole.UserRole)
        cert=self.db.fetchone("SELECT * FROM ptw_certificates WHERE id=?",(cid,))
        if not cert:return
        self.db.execute("UPDATE ptw_certificates SET permit_id=?,certificate_type=?,status=?,responsible_person=?,valid_until=?,verification_status=?,comments=?,updated_at=? WHERE id=?",(self.cert_permit.currentData(),self.cert_type.currentText(),self.cert_status.currentText(),self.cert_responsible.text().strip(),self.cert_valid_until.dateTime().toString("yyyy-MM-dd HH:mm"),self.cert_verification.currentText(),self.cert_comments.text().strip(),_now(),cid))
        self._event(cert["permit_id"],"Supporting certificate updated",details=cert["certificate_ref"]); self.editing_certificate_id=None; self.refresh_all()

    def _load_certificate_for_edit(self):
        row=self.cert_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Certificate","Select a certificate row first.")
        cid=self.cert_table.item(row,0).data(Qt.ItemDataRole.UserRole); c=self.db.fetchone("SELECT * FROM ptw_certificates WHERE id=?",(cid,))
        if not c:return
        self.editing_certificate_id=cid; self.cert_permit.setCurrentIndex(self.cert_permit.findData(c["permit_id"]))
        ix=self.cert_type.findText(c["certificate_type"]); self.cert_type.setCurrentIndex(ix if ix>=0 else 0)
        self.cert_responsible.setText(_s(c["responsible_person"]))
        dt=QDateTime.fromString(_s(c["valid_until"]),"yyyy-MM-dd HH:mm")
        if dt.isValid():self.cert_valid_until.setDateTime(dt)
        for w,key in [(self.cert_status,"status"),(self.cert_verification,"verification_status")]:
            ix=w.findText(_s(c[key])); w.setCurrentIndex(ix if ix>=0 else 0)
        self.cert_comments.setText(_s(c["comments"]))

    def _attach_certificate(self):
        row=self.cert_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Certificate","Select a certificate row first.")
        cid=self.cert_table.item(row,0).data(Qt.ItemDataRole.UserRole); path,_=QFileDialog.getOpenFileName(self,"Select Certificate File")
        if not path:return
        cert=self.db.fetchone("SELECT * FROM ptw_certificates WHERE id=?",(cid,)); dest=self._copy_attachment(path,cert["permit_id"])
        self.db.execute("UPDATE ptw_certificates SET attachment_path=?,updated_at=? WHERE id=?",(str(dest),_now(),cid)); self._event(cert["permit_id"],"Certificate evidence attached",details=cert["certificate_ref"],attachment=str(dest)); self.refresh_all()

    def _add_renewal(self):
        pid=self.renew_permit.currentData(); r=self._permit_row(pid)
        if not r:return QMessageBox.warning(self,"Permit Required","Select a permit.")
        start=self.renew_start.dateTime().toString("yyyy-MM-dd HH:mm"); end=self.renew_end.dateTime().toString("yyyy-MM-dd HH:mm"); reason=self.renew_reason.text().strip()
        if not reason or not start or not end:return QMessageBox.warning(self,"Required Information","Enter a reason and proposed start/end.")
        try:
            if datetime.fromisoformat(end)<=datetime.fromisoformat(start): raise ValueError()
        except Exception:return QMessageBox.warning(self,"Invalid Date","Use YYYY-MM-DD HH:MM and ensure proposed end is after start.")
        now=_now()
        self.db.execute("INSERT INTO ptw_renewals(permit_id,reason,previous_start,previous_end,proposed_start,proposed_end,status,responsible_person,reinspection,gas_retest,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(pid,reason,r["planned_start"],r["planned_end"],start,end,"Requested",self.renew_person.text().strip(),self.renew_reinspection.currentText(),self.renew_gas.currentText(),now))
        # A renewal is represented as a separate linked permit in the register; it remains Submitted until approved.
        renewed = {k: r[k] for k in r.keys() if k not in ("id", "permit_number")}
        renewed.update({"permit_number": self._next_number(), "original_permit_ref": r["permit_number"],
                        "renewal_of_permit_id": pid, "planned_start": start, "planned_end": end,
                        "status": "Submitted", "revision": int(r["revision"] or 1) + 1,
                        "approved_by": "", "approval_date": "",
                        "created_by": self._actor(), "created_at": now, "updated_at": now})
        cols = list(renewed.keys())
        cur = self.db.execute("INSERT INTO ptw_permits(" + ",".join(cols) + ") VALUES(" + ",".join("?" for _ in cols) + ")", [renewed[c] for c in cols])
        new_pid = cur.lastrowid
        self._event(pid,"Renewal requested",r["status"],r["status"],f"Renewal permit {renewed['permit_number']} created for {start} to {end}.")
        self._event(new_pid,"Renewal permit created", "", "Submitted", f"Renewal of {r['permit_number']}; awaiting review and approval.")
        self.db.conn.commit()
        QMessageBox.information(self,"Renewal Recorded",f"Renewal permit {renewed['permit_number']} has been added to the Permit Register with Submitted status. It must be approved before work can resume.")
        self.refresh_all()

    def _add_handover(self):
        pid=self.hand_permit.currentData(); r=self._permit_row(pid)
        if not r:return QMessageBox.warning(self,"Permit Required","Select a permit.")
        if r["status"]!="Issued / Active":return QMessageBox.warning(self,"Permit Not Active","Handover can only be recorded for an active permit.")
        if not self.hand_out.text().strip() or not self.hand_in.text().strip() or not self.hand_out_ack.isChecked() or not self.hand_in_ack.isChecked(): return QMessageBox.warning(self,"Acknowledgements Required","Enter outgoing/incoming responsible persons and record both acknowledgements.")
        if self.hand_gas.currentIndex()==2 or self.hand_iso.currentIndex()==2:return QMessageBox.warning(self,"Safety Verification Pending","Resolve or document required gas/isolation verification before handover.")
        status="Accepted" if self.hand_auth.isChecked() else "Pending Authority Confirmation"
        self.db.execute("INSERT INTO ptw_handovers(permit_id,outgoing_person,incoming_person,shift_name,work_completed,work_remaining,outstanding_hazards,gas_test_status,isolation_status,outgoing_ack,incoming_ack,authority_confirmation,status,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(pid,self.hand_out.text().strip(),self.hand_in.text().strip(),self.hand_shift.text().strip(),self.hand_done.toPlainText(),self.hand_remaining.toPlainText(),self.hand_hazards.toPlainText(),self.hand_gas.currentText(),self.hand_iso.currentText(),"Acknowledged","Accepted","Confirmed" if self.hand_auth.isChecked() else "Pending",status,_now()))
        self._event(pid,"Shift handover recorded",r["status"],r["status"],f"Outgoing: {self.hand_out.text().strip()}; incoming: {self.hand_in.text().strip()}; status: {status}"); self.refresh_all()

    def _suspension_action(self):
        pid=self.susp_permit.currentData(); permit=self._permit_row(pid)
        if not permit:
            return QMessageBox.warning(self,"Permit Required","Select a permit.")
        action=self.susp_action.currentText(); reason=self.susp_reason.toPlainText().strip(); evidence=self.susp_evidence.text().strip()
        if not reason:
            return QMessageBox.warning(self,"Reason Required","Enter the reason before submitting this action.")
        old_status=permit["status"]
        actor=self.susp_actor.text().strip() or self._actor()
        if action == "Suspend":
            if old_status not in ("Issued / Active", "On Hold"):
                return QMessageBox.warning(self,"Action Not Allowed","Only an issued/active or on-hold permit can be suspended.")
            new_status="Suspended"; event_type="Permit suspended"
        elif action == "On Hold":
            if old_status not in ("Issued / Active", "Suspended"):
                return QMessageBox.warning(self,"Action Not Allowed","Only an issued/active or suspended permit can be put on hold.")
            new_status="On Hold"; event_type="Permit placed on hold"
        else:
            if old_status not in ("Suspended", "On Hold"):
                return QMessageBox.warning(self,"Action Not Allowed","Only a suspended or on-hold permit can be restarted.")
            new_status="Issued / Active"; event_type="Permit restarted"
        self.db.execute("UPDATE ptw_permits SET status=?,updated_at=? WHERE id=?",(new_status,_now(),pid))
        self.db.execute("INSERT INTO ptw_events(permit_id,event_type,previous_status,new_status,details,actor,attachment_path,created_at) VALUES(?,?,?,?,?,?,?,?)",
                        (pid,event_type,old_status,new_status,reason,actor,evidence,_now()))
        self.db.conn.commit()
        self.susp_reason.clear(); self.susp_evidence.clear(); self.susp_actor.clear()
        self.refresh_all()
        QMessageBox.information(self,"PTW Status Updated",f"Permit {permit['permit_number']} status changed to {new_status}.")

    def _close_permit_simple(self):
        pid = self.close_permit.currentData()
        permit = self._permit_row(pid)
        if not permit:
            return QMessageBox.warning(self, "Permit Required", "Select a permit from the dropdown.")
        if not self.close_confirm_check.isChecked():
            return QMessageBox.warning(self, "Confirmation Required", "Tick the confirmation checkbox before closing the permit.")
        answer = QMessageBox.question(
            self, "Confirm Permit Closure",
            f"Close permit {permit['permit_number']}?\nIts status will become Closed, regardless of its current status.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if answer != QMessageBox.StandardButton.Yes:
            return
        now = _now()
        existing = self.db.fetchone("SELECT id FROM ptw_closures WHERE permit_id=? ORDER BY id DESC LIMIT 1", (pid,))
        if existing:
            self.db.execute("UPDATE ptw_closures SET work_complete='Completed',status='Closed',closure_comments=?,created_at=? WHERE id=?",
                            ("Permit closed using the simplified closeout action.", now, existing["id"]))
        else:
            self.db.execute("INSERT INTO ptw_closures(permit_id,work_complete,closure_comments,status,created_at) VALUES(?,?,?,?,?)",
                            (pid, "Completed", "Permit closed using the simplified closeout action.", "Closed", now))
        old_status = permit["status"]
        self.db.execute("UPDATE ptw_permits SET status='Closed',closure_date=?,updated_at=? WHERE id=?",
                        (datetime.now().strftime("%Y-%m-%d %H:%M"), now, pid))
        self._event(pid, "Permit closed", old_status, "Closed", "Permit closed using the simplified checkbox confirmation.")
        self.db.conn.commit()
        self.close_confirm_check.setChecked(False)
        self.refresh_all()
        QMessageBox.information(self, "Permit Closed", f"Permit {permit['permit_number']} has been closed successfully.")

    def _submit_closure(self):
        # Backward-compatible entry point for older PTW UI actions.
        return self._close_permit_simple()

    def _choose_into(self,line):
        path,_=QFileDialog.getOpenFileName(self,"Choose Evidence / Document")
        if path:line.setText(path)

    def _copy_attachment(self,path,pid):
        base=Path(os.environ.get("LOCALAPPDATA",Path.home()))/"HSE Management System"/"attachments"/"ptw"/str(pid)
        base.mkdir(parents=True,exist_ok=True); dest=base/f"{_stamp()}_{Path(path).name}"; shutil.copy2(path,dest); return dest

    def _add_document(self):
        pid=self.doc_permit.currentData(); path=self.doc_path.text().strip()
        if not pid:return QMessageBox.warning(self,"Permit Required","Select a permit.")
        if not path or not Path(path).is_file():return QMessageBox.warning(self,"File Required","Choose an existing file.")
        dest=self._copy_attachment(path,pid); self.db.execute("INSERT INTO ptw_documents(permit_id,document_type,file_path,description,uploaded_by,uploaded_at) VALUES(?,?,?,?,?,?)",(pid,self.doc_type.currentText(),str(dest),self.doc_desc.text().strip(),self._actor(),_now())); self._event(pid,"Safety document attached",details=self.doc_type.currentText()+": "+Path(path).name,attachment=str(dest)); self.doc_path.clear(); self.refresh_all()

    def _load_selected_document(self):
        row=self.doc_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Document","Select a document row first.")
        did=self.doc_table.item(row,0).data(Qt.ItemDataRole.UserRole); d=self.db.fetchone("SELECT * FROM ptw_documents WHERE id=?",(did,))
        if not d:return
        self.editing_document_id=did; self.doc_permit.setCurrentIndex(self.doc_permit.findData(d["permit_id"]))
        ix=self.doc_type.findText(_s(d["document_type"])); self.doc_type.setCurrentIndex(ix if ix>=0 else 0)
        self.doc_desc.setText(_s(d["description"])); self.doc_path.setText(_s(d["file_path"]))

    def _update_selected_document(self):
        if not self.editing_document_id:return QMessageBox.information(self,"Load Document","Load a document row for editing first.")
        d=self.db.fetchone("SELECT * FROM ptw_documents WHERE id=?",(self.editing_document_id,))
        if not d:return
        path=self.doc_path.text().strip()
        if not path or not Path(path).is_file():return QMessageBox.warning(self,"File Required","Choose an existing document file.")
        dest=self._copy_attachment(path,d["permit_id"]) if path!=d["file_path"] else Path(path)
        self.db.execute("UPDATE ptw_documents SET permit_id=?,document_type=?,file_path=?,description=?,uploaded_by=?,uploaded_at=? WHERE id=?",(self.doc_permit.currentData(),self.doc_type.currentText(),str(dest),self.doc_desc.text().strip(),self._actor(),_now(),self.editing_document_id))
        self._event(d["permit_id"],"PTW document updated",details=self.doc_type.currentText()); self.editing_document_id=None; self.refresh_all()

    def _delete_document(self):
        row=self.doc_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Document","Select a document row first.")
        did=self.doc_table.item(row,0).data(Qt.ItemDataRole.UserRole); d=self.db.fetchone("SELECT * FROM ptw_documents WHERE id=?",(did,))
        if not d:return
        if QMessageBox.question(self,"Delete Document","Delete the selected PTW attachment record?",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
        self.db.execute("DELETE FROM ptw_documents WHERE id=?",(did,)); self._event(d["permit_id"],"PTW document record deleted",details=d["document_type"]); self.refresh_all()

    def _open_document(self):
        row=self.doc_table.currentRow()
        if row<0:return QMessageBox.information(self,"Select Document","Select a document row first.")
        path=self.doc_table.item(row,2).text()
        if not Path(path).is_file():return QMessageBox.warning(self,"File Missing","The stored attachment cannot be found.")
        from PySide6.QtCore import QUrl
        from PySide6.QtGui import QDesktopServices
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def _export_selected(self):
        pid=self._selected_id(); r=self._permit_row(pid)
        if not r:return QMessageBox.information(self,"Select Permit","Select a permit row first.")
        path,_=QFileDialog.getSaveFileName(self,"Export Permit Record",f"{r['permit_number']}.txt","Text Files (*.txt)")
        if not path:return
        with open(path,"w",encoding="utf-8") as f:
            for k in r.keys(): f.write(f"{k.replace('_',' ').title()}: {_s(r[k])}\n")
            f.write("\nSUPPORTING CERTIFICATES\n")
            for c in self.db.fetchall("SELECT * FROM ptw_certificates WHERE permit_id=?",(pid,)): f.write(" | ".join(_s(c[k]) for k in ["certificate_ref","certificate_type","status","valid_until","attachment_path"])+"\n")
            f.write("\nDAILY REVALIDATION\n")
            for d in self.db.fetchall("SELECT work_date,day_status,comments,recorded_by,recorded_at FROM ptw_daily_revalidation WHERE permit_id=? ORDER BY work_date",(pid,)):
                f.write(f"{d['work_date']} | {d['day_status']} | {d['comments']} | {d['recorded_by']} | {d['recorded_at']}\n")
            f.write("\nATTACHMENTS\n")
            for d in self.db.fetchall("SELECT * FROM ptw_documents WHERE permit_id=?",(pid,)): f.write(f"{d['document_type']}: {d['file_path']}\n")
        QMessageBox.information(self,"Exported","Permit record and supporting-document index exported.")

    def _export_excel_register(self):
        path,_=QFileDialog.getSaveFileName(self,"Export PTW Register and Daily Revalidation","ptw_register_and_revalidation.xlsx","Excel Workbook (*.xlsx)")
        if not path:return
        if not path.lower().endswith(".xlsx"):path += ".xlsx"
        self.refresh_register()
        wb=Workbook(); ws=wb.active; ws.title="PTW Register"
        headers=[self.permit_table.horizontalHeaderItem(c).text() for c in range(self.permit_table.columnCount())]
        ws.append(["PERMIT TO WORK REGISTER"]); ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(headers))
        ws["A1"].font=Font(name="Arial",size=16,bold=True,color="FFFFFF"); ws["A1"].fill=PatternFill("solid",fgColor="17365D"); ws["A1"].alignment=Alignment(horizontal="center",vertical="center"); ws.row_dimensions[1].height=28
        ws.append(headers)
        for r in range(self.permit_table.rowCount()):
            ws.append([self.permit_table.item(r,c).text() if self.permit_table.item(r,c) else "" for c in range(self.permit_table.columnCount())])
        headfill=PatternFill("solid",fgColor="D9EAF7"); thin=Side(style="thin",color="B7C9D6")
        for sheet in [ws]:
            for cell in sheet[2]:
                cell.font=Font(bold=True,color="17365D"); cell.fill=headfill; cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cell.border=Border(bottom=thin)
            for row in sheet.iter_rows(min_row=3):
                for cell in row:cell.alignment=Alignment(vertical="top",wrap_text=True)
            sheet.freeze_panes="A3"; sheet.auto_filter.ref=f"A2:{sheet.cell(sheet.max_row,sheet.max_column).coordinate}"
            for col in sheet.columns:
                letter=col[0].column_letter; longest=max((len(str(c.value or "")) for c in col),default=10); sheet.column_dimensions[letter].width=min(max(longest+2,12),32)
        daily=wb.create_sheet("Daily Revalidation")
        dh=["Permit No.","Work Date","Day","Revalidation Mark","Daily Status","Daily PTW Holder Name","Permit Issuing Authority Name","Comments","Recorded By","Recorded At"]
        daily.append(["DAILY PTW REVALIDATION REGISTER"]); daily.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(dh))
        daily["A1"].font=Font(name="Arial",size=16,bold=True,color="FFFFFF"); daily["A1"].fill=PatternFill("solid",fgColor="17365D"); daily["A1"].alignment=Alignment(horizontal="center",vertical="center"); daily.row_dimensions[1].height=28
        daily.append(dh)
        records=self.db.fetchall("SELECT p.permit_number,d.work_date,d.day_status,d.permit_holder,d.permit_authority,d.comments,d.recorded_by,d.recorded_at FROM ptw_daily_revalidation d JOIN ptw_permits p ON p.id=d.permit_id ORDER BY p.permit_number,d.work_date")
        for d in records:
            try:day_name=datetime.fromisoformat(d["work_date"]).strftime("%A")
            except Exception:day_name=""
            status=_s(d["day_status"]); mark="✓" if status=="Revalidated" else "✗" if status in ("Suspended","On Hold","Rejected","No Work") else "—"
            daily.append([d["permit_number"],d["work_date"],day_name,mark,status,d["permit_holder"],d["permit_authority"],d["comments"],d["recorded_by"],d["recorded_at"]])
            fill="E2F0D9" if mark=="✓" else "FCE4D6" if mark=="✗" else "FFF2CC"
            for cell in daily[daily.max_row]:cell.fill=PatternFill("solid",fgColor=fill)
        for cell in daily[2]:
            cell.font=Font(bold=True,color="17365D"); cell.fill=headfill; cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cell.border=Border(bottom=thin)
        for row in daily.iter_rows(min_row=3):
            for cell in row:cell.alignment=Alignment(vertical="top",wrap_text=True)
        daily.freeze_panes="A3"; daily.auto_filter.ref=f"A2:{daily.cell(daily.max_row,daily.max_column).coordinate}"
        for col in daily.columns:
            letter=col[0].column_letter; longest=max((len(str(c.value or "")) for c in col),default=10); daily.column_dimensions[letter].width=min(max(longest+2,12),36)
        wb.save(path)
        QMessageBox.information(self,"Excel Exported",f"PTW Register and Daily Revalidation are saved in separate worksheets. Revalidated = ✓, non-revalidated outcomes = ✗, pending = —.\\n{path}")

    def _export_csv(self):
        return self._export_excel_register()

    def _show_history(self):
        pid=self._selected_id()
        if not pid:return QMessageBox.information(self,"Select Permit","Select a permit row first.")
        rows=self.db.fetchall("SELECT event_type,previous_status,new_status,details,actor,created_at FROM ptw_events WHERE permit_id=? ORDER BY id",(pid,))
        text="\n".join(f"{r['created_at']} | {r['actor']} | {r['event_type']} | {r['previous_status']} → {r['new_status']} | {r['details']}" for r in rows) or "No history yet."
        QMessageBox.information(self,"Permit History",text[:12000])

    def _save_config(self):
        name=self.cfg_type.text().strip()
        if not name:return QMessageBox.warning(self,"Type Name Required","Enter a permit type name.")
        now=_now(); old=self.db.fetchone("SELECT id FROM ptw_permit_types WHERE name=?",(name,))
        if old:
            self.db.execute("UPDATE ptw_permit_types SET workflow_profile=?,validity_hours=?,procedure_name=?,procedure_number=?,procedure_revision=?,effective_date=?,updated_at=? WHERE name=?",(self.cfg_profile.currentText(),self.cfg_validity.value(),self.cfg_proc.text().strip(),self.cfg_procno.text().strip(),self.cfg_procrev.text().strip(),self.cfg_effective.text().strip(),now,name))
        else:
            self.db.execute("INSERT INTO ptw_permit_types(name,active,workflow_profile,validity_hours,procedure_name,procedure_number,procedure_revision,effective_date,updated_at) VALUES(?,1,?,?,?,?,?,?,?)",(name,self.cfg_profile.currentText(),self.cfg_validity.value(),self.cfg_proc.text().strip(),self.cfg_procno.text().strip(),self.cfg_procrev.text().strip(),self.cfg_effective.text().strip(),now))
        for k,v in [("expiry_warning_hours",str(self.cfg_expiry.value())),("approval_roles",self.cfg_roles.text().strip()),("required_certificate_types",self.cfg_required_certs.text().strip())]: self.db.execute("INSERT INTO ptw_settings(key,value,updated_at,updated_by) VALUES(?,?,?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at,updated_by=excluded.updated_by",(k,v,now,self._actor()))
        self.db.execute("INSERT INTO ptw_events(permit_id,event_type,details,actor,created_at) SELECT id,'PTW configuration updated',?,?,? FROM ptw_permits WHERE id=(SELECT MIN(id) FROM ptw_permits)",(f"Permit type/configuration changed: {name}",self._actor(),now)) if self.db.fetchone("SELECT id FROM ptw_permits LIMIT 1") else None
        self.refresh_all(); QMessageBox.information(self,"Configuration Saved","PTW-specific configuration saved. Validate against the governing procedure before use.")



# ============================================================
# APPLICATION START
# ============================================================

def main():

    app = QApplication(sys.argv)

    app.setApplicationName(
        APP_NAME
    )

    app.setOrganizationName(
        "HSE"
    )

    window = MainWindow()

    window.show()

    sys.exit(
        app.exec()
    )


if __name__ == "__main__":
    main()
