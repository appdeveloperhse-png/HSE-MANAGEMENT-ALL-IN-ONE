import sys
import os
import csv
import shutil
import sqlite3
import zipfile
import logging
import json
import re
import calendar
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
        CREATE TABLE IF NOT EXISTS training_matrix_cells (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_name TEXT NOT NULL,
            designation TEXT NOT NULL,
            requirement TEXT DEFAULT 'Not Applicable',
            validity_months INTEGER DEFAULT 0,
            notes TEXT DEFAULT '',
            updated_at TEXT DEFAULT '',
            UNIQUE(course_name, designation)
        );
        CREATE INDEX IF NOT EXISTS idx_training_matrix_cells_course
            ON training_matrix_cells(course_name);
        CREATE INDEX IF NOT EXISTS idx_training_matrix_cells_desig
            ON training_matrix_cells(designation);
        """)

        # Training-only employee fields; additive migrations preserve all existing data.
        employee_columns = {r["name"] for r in self.conn.execute("PRAGMA table_info(employees)").fetchall()}
        if "date_joining" not in employee_columns:
            self.conn.execute("ALTER TABLE employees ADD COLUMN date_joining TEXT DEFAULT ''")
        if "safety_induction_date" not in employee_columns:
            self.conn.execute("ALTER TABLE employees ADD COLUMN safety_induction_date TEXT DEFAULT ''")
        if "photo_path" not in employee_columns:
            self.conn.execute("ALTER TABLE employees ADD COLUMN photo_path TEXT DEFAULT ''")
        # Older versions offered "Either"; the current Training Type choices are Internal or External.
        self.conn.execute("UPDATE training_courses SET provider_type='Internal' WHERE lower(COALESCE(provider_type,''))='either'")

        course_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(training_courses)").fetchall()}
        if "training_hours" not in course_cols:
            self.conn.execute("ALTER TABLE training_courses ADD COLUMN training_hours REAL DEFAULT 0")
        matrix_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(training_matrix_cells)").fetchall()}
        if "training_hours" not in matrix_cols:
            self.conn.execute("ALTER TABLE training_matrix_cells ADD COLUMN training_hours REAL DEFAULT 0")
        if "provider_type" not in matrix_cols:
            self.conn.execute("ALTER TABLE training_matrix_cells ADD COLUMN provider_type TEXT DEFAULT 'Internal'")

        # Training-only compliance extensions. All migrations are additive; legacy
        # records default to Verified so existing PTW/Audit eligibility behaviour is
        # not unexpectedly broken by upgrading the Training module.
        course_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(training_courses)").fetchall()}
        for col, definition in {
            "category": "TEXT DEFAULT 'General HSE'",
            "assessment_required": "INTEGER DEFAULT 0",
            "expiry_policy": "TEXT DEFAULT 'Validity Months'",
        }.items():
            if col not in course_cols:
                self.conn.execute(f"ALTER TABLE training_courses ADD COLUMN {col} {definition}")

        record_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(training_records)").fetchall()}
        for col, definition in {
            "verification_status": "TEXT DEFAULT 'Verified'",
            "verified_by": "TEXT DEFAULT ''",
            "verified_at": "TEXT DEFAULT ''",
            "assessment_result": "TEXT DEFAULT 'Not Required'",
            "assessment_score": "REAL",
        }.items():
            if col not in record_cols:
                self.conn.execute(f"ALTER TABLE training_records ADD COLUMN {col} {definition}")

        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS training_record_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            original_record_id INTEGER,
            employee_db_id INTEGER NOT NULL,
            course_id INTEGER NOT NULL,
            training_date TEXT NOT NULL,
            expiry_date TEXT DEFAULT '',
            provider TEXT DEFAULT '',
            certificate_no TEXT DEFAULT '',
            attachment_path TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT '',
            archived_at TEXT DEFAULT CURRENT_TIMESTAMP,
            archived_by TEXT DEFAULT 'Local User',
            archive_reason TEXT DEFAULT '',
            verification_status TEXT DEFAULT 'Verified',
            verified_by TEXT DEFAULT '',
            verified_at TEXT DEFAULT '',
            assessment_result TEXT DEFAULT 'Not Required',
            assessment_score REAL
        );
        CREATE TABLE IF NOT EXISTS training_audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entity_type TEXT NOT NULL,
            entity_id INTEGER,
            action TEXT NOT NULL,
            details TEXT DEFAULT '',
            changed_by TEXT DEFAULT 'Local User',
            changed_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS training_matrix_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_name TEXT NOT NULL,
            designation TEXT NOT NULL,
            requirement TEXT DEFAULT 'Not Applicable',
            validity_months INTEGER DEFAULT 0,
            notes TEXT DEFAULT '',
            updated_at TEXT DEFAULT '',
            archived_at TEXT DEFAULT CURRENT_TIMESTAMP,
            archive_reason TEXT DEFAULT 'Matrix updated'
        );
        CREATE INDEX IF NOT EXISTS idx_training_record_history_emp_course
            ON training_record_history(employee_db_id, course_id, training_date);
        CREATE INDEX IF NOT EXISTS idx_training_audit_entity
            ON training_audit_log(entity_type, entity_id, changed_at);
        """)

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

        # ---- Audit Management module v2: additive tables/columns (no other module is touched) ----
        _aud_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(audits)").fetchall()}
        for _col, _definition in {"aud_ver": "INTEGER DEFAULT 0", "procedures": "TEXT DEFAULT ''"}.items():
            if _col not in _aud_cols:
                self.conn.execute(f"ALTER TABLE audits ADD COLUMN {_col} {_definition}")
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS aud_findings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            audit_id INTEGER NOT NULL,
            seq INTEGER NOT NULL,
            finding_no TEXT UNIQUE,
            auditor_name TEXT DEFAULT '',
            clause_sno TEXT DEFAULT '',
            sub_clause TEXT DEFAULT '',
            checklist_pts TEXT DEFAULT '',
            evidence TEXT DEFAULT '',
            corrective_action TEXT DEFAULT '',
            responsible TEXT DEFAULT '',
            target_date TEXT DEFAULT '',
            capa_required TEXT DEFAULT 'No',
            capa_number TEXT DEFAULT '',
            status TEXT DEFAULT 'Open',
            route_stage TEXT DEFAULT 'Audit Register',
            procedure TEXT DEFAULT '',
            created_at TEXT,
            updated_at TEXT,
            FOREIGN KEY(audit_id) REFERENCES audits(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS aud_notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            audit_id INTEGER NOT NULL,
            doc_number TEXT UNIQUE,
            to_party TEXT DEFAULT '',
            from_party TEXT DEFAULT '',
            subject TEXT DEFAULT '',
            scope TEXT DEFAULT '',
            audit_date TEXT DEFAULT '',
            audit_time TEXT DEFAULT '',
            venue TEXT DEFAULT '',
            auditee TEXT DEFAULT '',
            audit_team TEXT DEFAULT '',
            reference_docs TEXT DEFAULT '',
            remarks TEXT DEFAULT '',
            created_at TEXT,
            updated_at TEXT,
            FOREIGN KEY(audit_id) REFERENCES audits(id) ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_aud_findings_audit ON aud_findings(audit_id);
        CREATE INDEX IF NOT EXISTS idx_aud_notifications_audit ON aud_notifications(audit_id);
        """)

        # Close-out / tracking workflow for Audit findings (additive only).
        _af_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(aud_findings)").fetchall()}
        for _col, _definition in {
            "closeout_status": "TEXT DEFAULT 'Open'",
            "closeout_comments": "TEXT DEFAULT ''",
            "closeout_evidence": "TEXT DEFAULT ''",
            "closeout_submitted_by": "TEXT DEFAULT ''",
            "closeout_submitted_at": "TEXT DEFAULT ''",
            "closeout_approved_by": "TEXT DEFAULT ''",
            "closeout_approved_at": "TEXT DEFAULT ''",
            "observation_text": "TEXT DEFAULT ''",
        }.items():
            if _col not in _af_cols:
                self.conn.execute(f"ALTER TABLE aud_findings ADD COLUMN {_col} {_definition}")
        self.conn.execute("""CREATE TABLE IF NOT EXISTS aud_finding_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            finding_id INTEGER NOT NULL,
            file_path TEXT,
            attachment_type TEXT DEFAULT 'Close-out Evidence',
            uploaded_by TEXT DEFAULT '',
            uploaded_at TEXT DEFAULT '',
            FOREIGN KEY(finding_id) REFERENCES aud_findings(id) ON DELETE CASCADE
        )""")

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
    "Internal Audit", "External Audit", "Client Audit", "Certification Audit",
    "Surveillance Audit", "Recertification Audit", "Completion Audit",
    "Compliance Audit", "Verification Audit", "Follow-up Audit", "Supplier Audit",
    "Contractor Audit", "Special Audit",
    # Legacy choices retained for compatibility with existing audit records.
    "Regulatory Audit", "Project Audit", "Other"
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
# AUDIT MANAGEMENT MODULE (v2) - DATA, ROUTING AND REPORT LOGIC
# Self-contained. Reads/writes only: audits (parent), aud_findings (child),
# aud_notifications, and inserts routed rows into the existing capa table.
# ============================================================

AUD_TYPES = [
    "Internal", "External", "ISO Certification Audit", "Surveillance Audit",
    "Recertification Audit", "Supplier/Vendor Audit", "Customer Audit",
    "Process Audit", "Product Audit", "Compliance Audit", "HSE Audit",
    "Environmental Audit", "Quality Audit",
    "Integrated Management System (IMS) Audit", "Follow-up Audit",
    "Special Audit", "Certificate", "Regulatory"
]
AUD_PROCEDURE_TYPES = ("Internal", "External")

def aud_is_internal_type(value):
    """Recognize both legacy 'Internal' and expanded 'Internal Audit' labels."""
    normalized = safe(value).strip().casefold()
    return normalized in ("internal", "internal audit")
AUD_FINDING_STATUSES = ["Open", "In Progress", "Closed"]
AUD_CONDUCTED_STATUSES = ("Submitted", "In Progress", "Completed")
AUD_SCHEDULED_STATUSES = ("Scheduled", "Submitted", "In Progress", "Completed")
AUD_STAGE_CAPA = "CAPA Module"
AUD_STAGE_CLOSURE = "Audit Report / Closure"
AUD_STAGE_REGISTER = "Audit Register"
AUD_CHECKLIST_HEADERS = [
    "Audit No.", "Audit Finding No.", "Auditor Name", "Clause S.No", "Sub Clause",
    "Checklist Pts", "Evidence", "Corrective Action", "Responsible", "Target Date",
    "CAPA (Yes/No)", "Status"
]
AUD_REGISTER_HEADERS = AUD_CHECKLIST_HEADERS + ["CAPA No.", "Routed To"]

_AUD_LIBRARY_CACHE = {}


def audit_procedure_library():
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
    return expanded


def aud_procedure_library():
    if not _AUD_LIBRARY_CACHE:
        _AUD_LIBRARY_CACHE.update(audit_procedure_library())
    return _AUD_LIBRARY_CACHE

def aud_iso_standard_checklist(standard):
    """Return an ordered working checklist for the selected ISO management-system standard."""
    standard = safe(standard).strip()
    if standard == "ISO 9001":
        clauses = {
            "4": ["4.1 Understanding the organization and its context",
                  "4.2 Understanding the needs and expectations of interested parties",
                  "4.3 Determining the scope of the quality management system",
                  "4.4 Quality management system and its processes"],
            "5": ["5.1 Leadership and commitment", "5.2 Quality policy",
                  "5.3 Organizational roles, responsibilities and authorities"],
            "6": ["6.1 Actions to address risks and opportunities",
                  "6.2 Quality objectives and planning to achieve them",
                  "6.3 Planning of changes"],
            "7": ["7.1 Resources", "7.2 Competence", "7.3 Awareness",
                  "7.4 Communication", "7.5 Documented information"],
            "8": ["8.1 Operational planning and control",
                  "8.2 Requirements for products and services",
                  "8.3 Design and development of products and services",
                  "8.4 Control of externally provided processes, products and services",
                  "8.5 Production and service provision", "8.6 Release of products and services",
                  "8.7 Control of nonconforming outputs"],
            "9": ["9.1 Monitoring, measurement, analysis and evaluation",
                  "9.2 Internal audit", "9.3 Management review"],
            "10": ["10.1 General", "10.2 Nonconformity and corrective action",
                   "10.3 Continual improvement"],
        }
    else:
        clauses = ISO_CLAUSES.get(standard, {})
    rows = []
    for clause_no, titles in clauses.items():
        for title in titles:
            # Keep the checklist in clause order and require an auditable observation/evidence entry.
            rows.append({
                "clause_sno": clause_no,
                "sub_clause": title,
                "checklist_pts": f"Verify implementation and retain objective evidence for: {title}",
                "procedure": standard,
            })
    return rows




def aud_ts():
    return datetime.now().isoformat(timespec="seconds")


def aud_parse_procedures(value):
    try:
        data = json.loads(safe(value) or "[]")
        return [str(x) for x in data] if isinstance(data, list) else []
    except Exception:
        return []


def aud_get(audit_id):
    return db.fetchone("SELECT * FROM audits WHERE id=?", (audit_id,))


def aud_list(statuses=None):
    """Return audits belonging to this workflow, including legacy rows submitted through it."""
    rows = db.fetchall("SELECT * FROM audits ORDER BY id DESC")
    allowed = set(statuses) if statuses else None
    result = []
    for row in rows:
        status = safe(row["status"]).strip()
        audit_type = safe(row["audit_type"]).strip()
        is_audit_v2 = int(row["aud_ver"] or 0) == 2
        # Older rows may have been submitted through the Audit Management screen
        # before the v2 marker was added. Include those in the relevant pickers.
        legacy_workflow_row = audit_type in AUD_TYPES and status in (
            "Draft", "Scheduled", "Submitted", "In Progress", "Completed"
        )
        if not (is_audit_v2 or legacy_workflow_row):
            continue
        if allowed is None or status in allowed:
            result.append(row)
    return result


# ------------------------------------------------------------
# Scheduling: Draft -> Scheduled (Save) -> Submitted -> In Progress -> Completed
# ------------------------------------------------------------
def aud_save_audit(audit_id, data, status):
    """Create or update the PARENT audit. status = Draft | Scheduled | Submitted."""
    audit_type = safe(data.get("audit_type")).strip()
    if audit_type not in AUD_TYPES:
        raise ValueError("Select an Audit Type first.")
    procedures = [safe(p) for p in (data.get("procedures") or [])] if aud_is_internal_type(audit_type) else []
    audit_date = safe(data.get("audit_date")).strip()
    location = safe(data.get("location")).strip()
    auditor = safe(data.get("auditor")).strip()
    if status in ("Scheduled", "Submitted"):
        missing = []
        if aud_is_internal_type(audit_type) and not procedures:
            missing.append("at least one Procedure")
        if not audit_date:
            missing.append("Date")
        if not location:
            missing.append("Location")
        if not auditor:
            missing.append("Auditor")
        if missing:
            raise ValueError("Please provide: " + ", ".join(missing) + ".")
    existing = aud_get(audit_id) if audit_id else None
    if existing and safe(existing["status"]) not in ("Draft", "Scheduled"):
        raise ValueError("This audit has already been submitted and can no longer be re-scheduled.")
    now = aud_ts()
    title = safe(data.get("title")).strip()
    scope = safe(data.get("scope")).strip()
    if existing:
        db.execute(
            "UPDATE audits SET audit_type=?,audit_date=?,location=?,auditor=?,lead_auditor=?,title=?,scope=?,"
            "procedures=?,status=?,updated_at=?,submitted_at=?,aud_ver=2 WHERE id=?",
            (audit_type, audit_date, location, auditor, auditor, title, scope, json.dumps(procedures),
             status, now, now if status == "Submitted" else "", existing["id"])
        )
        return int(existing["id"])
    number = document_number("AUD")
    db.execute(
        "INSERT INTO audits(number,audit_date,audit_type,project,location,auditor,lead_auditor,scope,status,"
        "created_at,title,created_by,updated_at,submitted_at,aud_ver,procedures) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (number, audit_date, audit_type, db.setting("project_name", ""), location, auditor, auditor, scope, status,
         now, title, "", now, now if status == "Submitted" else "", 2, json.dumps(procedures))
    )
    return int(db.fetchone("SELECT id FROM audits WHERE number=?", (number,))["id"])


def aud_set_auditee(audit_id, auditee, audit_type=None, procedures=None, standard=None):
    audit = aud_get(audit_id)
    if not audit:
        raise ValueError("Audit not found.")
    new_type = safe(audit_type).strip() if audit_type in AUD_TYPES else safe(audit["audit_type"])
    procs = json.dumps([safe(p) for p in procedures]) if procedures is not None else safe(audit["procedures"]) or "[]"
    selected_standard = safe(standard).strip() if standard is not None else safe(audit["standard"])
    db.execute("UPDATE audits SET auditee=?,audit_type=?,procedures=?,standard=?,updated_at=?,aud_ver=2 WHERE id=?",
               (safe(auditee).strip(), new_type, procs, selected_standard, aud_ts(), audit_id))


def aud_delete_audit(audit_id):
    """Deletes the audit with its findings and notifications. CAPA records already raised stay in the CAPA module."""
    db.execute("DELETE FROM aud_findings WHERE audit_id=?", (audit_id,))
    db.execute("DELETE FROM aud_notifications WHERE audit_id=?", (audit_id,))
    db.execute("DELETE FROM audits WHERE id=? AND COALESCE(aud_ver,0)=2", (audit_id,))


def aud_set_status(audit_id, status):
    db.execute("UPDATE audits SET status=?,updated_at=? WHERE id=?", (status, aud_ts(), audit_id))


# ------------------------------------------------------------
# Findings: Parent (Audit No.) / Child (Audit Finding No.)
# ------------------------------------------------------------
def aud_findings(audit_id):
    return db.fetchall("SELECT * FROM aud_findings WHERE audit_id=? ORDER BY seq,id", (audit_id,))


def aud_next_finding_no(audit_id):
    audit = aud_get(audit_id)
    row = db.fetchone("SELECT COALESCE(MAX(seq),0) AS n FROM aud_findings WHERE audit_id=?", (audit_id,))
    seq = int(row["n"]) + 1
    return seq, f"{safe(audit['number'])}-F{seq:02d}"


def aud_row_has_content(r):
    """Template (pre-loaded) rows are only persisted once assessed; manual rows once anything is typed."""
    if r.get("id"):
        return True
    evidence = safe(r.get("evidence")).strip()
    action = safe(r.get("corrective_action")).strip()
    # Checklist template text is not itself a finding. Count/persist only an
    # actual observation entered by the auditor (evidence/finding) or action.
    return bool(evidence or action or safe(r.get("capa_required")) == "Yes")


def aud_save_findings(audit_id, rows, removed_ids=()):
    """Persist checklist rows and apply the CAPA Yes/No routing. Returns a summary dict."""
    audit = aud_get(audit_id)
    if not audit:
        raise ValueError("Audit not found.")
    todo = [(i, r) for i, r in enumerate(rows) if aud_row_has_content(r)]
    errors = []
    for i, r in todo:
        if safe(r.get("capa_required")) == "Yes":
            missing = [label for key, label in (("corrective_action", "Corrective Action"), ("responsible", "Responsible"))
                       if not safe(r.get(key)).strip()]
            if missing:
                errors.append(f"Row {i + 1}: CAPA = Yes requires " + " and ".join(missing) + ".")
    if errors:
        raise ValueError("\n".join(errors))
    summary = {"saved": [], "capa_created": [], "register_only": 0, "notes": []}
    for fid in removed_ids:
        f = db.fetchone("SELECT capa_number FROM aud_findings WHERE id=? AND audit_id=?", (fid, audit_id))
        if f and not safe(f["capa_number"]):
            db.execute("DELETE FROM aud_findings WHERE id=?", (fid,))
    now = aud_ts()
    for i, r in todo:
        capa_flag = "Yes" if safe(r.get("capa_required")) == "Yes" else "No"
        status = safe(r.get("status")) if safe(r.get("status")) in AUD_FINDING_STATUSES else "Open"
        existing = db.fetchone("SELECT * FROM aud_findings WHERE id=? AND audit_id=?", (r["id"], audit_id)) if r.get("id") else None
        if existing and safe(existing["capa_number"]) and capa_flag == "No":
            capa_flag = "Yes"
            summary["notes"].append(f"{safe(existing['finding_no'])}: a CAPA ({safe(existing['capa_number'])}) already exists, so CAPA stays 'Yes'.")
        values = (safe(r.get("auditor_name")).strip(), safe(r.get("clause_sno")).strip(), safe(r.get("sub_clause")).strip(),
                  safe(r.get("checklist_pts")).strip(), safe(r.get("evidence")).strip(), safe(r.get("corrective_action")).strip(),
                  safe(r.get("responsible")).strip(), safe(r.get("target_date")).strip(), capa_flag, status,
                  safe(r.get("procedure")).strip())
        if existing:
            db.execute(
                "UPDATE aud_findings SET auditor_name=?,clause_sno=?,sub_clause=?,checklist_pts=?,evidence=?,corrective_action=?,"
                "responsible=?,target_date=?,capa_required=?,status=?,procedure=?,updated_at=? WHERE id=?",
                values + (now, existing["id"]))
            fid = int(existing["id"])
        else:
            seq, finding_no = aud_next_finding_no(audit_id)
            cur = db.execute(
                "INSERT INTO aud_findings(audit_id,seq,finding_no,auditor_name,clause_sno,sub_clause,checklist_pts,evidence,"
                "corrective_action,responsible,target_date,capa_required,status,procedure,created_at,updated_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (audit_id, seq, finding_no) + values + (now, now))
            fid = int(cur.lastrowid)
        routed = aud_route_finding(fid)
        if routed.get("capa_created"):
            summary["capa_created"].append(routed["capa_created"])
        if capa_flag == "No":
            summary["register_only"] += 1
        summary["saved"].append((i, fid))
    if summary["saved"] and safe(audit["status"]) in ("Submitted", "Scheduled"):
        aud_set_status(audit_id, "In Progress")
    return summary


def aud_route_finding(fid):
    """
    CAPA ROUTING (strictly on the CAPA Yes/No field)
      CAPA = Yes -> raise/keep a record in the CAPA module FIRST; stage 'CAPA Module'.
                    When that CAPA is Closed, aud_sync_capa_stages() moves it to 'Audit Report / Closure'.
      CAPA = No  -> bypass the CAPA module entirely; stage 'Audit Register' only.
    """
    f = db.fetchone("SELECT * FROM aud_findings WHERE id=?", (fid,))
    audit = aud_get(f["audit_id"])
    result = {}
    if safe(f["capa_required"]) == "Yes":
        if not safe(f["capa_number"]):
            capa_no = next_number("HSE-CAPA", "capa")
            finding_text = f"{safe(audit['number'])} / {safe(f['finding_no'])}: {safe(f['checklist_pts'])}"
            if safe(f["evidence"]):
                finding_text += f" | Evidence: {safe(f['evidence'])}"
            db.execute(
                "INSERT INTO capa(number,source,reference_number,finding,root_cause,corrective_action,preventive_action,"
                "responsible,priority,target_date,status,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                (capa_no, "Audit", safe(f["finding_no"]), finding_text, "", safe(f["corrective_action"]), "",
                 safe(f["responsible"]), "Medium", safe(f["target_date"]), "Open", aud_ts()))
            db.execute("UPDATE aud_findings SET capa_number=?,route_stage=?,status=CASE WHEN status='Closed' THEN 'In Progress' ELSE status END WHERE id=?",
                       (capa_no, AUD_STAGE_CAPA, fid))
            result["capa_created"] = capa_no
        else:
            aud_sync_capa_stages(None, fid)
    else:
        db.execute("UPDATE aud_findings SET route_stage=? WHERE id=?", (AUD_STAGE_REGISTER, fid))
    return result


def aud_sync_capa_stages(audit_id=None, finding_id=None):
    """Move CAPA=Yes findings on to 'Audit Report / Closure' once their CAPA is Closed in the CAPA module."""
    sql = "SELECT * FROM aud_findings WHERE capa_required='Yes' AND COALESCE(capa_number,'')<>''"
    params = []
    if audit_id:
        sql += " AND audit_id=?"
        params.append(audit_id)
    if finding_id:
        sql += " AND id=?"
        params.append(finding_id)
    for f in db.fetchall(sql, tuple(params)):
        capa = db.fetchone("SELECT status FROM capa WHERE number=?", (f["capa_number"],))
        if capa is None:
            # CAPA record no longer exists: clear the link so the next checklist save re-raises it.
            db.execute("UPDATE aud_findings SET capa_number='',route_stage=? WHERE id=?", (AUD_STAGE_CAPA, f["id"]))
        elif safe(capa["status"]) == "Closed":
            if safe(f["route_stage"]) != AUD_STAGE_CLOSURE or safe(f["status"]) != "Closed":
                db.execute("UPDATE aud_findings SET route_stage=?,status='Closed',updated_at=? WHERE id=?",
                           (AUD_STAGE_CLOSURE, aud_ts(), f["id"]))
        else:
            new_status = "In Progress" if safe(f["status"]) == "Closed" else safe(f["status"])
            if safe(f["route_stage"]) != AUD_STAGE_CAPA or new_status != safe(f["status"]):
                db.execute("UPDATE aud_findings SET route_stage=?,status=?,updated_at=? WHERE id=?",
                           (AUD_STAGE_CAPA, new_status, aud_ts(), f["id"]))


def aud_metrics():
    marks = ",".join("?" for _ in AUD_CONDUCTED_STATUSES)
    conducted = db.fetchone(f"SELECT COUNT(*) AS n FROM audits WHERE COALESCE(aud_ver,0)=2 AND status IN ({marks})",
                            AUD_CONDUCTED_STATUSES)["n"]
    base = "FROM aud_findings f JOIN audits a ON a.id=f.audit_id WHERE COALESCE(a.aud_ver,0)=2"
    substantive = " AND (TRIM(COALESCE(f.evidence,''))<>'' OR TRIM(COALESCE(f.corrective_action,''))<>'')"
    findings = db.fetchone(f"SELECT COUNT(*) AS n {base}{substantive}")["n"]
    open_f = db.fetchone(f"SELECT COUNT(*) AS n {base} AND f.status<>'Closed'{substantive}")["n"]
    capa_pending = db.fetchone(f"SELECT COUNT(*) AS n {base} AND f.route_stage=?{substantive}", (AUD_STAGE_CAPA,))["n"]
    return {"conducted": conducted, "findings": findings, "open": open_f, "capa": capa_pending}



def aud_trained_people(role="auditor"):
    """Return employees with a current Internal Audit / Auditor training record (mirrors PTW training gate)."""
    # Course names that qualify a person to conduct or receive internal audits.
    like_terms = (
        "%internal audit%", "%auditor%", "%lead auditor%", "%iso audit%",
        "%audit training%", "%hse audit%", "%management system audit%"
    )
    like_sql = " OR ".join(["lower(COALESCE(c.course_name,'')) LIKE ?" for _ in like_terms])
    rows = db.fetchall(
        f"""SELECT DISTINCT e.id AS employee_db_id, e.employee_id, e.name, e.designation, e.company,
                  c.course_name, tr.training_date, tr.expiry_date
           FROM employees e
           JOIN training_records tr ON tr.employee_db_id=e.id
           JOIN training_courses c ON c.id=tr.course_id
           WHERE COALESCE(e.active,1)=1
             AND ({like_sql})
             AND (COALESCE(trim(tr.expiry_date),'')='' OR trim(tr.expiry_date)>=?)
           ORDER BY e.name COLLATE NOCASE""",
        (*like_terms, today())
    )
    return rows


def aud_is_trained_auditor(name_or_id):
    """True if the given name or employee_id matches a currently trained auditor."""
    needle = safe(name_or_id).strip().casefold()
    if not needle:
        return False
    for r in aud_trained_people("auditor"):
        if needle in (safe(r["name"]).casefold(), safe(r["employee_id"]).casefold()):
            return True
    return False


def aud_days_remaining(target_date, status="Open"):
    """Return (label, overdue_bool) for close-out tracking."""
    if not safe(target_date).strip():
        return "—", False
    st = safe(status)
    if st in ("Closed", "Cancelled", "Approved"):
        return "Closed", False
    try:
        target = date.fromisoformat(safe(target_date)[:10])
    except Exception:
        return "—", False
    delta = (target - date.today()).days
    if delta < 0:
        return f"Overdue by {abs(delta)} day(s)", True
    if delta == 0:
        return "Due today", False
    return f"{delta} day(s) remaining", False


def aud_tracking_rows(include_closed=False):
    """Open (or all) audit findings for the Tracking Register."""
    sql = """
        SELECT f.*, a.number AS audit_number, a.audit_type, a.audit_date, a.location, a.auditor AS audit_auditor
        FROM aud_findings f
        JOIN audits a ON a.id=f.audit_id
        WHERE COALESCE(a.aud_ver,0)=2
          AND (TRIM(COALESCE(f.evidence,''))<>'' OR TRIM(COALESCE(f.corrective_action,''))<>''
               OR TRIM(COALESCE(f.observation_text,''))<>'' OR TRIM(COALESCE(f.checklist_pts,''))<>'')
    """
    if not include_closed:
        sql += " AND COALESCE(f.closeout_status,'Open') NOT IN ('Closed','Approved') AND COALESCE(f.status,'Open')<>'Closed' "
    sql += " ORDER BY a.audit_date DESC, f.seq, f.id"
    return db.fetchall(sql)


def aud_finding_attachments(finding_id):
    return db.fetchall("SELECT * FROM aud_finding_attachments WHERE finding_id=? ORDER BY id", (finding_id,))


def aud_add_finding_attachment(finding_id, source_path, uploaded_by="", attachment_type="Close-out Evidence"):
    source = Path(source_path)
    if not source.exists():
        raise ValueError("Attachment file not found.")
    finding = db.fetchone("SELECT finding_no FROM aud_findings WHERE id=?", (finding_id,))
    prefix = safe(finding["finding_no"]) if finding else f"FINDING-{finding_id}"
    dest = ATTACH_DIR / f"{prefix}_{source.name}"
    counter = 1
    while dest.exists():
        dest = ATTACH_DIR / f"{prefix}_{counter}_{source.name}"
        counter += 1
    shutil.copy2(source, dest)
    db.execute(
        "INSERT INTO aud_finding_attachments(finding_id,file_path,attachment_type,uploaded_by,uploaded_at) VALUES(?,?,?,?,?)",
        (finding_id, str(dest), attachment_type, safe(uploaded_by), aud_ts())
    )
    return str(dest)


def aud_submit_closeout(finding_id, comments, evidence_paths, submitted_by):
    """Responsible person submits close-out for senior review."""
    f = db.fetchone("SELECT * FROM aud_findings WHERE id=?", (finding_id,))
    if not f:
        raise ValueError("Finding not found.")
    if safe(f["status"]) == "Closed" or safe(f["closeout_status"]) in ("Closed", "Approved"):
        raise ValueError("This finding is already closed.")
    saved = []
    for p in evidence_paths or []:
        if safe(p).strip():
            saved.append(aud_add_finding_attachment(finding_id, p, submitted_by, "Close-out Evidence"))
    evidence_text = "; ".join(Path(x).name for x in saved)
    existing_ev = safe(f["closeout_evidence"]) if "closeout_evidence" in f.keys() else ""
    if existing_ev and evidence_text:
        evidence_text = existing_ev + "; " + evidence_text
    elif existing_ev:
        evidence_text = existing_ev
    db.execute(
        """UPDATE aud_findings SET closeout_status=?, closeout_comments=?, closeout_evidence=?,
           closeout_submitted_by=?, closeout_submitted_at=?, status=?, updated_at=? WHERE id=?""",
        ("Submitted for Review", safe(comments), evidence_text, safe(submitted_by), aud_ts(),
         "In Progress", aud_ts(), finding_id)
    )
    return True


def aud_approve_closeout(finding_id, approved_by, comments=""):
    """Senior approval closes the finding and syncs linked CAPA / observation if present."""
    f = db.fetchone("SELECT * FROM aud_findings WHERE id=?", (finding_id,))
    if not f:
        raise ValueError("Finding not found.")
    if safe(f["closeout_status"]) not in ("Submitted for Review", "Pending Approval"):
        raise ValueError("Only findings submitted for review can be approved.")
    note = safe(comments).strip()
    merged = safe(f["closeout_comments"])
    if note:
        merged = (merged + "\n" if merged else "") + f"Approval: {note}"
    db.execute(
        """UPDATE aud_findings SET closeout_status=?, status=?, closeout_approved_by=?, closeout_approved_at=?,
           closeout_comments=?, updated_at=? WHERE id=?""",
        ("Closed", "Closed", safe(approved_by), aud_ts(), merged, aud_ts(), finding_id)
    )
    # Sync linked CAPA (same pattern as observation CAPA completion).
    capa_no = safe(f["capa_number"])
    if capa_no:
        capa = db.fetchone("SELECT * FROM capa WHERE number=?", (capa_no,))
        if capa:
            db.execute(
                """UPDATE capa SET status=?, approval_status=?, approved_by=?, approved_date=?,
                   verification=?, action_status=? WHERE id=?""",
                ("Closed", "Approved", safe(approved_by), today(),
                 f"Closed via Audit finding {safe(f['finding_no'])} close-out approval.",
                 "CAPA Action Complete", capa["id"])
            )
            try:
                sync_observation_from_capa(capa)
            except Exception:
                logging.exception("Observation sync from CAPA after audit close-out failed")
    return True


def aud_reject_closeout(finding_id, reviewer, comments=""):
    f = db.fetchone("SELECT * FROM aud_findings WHERE id=?", (finding_id,))
    if not f:
        raise ValueError("Finding not found.")
    note = safe(comments).strip() or "Returned for further close-out evidence."
    merged = safe(f["closeout_comments"])
    merged = (merged + "\n" if merged else "") + f"Returned by {safe(reviewer)}: {note}"
    db.execute(
        """UPDATE aud_findings SET closeout_status=?, status=?, closeout_comments=?, updated_at=? WHERE id=?""",
        ("Open", "Open", merged, aud_ts(), finding_id)
    )
    return True


def aud_summary_rows():
    """Per-audit open-observation percentage for the Summary Register."""
    audits = db.fetchall(
        """SELECT a.id, a.number, a.audit_type, a.audit_date, a.location, a.auditor, a.auditee, a.status
           FROM audits a WHERE COALESCE(a.aud_ver,0)=2 ORDER BY a.audit_date DESC, a.id DESC"""
    )
    out = []
    for a in audits:
        findings = db.fetchall(
            """SELECT id, status, closeout_status FROM aud_findings
               WHERE audit_id=? AND (TRIM(COALESCE(evidence,''))<>'' OR TRIM(COALESCE(corrective_action,''))<>''
                     OR TRIM(COALESCE(observation_text,''))<>'')""",
            (a["id"],)
        )
        total = len(findings)
        open_n = sum(1 for f in findings if safe(f["status"]) != "Closed" and safe(f["closeout_status"]) not in ("Closed", "Approved"))
        closed_n = total - open_n
        pct_open = round(100.0 * open_n / total, 1) if total else 0.0
        pct_closed = round(100.0 * closed_n / total, 1) if total else 100.0 if total == 0 else 0.0
        out.append({
            "id": a["id"], "number": a["number"], "audit_type": a["audit_type"], "audit_date": a["audit_date"],
            "location": a["location"], "auditor": a["auditor"], "auditee": a["auditee"], "status": a["status"],
            "total": total, "open_n": open_n, "closed_n": closed_n, "pct_open": pct_open, "pct_closed": pct_closed
        })
    return out


def aud_cumulative_findings():
    """All audit findings across audits for the Cumulative Register."""
    return db.fetchall(
        """SELECT f.*, a.number AS audit_number, a.audit_type, a.audit_date, a.location, a.auditor AS audit_auditor, a.auditee
           FROM aud_findings f
           JOIN audits a ON a.id=f.audit_id
           WHERE COALESCE(a.aud_ver,0)=2
             AND (TRIM(COALESCE(f.evidence,''))<>'' OR TRIM(COALESCE(f.corrective_action,''))<>''
                  OR TRIM(COALESCE(f.observation_text,''))<>'' OR TRIM(COALESCE(f.checklist_pts,''))<>'')
           ORDER BY a.audit_date DESC, f.seq, f.id"""
    )



def aud_register_table(audit_id):
    """Complete Audit Register rows (parent number repeated on every child row)."""
    aud_sync_capa_stages(audit_id)
    audit = aud_get(audit_id)
    rows = []
    for f in aud_findings(audit_id):
        if not (safe(f["evidence"]).strip() or safe(f["corrective_action"]).strip()):
            continue
        rows.append([safe(audit["number"]), safe(f["finding_no"]), safe(f["auditor_name"]), safe(f["clause_sno"]),
                     safe(f["sub_clause"]), safe(f["checklist_pts"]), safe(f["evidence"]), safe(f["corrective_action"]),
                     safe(f["responsible"]), safe(f["target_date"]), safe(f["capa_required"]), safe(f["status"]),
                     safe(f["capa_number"]), safe(f["route_stage"])])
    return AUD_REGISTER_HEADERS, rows


# ------------------------------------------------------------
# Notifications
# ------------------------------------------------------------
def aud_next_notification_no():
    prefix = f"{document_prefix()}-ANOT-{datetime.now().year}-"
    highest = 0
    for r in db.fetchall("SELECT doc_number FROM aud_notifications WHERE doc_number LIKE ?", (prefix + "%",)):
        try:
            highest = max(highest, int(safe(r["doc_number"]).split("-")[-1]))
        except Exception:
            pass
    return f"{prefix}{highest + 1:04d}"


def aud_get_notification(audit_id):
    return db.fetchone("SELECT * FROM aud_notifications WHERE audit_id=? ORDER BY id DESC LIMIT 1", (audit_id,))


def aud_save_notification(audit_id, data):
    audit = aud_get(audit_id)
    if not audit or safe(audit["status"]) not in AUD_SCHEDULED_STATUSES:
        raise ValueError("Notifications can only be created for scheduled audits.")
    missing = [label for key, label in (("to_party", "To"), ("from_party", "From"), ("subject", "Subject"), ("scope", "Scope"))
               if not safe(data.get(key)).strip()]
    if missing:
        raise ValueError("Please complete: " + ", ".join(missing) + ".")
    now = aud_ts()
    cols = ("to_party", "from_party", "subject", "scope", "audit_date", "audit_time", "venue", "auditee",
            "audit_team", "reference_docs", "remarks")
    values = tuple(safe(data.get(c)).strip() for c in cols)
    existing = aud_get_notification(audit_id)
    if existing:
        db.execute("UPDATE aud_notifications SET " + ",".join(f"{c}=?" for c in cols) + ",updated_at=? WHERE id=?",
                   values + (now, existing["id"]))
        return int(existing["id"])
    cur = db.execute(
        "INSERT INTO aud_notifications(audit_id,doc_number," + ",".join(cols) + ",created_at,updated_at) VALUES(?,?," +
        ",".join("?" for _ in cols) + ",?,?)",
        (audit_id, aud_next_notification_no()) + values + (now, now))
    return int(cur.lastrowid)


# ------------------------------------------------------------
# Reports: Excel / Word (Audit Register) and Word (Audit Notification invite)
# ------------------------------------------------------------
def aud_export_excel(audit_id, path):
    audit = aud_get(audit_id)
    headers, rows = aud_register_table(audit_id)
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Register"
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    navy = PatternFill("solid", fgColor="17365D")
    light = PatternFill("solid", fgColor="DCE6F1")
    ncols = len(headers)
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=ncols)
    ws["B1"] = f"{company_name()} — AUDIT REGISTER"
    ws["B1"].font = Font(bold=True, size=16, color="17365D")
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 48
    logo = report_logo_path()
    if logo:
        try:
            img = XLImage(str(logo))
            w0, h0 = img.width, img.height
            img.height = 46
            img.width = int(w0 * 46 / max(h0, 1))
            ws.add_image(img, "A1")
        except Exception:
            logging.exception("Audit Excel logo embedding failed")
    info = [("Audit No.", audit["number"]), ("Audit Type", audit["audit_type"]), ("Audit Date", audit["audit_date"]),
            ("Location", audit["location"]), ("Auditor", audit["auditor"]), ("Auditee", audit["auditee"]),
            ("Status", audit["status"]), ("Procedures", ", ".join(aud_parse_procedures(audit["procedures"]))),
            ("ISO Standard", safe(audit["standard"]) or "—")]
    r = 3
    for i in range(0, len(info), 2):
        for j in range(2):
            if i + j < len(info):
                label, value = info[i + j]
                lc = ws.cell(row=r, column=1 + j * 4, value=label)
                lc.font = Font(bold=True)
                lc.fill = light
                lc.border = border
                ws.merge_cells(start_row=r, start_column=2 + j * 4, end_row=r, end_column=4 + j * 4)
                vc = ws.cell(row=r, column=2 + j * 4, value=xml_safe(value))
                vc.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    header_row = r + 1
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = navy
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 30
    for ri, row in enumerate(rows, header_row + 1):
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=c, value=xml_safe(val))
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    widths = [20, 24, 16, 10, 16, 40, 28, 32, 16, 12, 9, 12, 18, 22]
    from openpyxl.utils import get_column_letter
    for c, wd in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = wd
    total = len(rows)
    closed = sum(1 for x in rows if x[11] == "Closed")
    ws.cell(row=header_row + total + 2, column=1, value=f"Total findings: {total}   |   Open: {total - closed}   |   Closed: {closed}").font = Font(bold=True)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=3)
    if total:
        ws.auto_filter.ref = f"A{header_row}:{ws.cell(row=header_row + total, column=ncols).coordinate}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 0
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def aud_export_word(audit_id, path):
    from docx.shared import RGBColor
    audit = aud_get(audit_id)
    if not audit:
        raise ValueError("Audit record not found.")
    headers, rows = aud_register_table(audit_id)
    document = Document()
    set_docx_landscape(document)
    add_docx_report_header(document, "AUDIT REGISTER", "AUD", safe(audit["number"]))
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"AUDIT REGISTER — {xml_safe(audit['number'])}")
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor(0x17, 0x36, 0x5D)
    info = [("Audit No.", audit["number"]), ("Audit Type", audit["audit_type"]), ("Audit Date", audit["audit_date"]),
            ("Location", audit["location"]), ("Auditor", audit["auditor"]), ("Auditee", audit["auditee"]),
            ("Status", audit["status"]), ("Procedures", ", ".join(aud_parse_procedures(audit["procedures"])) or "—"),
            ("ISO Standard", safe(audit["standard"]) or "—"), ("Scope", safe(audit["scope"]) or "—")]
    # Pad to even count so the 2-column layout never indexes past the end.
    if len(info) % 2:
        info.append(("", ""))
    t = document.add_table(rows=0, cols=4)
    t.style = "Table Grid"
    t.autofit = False
    for i in range(0, len(info), 2):
        cells = t.add_row().cells
        for j in range(2):
            if i + j >= len(info):
                break
            label, value = info[i + j]
            if not label and not value:
                continue
            lc, vc = cells[j * 2], cells[j * 2 + 1]
            lc.text = xml_safe(label)
            vc.text = xml_safe(value) or "—"
            add_docx_cell_shading(lc, "DCE6F1")
            for run in lc.paragraphs[0].runs:
                run.bold = True
            for cell, wd in ((lc, 1.3), (vc, 3.8)):
                cell.width = Inches(wd)
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(8)
    document.add_paragraph()
    table = document.add_table(rows=1, cols=len(headers))
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h
    if not rows:
        cells = table.add_row().cells
        cells[0].text = "No audit findings have been recorded for this audit."
    for row in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            if i < len(cells):
                cells[i].text = xml_safe(val)
    style_docx_table(table, header=True, font_size=7)
    # Safe column widths — never index past available cells/headers.
    widths = [0.95, 1.1, 0.7, 0.45, 0.65, 1.55, 1.0, 1.15, 0.7, 0.65, 0.4, 0.5, 0.75, 0.9]
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths):
                cell.width = Inches(widths[i])
    total = len(rows)
    closed = sum(1 for x in rows if len(x) > 11 and x[11] == "Closed")
    capa_yes = sum(1 for x in rows if len(x) > 10 and x[10] == "Yes")
    sp = document.add_paragraph()
    sr = sp.add_run(f"Total findings: {total}   |   Open: {total - closed}   |   Closed: {closed}   |   "
                    f"CAPA routed: {capa_yes}")
    sr.bold = True
    sr.font.size = Pt(9)

    # Embed actual evidence images (not only file names) after the findings table.
    document.add_heading("Evidence / Attachments (embedded)", level=2)
    findings = aud_findings(audit_id)
    image_ext = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
    embedded_any = False
    for f in findings:
        # Paths may appear in the evidence text and/or finding attachment table.
        paths = []
        ev = safe(f["evidence"]) if "evidence" in f.keys() else ""
        for part in re.split(r"[;\n]+", ev):
            part = part.strip()
            if part and Path(part).exists() and Path(part).suffix.lower() in image_ext:
                paths.append(part)
        try:
            for att in aud_finding_attachments(int(f["id"])):
                p = Path(safe(att["file_path"]))
                if p.exists() and p.suffix.lower() in image_ext:
                    paths.append(str(p))
        except Exception:
            pass
        # Also audit-level attachments linked to this finding if any.
        try:
            for att in db.fetchall(
                "SELECT file_path FROM audit_attachments WHERE audit_id=? AND (finding_id=? OR finding_id IS NULL)",
                (audit_id, f["id"])
            ):
                p = Path(safe(att["file_path"]))
                if p.exists() and p.suffix.lower() in image_ext:
                    paths.append(str(p))
        except Exception:
            pass
        # Unique preserve order
        seen = set()
        uniq = []
        for p in paths:
            key = str(Path(p).resolve()) if Path(p).exists() else p
            if key not in seen:
                seen.add(key)
                uniq.append(p)
        if not uniq:
            continue
        embedded_any = True
        document.add_paragraph(
            f"Finding {xml_safe(f['finding_no'])}: {xml_safe(safe(f['observation_text']) or safe(f['checklist_pts']) or '')[:120]}"
        )
        for p in uniq:
            try:
                pp = document.add_paragraph()
                pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pp.add_run().add_picture(str(p), width=Inches(3.2))
                cap = document.add_paragraph(Path(p).name)
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                logging.exception("Could not embed audit evidence image %s", p)
                document.add_paragraph(f"(Could not embed image: {Path(p).name})")
    if not embedded_any:
        document.add_paragraph("No image evidence was available to embed for this audit.")

    add_docx_signatures(document)
    add_docx_footer(document)
    save_docx_validated(document, path)



def aud_export_pdf(audit_id, path):
    """Generate a detailed, landscape PDF Audit Register with audit metadata and every saved finding."""
    audit = aud_get(audit_id)
    if not audit:
        raise ValueError("Audit record not found.")
    headers, rows = aud_register_table(audit_id)
    styles = getSampleStyleSheet()
    story = [
        Paragraph(f"<b>{xml_safe(company_name())}</b>", styles["Title"]),
        Paragraph(f"<b>DETAILED AUDIT REGISTER — {xml_safe(audit['number'])}</b>", styles["Heading2"]),
        Spacer(1, 8),
    ]
    metadata = [
        ["Audit Number", safe(audit["number"]), "Audit Type", safe(audit["audit_type"])],
        ["Audit Date", safe(audit["audit_date"]), "Status", safe(audit["status"])],
        ["Location", safe(audit["location"]), "Auditor", safe(audit["auditor"])],
        ["Auditee", safe(audit["auditee"]), "ISO Standard", safe(audit["standard"]) or "—"],
        ["Procedures", ", ".join(aud_parse_procedures(audit["procedures"])) or "—", "Scope", safe(audit["scope"]) or "—"],
    ]
    meta_table = Table(metadata, colWidths=[78, 175, 78, 175])
    meta_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#DCE6F1")),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#DCE6F1")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.extend([meta_table, Spacer(1, 10)])
    closed = sum(1 for row in rows if row[11] == "Closed")
    story.append(Paragraph(
        f"Total checklist findings: {len(rows)} &nbsp;&nbsp; Open / other: {len(rows)-closed} "
        f"&nbsp;&nbsp; Closed: {closed} &nbsp;&nbsp; CAPA linked: {sum(1 for row in rows if row[10]=='Yes')}",
        styles["Normal"]))
    story.append(Spacer(1, 6))
    wrapped_headers = [Paragraph(f"<b>{xml_safe(h)}</b>", styles["BodyText"]) for h in headers]
    data = [wrapped_headers]
    for row in rows:
        data.append([Paragraph(xml_safe(value).replace("\\n", "<br/>"), styles["BodyText"]) for value in row])
    if len(data) == 1:
        data.append([Paragraph("No checklist findings have been saved.", styles["BodyText"])] + [""] * (len(headers)-1))
    widths = [72, 82, 55, 40, 58, 125, 95, 105, 58, 48, 42, 48, 70, 72]
    table = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17365D")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.extend([table, Spacer(1, 16)])
    signature = Table([
        ["Prepared / Lead Auditor", "Reviewed By", "Approved By"],
        ["Name: ______________________\\nSignature: ___________________\\nDate: ________________________"] * 3
    ], colWidths=[190, 190, 190])
    signature.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCE6F1")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
    ]))
    story.append(signature)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(path), pagesize=landscape(A4), rightMargin=22, leftMargin=22, topMargin=24, bottomMargin=24)
    doc.build(story)


def aud_export_notification_word(audit_id, path):
    """Professional 'Invites Document': logo, 'Audit Notification' heading, document number, details table."""
    from docx.shared import RGBColor
    audit = aud_get(audit_id)
    note = aud_get_notification(audit_id)
    if not audit or not note:
        raise ValueError("Save the notification before exporting.")
    document = Document()
    section = document.sections[0]
    section.left_margin = section.right_margin = Inches(0.9)
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    lp = document.add_paragraph()
    lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo = report_logo_path()
    if logo:
        try:
            lp.add_run().add_picture(str(logo), width=Inches(1.7))
        except Exception:
            logging.exception("Notification logo failed")
            lp.add_run(xml_safe(company_name())).bold = True
    else:
        cr = lp.add_run(xml_safe(company_name()))
        cr.bold = True
        cr.font.size = Pt(16)
    if logo:
        cp = document.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c = cp.add_run(xml_safe(company_name()))
        c.bold = True
        c.font.size = Pt(11)
    tp = document.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run("Audit Notification")
    tr.bold = True
    tr.font.size = Pt(22)
    tr.font.color.rgb = RGBColor(0x17, 0x36, 0x5D)
    dp = document.add_paragraph()
    dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    dr = dp.add_run(f"Document Number: {xml_safe(note['doc_number'])}")
    dr.bold = True
    dr.font.size = Pt(11)
    document.add_paragraph()
    audit_date = safe(note["audit_date"]) or safe(audit["audit_date"])
    pairs = [
        ("Document Number", note["doc_number"]), ("Audit Number", audit["number"]), ("Audit Type", audit["audit_type"]),
        ("To", note["to_party"]), ("From", note["from_party"]), ("Subject", note["subject"]), ("Scope", note["scope"]),
        ("Audit Date", audit_date), ("Time", note["audit_time"]), ("Venue / Location", note["venue"] or audit["location"]),
        ("Auditee", note["auditee"] or audit["auditee"]), ("Audit Team / Auditor", note["audit_team"] or audit["auditor"]),
        ("Procedures / Documents for Review", note["reference_docs"]), ("Remarks", note["remarks"]),
        ("Date of Issue", safe(note["updated_at"])[:10]),
    ]
    table = document.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for label, value in pairs:
        if label in ("Remarks", "Procedures / Documents for Review", "Time") and not safe(value).strip():
            continue
        cells = table.add_row().cells
        cells[0].text = xml_safe(label)
        cells[1].text = xml_safe(value) or "—"
        add_docx_cell_shading(cells[0], "17365D")
        for run in cells[0].paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
        for cell, wd in ((cells[0], 2.1), (cells[1], 4.6)):
            cell.width = Inches(wd)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for para in cell.paragraphs:
                para.paragraph_format.space_before = Pt(3)
                para.paragraph_format.space_after = Pt(3)
                for run in para.runs:
                    run.font.size = Pt(10)
    add_docx_footer(document)
    save_docx_validated(document, path)


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
    # AUDIT MANAGEMENT MODULE (v2) - desktop layout
    # Tabs: Dashboard | Schedule | Notifications | Audit | Master Register |
    #       Findings Register (tracking+cumulative) | Summary | Close Out
    # ========================================================
    def audits(self):
        w, layout = self.page("Audit Management")
        try:
            tabs = QTabWidget()
            tabs.setTabsClosable(False)
            tabs.setMovable(False)
            layout.addWidget(tabs, 1)
            ctx = {"tabs": tabs}
            self._last_audit_ctx = ctx
            dash_tab = self._aud_build_dashboard(ctx)
            sched_tab = self._aud_build_schedule(ctx)
            note_tab = self._aud_build_notifications(ctx)
            check_tab = self._aud_build_checklist(ctx)
            saved_tab = self._aud_build_saved_for_later(ctx)
            register_tab = self._aud_build_master_register(ctx)
            findings_tab = self._aud_build_findings_register(ctx)
            summary_tab = self._aud_build_summary_register(ctx)
            closeout_tab = self._aud_build_closeout_section(ctx)
            # Order: Dashboard → Schedule → Notifications → Audit → Saved for later → Master → Findings → Summary → Close Out
            tabs.addTab(dash_tab, "Audit Dashboard")
            tabs.addTab(sched_tab, "Schedule Audit")
            tabs.addTab(note_tab, "Notifications")
            tabs.addTab(check_tab, "Audit")
            tabs.addTab(saved_tab, "Saved for later")
            tabs.addTab(register_tab, "Master Audit Register")
            tabs.addTab(findings_tab, "Audit Findings Register")
            tabs.addTab(summary_tab, "Audit Summary Register")
            tabs.addTab(closeout_tab, "Close Out")
            refreshers = {
                0: ctx.get("refresh_dashboard", lambda: None),
                1: ctx.get("refresh_schedule", lambda: None),
                2: ctx.get("refresh_notification_options", lambda: None),
                3: ctx.get("refresh_checklist_options", lambda: None),
                4: ctx.get("refresh_saved_for_later", lambda: None),
                5: ctx.get("refresh_master_register", lambda: None),
                6: ctx.get("refresh_findings_register", lambda: None),
                7: ctx.get("refresh_summary", lambda: None),
                8: ctx.get("refresh_closeout_section", lambda: None),
            }
            tabs.currentChanged.connect(lambda index: refreshers.get(index, lambda: None)())
            ctx["refresh_dashboard"]()
        except Exception as exc:
            logging.exception("Audit Management failed to load")
            err = QLabel(
                f"<b>Audit Management could not be loaded.</b><br><br>"
                f"{type(exc).__name__}: {exc}<br><br>"
                f"Please share this message so it can be fixed."
            )
            err.setWordWrap(True)
            err.setStyleSheet("color:#B00020;padding:12px;")
            layout.addWidget(err)
            QMessageBox.critical(self, "Audit Management", f"Audit section failed to load:\n\n{type(exc).__name__}: {exc}")

    def _aud_build_master_register(self, ctx):
        tab = QWidget()
        lay = QVBoxLayout(tab)
        search = QLineEdit(); search.setPlaceholderText("Search audit number, type, date, location, auditor, or status")
        lay.addWidget(search)
        table = self._aud_table(["Audit Number", "Audit Type", "Date", "Location", "Auditor", "Scope", "Status"])
        self._aud_readonly(table)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        lay.addWidget(table, 1)
        def load():
            needle = search.text().strip().casefold()
            records = [a for a in aud_list() if not needle or needle in " ".join(safe(a[k]) for k in ("number", "audit_type", "audit_date", "location", "auditor", "scope", "status")).casefold()]
            table.setRowCount(len(records))
            for r, a in enumerate(records):
                for c, key in enumerate(("number", "audit_type", "audit_date", "location", "auditor", "scope", "status")):
                    table.setItem(r, c, self._aud_item(a[key], a["id"], link=(c == 0)))
        def open_record(row, col=0):
            item = table.item(row, 0)
            if item and item.data(Qt.ItemDataRole.UserRole):
                self._aud_register_dialog(int(item.data(Qt.ItemDataRole.UserRole)), load)
        table.cellClicked.connect(open_record)
        search.textChanged.connect(load)
        ctx["refresh_master_register"] = load
        load()
        return tab

    # ---------------- shared small helpers ----------------
    def _aud_table(self, headers, stretch_last=False):
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setVisible(False)
        table.setWordWrap(True)
        table.horizontalHeader().setStretchLastSection(stretch_last)
        return table

    def _aud_readonly(self, table):
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

    def _aud_item(self, text, audit_id=None, link=False):
        item = QTableWidgetItem(safe(text))
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        if audit_id is not None:
            item.setData(Qt.ItemDataRole.UserRole, int(audit_id))
        if link:
            font = QFont()
            font.setBold(True)
            font.setUnderline(True)
            item.setFont(font)
            item.setForeground(QBrush(QColor("#0B5ED7")))
        return item

    def _aud_date_edit(self, value=""):
        editor = QDateEdit()
        editor.setCalendarPopup(True)
        editor.setDisplayFormat("dd-MMM-yyyy")
        # Calendar range follows the local desktop clock — not year 2000.
        today = QDate.currentDate()
        editor.setMinimumDate(today.addYears(-5))
        editor.setMaximumDate(today.addYears(5))
        parsed = QDate.fromString(safe(value), "yyyy-MM-dd")
        editor.setDate(parsed if parsed.isValid() else today)
        return editor

    def _aud_blank_date_edit(self, value=""):
        """Date editor that can be left blank (shows empty at the minimum date)."""
        editor = QDateEdit()
        editor.setCalendarPopup(True)
        editor.setDisplayFormat("dd-MMM-yyyy")
        today = QDate.currentDate()
        # Keep a modest past window for historical target dates; default empty uses min sentinel.
        editor.setMinimumDate(today.addYears(-5))
        editor.setMaximumDate(today.addYears(5))
        editor.setSpecialValueText(" ")
        parsed = QDate.fromString(safe(value), "yyyy-MM-dd")
        if parsed.isValid() and parsed >= editor.minimumDate():
            editor.setDate(parsed)
        else:
            editor.setDate(editor.minimumDate())
        return editor

    def _aud_date_text(self, editor):
        if editor.date() == editor.minimumDate() and editor.specialValueText():
            return ""
        return editor.date().toString("yyyy-MM-dd")

    # ============================================================
    # TAB 1 - AUDIT DASHBOARD
    # ============================================================
    def _aud_build_dashboard(self, ctx):
        tab = QWidget()
        lay = QVBoxLayout(tab)
        cards = QHBoxLayout()
        labels = {}
        for key, title in [("conducted", "Total Audits Conducted"), ("findings", "Total Findings"),
                           ("open", "Open Findings"), ("capa", "Findings Pending in CAPA")]:
            box = QGroupBox(title)
            bl = QVBoxLayout(box)
            lab = QLabel("0")
            lab.setStyleSheet("font-size:30px;font-weight:bold;color:#17365D;")
            lab.setAlignment(Qt.AlignmentFlag.AlignCenter)
            bl.addWidget(lab)
            cards.addWidget(box)
            labels[key] = lab
        lay.addLayout(cards)
        top = QHBoxLayout()
        top.addWidget(QLabel("<b>Audit Summary</b> — click an Audit Number to open its complete Audit Register."), 1)
        search = QLineEdit()
        search.setPlaceholderText("Search audit number, type or status...")
        search.setFixedWidth(320)
        refresh_btn = QPushButton("Refresh")
        top.addWidget(search)
        top.addWidget(refresh_btn)
        lay.addLayout(top)
        table = self._aud_table(["Audit Number", "Audit Type", "Audit Date", "Status"])
        self._aud_readonly(table)
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        lay.addWidget(table, 1)

        def load():
            aud_sync_capa_stages()
            m = aud_metrics()
            for key, lab in labels.items():
                lab.setText(str(m[key]))
            needle = search.text().strip().lower()
            rows = [a for a in aud_list(AUD_CONDUCTED_STATUSES)
                    if not needle or needle in " ".join([safe(a["number"]), safe(a["audit_type"]), safe(a["status"])]).lower()]
            table.setRowCount(len(rows))
            for ri, a in enumerate(rows):
                # Summary level shows ONLY the parent Audit Number - never finding numbers.
                table.setItem(ri, 0, self._aud_item(a["number"], a["id"], link=True))
                table.setItem(ri, 1, self._aud_item(a["audit_type"], a["id"]))
                table.setItem(ri, 2, self._aud_item(a["audit_date"], a["id"]))
                table.setItem(ri, 3, self._aud_item(a["status"], a["id"]))

        def open_register(row, col=0):
            item = table.item(row, 0)
            if item is not None and item.data(Qt.ItemDataRole.UserRole):
                self._aud_register_dialog(int(item.data(Qt.ItemDataRole.UserRole)), load)

        table.cellClicked.connect(lambda r, c: open_register(r) if c == 0 else None)
        table.cellDoubleClicked.connect(open_register)
        search.textChanged.connect(load)
        refresh_btn.clicked.connect(load)
        ctx["refresh_dashboard"] = load
        return tab

    # ---------------- Audit Register (drill-down / UI report view) ----------------
    def _aud_register_dialog(self, audit_id, on_change=None):
        aud_sync_capa_stages(audit_id)
        audit = aud_get(audit_id)
        if not audit:
            QMessageBox.warning(self, "Audit Register", "The selected audit could not be found.")
            return
        d = QDialog(self)
        d.setWindowTitle(f"Audit Register — {safe(audit['number'])}")
        d.setWindowFlags(d.windowFlags() | Qt.WindowType.WindowMinMaxButtonsHint | Qt.WindowType.WindowCloseButtonHint)
        d.resize(1380, 780)
        d.setMinimumSize(900, 560)
        lay = QVBoxLayout(d)
        title = QLabel(f"AUDIT REGISTER — {safe(audit['number'])}")
        title.setStyleSheet("font-size:22px;font-weight:bold;color:#17365D;")
        lay.addWidget(title)
        info = QLabel()
        info.setTextFormat(Qt.TextFormat.RichText)
        info.setWordWrap(True)
        lay.addWidget(info)
        summary = QLabel()
        summary.setStyleSheet("font-weight:bold;")
        lay.addWidget(summary)
        table = self._aud_table(AUD_REGISTER_HEADERS)
        self._aud_readonly(table)
        lay.addWidget(table, 1)
        btns = QHBoxLayout()
        excel_btn = QPushButton("Export Excel")
        word_btn = QPushButton("Export Word")
        pdf_btn = QPushButton("Export PDF")
        edit_btn = QPushButton("Edit Submitted Audit")
        close_btn = QPushButton("Close")
        for b in (excel_btn, word_btn, pdf_btn, edit_btn, close_btn):
            b.setMinimumHeight(36)
        edit_btn.setStyleSheet("font-weight:bold;background:#17365D;color:white;")
        edit_btn.setToolTip("Open this audit in the Audit tab so findings can be edited and re-saved.")
        btns.addWidget(excel_btn)
        btns.addWidget(word_btn)
        btns.addWidget(pdf_btn)
        btns.addStretch(1)
        btns.addWidget(edit_btn)
        btns.addWidget(close_btn)
        lay.addLayout(btns)
        widths = [150, 170, 110, 80, 110, 260, 190, 210, 110, 95, 60, 90, 130, 160]

        def load():
            aud_sync_capa_stages(audit_id)
            a = aud_get(audit_id)
            procs = ", ".join(aud_parse_procedures(a["procedures"])) or "—"
            info.setText(
                f"<table cellspacing='6'><tr>"
                f"<td><b>Audit Type:</b> {safe(a['audit_type'])}</td><td><b>Date:</b> {safe(a['audit_date'])}</td>"
                f"<td><b>Location:</b> {safe(a['location'])}</td><td><b>Auditor:</b> {safe(a['auditor'])}</td>"
                f"<td><b>Auditee:</b> {safe(a['auditee']) or '—'}</td><td><b>Status:</b> {safe(a['status'])}</td></tr>"
                f"<tr><td colspan='6'><b>Procedures:</b> {procs} &nbsp; <b>ISO Standard:</b> {safe(a['standard']) or '—'}</td></tr></table>")
            headers, rows = aud_register_table(audit_id)
            closed = sum(1 for r in rows if r[11] == "Closed")
            total = len(rows)
            open_count = total - closed
            open_pct = (open_count * 100.0 / total) if total else 0.0
            closed_pct = (closed * 100.0 / total) if total else 0.0
            summary.setText(f"Total findings: {total}   |   Open: {open_count} ({open_pct:.1f}%)   |   "
                            f"Closed: {closed} ({closed_pct:.1f}%)   |   CAPA routed: {sum(1 for r in rows if r[10] == 'Yes')}")
            table.setRowCount(len(rows))
            for ri, row in enumerate(rows):
                for ci, val in enumerate(row):
                    table.setItem(ri, ci, self._aud_item(val))
            for ci, wd in enumerate(widths):
                table.setColumnWidth(ci, wd)
            table.resizeRowsToContents()
            edit_btn.setEnabled(safe(a["status"]) in ("Submitted", "In Progress", "Completed", "Scheduled"))

        def export_excel():
            path, _ = QFileDialog.getSaveFileName(d, "Export Audit Register (Excel)",
                                                  str(REPORT_DIR / f"{safe(audit['number'])}_Audit_Register.xlsx"),
                                                  "Excel Workbook (*.xlsx)")
            if not path:
                return
            try:
                aud_export_excel(audit_id, path)
                self.show_export_success(path)
            except Exception as exc:
                logging.exception("Audit Excel export failed")
                QMessageBox.critical(d, "Export Failed", f"Unable to create the Excel report.\n\n{exc}")

        def export_word():
            path, _ = QFileDialog.getSaveFileName(d, "Export Audit Register (Word)",
                                                  str(REPORT_DIR / f"{safe(audit['number'])}_Audit_Register.docx"),
                                                  "Word Document (*.docx)")
            if not path:
                return
            try:
                aud_export_word(audit_id, path)
                self.show_export_success(path)
            except Exception as exc:
                logging.exception("Audit Word export failed")
                QMessageBox.critical(d, "Export Failed", f"Unable to create the Word report.\n\n{exc}")

        def export_pdf():
            path, _ = QFileDialog.getSaveFileName(d, "Export Audit Register (PDF)",
                                                  str(REPORT_DIR / f"{safe(audit['number'])}_Audit_Register.pdf"),
                                                  "PDF Document (*.pdf)")
            if not path:
                return
            try:
                aud_export_pdf(audit_id, path)
                self.show_export_success(path)
            except Exception as exc:
                logging.exception("Audit PDF export failed")
                QMessageBox.critical(d, "Export Failed", f"Unable to create the PDF report.\\n\\n{exc}")

        def edit_submitted():
            d.accept()
            # Open the Audit conduct tab with this audit selected for editing.
            try:
                if "open_checklist" in (getattr(self, "_aud_ctx", {}) or {}):
                    pass
            except Exception:
                pass
            # Prefer live context from the current Audit Management tabs if available.
            parent_tabs = None
            for w in self.findChildren(QTabWidget):
                # Audit Management tab widget has 8+ tabs including "Audit"
                try:
                    labels = [w.tabText(i) for i in range(w.count())]
                    if "Audit" in labels and "Schedule Audit" in labels:
                        parent_tabs = w
                        break
                except Exception:
                    continue
            if parent_tabs is not None:
                for i in range(parent_tabs.count()):
                    if parent_tabs.tabText(i) == "Audit":
                        parent_tabs.setCurrentIndex(i)
                        break
            # Call open_checklist via stored ctx on MainWindow if set
            ctx = getattr(self, "_last_audit_ctx", None)
            if ctx and callable(ctx.get("open_checklist")):
                ctx["open_checklist"](audit_id)
                if callable(ctx.get("refresh_checklist_options")):
                    ctx["refresh_checklist_options"]()
            else:
                QMessageBox.information(
                    self, "Edit Audit",
                    "Open the Audit tab and select this audit from the list to continue editing."
                )

        excel_btn.clicked.connect(export_excel)
        word_btn.clicked.connect(export_word)
        pdf_btn.clicked.connect(export_pdf)
        edit_btn.clicked.connect(edit_submitted)
        close_btn.clicked.connect(d.accept)
        load()
        d.exec()
        if on_change:
            on_change()

    # ============================================================
    # TAB 2 - AUDIT SCHEDULING
    # ============================================================
    def _aud_build_schedule(self, ctx):
        tab = QWidget()
        lay = QVBoxLayout(tab)
        state = {"id": None}
        box = QGroupBox("Schedule an Audit")
        form_lay = QGridLayout(box)
        number_label = QLabel("Audit Number: assigned automatically on first save")
        number_label.setStyleSheet("font-weight:bold;color:#17365D;")
        form_lay.addWidget(number_label, 0, 0, 1, 4)
        type_combo = QComboBox()
        type_combo.addItem("Select audit type...")
        type_combo.addItems(AUD_TYPES)
        form_lay.addWidget(QLabel("Audit Type:"), 1, 0)
        form_lay.addWidget(type_combo, 1, 1)
        proc_label = QLabel("Internal Audit Procedures:")
        proc_combo = QComboBox()
        proc_combo.addItem("Select procedure...")
        proc_combo.addItems(list(aud_procedure_library().keys()))
        proc_list = QListWidget()
        proc_list.setMaximumHeight(92)
        proc_add_btn = QPushButton("Add Procedure")
        proc_remove_btn = QPushButton("Remove Selected")
        proc_picker = QWidget()
        proc_picker_lay = QVBoxLayout(proc_picker)
        proc_picker_lay.setContentsMargins(0, 0, 0, 0)
        proc_picker_row = QHBoxLayout()
        proc_picker_row.addWidget(proc_combo, 1)
        proc_picker_row.addWidget(proc_add_btn)
        proc_picker_row.addWidget(proc_remove_btn)
        proc_picker_lay.addLayout(proc_picker_row)
        proc_picker_lay.addWidget(proc_list)
        form_lay.addWidget(proc_label, 2, 0, Qt.AlignmentFlag.AlignTop)
        form_lay.addWidget(proc_picker, 2, 1, 1, 3)
        date_edit = self._aud_date_edit()
        location_edit = QLineEdit()
        auditor_edit = QLineEdit()
        auditor_combo = QComboBox()
        auditor_combo.setMinimumWidth(240)
        auditor_combo.addItem("Select trained auditor (Internal Audit)", "")
        title_edit = QLineEdit()
        scope_edit = QTextEdit()
        scope_edit.setMaximumHeight(60)
        details = QWidget()
        dl = QGridLayout(details)
        dl.setContentsMargins(0, 0, 0, 0)
        dl.addWidget(QLabel("Date:"), 0, 0)
        dl.addWidget(date_edit, 0, 1)
        dl.addWidget(QLabel("Location:"), 0, 2)
        dl.addWidget(location_edit, 0, 3)
        dl.addWidget(QLabel("Auditor:"), 1, 0)
        auditor_stack = QWidget()
        auditor_stack_lay = QHBoxLayout(auditor_stack)
        auditor_stack_lay.setContentsMargins(0, 0, 0, 0)
        auditor_stack_lay.addWidget(auditor_edit)
        auditor_stack_lay.addWidget(auditor_combo)
        dl.addWidget(auditor_stack, 1, 1)
        dl.addWidget(QLabel("Audit Title:"), 1, 2)
        dl.addWidget(title_edit, 1, 3)
        dl.addWidget(QLabel("Scope:"), 2, 0, Qt.AlignmentFlag.AlignTop)
        dl.addWidget(scope_edit, 2, 1, 1, 3)
        form_lay.addWidget(details, 3, 0, 1, 4)
        train_hint = QLabel("Internal audits may only be conducted by personnel with a current Internal Audit / Auditor training record in the Training Register.")
        train_hint.setWordWrap(True)
        train_hint.setStyleSheet("color:#5A6A7A;font-size:11px;")
        form_lay.addWidget(train_hint, 4, 0, 1, 4)
        form_lay.setColumnStretch(1, 1)
        form_lay.setColumnStretch(3, 1)
        lay.addWidget(box)
        btn_row = QHBoxLayout()
        save_btn = QPushButton("Save for later")
        submit_btn = QPushButton("Submit")
        draft_btn = QPushButton("Draft")
        clear_btn = QPushButton("New / Clear")
        for b in (save_btn, submit_btn, draft_btn, clear_btn):
            b.setMinimumHeight(38)
            b.setMinimumWidth(120)
        submit_btn.setStyleSheet("font-weight:bold;background:#17365D;color:white;")
        save_btn.setToolTip("Save the schedule so it can be edited later (status: Scheduled).")
        submit_btn.setToolTip("Submit the schedule. The audit then appears in the Audit tab for conducting.")
        btn_row.addStretch(1)
        for b in (draft_btn, save_btn, submit_btn, clear_btn):
            btn_row.addWidget(b)
        lay.addLayout(btn_row)

        lay.addWidget(QLabel("<b>All audits</b> — Draft and Saved-for-later (Scheduled) audits can be edited here. Once Submitted they continue in the <b>Audit</b> tab."))
        table = self._aud_table(["Audit Number", "Audit Type", "Date", "Location", "Auditor", "Status"])
        self._aud_readonly(table)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        lay.addWidget(table, 1)
        row_btns = QHBoxLayout()
        edit_btn = QPushButton("Edit Selected (Draft / Save for later)")
        delete_btn = QPushButton("Delete Selected Audit")
        row_btns.addWidget(edit_btn)
        row_btns.addWidget(delete_btn)
        row_btns.addStretch(1)
        lay.addLayout(row_btns)

        def checked_procedures():
            return [proc_list.item(i).text() for i in range(proc_list.count())]

        def add_selected_procedure():
            name = proc_combo.currentText()
            if proc_combo.currentIndex() <= 0:
                return
            if not any(proc_list.item(i).text() == name for i in range(proc_list.count())):
                proc_list.addItem(name)
            proc_combo.setCurrentIndex(0)

        def remove_selected_procedure():
            row = proc_list.currentRow()
            if row >= 0:
                proc_list.takeItem(row)

        def refresh_trained_auditors(selected_name=""):
            auditor_combo.blockSignals(True)
            auditor_combo.clear()
            auditor_combo.addItem("Select trained auditor (Internal Audit)", "")
            for r in aud_trained_people("auditor"):
                label = f'{safe(r["name"])}  |  ID: {safe(r["employee_id"]) or "—"}  |  {safe(r["designation"]) or "—"}'
                auditor_combo.addItem(label, safe(r["name"]))
            if selected_name:
                ix = auditor_combo.findData(selected_name)
                if ix < 0:
                    # Legacy free-text auditor still shown for historical records.
                    auditor_combo.addItem(selected_name + " (not in training register)", selected_name)
                    ix = auditor_combo.findData(selected_name)
                auditor_combo.setCurrentIndex(ix if ix >= 0 else 0)
            auditor_combo.blockSignals(False)

        def update_visibility():
            is_internal = aud_is_internal_type(type_combo.currentText())
            proc_label.setVisible(is_internal)
            proc_picker.setVisible(is_internal)
            has_type = type_combo.currentIndex() > 0
            # Date / Location / Auditor are requested after the type (and, for Internal, the procedures) are chosen.
            details.setEnabled(has_type)
            # Internal audits: trained auditor dropdown only (Training Register gate, same idea as PTW).
            auditor_combo.setVisible(is_internal)
            auditor_edit.setVisible(not is_internal)
            if is_internal:
                refresh_trained_auditors(auditor_combo.currentData() or auditor_edit.text().strip())
            train_hint.setVisible(is_internal)

        def clear_form():
            state["id"] = None
            number_label.setText("Audit Number: assigned automatically on first save")
            type_combo.setCurrentIndex(0)
            proc_list.clear()
            proc_combo.setCurrentIndex(0)
            date_edit.setDate(QDate.currentDate())
            location_edit.clear()
            auditor_edit.clear()
            refresh_trained_auditors()
            title_edit.clear()
            scope_edit.clear()
            update_visibility()

        def load_table():
            rows = aud_list()
            table.setRowCount(len(rows))
            for ri, a in enumerate(rows):
                for ci, key in enumerate(["number", "audit_type", "audit_date", "location", "auditor", "status"]):
                    table.setItem(ri, ci, self._aud_item(a[key], a["id"]))

        def collect():
            is_internal = aud_is_internal_type(type_combo.currentText() if type_combo.currentIndex() > 0 else "")
            auditor_name = (safe(auditor_combo.currentData()).strip() if is_internal else auditor_edit.text().strip())
            return {
                "audit_type": type_combo.currentText() if type_combo.currentIndex() > 0 else "",
                "procedures": checked_procedures(),
                "audit_date": date_edit.date().toString("yyyy-MM-dd"),
                "location": location_edit.text(),
                "auditor": auditor_name,
                "title": title_edit.text(),
                "scope": scope_edit.toPlainText(),
            }

        def do_save(status):
            data = collect()
            if aud_is_internal_type(data.get("audit_type")):
                if not data.get("auditor"):
                    QMessageBox.warning(tab, "Trained Auditor Required",
                        "Select a trained auditor from the Training Register before scheduling or submitting an Internal Audit.")
                    return
                if not aud_is_trained_auditor(data["auditor"]):
                    QMessageBox.warning(tab, "Trained Auditor Required",
                        f"Auditor '{data['auditor']}' does not have a current Internal Audit / Auditor training record.\n\n"
                        "Only trained auditors may conduct an Internal Audit (same control as PTW holder training).")
                    return
            try:
                audit_id = aud_save_audit(state["id"], data, status)
            except ValueError as exc:
                QMessageBox.warning(tab, "Audit Scheduling", str(exc))
                return
            except Exception as exc:
                logging.exception("Audit save failed")
                QMessageBox.critical(tab, "Audit Scheduling", f"Could not save the audit.\n\n{exc}")
                return
            audit = aud_get(audit_id)
            number = safe(audit["number"])
            if status == "Submitted":
                # No auditee popup — checklist is opened ready in the Audit tab.
                clear_form()
                load_table()
                if "refresh_checklist_options" in ctx:
                    ctx["refresh_checklist_options"]()
                if "open_checklist" in ctx:
                    ctx["open_checklist"](audit_id)
                # Tab order: Dashboard=0, Schedule=1, Notifications=2, Audit=3, Saved for later=4, ...
                try:
                    for i in range(ctx["tabs"].count()):
                        if ctx["tabs"].tabText(i) == "Audit":
                            ctx["tabs"].setCurrentIndex(i)
                            break
                except Exception:
                    ctx["tabs"].setCurrentIndex(3)
                return
            label = "Saved for later (Scheduled)" if status == "Scheduled" else status
            QMessageBox.information(tab, "Audit Scheduling",
                                    f"Audit {number} {label}." + (
                                        "\nIt appears under Saved for later and Notifications, and can be edited later." if status == "Scheduled" else ""))
            clear_form()
            load_table()
            if "refresh_saved_for_later" in ctx:
                ctx["refresh_saved_for_later"]()

        def edit_selected():
            row = table.currentRow()
            item = table.item(row, 0) if row >= 0 else None
            if item is None:
                QMessageBox.information(tab, "Select Audit", "Select a Draft or Scheduled audit first.")
                return
            audit = aud_get(int(item.data(Qt.ItemDataRole.UserRole)))
            if safe(audit["status"]) not in ("Draft", "Scheduled"):
                QMessageBox.information(tab, "Audit Submitted",
                                        "Submitted audits can no longer be re-scheduled. Continue them in the Audit tab.")
                return
            state["id"] = int(audit["id"])
            number_label.setText(f"Audit Number: {safe(audit['number'])}   (editing {safe(audit['status'])})")
            type_combo.setCurrentText(safe(audit["audit_type"]))
            selected = aud_parse_procedures(audit["procedures"])
            proc_list.clear()
            for name in selected:
                proc_list.addItem(name)
            parsed = QDate.fromString(safe(audit["audit_date"]), "yyyy-MM-dd")
            date_edit.setDate(parsed if parsed.isValid() else QDate.currentDate())
            location_edit.setText(safe(audit["location"]))
            auditor_edit.setText(safe(audit["auditor"]))
            refresh_trained_auditors(safe(audit["auditor"]))
            title_edit.setText(safe(audit["title"]))
            scope_edit.setPlainText(safe(audit["scope"]))
            update_visibility()

        def delete_selected():
            row = table.currentRow()
            item = table.item(row, 0) if row >= 0 else None
            if item is None:
                QMessageBox.information(tab, "Select Audit", "Select an audit to delete first.")
                return
            audit = aud_get(int(item.data(Qt.ItemDataRole.UserRole)))
            count = db.fetchone("SELECT COUNT(*) AS n FROM aud_findings WHERE audit_id=?", (audit["id"],))["n"]
            if QMessageBox.question(
                    tab, "Delete Audit",
                    f"Delete audit {safe(audit['number'])}?\n\nIts {count} finding(s) and any notification will be removed. "
                    f"CAPA records already raised stay in the CAPA module.") != QMessageBox.StandardButton.Yes:
                return
            try:
                aud_delete_audit(int(audit["id"]))
            except Exception as exc:
                logging.exception("Audit delete failed")
                QMessageBox.critical(tab, "Delete Failed", str(exc))
                return
            if state["id"] == int(audit["id"]):
                clear_form()
            load_table()

        type_combo.currentIndexChanged.connect(lambda *_: update_visibility())
        proc_add_btn.clicked.connect(add_selected_procedure)
        proc_remove_btn.clicked.connect(remove_selected_procedure)
        draft_btn.clicked.connect(lambda: do_save("Draft"))
        save_btn.clicked.connect(lambda: do_save("Scheduled"))
        submit_btn.clicked.connect(lambda: do_save("Submitted"))
        clear_btn.clicked.connect(clear_form)
        edit_btn.clicked.connect(edit_selected)
        delete_btn.clicked.connect(delete_selected)
        table.cellDoubleClicked.connect(lambda *_: edit_selected())
        update_visibility()
        load_table()
        ctx["refresh_schedule"] = load_table
        return tab

    # ============================================================
    # TAB - AUDIT (conduct checklist; Save for later / Submit)
    # ============================================================
    def _aud_build_checklist(self, ctx):
        tab = QWidget()
        lay = QVBoxLayout(tab)
        state = {"audit_id": None, "removed": [], "loading": False}
        top = QGridLayout()
        audit_combo = QComboBox()
        audit_combo.setMinimumWidth(320)
        auditee_edit = QLineEdit()
        auditee_edit.setPlaceholderText("Enter the Auditee")
        auditee_combo = QComboBox()
        auditee_combo.setMinimumWidth(260)
        auditee_combo.addItem("Select trained auditee (Internal Audit)", "")
        type_combo = QComboBox()
        type_combo.addItem("Select type of audit...")
        type_combo.addItems(AUD_TYPES)
        top.addWidget(QLabel("Select Audit:"), 0, 0)
        top.addWidget(audit_combo, 0, 1)
        top.addWidget(QLabel("Auditee:"), 0, 2)
        auditee_wrap = QWidget()
        aw = QHBoxLayout(auditee_wrap); aw.setContentsMargins(0,0,0,0)
        aw.addWidget(auditee_edit); aw.addWidget(auditee_combo)
        top.addWidget(auditee_wrap, 0, 3)
        top.addWidget(QLabel("Type of Audit:"), 0, 4)
        top.addWidget(type_combo, 0, 5)
        top.setColumnStretch(3, 1)
        lay.addLayout(top)
        auditee_train_hint = QLabel("For Internal Audits the Auditee is selected from personnel with a current Internal Audit / related training record.")
        auditee_train_hint.setWordWrap(True)
        auditee_train_hint.setStyleSheet("color:#5A6A7A;font-size:11px;")
        lay.addWidget(auditee_train_hint)

        standard_label = QLabel("ISO Checklist Standard:")
        standard_combo = QComboBox()
        standard_combo.addItem("Select ISO standard / IMS...")
        standard_combo.addItems(["ISO 45001", "ISO 9001", "ISO 14001", "IMS"])
        top.addWidget(standard_label, 1, 0)
        top.addWidget(standard_combo, 1, 1)
        proc_box = QGroupBox("Scheduled procedures (from Schedule Audit — read only)")
        pl = QVBoxLayout(proc_box)
        proc_list = QListWidget()
        proc_list.setMaximumHeight(120)
        proc_list.setToolTip("Procedures were chosen when this audit was scheduled. The checklist is loaded from those procedures only.")
        pl.addWidget(proc_list)
        lay.addWidget(proc_box)
        hint = QLabel("")
        hint.setStyleSheet("color:#555;")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        table = self._aud_table(AUD_CHECKLIST_HEADERS)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        for ci, wd in enumerate([150, 170, 110, 80, 110, 260, 190, 210, 110, 110, 90, 100]):
            table.setColumnWidth(ci, wd)
        lay.addWidget(table, 1)

        btns = QHBoxLayout()
        add_btn = QPushButton("+ Add Blank Row")
        del_btn = QPushButton("Delete Selected Row")
        meta_btn = QPushButton("Save Auditee && Type")
        save_later_btn = QPushButton("Save for later")
        submit_btn = QPushButton("Submit")
        for b in (add_btn, del_btn, meta_btn, save_later_btn, submit_btn):
            b.setMinimumHeight(38)
        save_later_btn.setToolTip("Save checklist progress so the audit can be edited later (status: In Progress).")
        submit_btn.setToolTip("Submit the completed audit. Findings then appear in Audit Findings Register / Close Out.")
        submit_btn.setStyleSheet("font-weight:bold;background:#17365D;color:white;")
        btns.addWidget(add_btn)
        btns.addWidget(del_btn)
        btns.addStretch(1)
        btns.addWidget(meta_btn)
        btns.addWidget(save_later_btn)
        btns.addWidget(submit_btn)
        lay.addLayout(btns)

        # ---- row helpers ----
        def row_meta(row):
            item = table.item(row, 1)
            try:
                return json.loads(safe(item.data(Qt.ItemDataRole.UserRole)) or "{}") if item else {}
            except Exception:
                return {}

        def set_meta(row, meta):
            item = table.item(row, 1)
            if item is not None:
                item.setData(Qt.ItemDataRole.UserRole, json.dumps(meta))

        def cell_text(row, col):
            item = table.item(row, col)
            return item.text() if item else ""

        def readonly_item(text):
            it = QTableWidgetItem(safe(text))
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            return it

        def add_row(rec):
            row = table.rowCount()
            table.insertRow(row)
            audit = aud_get(state["audit_id"])
            table.setItem(row, 0, readonly_item(audit["number"]))
            fno = readonly_item(rec.get("finding_no") or "(assigned on save)")
            table.setItem(row, 1, fno)
            set_meta(row, {"id": rec.get("id"), "template": bool(rec.get("template")), "procedure": safe(rec.get("procedure"))})
            for col, key in ((2, "auditor_name"), (3, "clause_sno"), (4, "sub_clause"), (5, "checklist_pts"),
                             (6, "evidence"), (7, "corrective_action"), (8, "responsible")):
                table.setItem(row, col, QTableWidgetItem(safe(rec.get(key))))
            if safe(rec.get("procedure")):
                table.item(row, 5).setToolTip(f"Procedure: {safe(rec.get('procedure'))}")
            table.setCellWidget(row, 9, self._aud_blank_date_edit(rec.get("target_date")))
            capa = QComboBox()
            capa.addItems(["No", "Yes"])
            capa.setCurrentText("Yes" if safe(rec.get("capa_required")) == "Yes" else "No")
            table.setCellWidget(row, 10, capa)
            status = QComboBox()
            status.addItems(AUD_FINDING_STATUSES)
            status.setCurrentText(safe(rec.get("status")) if safe(rec.get("status")) in AUD_FINDING_STATUSES else "Open")
            table.setCellWidget(row, 11, status)
            lock_capa(row, rec)
            table.setRowHeight(row, 52)

        def lock_capa(row, rec):
            capa = table.cellWidget(row, 10)
            if capa is None:
                return
            if safe(rec.get("capa_number")):
                capa.setCurrentText("Yes")
                capa.setEnabled(False)
                capa.setToolTip(f"Routed to {safe(rec.get('route_stage')) or AUD_STAGE_CAPA} — CAPA No. {safe(rec.get('capa_number'))}")
            else:
                capa.setEnabled(True)
                capa.setToolTip("Yes = route to the CAPA module first, then to Audit Report / Closure.  No = Audit Register only.")

        def rec_from_db(f):
            return {"id": f["id"], "finding_no": f["finding_no"], "auditor_name": f["auditor_name"], "clause_sno": f["clause_sno"],
                    "sub_clause": f["sub_clause"], "checklist_pts": f["checklist_pts"], "evidence": f["evidence"],
                    "corrective_action": f["corrective_action"], "responsible": f["responsible"], "target_date": f["target_date"],
                    "capa_required": f["capa_required"], "status": f["status"], "procedure": f["procedure"],
                    "capa_number": f["capa_number"], "route_stage": f["route_stage"], "template": False}

        def collect_rows():
            rows = []
            for r in range(table.rowCount()):
                meta = row_meta(r)
                capa = table.cellWidget(r, 10)
                status = table.cellWidget(r, 11)
                target = table.cellWidget(r, 9)
                rows.append({
                    "id": meta.get("id"), "template": meta.get("template"), "procedure": meta.get("procedure"),
                    "auditor_name": cell_text(r, 2), "clause_sno": cell_text(r, 3), "sub_clause": cell_text(r, 4),
                    "checklist_pts": cell_text(r, 5), "evidence": cell_text(r, 6), "corrective_action": cell_text(r, 7),
                    "responsible": cell_text(r, 8), "target_date": self._aud_date_text(target) if target else "",
                    "capa_required": capa.currentText() if capa else "No", "status": status.currentText() if status else "Open"})
            return rows

        def add_template_rows(proc):
            audit = aud_get(state["audit_id"])
            clauses, items = aud_procedure_library().get(proc, ("", []))
            present = {(row_meta(r).get("procedure"), cell_text(r, 5)) for r in range(table.rowCount())}
            for n, text in enumerate(items, 1):
                if (proc, text) in present:
                    continue
                add_row({"template": True, "procedure": proc, "auditor_name": safe(audit["auditor"]),
                         "clause_sno": str(n), "sub_clause": clauses, "checklist_pts": text})

        def add_iso_template_rows(standard):
            if not standard or standard_combo.currentIndex() <= 0 or state["audit_id"] is None:
                return
            audit = aud_get(state["audit_id"])
            present = {(row_meta(r).get("procedure"), cell_text(r, 5)) for r in range(table.rowCount())}
            for rec in aud_iso_standard_checklist(standard):
                if (standard, rec["checklist_pts"]) in present:
                    continue
                add_row({"template": True, "procedure": standard, "auditor_name": safe(audit["auditor"]),
                         "clause_sno": rec["clause_sno"], "sub_clause": rec["sub_clause"],
                         "checklist_pts": rec["checklist_pts"]})

        def remove_template_rows(proc):
            for r in range(table.rowCount() - 1, -1, -1):
                meta = row_meta(r)
                if meta.get("template") and not meta.get("id") and meta.get("procedure") == proc:
                    table.removeRow(r)

        def checked_procs():
            # All listed procedures are the ones scheduled for this audit (read-only list).
            return [proc_list.item(i).text() for i in range(proc_list.count())]

        def refresh_trained_auditees(selected_name=""):
            auditee_combo.blockSignals(True)
            auditee_combo.clear()
            auditee_combo.addItem("Select trained auditee (Internal Audit)", "")
            for r in aud_trained_people("auditee"):
                label = f'{safe(r["name"])}  |  ID: {safe(r["employee_id"]) or "—"}  |  {safe(r["designation"]) or "—"}'
                auditee_combo.addItem(label, safe(r["name"]))
            if selected_name:
                ix = auditee_combo.findData(selected_name)
                if ix < 0:
                    auditee_combo.addItem(selected_name + " (manual / not in training register)", selected_name)
                    ix = auditee_combo.findData(selected_name)
                auditee_combo.setCurrentIndex(ix if ix >= 0 else 0)
            auditee_combo.blockSignals(False)

        def current_auditee_value():
            if aud_is_internal_type(type_combo.currentText()):
                return safe(auditee_combo.currentData()).strip() or auditee_edit.text().strip()
            return auditee_edit.text().strip()

        def update_type_ui():
            has_type = type_combo.currentIndex() > 0 and state["audit_id"] is not None
            uses_procs = aud_is_internal_type(type_combo.currentText())
            standard_label.setVisible(has_type and not uses_procs)
            standard_combo.setVisible(has_type and not uses_procs)
            proc_box.setVisible(uses_procs and state["audit_id"] is not None)
            table.setEnabled(has_type and (uses_procs or standard_combo.currentIndex() > 0))
            auditee_combo.setVisible(uses_procs)
            auditee_edit.setVisible(not uses_procs)
            auditee_train_hint.setVisible(uses_procs)
            if uses_procs:
                refresh_trained_auditees(current_auditee_value() or auditee_edit.text().strip())
            for b in (add_btn, del_btn, meta_btn, save_later_btn, submit_btn):
                b.setEnabled(state["audit_id"] is not None)
            if state["audit_id"] is None:
                hint.setText("Select a submitted audit to begin.")
            elif type_combo.currentIndex() == 0:
                hint.setText("Select the Type of Audit to open the checklist.")
            elif uses_procs:
                hint.setText("Checklist is loaded from the procedures scheduled for this audit. Use '+ Add Blank Row' for extra rows. "
                             "Pre-loaded checklist points are saved as findings once Evidence or a Corrective Action is entered (or CAPA = Yes).")
            else:
                hint.setText("Select ISO 45001, ISO 9001 or ISO 14001. The ordered clause checklist will load; enter an observation/evidence for each assessed item.")

        def on_proc_changed(item):
            # Procedures are fixed from Schedule Audit; list is read-only.
            return

        def on_type_changed(*_):
            if state["loading"]:
                return
            update_type_ui()
            if state["audit_id"] is not None and not aud_is_internal_type(type_combo.currentText()):
                standard_combo.setCurrentIndex(0)
                update_type_ui()

        def load_audit(audit_id):
            state["loading"] = True
            state["audit_id"] = audit_id
            state["removed"] = []
            table.setRowCount(0)
            if not audit_id:
                auditee_edit.clear()
                refresh_trained_auditees()
                type_combo.setCurrentIndex(0)
                standard_combo.setCurrentIndex(0)
                proc_list.clear()
                state["loading"] = False
                update_type_ui()
                return
            aud_sync_capa_stages(audit_id)
            audit = aud_get(audit_id)
            auditee_edit.setText(safe(audit["auditee"]))
            refresh_trained_auditees(safe(audit["auditee"]))
            # Type / procedures come from the Schedule Audit step — shown read-only here.
            idx = type_combo.findText(safe(audit["audit_type"]))
            type_combo.setCurrentIndex(idx if idx > 0 else 0)
            type_combo.setEnabled(False)  # locked to scheduled type
            standard_idx = standard_combo.findText(safe(audit["standard"]))
            standard_combo.setCurrentIndex(standard_idx if standard_idx > 0 else 0)
            selected = aud_parse_procedures(audit["procedures"])
            proc_list.clear()
            for name in selected:
                item = QListWidgetItem(name)
                item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)  # not checkable / not editable
                proc_list.addItem(item)
            update_type_ui()
            for f in aud_findings(audit_id):
                add_row(rec_from_db(f))
            if aud_is_internal_type(type_combo.currentText()):
                for proc in selected:
                    add_template_rows(proc)
            elif standard_combo.currentIndex() > 0:
                add_iso_template_rows(standard_combo.currentText())
            state["loading"] = False
            type_combo.setEnabled(True)

        def audit_selected(*_):
            load_audit(audit_combo.currentData())

        def fill_audit_combo(select_id=None):
            current = select_id if select_id is not None else audit_combo.currentData()
            audit_combo.blockSignals(True)
            audit_combo.clear()
            audit_combo.addItem("Select audit...", None)
            for a in aud_list(AUD_CONDUCTED_STATUSES):
                audit_combo.addItem(f"{safe(a['number'])} — {safe(a['audit_type'])} — {safe(a['audit_date'])}", int(a["id"]))
            idx = audit_combo.findData(current) if current is not None else 0
            audit_combo.setCurrentIndex(idx if idx >= 0 else 0)
            audit_combo.blockSignals(False)

        def refresh_options():
            fill_audit_combo()
            if audit_combo.currentData() is None and state["audit_id"] is not None:
                load_audit(None)

        def open_checklist(audit_id):
            fill_audit_combo(audit_id)
            load_audit(audit_combo.currentData())

        def persist_meta():
            if not current_auditee_value():
                QMessageBox.warning(tab, "Auditee Required", "Please enter the Auditee before saving.")
                auditee_edit.setFocus()
                return False
            if type_combo.currentIndex() == 0:
                QMessageBox.warning(tab, "Type of Audit", "Please select the Type of Audit.")
                return False
            is_internal = aud_is_internal_type(type_combo.currentText())
            procs = checked_procs() if is_internal else []
            standard = "" if is_internal else standard_combo.currentText()
            if not is_internal and standard_combo.currentIndex() == 0:
                QMessageBox.warning(tab, "ISO Standard Required", "Select ISO 45001, ISO 9001 or ISO 14001 to load the checklist.")
                return False
            aud_set_auditee(state["audit_id"], current_auditee_value(), type_combo.currentText(), procs, standard)
            return True

        def save_meta():
            if state["audit_id"] is None:
                return
            if persist_meta():
                QMessageBox.information(tab, "Saved", "Auditee and Type of Audit saved.")

        def refresh_row(row, fid):
            f = db.fetchone("SELECT * FROM aud_findings WHERE id=?", (fid,))
            rec = rec_from_db(f)
            table.item(row, 1).setText(safe(f["finding_no"]))
            set_meta(row, {"id": int(f["id"]), "template": False, "procedure": safe(f["procedure"])})
            status = table.cellWidget(row, 11)
            if status is not None:
                status.setCurrentText(safe(f["status"]) if safe(f["status"]) in AUD_FINDING_STATUSES else "Open")
            lock_capa(row, rec)

        def attach_evidence_to_row(row, col):
            # Evidence cell supports choosing one or more files; paths are stored
            # with the finding evidence so they remain available to exports/register.
            if col != 6 or row < 0:
                return
            paths, _ = QFileDialog.getOpenFileNames(
                tab, "Attach Audit Evidence", "",
                "Evidence Files (*.png *.jpg *.jpeg *.bmp *.gif *.pdf *.doc *.docx *.xls *.xlsx *.txt *.mp4 *.avi *.mov);;All Files (*)")
            if not paths:
                return
            current = cell_text(row, 6).strip()
            joined = "; ".join(paths)
            table.item(row, 6).setText((current + "\n" if current else "") + joined)
            table.item(row, 6).setToolTip("Attached evidence: " + joined)

        table.cellDoubleClicked.connect(attach_evidence_to_row)

        def save_checklist(submit=False):
            if state["audit_id"] is None:
                return
            if not persist_meta():
                return
            try:
                result = aud_save_findings(state["audit_id"], collect_rows(), state["removed"])
            except ValueError as exc:
                QMessageBox.warning(tab, "Audit", str(exc))
                return
            except Exception as exc:
                logging.exception("Audit checklist save failed")
                QMessageBox.critical(tab, "Audit", f"Could not save the checklist.\n\n{exc}")
                return
            state["removed"] = []
            for row, fid in result["saved"]:
                refresh_row(row, fid)
            # Save for later keeps the audit editable; Submit finalises for registers / close-out.
            try:
                if submit:
                    aud_set_status(state["audit_id"], "Completed")
                else:
                    cur = aud_get(state["audit_id"])
                    if safe(cur["status"]) in ("Submitted", "Scheduled"):
                        aud_set_status(state["audit_id"], "In Progress")
            except Exception:
                logging.exception("Audit status update failed")
            msg = [f"{len(result['saved'])} finding(s) saved."]
            if result["capa_created"]:
                msg.append("Routed to the CAPA module first (then Audit Report / Closure): " + ", ".join(result["capa_created"]))
            if result["register_only"]:
                msg.append(f"{result['register_only']} finding(s) with CAPA = No were routed to the Audit Register only.")
            msg.extend(result["notes"])
            if submit:
                msg.append("Audit submitted. Findings are available in Audit Findings Register and Close Out.")
                QMessageBox.information(tab, "Audit Submitted", "\n".join(msg))
            else:
                msg.append("Saved for later — you can continue editing this audit from the Audit tab.")
                QMessageBox.information(tab, "Saved for later", "\n".join(msg))
            ctx["refresh_dashboard"]()
            if "refresh_findings_register" in ctx:
                ctx["refresh_findings_register"]()
            if "refresh_closeout_section" in ctx:
                ctx["refresh_closeout_section"]()
            if "refresh_summary" in ctx:
                ctx["refresh_summary"]()
            if "refresh_saved_for_later" in ctx:
                ctx["refresh_saved_for_later"]()

        def delete_row():
            row = table.currentRow()
            if row < 0:
                QMessageBox.information(tab, "Select Row", "Select a row to delete first.")
                return
            meta = row_meta(row)
            if meta.get("id"):
                f = db.fetchone("SELECT capa_number FROM aud_findings WHERE id=?", (meta["id"],))
                if f and safe(f["capa_number"]):
                    QMessageBox.warning(tab, "Cannot Delete", f"This finding has CAPA {safe(f['capa_number'])} in the CAPA module and cannot be deleted here.")
                    return
                state["removed"].append(int(meta["id"]))
            table.removeRow(row)

        def add_blank():
            if state["audit_id"] is None:
                return
            add_row({"template": False, "auditor_name": safe(aud_get(state["audit_id"])["auditor"])})

        audit_combo.currentIndexChanged.connect(audit_selected)
        type_combo.currentIndexChanged.connect(on_type_changed)
        standard_combo.currentIndexChanged.connect(
            lambda *_: (add_iso_template_rows(standard_combo.currentText()), update_type_ui())
            if not state["loading"] and not aud_is_internal_type(type_combo.currentText()) else None
        )
        proc_list.itemChanged.connect(on_proc_changed)
        add_btn.clicked.connect(add_blank)
        del_btn.clicked.connect(delete_row)
        meta_btn.clicked.connect(save_meta)
        save_later_btn.clicked.connect(lambda: save_checklist(False))
        submit_btn.clicked.connect(lambda: save_checklist(True))
        fill_audit_combo()
        update_type_ui()
        ctx["refresh_checklist_options"] = refresh_options
        ctx["open_checklist"] = open_checklist
        return tab

    # ============================================================
    # TAB 4 - AUDIT NOTIFICATIONS (+ Word "Invites Document")
    # ============================================================

    # ============================================================
    # TAB - AUDIT TRACKING REGISTER (open findings + close-out workflow)
    # ============================================================
    def _aud_build_findings_register(self, ctx):
        """Combined Audit Tracking + Cumulative Findings register with due dates and full details."""
        tab = QWidget()
        lay = QVBoxLayout(tab)
        intro = QLabel(
            "Complete audit findings register (open and closed). "
            "Shows observation, corrective action, responsible person, target / due date, days remaining, "
            "close-out status and CAPA. Double-click a row to open the Close Out workspace."
        )
        intro.setWordWrap(True)
        lay.addWidget(intro)
        filters = QHBoxLayout()
        search = QLineEdit(); search.setPlaceholderText("Search audit no., observation, corrective action, responsible…")
        status_filter = QComboBox()
        status_filter.addItems(["All statuses", "Open only", "Closed only", "Submitted for Review", "Overdue"])
        refresh_btn = QPushButton("Refresh")
        filters.addWidget(search, 2)
        filters.addWidget(status_filter)
        filters.addWidget(refresh_btn)
        lay.addLayout(filters)
        headers = ["S.No.", "Audit No.", "Finding No.", "Type", "Date", "Observation / Finding",
                   "Corrective Action", "Responsible", "Target / Due Date", "Days Remaining",
                   "Close-out Status", "CAPA No.", "Auditor"]
        table = self._aud_table(headers)
        self._aud_readonly(table)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.horizontalHeader().setStretchLastSection(True)
        for ci, wd in enumerate([50, 110, 110, 100, 90, 220, 200, 120, 110, 100, 130, 100, 110]):
            table.setColumnWidth(ci, wd)
        lay.addWidget(table, 1)
        open_btn = QPushButton("Open Close-out Workspace")
        open_btn.setMinimumHeight(36)
        open_btn.setStyleSheet("font-weight:bold;background:#17365D;color:white;")
        lay.addWidget(open_btn)

        def load():
            needle = search.text().strip().casefold()
            mode = status_filter.currentText()
            # Prefer cumulative (all findings) when available; fall back to tracking rows.
            try:
                rows = aud_cumulative_findings()
            except Exception:
                rows = aud_tracking_rows(include_closed=True)
            data = []
            for f in rows:
                obs = safe(f["observation_text"]) if "observation_text" in f.keys() else ""
                if not obs:
                    obs = safe(f["evidence"]) if "evidence" in f.keys() else ""
                if not obs:
                    obs = safe(f["checklist_pts"]) if "checklist_pts" in f.keys() else ""
                action = safe(f["corrective_action"]) if "corrective_action" in f.keys() else ""
                co_status = safe(f["closeout_status"]) if "closeout_status" in f.keys() else ""
                status = co_status or (safe(f["status"]) if "status" in f.keys() else "Open")
                days_label, overdue = aud_days_remaining(
                    f["target_date"] if "target_date" in f.keys() else "", status)
                if mode == "Open only" and status.lower() in ("closed", "approved", "capa action complete"):
                    continue
                if mode == "Closed only" and status.lower() not in ("closed", "approved"):
                    continue
                if mode == "Submitted for Review" and "submit" not in status.lower():
                    continue
                if mode == "Overdue" and not overdue:
                    continue
                blob = " ".join([
                    safe(f[k]) if k in f.keys() else ""
                    for k in ("audit_number", "finding_no", "evidence", "corrective_action",
                              "responsible", "closeout_status", "capa_number", "observation_text",
                              "auditor_name", "audit_type")
                ]).casefold()
                if needle and needle not in blob and needle not in obs.casefold() and needle not in action.casefold():
                    continue
                data.append((f, obs, action, days_label, overdue, status))
            table.setRowCount(len(data))
            for i, (f, obs, action, days_label, overdue, status) in enumerate(data):
                vals = [
                    str(i + 1),
                    safe(f["audit_number"]) if "audit_number" in f.keys() else "",
                    safe(f["finding_no"]) if "finding_no" in f.keys() else "",
                    safe(f["audit_type"]) if "audit_type" in f.keys() else "",
                    safe(f["audit_date"]) if "audit_date" in f.keys() else "",
                    obs, action,
                    safe(f["responsible"]) if "responsible" in f.keys() else "",
                    safe(f["target_date"]) if "target_date" in f.keys() else "",
                    days_label, status,
                    safe(f["capa_number"]) if "capa_number" in f.keys() else "—",
                    safe(f["auditor_name"]) if "auditor_name" in f.keys() else (safe(f["audit_auditor"]) if "audit_auditor" in f.keys() else ""),
                ]
                fid = f["id"] if "id" in f.keys() else None
                for c, v in enumerate(vals):
                    item = self._aud_item(v, fid)
                    if overdue and c == 9:
                        item.setForeground(QBrush(QColor("#B00020")))
                        item.setFont(QFont("Arial", 9, QFont.Weight.Bold))
                    table.setItem(i, c, item)
            table.resizeRowsToContents()

        def open_workspace(row=None, col=None):
            r = table.currentRow()
            item = table.item(r, 0) if r >= 0 else None
            if item is None:
                QMessageBox.information(tab, "Select Finding", "Select a findings row first.")
                return
            fid = item.data(Qt.ItemDataRole.UserRole)
            if not fid:
                return
            self._aud_closeout_dialog(int(fid), on_change=load)

        table.cellDoubleClicked.connect(open_workspace)
        open_btn.clicked.connect(open_workspace)
        search.textChanged.connect(lambda *_: load())
        status_filter.currentIndexChanged.connect(lambda *_: load())
        refresh_btn.clicked.connect(load)
        ctx["refresh_findings_register"] = load
        # Keep legacy keys so any remaining callers do not break.
        ctx["refresh_tracking"] = load
        ctx["refresh_cumulative"] = load
        load()
        return tab

    def _aud_build_tracking_register(self, ctx):
        # Backward-compatible alias — combined findings register is the single source of truth.
        return self._aud_build_findings_register(ctx)

    def _aud_closeout_dialog(self, finding_id, on_change=None):
        """Scrollable close-out workspace: details + evidence + submit / senior approve."""
        f = db.fetchone(
            """SELECT f.*, a.number AS audit_number, a.audit_type, a.audit_date, a.location,
                      a.auditor AS audit_auditor, a.auditee
               FROM aud_findings f JOIN audits a ON a.id=f.audit_id WHERE f.id=?""",
            (finding_id,)
        )
        if not f:
            QMessageBox.warning(self, "Finding", "The selected finding could not be found.")
            return
        d = QDialog(self)
        d.setWindowTitle(f"Close-out — {safe(f['finding_no'])} / {safe(f['audit_number'])}")
        d.setWindowFlags(d.windowFlags() | Qt.WindowType.WindowMinMaxButtonsHint | Qt.WindowType.WindowCloseButtonHint)
        d.resize(980, 760)
        d.setMinimumSize(720, 520)
        root = QVBoxLayout(d)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        body = QWidget(); lay = QVBoxLayout(body)

        title = QLabel(f"AUDIT FINDING CLOSE-OUT — {safe(f['finding_no'])}")
        title.setStyleSheet("font-size:18px;font-weight:bold;color:#17365D;")
        lay.addWidget(title)

        info = QLabel(
            f"<b>Audit No.:</b> {safe(f['audit_number'])} &nbsp;|&nbsp; "
            f"<b>Type:</b> {safe(f['audit_type'])} &nbsp;|&nbsp; "
            f"<b>Date:</b> {safe(f['audit_date'])}<br>"
            f"<b>Location:</b> {safe(f['location'])} &nbsp;|&nbsp; "
            f"<b>Auditor:</b> {safe(f['audit_auditor'])} &nbsp;|&nbsp; "
            f"<b>Auditee:</b> {safe(f['auditee']) or '—'}<br>"
            f"<b>Status:</b> {safe(f['status'])} &nbsp;|&nbsp; "
            f"<b>Close-out Status:</b> {safe(f['closeout_status']) or 'Open'} &nbsp;|&nbsp; "
            f"<b>CAPA:</b> {safe(f['capa_number']) or 'Not linked'}"
        )
        info.setTextFormat(Qt.TextFormat.RichText); info.setWordWrap(True)
        lay.addWidget(info)

        def boxed(heading, text_value):
            g = QGroupBox(heading)
            gl = QVBoxLayout(g)
            lab = QLabel(xml_safe(text_value) or "—")
            lab.setWordWrap(True)
            gl.addWidget(lab)
            lay.addWidget(g)

        boxed("Observation / Finding", safe(f["observation_text"]) or safe(f["evidence"]) or safe(f["checklist_pts"]))
        boxed("Evidence recorded during audit", safe(f["evidence"]))
        boxed("Corrective Action", safe(f["corrective_action"]))
        days_label, overdue = aud_days_remaining(f["target_date"], safe(f["closeout_status"]) or safe(f["status"]))
        boxed("Responsible / Target", f"Responsible: {safe(f['responsible']) or '—'}    Target Date: {safe(f['target_date']) or '—'}    {days_label}")

        evidence_box = QGroupBox("Close-out Evidence")
        el = QVBoxLayout(evidence_box)
        evidence_list = QListWidget()
        for a in aud_finding_attachments(finding_id):
            evidence_list.addItem(f"{Path(safe(a['file_path'])).name}  ({safe(a['attachment_type'])})")
        if evidence_list.count() == 0 and safe(f["closeout_evidence"]):
            for name in safe(f["closeout_evidence"]).split(";"):
                if name.strip():
                    evidence_list.addItem(name.strip())
        el.addWidget(evidence_list)
        pending_paths = []
        path_row = QHBoxLayout()
        path_edit = QLineEdit(); path_edit.setPlaceholderText("Selected file path…")
        browse = QPushButton("Browse…")
        add_path = QPushButton("Add File")
        def choose():
            p, _ = QFileDialog.getOpenFileName(d, "Select Close-out Evidence")
            if p: path_edit.setText(p)
        def add_file():
            p = path_edit.text().strip()
            if not p or not Path(p).exists():
                QMessageBox.warning(d, "File Required", "Choose an existing evidence file.")
                return
            pending_paths.append(p)
            evidence_list.addItem(f"[new] {Path(p).name}")
            path_edit.clear()
        browse.clicked.connect(choose); add_path.clicked.connect(add_file)
        path_row.addWidget(path_edit, 1); path_row.addWidget(browse); path_row.addWidget(add_path)
        el.addLayout(path_row)
        lay.addWidget(evidence_box)

        comments = QTextEdit()
        comments.setPlaceholderText("Close-out comments / verification notes")
        comments.setPlainText(safe(f["closeout_comments"]))
        comments.setMaximumHeight(120)
        lay.addWidget(QLabel("<b>Close-out Comments</b>"))
        lay.addWidget(comments)

        submitter = QLineEdit(); submitter.setPlaceholderText("Submitted by (responsible person)")
        submitter.setText(safe(f["closeout_submitted_by"]) or safe(f["responsible"]))
        approver = QLineEdit(); approver.setPlaceholderText("Senior approver name")
        approver.setText(safe(f["closeout_approved_by"]))
        form = QFormLayout()
        form.addRow("Submitted By:", submitter)
        form.addRow("Senior Approver:", approver)
        lay.addLayout(form)

        status_lab = QLabel(f"Current close-out status: <b>{safe(f['closeout_status']) or 'Open'}</b>")
        status_lab.setTextFormat(Qt.TextFormat.RichText)
        lay.addWidget(status_lab)

        scroll.setWidget(body)
        root.addWidget(scroll, 1)
        btns = QHBoxLayout()
        save_btn = QPushButton("Save for later")
        submit_btn = QPushButton("Submit for review")
        approve_btn = QPushButton("Approve")
        reject_btn = QPushButton("Return for Correction")
        close_btn = QPushButton("Close Window")
        save_btn.setToolTip("Save comments and evidence without submitting for review.")
        submit_btn.setToolTip("Submit this close-out for senior review.")
        approve_btn.setToolTip("Senior approve and close the finding.")
        submit_btn.setStyleSheet("font-weight:bold;background:#17365D;color:white;")
        approve_btn.setStyleSheet("font-weight:bold;background:#1B5E20;color:white;")
        for b in (save_btn, submit_btn, approve_btn, reject_btn, close_btn):
            b.setMinimumHeight(36)
            btns.addWidget(b)
        root.addLayout(btns)

        def save_draft_only():
            try:
                for p in list(pending_paths):
                    aud_add_finding_attachment(finding_id, p, submitter.text().strip() or "User", "Close-out Evidence")
                pending_paths.clear()
                db.execute(
                    "UPDATE aud_findings SET closeout_comments=?, updated_at=? WHERE id=?",
                    (comments.toPlainText().strip(), aud_ts(), finding_id)
                )
                QMessageBox.information(d, "Saved for later", "Close-out comments and evidence were saved. You can continue editing later.")
                if on_change: on_change()
            except Exception as exc:
                logging.exception("Close-out draft save failed")
                QMessageBox.critical(d, "Save Failed", str(exc))

        def do_submit():
            if not submitter.text().strip():
                QMessageBox.warning(d, "Submitter Required", "Enter the name of the person submitting the close-out.")
                return
            try:
                aud_submit_closeout(finding_id, comments.toPlainText().strip(), list(pending_paths), submitter.text().strip())
                pending_paths.clear()
                QMessageBox.information(d, "Submitted for Review",
                    "Close-out has been submitted for senior review. It will show as Closed only after approval.")
                if on_change: on_change()
                d.accept()
            except Exception as exc:
                logging.exception("Close-out submit failed")
                QMessageBox.warning(d, "Submit Failed", str(exc))

        def do_approve():
            if not approver.text().strip():
                QMessageBox.warning(d, "Approver Required", "Enter the senior approver name.")
                return
            try:
                # Ensure any pending files are stored first
                for p in list(pending_paths):
                    aud_add_finding_attachment(finding_id, p, approver.text().strip(), "Close-out Evidence")
                pending_paths.clear()
                # Allow direct approve only when already submitted, otherwise force submit path
                cur = db.fetchone("SELECT closeout_status FROM aud_findings WHERE id=?", (finding_id,))
                if safe(cur["closeout_status"]) not in ("Submitted for Review", "Pending Approval"):
                    aud_submit_closeout(finding_id, comments.toPlainText().strip(), [], submitter.text().strip() or approver.text().strip())
                aud_approve_closeout(finding_id, approver.text().strip(), comments.toPlainText().strip())
                QMessageBox.information(d, "Closed",
                    "Finding closed after senior approval. Audit Summary Register percentages have been updated. "
                    "Linked CAPA / observation records were synced where applicable.")
                if on_change: on_change()
                d.accept()
            except Exception as exc:
                logging.exception("Close-out approve failed")
                QMessageBox.warning(d, "Approval Failed", str(exc))

        def do_reject():
            if not approver.text().strip():
                QMessageBox.warning(d, "Reviewer Required", "Enter the reviewer name.")
                return
            try:
                aud_reject_closeout(finding_id, approver.text().strip(), comments.toPlainText().strip())
                QMessageBox.information(d, "Returned", "Close-out returned for further action.")
                if on_change: on_change()
                d.accept()
            except Exception as exc:
                QMessageBox.warning(d, "Return Failed", str(exc))

        save_btn.clicked.connect(save_draft_only)
        submit_btn.clicked.connect(do_submit)
        approve_btn.clicked.connect(do_approve)
        reject_btn.clicked.connect(do_reject)
        close_btn.clicked.connect(d.reject)
        d.exec()

    # ============================================================
    # TAB - AUDIT SUMMARY REGISTER (% open observations per audit)
    # ============================================================
    def _aud_build_summary_register(self, ctx):
        tab = QWidget()
        lay = QVBoxLayout(tab)
        intro = QLabel(
            "Audit Summary Register — percentage of open observations for each audit. "
            "Values update automatically when findings are closed through the Close Out / Findings Register workflow."
        )
        intro.setWordWrap(True)
        lay.addWidget(intro)
        search = QLineEdit(); search.setPlaceholderText("Search audit number, type, location, auditor…")
        lay.addWidget(search)
        headers = ["S.No.", "Audit No.", "Type", "Date", "Location", "Auditor", "Auditee",
                   "Total Findings", "Open", "Closed", "% Open", "% Closed", "Audit Status"]
        table = self._aud_table(headers)
        self._aud_readonly(table)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        lay.addWidget(table, 1)
        refresh_btn = QPushButton("Refresh Summary")
        lay.addWidget(refresh_btn)

        def load():
            needle = search.text().strip().casefold()
            rows = aud_summary_rows()
            filtered = []
            for r in rows:
                blob = " ".join(safe(r[k]) for k in ("number", "audit_type", "location", "auditor", "auditee", "status")).casefold()
                if needle and needle not in blob:
                    continue
                filtered.append(r)
            table.setRowCount(len(filtered))
            for i, r in enumerate(filtered):
                vals = [
                    str(i + 1), safe(r["number"]), safe(r["audit_type"]), safe(r["audit_date"]),
                    safe(r["location"]), safe(r["auditor"]), safe(r["auditee"]) or "—",
                    str(r["total"]), str(r["open_n"]), str(r["closed_n"]),
                    f"{r['pct_open']}%", f"{r['pct_closed']}%", safe(r["status"])
                ]
                for c, v in enumerate(vals):
                    item = self._aud_item(v, r["id"])
                    if c == 10 and r["pct_open"] > 0:
                        item.setForeground(QBrush(QColor("#B00020")))
                    elif c == 11 and r["pct_closed"] == 100 and r["total"] > 0:
                        item.setForeground(QBrush(QColor("#1B5E20")))
                    table.setItem(i, c, item)

        search.textChanged.connect(lambda *_: load())
        refresh_btn.clicked.connect(load)
        ctx["refresh_summary"] = load
        load()
        return tab

    # ============================================================
    # TAB - CUMULATIVE AUDIT FINDINGS REGISTER
    # ============================================================
    def _aud_build_cumulative_register(self, ctx):
        tab = QWidget()
        lay = QVBoxLayout(tab)
        intro = QLabel("Cumulative register of all audit findings across every audit (open and closed).")
        intro.setWordWrap(True)
        lay.addWidget(intro)
        search = QLineEdit(); search.setPlaceholderText("Search finding no., audit no., observation, status, CAPA…")
        lay.addWidget(search)
        headers = ["S.No.", "Audit No.", "Finding No.", "Audit Date", "Type", "Observation / Evidence",
                   "Corrective Action", "Responsible", "Target Date", "Status", "Close-out Status", "CAPA No.", "Auditor"]
        table = self._aud_table(headers)
        self._aud_readonly(table)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
        lay.addWidget(table, 1)
        row = QHBoxLayout()
        refresh_btn = QPushButton("Refresh")
        open_btn = QPushButton("Open Close-out Workspace")
        row.addWidget(refresh_btn); row.addWidget(open_btn); row.addStretch(1)
        lay.addLayout(row)

        def load():
            needle = search.text().strip().casefold()
            rows = aud_cumulative_findings()
            data = []
            for f in rows:
                obs = safe(f["observation_text"]) or safe(f["evidence"]) or safe(f["checklist_pts"])
                blob = " ".join([
                    safe(f["audit_number"]), safe(f["finding_no"]), obs, safe(f["corrective_action"]),
                    safe(f["status"]), safe(f["closeout_status"]), safe(f["capa_number"]), safe(f["responsible"])
                ]).casefold()
                if needle and needle not in blob:
                    continue
                data.append((f, obs))
            table.setRowCount(len(data))
            for i, (f, obs) in enumerate(data):
                vals = [
                    str(i + 1), safe(f["audit_number"]), safe(f["finding_no"]), safe(f["audit_date"]),
                    safe(f["audit_type"]), obs, safe(f["corrective_action"]), safe(f["responsible"]),
                    safe(f["target_date"]), safe(f["status"]), safe(f["closeout_status"]) or "Open",
                    safe(f["capa_number"]) or "—", safe(f["audit_auditor"]) or safe(f["auditor_name"])
                ]
                for c, v in enumerate(vals):
                    table.setItem(i, c, self._aud_item(v, f["id"]))

        def open_workspace():
            r = table.currentRow()
            item = table.item(r, 0) if r >= 0 else None
            if item is None:
                QMessageBox.information(tab, "Select Finding", "Select a finding row first.")
                return
            self._aud_closeout_dialog(int(item.data(Qt.ItemDataRole.UserRole)), on_change=load)

        table.cellDoubleClicked.connect(lambda *_: open_workspace())
        open_btn.clicked.connect(open_workspace)
        search.textChanged.connect(lambda *_: load())
        refresh_btn.clicked.connect(load)
        ctx["refresh_cumulative"] = load
        load()
        return tab



    def _aud_build_saved_for_later(self, ctx):
        """Audits saved for later (Scheduled / In Progress / Draft) — continue editing and submit."""
        tab = QWidget()
        lay = QVBoxLayout(tab)
        intro = QLabel(
            "Audits that were <b>Saved for later</b> appear here. "
            "Select a row and click <b>Continue / Edit</b> to open the Audit checklist, then use "
            "<b>Save for later</b> or <b>Submit</b> when ready."
        )
        intro.setWordWrap(True)
        lay.addWidget(intro)
        search = QLineEdit(); search.setPlaceholderText("Search audit number, type, location, auditor, status…")
        lay.addWidget(search)
        table = self._aud_table(["Audit Number", "Audit Type", "Date", "Location", "Auditor", "Status", "Procedures"])
        self._aud_readonly(table)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        lay.addWidget(table, 1)
        row = QHBoxLayout()
        continue_btn = QPushButton("Continue / Edit in Audit tab")
        continue_btn.setMinimumHeight(36)
        continue_btn.setStyleSheet("font-weight:bold;background:#17365D;color:white;")
        refresh_btn = QPushButton("Refresh")
        row.addWidget(continue_btn)
        row.addWidget(refresh_btn)
        row.addStretch(1)
        lay.addLayout(row)

        def load():
            needle = search.text().strip().casefold()
            # Draft + Scheduled (save for later from schedule) + In Progress (save for later from Audit)
            rows = aud_list(("Draft", "Scheduled", "In Progress"))
            data = []
            for a in rows:
                blob = " ".join(safe(a[k]) for k in ("number", "audit_type", "audit_date", "location", "auditor", "status")).casefold()
                if needle and needle not in blob:
                    continue
                data.append(a)
            table.setRowCount(len(data))
            for i, a in enumerate(data):
                procs = ", ".join(aud_parse_procedures(a["procedures"])) or "—"
                vals = [safe(a["number"]), safe(a["audit_type"]), safe(a["audit_date"]),
                        safe(a["location"]), safe(a["auditor"]), safe(a["status"]), procs]
                for c, v in enumerate(vals):
                    table.setItem(i, c, self._aud_item(v, a["id"]))
            table.resizeRowsToContents()

        def continue_edit():
            r = table.currentRow()
            item = table.item(r, 0) if r >= 0 else None
            if item is None:
                QMessageBox.information(tab, "Select Audit", "Select a saved audit first.")
                return
            audit_id = item.data(Qt.ItemDataRole.UserRole)
            audit = aud_get(int(audit_id))
            if not audit:
                return
            status = safe(audit["status"])
            # Draft / Scheduled → open Schedule for edit; In Progress / Submitted → open Audit checklist
            if status in ("Draft", "Scheduled"):
                try:
                    for i in range(ctx["tabs"].count()):
                        if ctx["tabs"].tabText(i) == "Schedule Audit":
                            ctx["tabs"].setCurrentIndex(i)
                            break
                    if "refresh_schedule" in ctx:
                        ctx["refresh_schedule"]()
                    QMessageBox.information(
                        tab, "Continue in Schedule",
                        f"Audit {safe(audit['number'])} is still being scheduled ({status}). "
                        "Open it from the Schedule Audit list using Edit Selected."
                    )
                except Exception:
                    logging.exception("Continue schedule failed")
                return
            if "open_checklist" in ctx:
                ctx["open_checklist"](int(audit_id))
            if "refresh_checklist_options" in ctx:
                ctx["refresh_checklist_options"]()
            try:
                for i in range(ctx["tabs"].count()):
                    if ctx["tabs"].tabText(i) == "Audit":
                        ctx["tabs"].setCurrentIndex(i)
                        break
            except Exception:
                pass

        table.cellDoubleClicked.connect(lambda *_: continue_edit())
        continue_btn.clicked.connect(continue_edit)
        search.textChanged.connect(lambda *_: load())
        refresh_btn.clicked.connect(load)
        ctx["refresh_saved_for_later"] = load
        load()
        return tab


    def _aud_build_closeout_section(self, ctx):
        """Close Out: audit list with open observation counts → findings → close-out workspace."""
        tab = QWidget()
        lay = QVBoxLayout(tab)
        intro = QLabel(
            "Close Out — select an audit to see its open observations / findings. "
            "Click a finding to open the Close Out workspace (evidence, Save for later, Submit for review, Approve)."
        )
        intro.setWordWrap(True)
        lay.addWidget(intro)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left: audits with open finding counts
        left = QWidget(); left_lay = QVBoxLayout(left); left_lay.setContentsMargins(0, 0, 0, 0)
        left_lay.addWidget(QLabel("<b>Audits</b>"))
        search = QLineEdit(); search.setPlaceholderText("Search audit number, type, location…")
        left_lay.addWidget(search)
        audit_table = self._aud_table(["Audit No.", "Type", "Date", "Location", "Open Observations", "Status"])
        self._aud_readonly(audit_table)
        audit_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        left_lay.addWidget(audit_table, 1)
        splitter.addWidget(left)

        # Right: findings for selected audit
        right = QWidget(); right_lay = QVBoxLayout(right); right_lay.setContentsMargins(0, 0, 0, 0)
        right_title = QLabel("<b>Findings for selected audit</b>")
        right_lay.addWidget(right_title)
        finding_table = self._aud_table(
            ["Finding No.", "Observation / Finding", "Corrective Action", "Responsible",
             "Target Date", "Days Remaining", "Close-out Status", "CAPA No."]
        )
        self._aud_readonly(finding_table)
        finding_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        finding_table.horizontalHeader().setStretchLastSection(True)
        right_lay.addWidget(finding_table, 1)
        open_btn = QPushButton("Open Close Out Workspace")
        open_btn.setMinimumHeight(36)
        open_btn.setStyleSheet("font-weight:bold;background:#17365D;color:white;")
        right_lay.addWidget(open_btn)
        splitter.addWidget(right)
        splitter.setSizes([420, 680])
        lay.addWidget(splitter, 1)

        state = {"audit_id": None}

        def open_count_for(audit_id):
            rows = db.fetchall(
                "SELECT closeout_status, status FROM aud_findings WHERE audit_id=?", (audit_id,)
            )
            open_n = 0
            for r in rows:
                st = (safe(r["closeout_status"]) or safe(r["status"]) or "Open").lower()
                if st not in ("closed", "approved"):
                    open_n += 1
            return open_n

        def load_audits():
            needle = search.text().strip().casefold()
            audits = aud_list(AUD_CONDUCTED_STATUSES + ("Completed",))
            # de-dupe by id
            seen = set(); records = []
            for a in audits:
                if a["id"] in seen:
                    continue
                seen.add(a["id"])
                records.append(a)
            data = []
            for a in records:
                blob = " ".join(safe(a[k]) for k in ("number", "audit_type", "audit_date", "location", "status")).casefold()
                if needle and needle not in blob:
                    continue
                data.append((a, open_count_for(int(a["id"]))))
            audit_table.setRowCount(len(data))
            for i, (a, open_n) in enumerate(data):
                vals = [safe(a["number"]), safe(a["audit_type"]), safe(a["audit_date"]),
                        safe(a["location"]), str(open_n), safe(a["status"])]
                for c, v in enumerate(vals):
                    item = self._aud_item(v, a["id"])
                    if c == 4 and open_n > 0:
                        item.setForeground(QBrush(QColor("#B00020")))
                        item.setFont(QFont("Arial", 9, QFont.Weight.Bold))
                    audit_table.setItem(i, c, item)
            audit_table.resizeRowsToContents()

        def load_findings(audit_id):
            state["audit_id"] = audit_id
            finding_table.setRowCount(0)
            if not audit_id:
                right_title.setText("<b>Findings for selected audit</b>")
                return
            a = aud_get(audit_id)
            right_title.setText(f"<b>Findings — {safe(a['number']) if a else ''}</b>")
            rows = db.fetchall(
                "SELECT * FROM aud_findings WHERE audit_id=? ORDER BY seq, id", (audit_id,)
            )
            finding_table.setRowCount(len(rows))
            for i, f in enumerate(rows):
                obs = safe(f["observation_text"]) or safe(f["evidence"]) or safe(f["checklist_pts"])
                status = safe(f["closeout_status"]) or safe(f["status"]) or "Open"
                days_label, overdue = aud_days_remaining(f["target_date"], status)
                vals = [
                    safe(f["finding_no"]), obs, safe(f["corrective_action"]), safe(f["responsible"]),
                    safe(f["target_date"]), days_label, status, safe(f["capa_number"]) or "—"
                ]
                for c, v in enumerate(vals):
                    item = self._aud_item(v, f["id"])
                    if overdue and c == 5:
                        item.setForeground(QBrush(QColor("#B00020")))
                        item.setFont(QFont("Arial", 9, QFont.Weight.Bold))
                    finding_table.setItem(i, c, item)
            finding_table.resizeRowsToContents()

        def on_audit_clicked(row, col=None):
            item = audit_table.item(row, 0)
            if item is None:
                return
            load_findings(item.data(Qt.ItemDataRole.UserRole))

        def open_workspace():
            r = finding_table.currentRow()
            item = finding_table.item(r, 0) if r >= 0 else None
            if item is None:
                QMessageBox.information(tab, "Select Finding", "Select a finding row first.")
                return
            fid = item.data(Qt.ItemDataRole.UserRole)
            def after():
                if state["audit_id"]:
                    load_findings(state["audit_id"])
                load_audits()
                if "refresh_findings_register" in ctx:
                    ctx["refresh_findings_register"]()
                if "refresh_summary" in ctx:
                    ctx["refresh_summary"]()
            self._aud_closeout_dialog(int(fid), on_change=after)

        audit_table.cellClicked.connect(on_audit_clicked)
        audit_table.cellDoubleClicked.connect(on_audit_clicked)
        finding_table.cellDoubleClicked.connect(lambda *_: open_workspace())
        open_btn.clicked.connect(open_workspace)
        search.textChanged.connect(lambda *_: load_audits())
        ctx["refresh_closeout_section"] = load_audits
        load_audits()
        return tab


    def _aud_build_notifications(self, ctx):
        tab = QWidget()
        lay = QVBoxLayout(tab)
        state = {"audit_id": None, "loading": False}
        top = QHBoxLayout()
        audit_combo = QComboBox()
        audit_combo.setMinimumWidth(380)
        top.addWidget(QLabel("<b>Select Scheduled Audit:</b>"))
        top.addWidget(audit_combo)
        top.addStretch(1)
        lay.addLayout(top)

        box = QGroupBox("Audit Notification Details")
        form = QFormLayout(box)
        doc_label = QLabel("—")
        doc_label.setStyleSheet("font-weight:bold;color:#17365D;")
        type_label = QLabel("—")
        to_edit = QLineEdit()
        from_edit = QLineEdit()
        subject_edit = QLineEdit()
        scope_edit = QTextEdit()
        scope_edit.setMaximumHeight(90)
        date_edit = self._aud_date_edit()
        time_edit = QLineEdit()
        time_edit.setPlaceholderText("e.g. 09:00 - 12:00")
        venue_edit = QLineEdit()
        auditee_edit = QLineEdit()
        team_edit = QLineEdit()
        refs_edit = QLineEdit()
        remarks_edit = QTextEdit()
        remarks_edit.setMaximumHeight(70)
        form.addRow("Document Number:", doc_label)
        form.addRow("Audit Type:", type_label)
        form.addRow("To *:", to_edit)
        form.addRow("From *:", from_edit)
        form.addRow("Subject *:", subject_edit)
        form.addRow("Scope *:", scope_edit)
        form.addRow("Audit Date:", date_edit)
        form.addRow("Time:", time_edit)
        form.addRow("Venue / Location:", venue_edit)
        form.addRow("Auditee:", auditee_edit)
        form.addRow("Audit Team / Auditor:", team_edit)
        form.addRow("Procedures / Documents for Review:", refs_edit)
        form.addRow("Remarks:", remarks_edit)
        lay.addWidget(box)
        btns = QHBoxLayout()
        save_btn = QPushButton("Save")
        export_btn = QPushButton("Export to Word")
        for b in (save_btn, export_btn):
            b.setMinimumHeight(38)
            b.setMinimumWidth(150)
        save_btn.setStyleSheet("font-weight:bold;background:#17365D;color:white;")
        btns.addStretch(1)
        btns.addWidget(save_btn)
        btns.addWidget(export_btn)
        lay.addLayout(btns)
        lay.addStretch(1)

        fields = [to_edit, from_edit, subject_edit, scope_edit, date_edit, time_edit, venue_edit, auditee_edit,
                  team_edit, refs_edit, remarks_edit]

        def clear_fields():
            for w in (to_edit, from_edit, subject_edit, scope_edit, time_edit, venue_edit, auditee_edit, team_edit, refs_edit, remarks_edit):
                w.clear()
            date_edit.setDate(QDate.currentDate())
            doc_label.setText("—")
            type_label.setText("—")

        def load_audit(*_):
            audit_id = audit_combo.currentData()
            state["audit_id"] = audit_id
            clear_fields()
            for w in fields:
                w.setEnabled(audit_id is not None)
            save_btn.setEnabled(audit_id is not None)
            export_btn.setEnabled(False)
            if audit_id is None:
                return
            audit = aud_get(audit_id)
            note = aud_get_notification(audit_id)
            type_label.setText(safe(audit["audit_type"]))
            if note:
                doc_label.setText(safe(note["doc_number"]))
                to_edit.setText(safe(note["to_party"]))
                from_edit.setText(safe(note["from_party"]))
                subject_edit.setText(safe(note["subject"]))
                scope_edit.setPlainText(safe(note["scope"]))
                parsed = QDate.fromString(safe(note["audit_date"]), "yyyy-MM-dd")
                date_edit.setDate(parsed if parsed.isValid() else QDate.currentDate())
                time_edit.setText(safe(note["audit_time"]))
                venue_edit.setText(safe(note["venue"]))
                auditee_edit.setText(safe(note["auditee"]))
                team_edit.setText(safe(note["audit_team"]))
                refs_edit.setText(safe(note["reference_docs"]))
                remarks_edit.setPlainText(safe(note["remarks"]))
                export_btn.setEnabled(True)
            else:
                doc_label.setText("Assigned on save")
                subject_edit.setText(f"Notification of {safe(audit['audit_type'])} Audit — {safe(audit['number'])}")
                scope_edit.setPlainText(safe(audit["scope"]))
                parsed = QDate.fromString(safe(audit["audit_date"]), "yyyy-MM-dd")
                date_edit.setDate(parsed if parsed.isValid() else QDate.currentDate())
                venue_edit.setText(safe(audit["location"]))
                auditee_edit.setText(safe(audit["auditee"]))
                team_edit.setText(safe(audit["auditor"]))
                refs_edit.setText(", ".join(aud_parse_procedures(audit["procedures"])))

        def fill_combo(select_id=None):
            current = select_id if select_id is not None else audit_combo.currentData()
            audit_combo.blockSignals(True)
            audit_combo.clear()
            audit_combo.addItem("Select a scheduled audit...", None)
            for a in aud_list(AUD_SCHEDULED_STATUSES):
                audit_combo.addItem(f"{safe(a['number'])} — {safe(a['audit_type'])} — {safe(a['audit_date'])}", int(a["id"]))
            idx = audit_combo.findData(current) if current is not None else 0
            audit_combo.setCurrentIndex(idx if idx >= 0 else 0)
            audit_combo.blockSignals(False)

        def refresh_options():
            previous = state["audit_id"]
            fill_combo()
            if audit_combo.currentData() != previous:
                load_audit()

        def save():
            if state["audit_id"] is None:
                return
            data = {"to_party": to_edit.text(), "from_party": from_edit.text(), "subject": subject_edit.text(),
                    "scope": scope_edit.toPlainText(), "audit_date": date_edit.date().toString("yyyy-MM-dd"),
                    "audit_time": time_edit.text(), "venue": venue_edit.text(), "auditee": auditee_edit.text(),
                    "audit_team": team_edit.text(), "reference_docs": refs_edit.text(), "remarks": remarks_edit.toPlainText()}
            try:
                aud_save_notification(state["audit_id"], data)
            except ValueError as exc:
                QMessageBox.warning(tab, "Audit Notification", str(exc))
                return
            except Exception as exc:
                logging.exception("Audit notification save failed")
                QMessageBox.critical(tab, "Audit Notification", f"Could not save the notification.\n\n{exc}")
                return
            note = aud_get_notification(state["audit_id"])
            doc_label.setText(safe(note["doc_number"]))
            export_btn.setEnabled(True)
            QMessageBox.information(tab, "Saved", f"Audit Notification {safe(note['doc_number'])} saved. You can now export it to Word.")

        def export_word():
            if state["audit_id"] is None:
                return
            note = aud_get_notification(state["audit_id"])
            if not note:
                QMessageBox.information(tab, "Save First", "Save the notification before exporting.")
                return
            path, _ = QFileDialog.getSaveFileName(tab, "Export Audit Notification",
                                                  str(REPORT_DIR / f"{safe(note['doc_number'])}_Audit_Notification.docx"),
                                                  "Word Document (*.docx)")
            if not path:
                return
            try:
                aud_export_notification_word(state["audit_id"], path)
                self.show_export_success(path)
            except Exception as exc:
                logging.exception("Audit notification export failed")
                QMessageBox.critical(tab, "Export Failed", f"Unable to create the Word document.\n\n{exc}")

        audit_combo.currentIndexChanged.connect(load_audit)
        save_btn.clicked.connect(save)
        export_btn.clicked.connect(export_word)
        fill_combo()
        load_audit()
        ctx["refresh_notification_options"] = refresh_options
        return tab

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
        """Training Management — Enrol · Employee Register · Matrix Grid · Add Course · Tracking.
        Other application modules are not modified.
        """
        w, layout = self.page("Training Management")
        intro = QLabel(
            "1 Enrol Employee  ·  2 Employee Register  ·  3 Training Matrix Grid  ·  "
            "Compliance Dashboard  ·  Enrolment  ·  Employee Register  ·  Training Matrix  ·  Course Catalogue  ·  Tracking Register. "
            "Training-only changes: certificate evidence, verification status, calendar-month expiry and record history."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet("color:#17365D; font-weight:600; padding:6px 8px; background:#E8F1FB; border-radius:4px;")
        layout.addWidget(intro)

        tabs = QTabWidget()
        tabs.setDocumentMode(True)
        tabs.setUsesScrollButtons(True)
        tabs.tabBar().setExpanding(False)
        tabs.tabBar().setMovable(False)
        tabs.setMinimumHeight(360)
        layout.addWidget(tabs, 1)

        MATRIX_DESIGNATIONS = [
            "Managing Director / MD", "CEO", "General Manager", "Project Director", "Project Manager",
            "Construction Manager", "Site Manager", "Operations Manager", "HSE Manager", "QA/QC Manager",
            "Superintendent", "Supervisor", "Foreman", "Chargehand", "Team Leader",
            "HSE Officer", "HSE Supervisor", "HSE Advisor", "Safety Representative",
            "Site Engineer", "Civil Engineer", "Mechanical Engineer", "Electrical Engineer", "Instrument Engineer",
            "Process Engineer", "QA/QC Engineer", "Welding Engineer", "Planning Engineer",
            "Mechanical Technician", "Electrical Technician", "Instrument Technician", "NDE Technician",
            "Welder", "Pipe Fitter", "Fabricator", "Plater", "Grinder", "Painter / Blaster",
            "Rigger", "Banksman / Signaler", "Lifting Supervisor", "Crane Operator",
            "Forklift Operator", "MEWP / Manlift Operator", "Heavy Equipment Operator", "Excavator Operator",
            "Scaffolder", "Scaffold Inspector", "Carpenter", "Mason / Bricklayer", "Steel Fixer",
            "Electrician", "Cable Jointer", "Lineman", "Insulator", "Millwright",
            "Driver", "Heavy Vehicle Driver", "Light Vehicle Driver",
            "Storekeeper", "Warehouse Supervisor", "Admin / Office Staff", "Document Controller",
            "Security Guard", "Helper", "General Worker", "Labourer", "General Helper", "Visitor", "Other"
        ]
        READY_COURSES = [
            ("HSE / Company Induction", 12),
            ("Site / Project HSE Induction", 12),
            ("Visitor Induction", 12),
            ("OSHA 10 Construction (or equivalent)", 60),
            ("OSHA 30 Construction (or equivalent)", 60),
            ("IOSH Working Safely", 36),
            ("IOSH Managing Safely", 36),
            ("Permit to Work (PTW)", 24),
            ("Work at Height / Fall Protection", 24),
            ("Scaffolding Safety (User)", 24),
            ("Scaffold Erector / Inspector", 24),
            ("Confined Space Entry", 24),
            ("Hot Work Safety", 24),
            ("H2S Awareness", 24),
            ("Fire Safety Awareness", 12),
            ("Basic Fire Fighting / Fire Watch", 24),
            ("First Aid / CPR", 36),
            ("Heat Stress Awareness", 12),
            ("Manual Handling", 24),
            ("Hand and Power Tools Safety", 24),
            ("PPE Awareness", 12),
            ("Electrical Safety", 24),
            ("Lockout / Tagout (LOTO)", 24),
            ("Lifting Operations Safety", 24),
            ("Rigging and Slinging", 24),
            ("Banksman / Signaler", 24),
            ("Excavation & Trenching Safety", 24),
            ("Chemical / HAZCOM / SDS", 24),
            ("Job Safety Analysis (JSA / JHA)", 24),
            ("Hazard Identification and Risk Assessment", 24),
            ("Incident Reporting and Investigation", 36),
            ("Emergency Response and Evacuation", 12),
            ("Environmental Awareness / Spill Response", 24),
            ("Defensive Driving", 36),
            ("Abrasive Blasting / Painting Safety", 24),
            ("Welding & Cutting Safety", 24),
            ("Radiography / NDT Safety", 24),
            ("Supervisor HSE Leadership", 36),
            ("HSE for Managers", 36),
            ("Internal HSE Auditor", 36),
        ]
        FIELD_BASE = [
            ("HSE / Company Induction", "Mandatory"),
            ("Site / Project HSE Induction", "Mandatory"),
            ("PPE Awareness", "Mandatory"),
            ("Fire Safety Awareness", "Mandatory"),
            ("Emergency Response and Evacuation", "Mandatory"),
            ("Manual Handling", "Mandatory"),
            ("Hand and Power Tools Safety", "Mandatory"),
            ("Hazard Identification and Risk Assessment", "Mandatory"),
            ("Job Safety Analysis (JSA / JHA)", "Mandatory"),
            ("Heat Stress Awareness", "Mandatory"),
            ("Incident Reporting and Investigation", "Mandatory"),
        ]
        TRADE_SEED = {
            "Scaffolder": [("Work at Height / Fall Protection", "Mandatory"), ("Scaffolding Safety (User)", "Mandatory"), ("Scaffold Erector / Inspector", "Mandatory"), ("H2S Awareness", "Mandatory"), ("Permit to Work (PTW)", "Mandatory")],
            "Scaffold Inspector": [("Work at Height / Fall Protection", "Mandatory"), ("Scaffolding Safety (User)", "Mandatory"), ("Scaffold Erector / Inspector", "Mandatory"), ("Permit to Work (PTW)", "Mandatory")],
            "Welder": [("Hot Work Safety", "Mandatory"), ("Welding & Cutting Safety", "Mandatory"), ("Work at Height / Fall Protection", "Mandatory"), ("Confined Space Entry", "Before Job"), ("Permit to Work (PTW)", "Mandatory"), ("Basic Fire Fighting / Fire Watch", "Mandatory")],
            "Pipe Fitter": [("Hot Work Safety", "Before Job"), ("Confined Space Entry", "Before Job"), ("Work at Height / Fall Protection", "Mandatory"), ("Permit to Work (PTW)", "Mandatory"), ("Lifting Operations Safety", "Before Job")],
            "Rigger": [("Lifting Operations Safety", "Mandatory"), ("Rigging and Slinging", "Mandatory"), ("Banksman / Signaler", "Mandatory"), ("Work at Height / Fall Protection", "Mandatory"), ("Permit to Work (PTW)", "Mandatory")],
            "Banksman / Signaler": [("Lifting Operations Safety", "Mandatory"), ("Banksman / Signaler", "Mandatory"), ("Rigging and Slinging", "Before Job")],
            "Lifting Supervisor": [("Lifting Operations Safety", "Mandatory"), ("Rigging and Slinging", "Mandatory"), ("Banksman / Signaler", "Mandatory"), ("Supervisor HSE Leadership", "Mandatory")],
            "Crane Operator": [("Lifting Operations Safety", "Mandatory"), ("Permit to Work (PTW)", "Mandatory")],
            "Electrician": [("Electrical Safety", "Mandatory"), ("Lockout / Tagout (LOTO)", "Mandatory"), ("Permit to Work (PTW)", "Mandatory"), ("Work at Height / Fall Protection", "Mandatory"), ("Confined Space Entry", "Before Job")],
            "Heavy Equipment Operator": [("Excavation & Trenching Safety", "Before Job"), ("Permit to Work (PTW)", "Mandatory"), ("Defensive Driving", "Before Job")],
            "Excavator Operator": [("Excavation & Trenching Safety", "Mandatory"), ("Permit to Work (PTW)", "Mandatory")],
            "Painter / Blaster": [("Abrasive Blasting / Painting Safety", "Mandatory"), ("Work at Height / Fall Protection", "Mandatory"), ("Chemical / HAZCOM / SDS", "Mandatory"), ("Permit to Work (PTW)", "Mandatory")],
            "HSE Officer": [("Permit to Work (PTW)", "Mandatory"), ("Work at Height / Fall Protection", "Mandatory"), ("Confined Space Entry", "Mandatory"), ("First Aid / CPR", "Mandatory"), ("Supervisor HSE Leadership", "Mandatory"), ("IOSH Managing Safely", "Optional")],
            "HSE Supervisor": [("Permit to Work (PTW)", "Mandatory"), ("Work at Height / Fall Protection", "Mandatory"), ("Confined Space Entry", "Mandatory"), ("First Aid / CPR", "Mandatory"), ("Supervisor HSE Leadership", "Mandatory")],
            "HSE Manager": [("HSE for Managers", "Mandatory"), ("IOSH Managing Safely", "Mandatory"), ("Internal HSE Auditor", "Mandatory"), ("Permit to Work (PTW)", "Mandatory")],
            "Supervisor": [("Permit to Work (PTW)", "Mandatory"), ("Supervisor HSE Leadership", "Mandatory"), ("Work at Height / Fall Protection", "Mandatory")],
            "Foreman": [("Permit to Work (PTW)", "Mandatory"), ("Supervisor HSE Leadership", "Mandatory"), ("Work at Height / Fall Protection", "Mandatory")],
            "Site Manager": [("HSE for Managers", "Mandatory"), ("IOSH Managing Safely", "Optional"), ("Permit to Work (PTW)", "Mandatory")],
            "Project Manager": [("HSE for Managers", "Mandatory"), ("IOSH Managing Safely", "Optional")],
            "Site Engineer": [("Permit to Work (PTW)", "Mandatory"), ("Work at Height / Fall Protection", "Mandatory"), ("Job Safety Analysis (JSA / JHA)", "Mandatory")],
            "Helper": [("Work at Height / Fall Protection", "Before Job"), ("Permit to Work (PTW)", "Before Job")],
            "General Helper": [("Work at Height / Fall Protection", "Before Job"), ("Permit to Work (PTW)", "Before Job")],
            "General Worker": [("Work at Height / Fall Protection", "Before Job")],
            "Labourer": [("Work at Height / Fall Protection", "Before Job")],
            "Visitor": [("Visitor Induction", "Mandatory"), ("PPE Awareness", "Mandatory"), ("Emergency Response and Evacuation", "Mandatory")],
            "Admin / Office Staff": [("HSE / Company Induction", "Mandatory"), ("Fire Safety Awareness", "Mandatory"), ("Emergency Response and Evacuation", "Mandatory")],
        }
        MATRIX_REQ_OPTIONS = ["Not Applicable", "Mandatory", "Optional", "Before Job"]
        MATRIX_REQ_SHORT = {"Not Applicable": "—", "Mandatory": "M", "Optional": "O", "Before Job": "B"}
        MATRIX_REQ_COLORS = {
            "Not Applicable": ("#777777", None),
            "Mandatory": ("#FFFFFF", "#B42318"),
            "Optional": ("#176B3A", "#DDF5E5"),
            "Before Job": ("#9A6700", "#FFF2CC"),
        }
        grid_state = {"courses": [], "validity": {}, "cells": {}, "dirty": False}

        def _style_req(item, req):
            req = safe(req) or "Not Applicable"
            if req not in MATRIX_REQ_OPTIONS:
                req = "Not Applicable"
            item.setText(MATRIX_REQ_SHORT.get(req, "—"))
            item.setData(Qt.ItemDataRole.UserRole, req)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setToolTip(req + " — double-click to change")
            fg, bg = MATRIX_REQ_COLORS.get(req, ("#000000", None))
            item.setForeground(QBrush(QColor(fg)))
            if bg:
                item.setBackground(QBrush(QColor(bg)))
            f = item.font(); f.setBold(req != "Not Applicable"); item.setFont(f)

        def load_grid_state():
            grid_state["courses"] = []
            grid_state["validity"] = {}
            grid_state["cells"] = {}
            rows = db.fetchall(
                "SELECT course_name, designation, requirement, validity_months "
                "FROM training_matrix_cells ORDER BY course_name COLLATE NOCASE, designation COLLATE NOCASE"
            )
            for row in rows:
                cname = safe(row["course_name"]).strip()
                desig = safe(row["designation"]).strip()
                if not cname:
                    continue
                if cname not in grid_state["courses"]:
                    grid_state["courses"].append(cname)
                if int(row["validity_months"] or 0):
                    grid_state["validity"][cname] = int(row["validity_months"] or 0)
                if desig:
                    grid_state["cells"][(cname, desig)] = safe(row["requirement"]) or "Not Applicable"
            for cname, months in READY_COURSES:
                grid_state["validity"].setdefault(cname, months)
            if not grid_state["courses"]:
                seed_ready_matrix(force=True)
            grid_state["dirty"] = False

        def seed_ready_matrix(force=False):
            for cname, months in READY_COURSES:
                if cname not in grid_state["courses"]:
                    grid_state["courses"].append(cname)
                grid_state["validity"][cname] = months
            office = {"Managing Director / MD", "CEO", "General Manager", "Admin / Office Staff", "Document Controller", "Visitor"}
            for desig in MATRIX_DESIGNATIONS:
                seeds = []
                if desig not in office:
                    seeds.extend(FIELD_BASE)
                seeds.extend(TRADE_SEED.get(desig, []))
                for key, items in TRADE_SEED.items():
                    if key != desig and key.lower() in desig.lower():
                        seeds.extend(items)
                for cname, req in seeds:
                    if cname not in grid_state["courses"]:
                        continue
                    k = (cname, desig)
                    if force or grid_state["cells"].get(k, "Not Applicable") in ("", "Not Applicable"):
                        grid_state["cells"][k] = req
            grid_state["dirty"] = True

        def save_grid_state():
            # Preserve the previous matrix configuration before applying edits.
            try:
                db.execute("""INSERT INTO training_matrix_history
                    (course_name, designation, requirement, validity_months, notes, updated_at, archive_reason)
                    SELECT course_name, designation, requirement, validity_months, notes, updated_at, 'Matrix revised'
                    FROM training_matrix_cells""")
            except Exception:
                logging.exception("Training matrix history snapshot failed")
                raise
            db.execute("DELETE FROM training_matrix_cells")
            for cname in grid_state["courses"]:
                months = int(grid_state["validity"].get(cname, 0) or 0)
                for desig in MATRIX_DESIGNATIONS:
                    req = grid_state["cells"].get((cname, desig), "Not Applicable")
                    db.execute(
                        "INSERT INTO training_matrix_cells (course_name, designation, requirement, validity_months, notes, updated_at) VALUES (?,?,?,?,?,?)",
                        (cname, desig, req, months, "", today())
                    )
                mand = [d for d in MATRIX_DESIGNATIONS if grid_state["cells"].get((cname, d)) in ("Mandatory", "Before Job")]
                opt = [d for d in MATRIX_DESIGNATIONS if grid_state["cells"].get((cname, d)) == "Optional"]
                trade_csv = ", ".join(mand + opt)
                existing = db.fetchone("SELECT id FROM training_courses WHERE lower(course_name)=lower(?) AND active=1 LIMIT 1", (cname,))
                if existing:
                    db.execute(
                        "UPDATE training_courses SET designation=?, validity_months=?, mandatory=?, requirement_source=? WHERE id=?",
                        (trade_csv, months, 1 if mand else 0, "Matrix Grid", existing["id"])
                    )
                else:
                    db.execute(
                        "INSERT INTO training_courses (course_name, designation, department, mandatory, provider_type, validity_months, requirement_source, active) VALUES (?,?,?,?,?,?,?,1)",
                        (cname, trade_csv, "", 1 if mand else 0, "Internal", months, "Matrix Grid")
                    )
            grid_state["dirty"] = False

        def requirements_for_trade(trade):
            trade = safe(trade).strip()
            out = []
            for cname in grid_state["courses"]:
                req = grid_state["cells"].get((cname, trade), "Not Applicable")
                if req == "Not Applicable":
                    for desig in MATRIX_DESIGNATIONS:
                        if trade and (desig.lower() == trade.lower() or trade.lower() in desig.lower()):
                            req = grid_state["cells"].get((cname, desig), req)
                            if req != "Not Applicable":
                                break
                if req in ("Mandatory", "Optional", "Before Job"):
                    out.append((cname, req, int(grid_state["validity"].get(cname, 0) or 0)))
            return out

        def latest_record(emp_db_id, course_name):
            course = db.fetchone("SELECT id FROM training_courses WHERE lower(course_name)=lower(?) ORDER BY active DESC, id DESC LIMIT 1", (course_name,))
            if not course:
                return None
            return db.fetchone(
                "SELECT * FROM training_records WHERE employee_db_id=? AND course_id=? ORDER BY training_date DESC, id DESC LIMIT 1",
                (emp_db_id, course["id"])
            )

        def ensure_course(course_name, validity):
            row = db.fetchone("SELECT id FROM training_courses WHERE lower(course_name)=lower(?) AND active=1 LIMIT 1", (course_name,))
            if row:
                db.execute("UPDATE training_courses SET validity_months=? WHERE id=?", (validity, row["id"]))
                return int(row["id"])
            cur = db.execute(
                "INSERT INTO training_courses (course_name, designation, department, mandatory, provider_type, validity_months, requirement_source, active) VALUES (?,?,?,?,?,?,?,1)",
                (course_name, "", "", 1, "Internal", validity, "Matrix Grid")
            )
            return int(cur.lastrowid)

        def add_calendar_months(source_date, months):
            """Add true calendar months, clamping month-end dates safely."""
            months = int(months or 0)
            month_index = source_date.month - 1 + months
            year = source_date.year + month_index // 12
            month = month_index % 12 + 1
            day = min(source_date.day, calendar.monthrange(year, month)[1])
            return date(year, month, day)

        def status_of(record, validity_months):
            if not record:
                return "Not Completed", None, None
            try:
                if "verification_status" in record.keys() and safe(record["verification_status"]).strip().lower() not in ("", "verified", "approved"):
                    return "Pending Verification", safe(record["training_date"]), safe(record["expiry_date"])
            except Exception:
                pass
            try:
                assessment_result = safe(record["assessment_result"]).strip().lower()
                if assessment_result == "fail":
                    return "Failed Assessment", safe(record["training_date"]), safe(record["expiry_date"])
                if assessment_result == "pending":
                    return "Assessment Pending", safe(record["training_date"]), safe(record["expiry_date"])
                course_meta = db.fetchone(
                    "SELECT certificate_required, assessment_required FROM training_courses WHERE id=?",
                    (record["course_id"],)
                )
                if course_meta and int(course_meta["certificate_required"] or 0) and (not safe(record["certificate_no"]) or not safe(record["attachment_path"])):
                    return "Certificate Missing", safe(record["training_date"]), safe(record["expiry_date"])
                if course_meta and int(course_meta["assessment_required"] or 0) and assessment_result != "pass":
                    return "Assessment Pending", safe(record["training_date"]), safe(record["expiry_date"])
            except Exception:
                pass
            expiry = safe(record["expiry_date"])
            train = safe(record["training_date"])
            if not expiry and train and int(validity_months or 0) > 0:
                try:
                    expiry = add_calendar_months(date.fromisoformat(train), int(validity_months)).isoformat()
                except Exception:
                    return "Invalid Date", train, ""
            # A zero-validity course is treated as non-expiring only when configured
            # with zero months. Legacy courses therefore remain compatible.
            if not expiry:
                return ("Valid", train, "") if int(validity_months or 0) == 0 else ("Invalid Date", train, "")
            try:
                days = (date.fromisoformat(expiry) - date.today()).days
            except Exception:
                return "Invalid Date", train, expiry
            if days < 0:
                return "Expired", train, expiry
            if days <= 30:
                return "Expiring Soon", train, expiry
            return "Valid", train, expiry

        def color_status(status):
            if status in ("Expired", "Not Completed", "Invalid Date", "Failed Assessment", "Certificate Missing"):
                return QColor("#FDE2E2"), QColor("#B42318")
            if status in ("Pending Verification", "Assessment Pending"):
                return QColor("#FFF4CE"), QColor("#9A6700")
            if status == "Failed Assessment":
                return QColor("#FDE2E2"), QColor("#B42318")
            if status == "Expiring Soon":
                return QColor("#FFF4CE"), QColor("#9A6700")
            if status == "Valid":
                return QColor("#DDF5E5"), QColor("#176B3A")
            return QColor("#F0F2F5"), QColor("#555555")

        # =====================================================================
        # TAB 1 — Enrol Employee (form + bulk only)
        # =====================================================================
        enrol_page = QWidget()
        enrol_layout = QVBoxLayout(enrol_page)
        enrol_layout.setContentsMargins(12, 12, 12, 12)
        enrol_layout.setSpacing(10)

        form_box = QGroupBox("Enrol Employee — ID, Name, Trade, Company")
        form = QGridLayout(form_box)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)

        emp_id_in = QLineEdit(); emp_id_in.setPlaceholderText("Employee / Badge ID *")
        emp_name_in = QLineEdit(); emp_name_in.setPlaceholderText("Full name *")
        emp_trade_in = QComboBox(); emp_trade_in.setEditable(True)
        emp_trade_in.addItem("Select trade / designation", "")
        for t in MATRIX_DESIGNATIONS:
            emp_trade_in.addItem(t, t)
        emp_company_in = QLineEdit(); emp_company_in.setPlaceholderText("Company / Subcontractor")
        emp_induction = QDateEdit(QDate.currentDate()); emp_induction.setCalendarPopup(True); emp_induction.setDisplayFormat("yyyy-MM-dd")
        emp_photo_path = {"value": ""}
        emp_photo_lbl = QLabel("No photo")
        emp_photo_lbl.setFixedSize(100, 100)
        emp_photo_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        emp_photo_lbl.setStyleSheet("border:1px solid #C8D5E3; background:#F7FAFC; border-radius:4px;")
        emp_photo_btn = QPushButton("Upload Photo")
        emp_photo_clear = QPushButton("Clear Photo")

        form.addWidget(QLabel("Employee ID *"), 0, 0); form.addWidget(emp_id_in, 0, 1)
        form.addWidget(QLabel("Full Name *"), 0, 2); form.addWidget(emp_name_in, 0, 3)
        form.addWidget(QLabel("Trade / Designation *"), 1, 0); form.addWidget(emp_trade_in, 1, 1)
        form.addWidget(QLabel("Company / Subcontractor"), 1, 2); form.addWidget(emp_company_in, 1, 3)
        form.addWidget(QLabel("Induction Date"), 2, 0); form.addWidget(emp_induction, 2, 1)
        photo_col = QVBoxLayout()
        photo_col.addWidget(emp_photo_lbl, 0, Qt.AlignmentFlag.AlignCenter)
        photo_btns = QHBoxLayout(); photo_btns.addWidget(emp_photo_btn); photo_btns.addWidget(emp_photo_clear)
        photo_col.addLayout(photo_btns)
        form.addLayout(photo_col, 0, 4, 3, 1)

        btn_row = QHBoxLayout()
        btn_save_draft = QPushButton("Save & Exit")
        btn_submit = QPushButton("Submit")
        btn_submit.setStyleSheet("font-weight:bold;")
        btn_clear = QPushButton("Clear Form")
        for b in (btn_save_draft, btn_submit, btn_clear):
            b.setMinimumHeight(36)
            btn_row.addWidget(b)
        btn_row.addStretch()
        form.addLayout(btn_row, 3, 0, 1, 5)
        enrol_layout.addWidget(form_box)

        bulk_box = QGroupBox("Bulk enrol workers (one per line: ID | Name | Trade | Company)")
        bulk_lay = QVBoxLayout(bulk_box)
        bulk_text = QTextEdit()
        bulk_text.setPlaceholderText("E-1001 | John Smith | Scaffolder | ABC Contracting")
        bulk_text.setMinimumHeight(160)
        bulk_btn = QPushButton("Submit Bulk Workers")
        bulk_btn.setMinimumHeight(36)
        bulk_lay.addWidget(bulk_text)
        bulk_lay.addWidget(bulk_btn)
        enrol_layout.addWidget(bulk_box, 1)

        selected_emp = {"id": None}

        def set_photo_preview(path_str):
            emp_photo_path["value"] = safe(path_str)
            p = Path(safe(path_str)) if path_str else None
            if p and p.is_file():
                pix = QPixmap(str(p))
                if not pix.isNull():
                    emp_photo_lbl.setPixmap(pix.scaled(100, 100, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
                    emp_photo_lbl.setText("")
                    return
            emp_photo_lbl.setPixmap(QPixmap())
            emp_photo_lbl.setText("No photo")

        def choose_photo():
            path, _ = QFileDialog.getOpenFileName(self, "Select Passport Photo", "", "Images (*.png *.jpg *.jpeg *.bmp)")
            if not path:
                return
            dest_dir = ATTACH_DIR / "employee_photos"
            dest_dir.mkdir(parents=True, exist_ok=True)
            src = Path(path)
            dest = dest_dir / f"{emp_id_in.text().strip() or 'emp'}_{src.name}"
            try:
                shutil.copy2(src, dest)
                set_photo_preview(str(dest))
            except Exception as exc:
                QMessageBox.warning(self, "Photo", f"Could not save photo.\\n{exc}")

        def clear_emp_form():
            selected_emp["id"] = None
            emp_id_in.clear(); emp_name_in.clear(); emp_company_in.clear()
            emp_trade_in.setCurrentIndex(0)
            emp_induction.setDate(QDate.currentDate())
            set_photo_preview("")

        def save_employee(submit=True):
            emp_id = emp_id_in.text().strip()
            name = emp_name_in.text().strip()
            trade = emp_trade_in.currentText().strip()
            if trade.lower().startswith("select"):
                trade = ""
            company = emp_company_in.text().strip()
            if not emp_id or not name:
                return QMessageBox.warning(self, "Enrol Employee", "Employee ID and Full Name are required.")
            if submit and not trade:
                return QMessageBox.warning(self, "Enrol Employee", "Select Trade / Designation before Submit.")
            induction = emp_induction.date().toString("yyyy-MM-dd")
            photo = emp_photo_path["value"]
            dup = db.fetchone(
                "SELECT id FROM employees WHERE lower(employee_id)=lower(?) AND id<>COALESCE(?, -1)",
                (emp_id, selected_emp["id"])
            )
            if dup:
                return QMessageBox.warning(self, "Enrol Employee", "That Employee ID already exists.")
            if selected_emp["id"] is None:
                db.execute(
                    "INSERT INTO employees (employee_id,name,designation,company,department,safety_induction_date,photo_path,active) VALUES (?,?,?,?,?,?,?,1)",
                    (emp_id, name, trade, company, "", induction, photo)
                )
            else:
                db.execute(
                    "UPDATE employees SET employee_id=?,name=?,designation=?,company=?,safety_induction_date=?,photo_path=?,active=1 WHERE id=?",
                    (emp_id, name, trade, company, induction, photo, selected_emp["id"])
                )
            clear_emp_form()
            refresh_register()
            load_tracker()
            if submit:
                QMessageBox.information(self, "Enrol Employee", "Employee submitted. Open Employee Register or Tracking Register to view.")
            else:
                QMessageBox.information(self, "Enrol Employee", "Draft saved.")

        def submit_bulk():
            lines = [ln.strip() for ln in bulk_text.toPlainText().splitlines() if ln.strip()]
            if not lines:
                return QMessageBox.warning(self, "Bulk Enrol", "Enter lines as: ID | Name | Trade | Company")
            added = updated = skipped = 0
            for ln in lines:
                parts = [p.strip() for p in re.split(r"[|\t;]", ln) if p.strip()]
                if len(parts) < 3:
                    skipped += 1
                    continue
                eid, name, trade = parts[0], parts[1], parts[2]
                company = parts[3] if len(parts) > 3 else ""
                existing = db.fetchone("SELECT id FROM employees WHERE lower(trim(COALESCE(employee_id,'')))=lower(trim(?))", (eid,))
                if existing:
                    db.execute("UPDATE employees SET name=?, designation=?, company=?, active=1 WHERE id=?", (name, trade, company, existing["id"]))
                    updated += 1
                else:
                    db.execute("INSERT INTO employees (employee_id,name,designation,company,department,active) VALUES (?,?,?,?,?,1)", (eid, name, trade, company, ""))
                    added += 1
            bulk_text.clear()
            refresh_register()
            load_tracker()
            QMessageBox.information(self, "Bulk Enrol", f"Added: {added} · Updated: {updated} · Skipped: {skipped}")

        emp_photo_btn.clicked.connect(choose_photo)
        emp_photo_clear.clicked.connect(lambda: set_photo_preview(""))
        btn_save_draft.clicked.connect(lambda: save_employee(submit=False))
        btn_submit.clicked.connect(lambda: save_employee(submit=True))
        btn_clear.clicked.connect(clear_emp_form)
        bulk_btn.clicked.connect(submit_bulk)
        tabs.addTab(enrol_page, "1. Enrol Employee")

        # =====================================================================
        # TAB 2 — Employee Register (list only)
        # =====================================================================
        reg_page = QWidget()
        reg_layout = QVBoxLayout(reg_page)
        reg_layout.setContentsMargins(8, 8, 8, 8)
        reg_layout.setSpacing(8)
        reg_tools = QHBoxLayout()
        reg_search = QLineEdit(); reg_search.setPlaceholderText("Search ID, name, trade, company…"); reg_search.setClearButtonEnabled(True)
        reg_refresh = QPushButton("Refresh")
        reg_edit = QPushButton("Edit Selected")
        reg_delete = QPushButton("Delete Selected")
        reg_export_xlsx = QPushButton("Export Excel")
        reg_export_docx = QPushButton("Export Word")
        for b in (reg_refresh, reg_edit, reg_delete, reg_export_xlsx, reg_export_docx):
            b.setMinimumHeight(34)
        reg_tools.addWidget(reg_search, 1)
        reg_tools.addWidget(reg_refresh); reg_tools.addWidget(reg_edit); reg_tools.addWidget(reg_delete)
        reg_tools.addWidget(reg_export_xlsx); reg_tools.addWidget(reg_export_docx)
        reg_layout.addLayout(reg_tools)

        reg_table = QTableWidget(0, 7)
        reg_table.setHorizontalHeaderLabels(["S.No.", "Employee ID", "Full Name", "Trade", "Company", "Induction", "Photo"])
        reg_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        reg_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        reg_table.setAlternatingRowColors(True)
        reg_table.verticalHeader().setVisible(False)
        reg_table.verticalHeader().setDefaultSectionSize(48)
        reg_table.horizontalHeader().setStretchLastSection(True)
        reg_layout.addWidget(reg_table, 1)

        def refresh_register():
            needle = reg_search.text().strip().lower()
            rows = db.fetchall(
                "SELECT id, employee_id, name, designation, company, safety_induction_date, photo_path "
                "FROM employees WHERE COALESCE(active,1)=1 ORDER BY name COLLATE NOCASE"
            )
            data = []
            for r in rows:
                blob = " ".join(safe(r[k]).lower() for k in ("employee_id", "name", "designation", "company"))
                if needle and needle not in blob:
                    continue
                data.append(r)
            reg_table.setRowCount(len(data))
            for i, r in enumerate(data):
                vals = [str(i + 1), safe(r["employee_id"]), safe(r["name"]), safe(r["designation"]),
                        safe(r["company"]), safe(r["safety_induction_date"])]
                for j, v in enumerate(vals):
                    it = QTableWidgetItem(v)
                    if j == 1:
                        it.setData(Qt.ItemDataRole.UserRole, int(r["id"]))
                    reg_table.setItem(i, j, it)
                photo_item = QTableWidgetItem()
                p = Path(safe(r["photo_path"])) if safe(r["photo_path"]) else None
                if p and p.is_file():
                    pix = QPixmap(str(p))
                    if not pix.isNull():
                        photo_item.setData(Qt.ItemDataRole.DecorationRole, pix.scaled(40, 40, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
                reg_table.setItem(i, 6, photo_item)
            reg_table.setColumnWidth(0, 56)
            reg_table.setColumnWidth(1, 110)
            reg_table.setColumnWidth(2, 180)
            reg_table.setColumnWidth(3, 150)
            reg_table.setColumnWidth(4, 150)
            reg_table.setColumnWidth(5, 110)

        def edit_selected_employee():
            row = reg_table.currentRow()
            if row < 0 or not reg_table.item(row, 1):
                return QMessageBox.information(self, "Edit", "Select an employee row first.")
            eid = reg_table.item(row, 1).data(Qt.ItemDataRole.UserRole)
            r = db.fetchone("SELECT * FROM employees WHERE id=?", (eid,))
            if not r:
                return
            selected_emp["id"] = eid
            emp_id_in.setText(safe(r["employee_id"]))
            emp_name_in.setText(safe(r["name"]))
            emp_company_in.setText(safe(r["company"]))
            trade = safe(r["designation"])
            ix = emp_trade_in.findText(trade)
            if ix < 0 and trade:
                emp_trade_in.addItem(trade, trade); ix = emp_trade_in.findText(trade)
            emp_trade_in.setCurrentIndex(ix if ix >= 0 else 0)
            try:
                d = QDate.fromString(safe(r["safety_induction_date"]), "yyyy-MM-dd")
                if d.isValid():
                    emp_induction.setDate(d)
            except Exception:
                pass
            set_photo_preview(safe(r["photo_path"]) if "photo_path" in r.keys() else "")
            tabs.setCurrentWidget(enrol_page)

        def delete_selected_employee():
            row = reg_table.currentRow()
            if row < 0 or not reg_table.item(row, 1):
                return QMessageBox.information(self, "Delete", "Select an employee row first.")
            eid = reg_table.item(row, 1).data(Qt.ItemDataRole.UserRole)
            r = db.fetchone("SELECT employee_id, name FROM employees WHERE id=?", (eid,))
            if not r:
                return
            if QMessageBox.question(
                self, "Delete Employee",
                f"Remove {safe(r['name'])} ({safe(r['employee_id'])}) from the register?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            ) != QMessageBox.StandardButton.Yes:
                return
            db.execute("UPDATE employees SET active=0 WHERE id=?", (eid,))
            refresh_register()
            load_tracker()

        def export_reg_excel():
            path, _ = QFileDialog.getSaveFileName(self, "Export Employee Register", "Employee_Register.xlsx", "Excel Workbook (*.xlsx)")
            if not path:
                return
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            wb = Workbook(); ws = wb.active; ws.title = "Employees"
            ws.append(["EMPLOYEE REGISTER"])
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
            ws["A1"].font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
            ws["A1"].fill = PatternFill("solid", fgColor="17365D")
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws.append(["S.No.", "Employee ID", "Full Name", "Trade", "Company", "Induction Date"])
            for cell in ws[2]:
                cell.font = Font(bold=True, color="17365D"); cell.fill = PatternFill("solid", fgColor="D9EAF7")
            for r in range(reg_table.rowCount()):
                ws.append([(reg_table.item(r, c).text() if reg_table.item(r, c) else "") for c in range(6)])
            wb.save(path)
            QMessageBox.information(self, "Export", f"Exported:\\n{path}")

        def export_reg_word():
            path, _ = QFileDialog.getSaveFileName(self, "Export Employee Register", "Employee_Register.docx", "Word Document (*.docx)")
            if not path:
                return
            if not path.lower().endswith(".docx"):
                path += ".docx"
            doc = Document()
            add_docx_report_header(doc, "EMPLOYEE REGISTER", "TRN", f"{document_prefix()}-EMP-{datetime.now().strftime('%Y%m%d%H%M%S')}")
            table = doc.add_table(rows=1, cols=6)
            for i, h in enumerate(["S.No.", "Employee ID", "Full Name", "Trade", "Company", "Induction"]):
                table.rows[0].cells[i].text = h
            for r in range(reg_table.rowCount()):
                cells = table.add_row().cells
                for c in range(6):
                    it = reg_table.item(r, c)
                    cells[c].text = xml_safe(it.text() if it else "")
            style_docx_table(table, header=True, font_size=8)
            add_docx_footer(doc)
            save_docx_validated(doc, path)
            QMessageBox.information(self, "Export", f"Exported:\\n{path}")

        reg_refresh.clicked.connect(refresh_register)
        reg_search.textChanged.connect(refresh_register)
        reg_edit.clicked.connect(edit_selected_employee)
        reg_delete.clicked.connect(delete_selected_employee)
        reg_export_xlsx.clicked.connect(export_reg_excel)
        reg_export_docx.clicked.connect(export_reg_word)
        reg_table.doubleClicked.connect(lambda *_: edit_selected_employee())
        refresh_register()
        tabs.addTab(reg_page, "2. Employee Register")

        # =====================================================================
        # TAB 3 — Training Matrix Grid (full screen)
        # =====================================================================
        grid_page = QWidget()
        grid_layout = QVBoxLayout(grid_page)
        grid_layout.setContentsMargins(6, 6, 6, 6)
        grid_layout.setSpacing(6)
        grid_tools = QHBoxLayout()
        btn_load_ready = QPushButton("Load Ready-to-Use Matrix")
        btn_save_grid = QPushButton("Save Matrix Grid")
        btn_reload_grid = QPushButton("Reload")
        btn_rename_course = QPushButton("Rename Selected Course")
        btn_delete_course = QPushButton("Delete Selected Course")
        grid_status = QLabel("")
        for b in (btn_load_ready, btn_save_grid, btn_reload_grid, btn_rename_course, btn_delete_course):
            b.setMinimumHeight(32)
            grid_tools.addWidget(b)
        grid_tools.addStretch()
        grid_tools.addWidget(grid_status)
        grid_layout.addLayout(grid_tools)
        matrix_grid = QTableWidget()
        matrix_grid.setAlternatingRowColors(True)
        matrix_grid.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        matrix_grid.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        matrix_grid.horizontalHeader().setDefaultSectionSize(88)
        matrix_grid.verticalHeader().setDefaultSectionSize(28)
        matrix_grid.verticalHeader().setVisible(False)
        grid_layout.addWidget(matrix_grid, 1)

        def render_matrix_grid():
            designations = list(MATRIX_DESIGNATIONS)
            courses = list(grid_state["courses"])
            matrix_grid.clear()
            matrix_grid.setRowCount(len(courses))
            matrix_grid.setColumnCount(2 + len(designations))
            matrix_grid.setHorizontalHeaderLabels(["Training Name", "Validity (months)"] + designations)
            for r, cname in enumerate(courses):
                name_item = QTableWidgetItem(cname)
                name_item.setBackground(QBrush(QColor("#EEF4FB")))
                f = name_item.font(); f.setBold(True); name_item.setFont(f)
                name_item.setToolTip("Double-click to rename this training course")
                matrix_grid.setItem(r, 0, name_item)
                months = int(grid_state["validity"].get(cname, 0) or 0)
                val_item = QTableWidgetItem(str(months))
                val_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                val_item.setBackground(QBrush(QColor("#F7FAFC")))
                val_item.setToolTip("Double-click to change validity (months)")
                matrix_grid.setItem(r, 1, val_item)
                for c, desig in enumerate(designations):
                    req = grid_state["cells"].get((cname, desig), "Not Applicable")
                    it = QTableWidgetItem()
                    _style_req(it, req)
                    matrix_grid.setItem(r, c + 2, it)
            matrix_grid.setColumnWidth(0, 260)
            matrix_grid.setColumnWidth(1, 120)
            for col in range(2, matrix_grid.columnCount()):
                matrix_grid.setColumnWidth(col, 90)
            grid_status.setText(f"{len(courses)} courses × {len(designations)} trades" + (" · Unsaved" if grid_state["dirty"] else ""))

        def on_grid_dbl(row, column):
            if row < 0 or row >= len(grid_state["courses"]):
                return
            old_name = grid_state["courses"][row]
            if column == 0:
                new_name, ok = QInputDialog.getText(self, "Rename Training", "Training name:", text=old_name)
                new_name = safe(new_name).strip()
                if not ok or not new_name or new_name == old_name:
                    return
                if new_name in grid_state["courses"]:
                    return QMessageBox.warning(self, "Rename Training", "A course with that name already exists.")
                grid_state["courses"][row] = new_name
                if old_name in grid_state["validity"]:
                    grid_state["validity"][new_name] = grid_state["validity"].pop(old_name)
                new_cells = {}
                for (cn, desig), req in list(grid_state["cells"].items()):
                    if cn == old_name:
                        new_cells[(new_name, desig)] = req
                    else:
                        new_cells[(cn, desig)] = req
                grid_state["cells"] = new_cells
                grid_state["dirty"] = True
                render_matrix_grid()
                return
            if column == 1:
                months, ok = QInputDialog.getInt(self, "Validity", f"Validity months for {old_name}:", int(grid_state["validity"].get(old_name, 0) or 0), 0, 120)
                if ok:
                    grid_state["validity"][old_name] = months
                    grid_state["dirty"] = True
                    render_matrix_grid()
                return
            desig = MATRIX_DESIGNATIONS[column - 2]
            current = grid_state["cells"].get((old_name, desig), "Not Applicable")
            choice, ok = QInputDialog.getItem(self, "Set requirement", f"{old_name}\\n{desig}", MATRIX_REQ_OPTIONS, max(0, MATRIX_REQ_OPTIONS.index(current) if current in MATRIX_REQ_OPTIONS else 0), False)
            if ok:
                grid_state["cells"][(old_name, desig)] = choice
                grid_state["dirty"] = True
                render_matrix_grid()

        def rename_selected_course():
            row = matrix_grid.currentRow()
            if row < 0:
                return QMessageBox.information(self, "Rename", "Select a course row first.")
            on_grid_dbl(row, 0)

        def delete_selected_course():
            row = matrix_grid.currentRow()
            if row < 0 or row >= len(grid_state["courses"]):
                return QMessageBox.information(self, "Delete Course", "Select a course row first.")
            cname = grid_state["courses"][row]
            if QMessageBox.question(self, "Delete Course", f"Remove '{cname}' from the matrix?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
                return
            grid_state["courses"].pop(row)
            grid_state["validity"].pop(cname, None)
            grid_state["cells"] = {k: v for k, v in grid_state["cells"].items() if k[0] != cname}
            grid_state["dirty"] = True
            render_matrix_grid()

        def do_load_ready():
            if QMessageBox.question(self, "Load Matrix", "Load ready-to-use matrix into empty cells?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes) != QMessageBox.StandardButton.Yes:
                return
            seed_ready_matrix(force=False)
            render_matrix_grid()

        def do_save_grid():
            save_grid_state()
            render_matrix_grid()
            load_tracker()
            QMessageBox.information(self, "Matrix Grid", "Matrix saved.")

        matrix_grid.cellDoubleClicked.connect(on_grid_dbl)
        btn_load_ready.clicked.connect(do_load_ready)
        btn_save_grid.clicked.connect(do_save_grid)
        btn_reload_grid.clicked.connect(lambda: (load_grid_state(), render_matrix_grid()))
        btn_rename_course.clicked.connect(rename_selected_course)
        btn_delete_course.clicked.connect(delete_selected_course)
        load_grid_state()
        render_matrix_grid()
        tabs.addTab(grid_page, "3. Training Matrix Grid")

        # =====================================================================
        # TAB 4 — Add Training Course (goes into matrix)
        # =====================================================================
        add_page = QWidget()
        add_layout = QVBoxLayout(add_page)
        add_layout.setContentsMargins(16, 16, 16, 16)
        add_box = QGroupBox("Add Training Course — appears on Matrix Grid")
        add_form = QGridLayout(add_box)
        add_form.setHorizontalSpacing(12)
        add_form.setVerticalSpacing(10)
        add_name = QLineEdit(); add_name.setPlaceholderText("Training / course name *")
        add_valid = QSpinBox(); add_valid.setRange(0, 120); add_valid.setValue(24); add_valid.setSuffix(" months")
        add_default_req = QComboBox(); add_default_req.addItems(["Not Applicable", "Mandatory", "Optional", "Before Job"])
        add_default_req.setCurrentText("Mandatory")
        add_apply_trade = QComboBox(); add_apply_trade.setEditable(True)
        add_apply_trade.addItem("All field trades (recommended)", "__ALL_FIELD__")
        add_apply_trade.addItem("No trades (set M/O/B on matrix)", "__NONE__")
        for t in MATRIX_DESIGNATIONS:
            add_apply_trade.addItem(t, t)
        add_btn = QPushButton("Add to Matrix Grid")
        add_btn.setMinimumHeight(38)
        add_btn.setStyleSheet("font-weight:bold;")
        add_form.addWidget(QLabel("Training Name *"), 0, 0); add_form.addWidget(add_name, 0, 1, 1, 3)
        add_form.addWidget(QLabel("Validity"), 1, 0); add_form.addWidget(add_valid, 1, 1)
        add_form.addWidget(QLabel("Default requirement"), 1, 2); add_form.addWidget(add_default_req, 1, 3)
        add_form.addWidget(QLabel("Apply to trade"), 2, 0); add_form.addWidget(add_apply_trade, 2, 1, 1, 3)
        add_form.addWidget(add_btn, 3, 0, 1, 4)
        add_layout.addWidget(add_box)
        add_hint = QLabel(
            "After adding, open Tab 3 (Training Matrix Grid) to fine-tune M / O / B per trade, "
            "then click Save Matrix Grid. Tracking Register uses the saved matrix."
        )
        add_hint.setWordWrap(True)
        add_layout.addWidget(add_hint)
        add_layout.addStretch()

        def add_course_to_matrix():
            cname = add_name.text().strip()
            if not cname:
                return QMessageBox.warning(self, "Add Training", "Enter a training name.")
            if cname not in grid_state["courses"]:
                grid_state["courses"].append(cname)
            grid_state["validity"][cname] = int(add_valid.value())
            req = add_default_req.currentText()
            target = add_apply_trade.currentData() or add_apply_trade.currentText()
            office = {"Managing Director / MD", "CEO", "General Manager", "Admin / Office Staff", "Document Controller", "Visitor"}
            if target == "__ALL_FIELD__":
                for desig in MATRIX_DESIGNATIONS:
                    if desig not in office:
                        grid_state["cells"][(cname, desig)] = req
                    else:
                        grid_state["cells"].setdefault((cname, desig), "Not Applicable")
            elif target == "__NONE__":
                for desig in MATRIX_DESIGNATIONS:
                    grid_state["cells"].setdefault((cname, desig), "Not Applicable")
            else:
                for desig in MATRIX_DESIGNATIONS:
                    grid_state["cells"].setdefault((cname, desig), "Not Applicable")
                grid_state["cells"][(cname, target)] = req
            grid_state["dirty"] = True
            render_matrix_grid()
            add_name.clear()
            tabs.setCurrentWidget(grid_page)
            QMessageBox.information(self, "Add Training", f"'{cname}' added to the matrix. Set M/O/B cells if needed, then Save Matrix Grid.")

        add_btn.clicked.connect(add_course_to_matrix)
        tabs.addTab(add_page, "4. Add Training Course")

        # =====================================================================
        # TAB 5 — Compliance Dashboard (Training module only)
        # =====================================================================
        compliance_page = QWidget()
        compliance_layout = QVBoxLayout(compliance_page)
        compliance_layout.setContentsMargins(10, 10, 10, 10)
        compliance_layout.setSpacing(8)
        compliance_title = QLabel("COMPANY TRAINING COMPLIANCE OVERVIEW")
        compliance_title.setStyleSheet("font-size:16px;font-weight:700;color:#17365D;padding:5px;")
        compliance_layout.addWidget(compliance_title)
        compliance_kpis = QGridLayout()
        compliance_values = {}
        for idx, (key, caption) in enumerate([
            ("employees", "Active Employees"), ("compliance", "Mandatory Compliance"),
            ("expired", "Expired Requirements"), ("expiring", "Expiring ≤ 30 Days"),
            ("missing", "Not Completed"), ("pending", "Pending Verification"),
        ]):
            card = QGroupBox(caption)
            card_layout = QVBoxLayout(card)
            value_label = QLabel("—")
            value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            value_label.setStyleSheet("font-size:22px;font-weight:700;color:#17365D;padding:7px;")
            card_layout.addWidget(value_label)
            compliance_values[key] = value_label
            compliance_kpis.addWidget(card, idx // 3, idx % 3)
        compliance_layout.addLayout(compliance_kpis)
        compliance_hint = QLabel("Compliance is calculated from applicable Mandatory requirements. Before Job requirements are tracked separately in the employee register.")
        compliance_hint.setWordWrap(True)
        compliance_hint.setStyleSheet("color:#475467;padding:5px;background:#F2F4F7;border-radius:4px;")
        compliance_layout.addWidget(compliance_hint)
        compliance_refresh = QPushButton("Refresh Compliance")
        compliance_refresh.setMinimumHeight(34)
        compliance_layout.addWidget(compliance_refresh, 0, Qt.AlignmentFlag.AlignLeft)
        compliance_table = QTableWidget()
        compliance_table.setColumnCount(6)
        compliance_table.setHorizontalHeaderLabels(["Employee ID", "Name", "Company", "Trade", "Non-compliant", "Expired"])
        compliance_table.setAlternatingRowColors(True)
        compliance_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        compliance_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        compliance_table.setMinimumHeight(180)
        compliance_layout.addWidget(compliance_table, 1)

        # =====================================================================
        # TAB 6 — Tracking Register (full screen + actions only)
        # =====================================================================
        track_page = QWidget()
        track_layout = QVBoxLayout(track_page)
        track_layout.setContentsMargins(6, 6, 6, 6)
        track_layout.setSpacing(6)
        track_tools = QHBoxLayout()
        track_search = QLineEdit(); track_search.setPlaceholderText("Search ID, name, trade, company…"); track_search.setClearButtonEnabled(True)
        track_refresh = QPushButton("Refresh")
        track_edit = QPushButton("Edit Selected Cell")
        track_delete = QPushButton("Archive Selected Record")
        track_export_xlsx = QPushButton("Export Excel")
        track_export_docx = QPushButton("Export Word")
        for b in (track_refresh, track_edit, track_delete, track_export_xlsx, track_export_docx):
            b.setMinimumHeight(32)
        track_tools.addWidget(track_search, 1)
        track_tools.addWidget(track_refresh); track_tools.addWidget(track_edit); track_tools.addWidget(track_delete)
        track_tools.addWidget(track_export_xlsx); track_tools.addWidget(track_export_docx)
        track_layout.addLayout(track_tools)

        track_table = QTableWidget()
        track_table.setAlternatingRowColors(True)
        track_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        track_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        track_table.horizontalHeader().setDefaultSectionSize(130)
        track_table.verticalHeader().setDefaultSectionSize(52)
        track_table.setWordWrap(True)
        track_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        track_table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        track_layout.addWidget(track_table, 1)
        track_state = {"employees": [], "courses": []}

        def load_tracker():
            load_grid_state()
            employees = db.fetchall(
                "SELECT id, employee_id, name, designation, company FROM employees WHERE COALESCE(active,1)=1 ORDER BY name COLLATE NOCASE"
            )
            needle = track_search.text().strip().lower()
            if needle:
                employees = [e for e in employees if needle in " ".join(safe(e[k]).lower() for k in ("employee_id", "name", "designation", "company"))]
            course_set = []
            for e in employees:
                for cname, req, validity in requirements_for_trade(e["designation"]):
                    if cname not in course_set:
                        course_set.append(cname)
            if not course_set:
                course_set = list(grid_state["courses"])[:25]
            track_state["employees"] = employees
            track_state["courses"] = course_set
            headers = ["S.No.", "ID", "Name", "Trade", "Company", "Mandatory %"] + course_set
            track_table.clear()
            track_table.setColumnCount(len(headers))
            track_table.setHorizontalHeaderLabels(headers)
            track_table.setRowCount(len(employees))
            for r, emp in enumerate(employees):
                reqs = requirements_for_trade(emp["designation"])
                mand = [x for x in reqs if x[1] == "Mandatory"]
                done = 0
                for cname, req, validity in mand:
                    st, _, _ = status_of(latest_record(emp["id"], cname), validity)
                    if st in ("Valid", "Expiring Soon"):
                        done += 1
                pct = int(round(100 * done / len(mand))) if mand else None
                pct_text = "N/A" if pct is None else f"{pct}% ({done}/{len(mand)})"
                base = [str(r + 1), safe(emp["employee_id"]), safe(emp["name"]), safe(emp["designation"]), safe(emp["company"]), pct_text]
                for c, val in enumerate(base):
                    it = QTableWidgetItem(val)
                    if c < 5:
                        it.setBackground(QBrush(QColor("#EEF4FB")))
                    if c == 2:
                        f = it.font(); f.setBold(True); it.setFont(f)
                    if c == 5 and pct is not None:
                        bg, fg = color_status("Valid" if pct >= 100 else ("Expiring Soon" if pct >= 50 else "Not Completed"))
                        it.setBackground(QBrush(bg)); it.setForeground(QBrush(fg))
                        f = it.font(); f.setBold(True); it.setFont(f)
                    if c == 1:
                        it.setData(Qt.ItemDataRole.UserRole, int(emp["id"]))
                    track_table.setItem(r, c, it)
                req_map = {x[0]: x for x in reqs}
                for offset, cname in enumerate(course_set):
                    col = 6 + offset
                    if cname not in req_map:
                        it = QTableWidgetItem("—")
                        it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                        bg, fg = color_status("")
                        it.setBackground(QBrush(bg)); it.setForeground(QBrush(fg))
                        track_table.setItem(r, col, it)
                        continue
                    req, validity = req_map[cname][1], req_map[cname][2]
                    rec = latest_record(emp["id"], cname)
                    st, train, expiry = status_of(rec, validity)
                    days = None
                    if expiry:
                        try:
                            days = (date.fromisoformat(expiry) - date.today()).days
                        except Exception:
                            days = None
                    days_txt = "" if days is None else (f"{days}d left" if days >= 0 else f"{abs(days)}d overdue")
                    short = {"Mandatory": "M", "Optional": "O", "Before Job": "B"}.get(req, "")
                    lines = [f"{short} · {st}"]
                    if train:
                        lines.append(train)
                    if expiry:
                        lines.append("Exp " + expiry)
                    if days_txt:
                        lines.append(days_txt)
                    it = QTableWidgetItem("\\n".join(lines))
                    it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    it.setData(Qt.ItemDataRole.UserRole, {"emp_id": int(emp["id"]), "course": cname, "validity": validity, "req": req, "record_id": int(rec["id"]) if rec else None})
                    bg, fg = color_status(st)
                    it.setBackground(QBrush(bg)); it.setForeground(QBrush(fg))
                    track_table.setItem(r, col, it)
            track_table.setColumnWidth(0, 48)
            track_table.setColumnWidth(1, 90)
            track_table.setColumnWidth(2, 140)
            track_table.setColumnWidth(3, 120)
            track_table.setColumnWidth(4, 120)
            track_table.setColumnWidth(5, 110)

        def refresh_compliance_dashboard():
            employees = db.fetchall(
                "SELECT id, employee_id, name, designation, company FROM employees WHERE COALESCE(active,1)=1 ORDER BY name COLLATE NOCASE"
            )
            total_required = 0
            compliant_required = 0
            expired_count = 0
            expiring_count = 0
            missing_count = 0
            pending_count = 0
            exceptions = []
            for emp in employees:
                reqs = requirements_for_trade(emp["designation"])
                non_compliant = 0
                emp_expired = 0
                for cname, req, validity in reqs:
                    if req != "Mandatory":
                        continue
                    total_required += 1
                    rec = latest_record(emp["id"], cname)
                    st, _, _ = status_of(rec, validity)
                    if st in ("Valid", "Expiring Soon"):
                        compliant_required += 1
                    else:
                        non_compliant += 1
                        if st == "Expired":
                            expired_count += 1
                            emp_expired += 1
                        elif st == "Pending Verification":
                            pending_count += 1
                        else:
                            missing_count += 1
                    if st == "Expiring Soon":
                        expiring_count += 1
                if non_compliant:
                    exceptions.append((emp, non_compliant, emp_expired))
            compliance_pct = (100.0 * compliant_required / total_required) if total_required else 100.0
            compliance_values["employees"].setText(str(len(employees)))
            compliance_values["compliance"].setText(f"{compliance_pct:.1f}%")
            compliance_values["expired"].setText(str(expired_count))
            compliance_values["expiring"].setText(str(expiring_count))
            compliance_values["missing"].setText(str(missing_count))
            compliance_values["pending"].setText(str(pending_count))
            compliance_table.setRowCount(len(exceptions))
            for row_idx, (emp, non_compliant, emp_expired) in enumerate(exceptions):
                values = [safe(emp["employee_id"]), safe(emp["name"]), safe(emp["company"]), safe(emp["designation"]), str(non_compliant), str(emp_expired)]
                for col_idx, value in enumerate(values):
                    compliance_table.setItem(row_idx, col_idx, QTableWidgetItem(value))
            compliance_table.sortItems(4, Qt.SortOrder.DescendingOrder)

        compliance_refresh.clicked.connect(lambda: (load_grid_state(), refresh_compliance_dashboard()))

        def log_or_edit_cell(row=None, column=None):
            if row is None:
                row = track_table.currentRow()
            if column is None:
                column = track_table.currentColumn()
            if row < 0 or column is None or column < 6:
                return QMessageBox.information(self, "Tracking", "Select a training cell (not identity columns).")
            item = track_table.item(row, column)
            if not item:
                return
            data = item.data(Qt.ItemDataRole.UserRole)
            if not isinstance(data, dict):
                return
            emp_id = data["emp_id"]
            cname = data["course"]
            validity = int(data.get("validity") or 0)
            req = data.get("req", "")
            rec = latest_record(emp_id, cname)
            dlg = QDialog(self)
            dlg.setWindowTitle(f"Log Training — {cname}")
            dlg.setMinimumWidth(360)
            form = QFormLayout(dlg)
            form.addRow(QLabel(f"Requirement: {req} · Validity: {validity} months"))
            d_train = QDateEdit(QDate.currentDate()); d_train.setCalendarPopup(True); d_train.setDisplayFormat("yyyy-MM-dd")
            if rec and safe(rec["training_date"]):
                td = QDate.fromString(safe(rec["training_date"]), "yyyy-MM-dd")
                if td.isValid():
                    d_train.setDate(td)
            d_exp = QDateEdit(QDate.currentDate()); d_exp.setCalendarPopup(True); d_exp.setDisplayFormat("yyyy-MM-dd")
            no_expiry = QCheckBox("No expiry — this course does not expire")
            no_expiry.setChecked(bool(not validity and not (rec and safe(rec["expiry_date"]))))
            d_exp.setEnabled(not no_expiry.isChecked())
            no_expiry.toggled.connect(d_exp.setDisabled)
            def auto_exp(*_):
                if validity and not no_expiry.isChecked():
                    try:
                        base = date(d_train.date().year(), d_train.date().month(), d_train.date().day())
                        exp = add_calendar_months(base, int(validity))
                        d_exp.setDate(QDate(exp.year, exp.month, exp.day))
                    except Exception:
                        pass
            d_train.dateChanged.connect(auto_exp)
            auto_exp()
            if rec and safe(rec["expiry_date"]):
                ed = QDate.fromString(safe(rec["expiry_date"]), "yyyy-MM-dd")
                if ed.isValid():
                    d_exp.setDate(ed)
                    no_expiry.setChecked(False)
            trainer = QLineEdit(safe(rec["provider"]) if rec else "")
            cert_no = QLineEdit(safe(rec["certificate_no"]) if rec else "")
            notes = QLineEdit(safe(rec["notes"]) if rec else "")
            attachment_path = QLineEdit(safe(rec["attachment_path"]) if rec else "")
            attachment_path.setReadOnly(True)
            attachment_btn = QPushButton("Browse…")
            attachment_row = QHBoxLayout()
            attachment_row.addWidget(attachment_path, 1)
            attachment_row.addWidget(attachment_btn)
            def choose_training_attachment():
                selected, _ = QFileDialog.getOpenFileName(
                    dlg, "Select Training Certificate / Evidence", "",
                    "Evidence Files (*.pdf *.png *.jpg *.jpeg *.doc *.docx);;All Files (*)"
                )
                if not selected:
                    return
                try:
                    source = Path(selected)
                    evidence_dir = ATTACH_DIR / "training"
                    evidence_dir.mkdir(parents=True, exist_ok=True)
                    destination = evidence_dir / f"training_{emp_id}_{re.sub(r'[^A-Za-z0-9_-]+', '_', cname)[:45]}_{datetime.now().strftime('%Y%m%d%H%M%S')}{source.suffix.lower()}"
                    shutil.copy2(source, destination)
                    attachment_path.setText(str(destination))
                except Exception as exc:
                    QMessageBox.critical(dlg, "Attachment Error", f"Unable to copy evidence file.\n\n{exc}")
            attachment_btn.clicked.connect(choose_training_attachment)
            verification = QComboBox()
            verification.addItems(["Verified", "Pending Verification"])
            current_verification = safe(rec["verification_status"]) if rec and "verification_status" in rec.keys() else "Verified"
            verification.setCurrentText(current_verification if current_verification in ("Verified", "Pending Verification") else "Verified")
            assessment = QComboBox()
            assessment.addItems(["Not Required", "Pending", "Pass", "Fail"])
            current_assessment = safe(rec["assessment_result"]) if rec and "assessment_result" in rec.keys() else "Not Required"
            assessment.setCurrentText(current_assessment if current_assessment in ("Not Required", "Pending", "Pass", "Fail") else "Not Required")
            form.addRow("Date of Training *", d_train)
            form.addRow("Expiry Date", d_exp)
            form.addRow("Expiry Policy", no_expiry)
            form.addRow("Trainer / Provider", trainer)
            form.addRow("Certificate Number", cert_no)
            form.addRow("Certificate / Evidence", attachment_row)
            form.addRow("Verification Status", verification)
            form.addRow("Assessment Result", assessment)
            form.addRow("Notes", notes)
            dlg.resize(620, 480)
            dlg.setMinimumSize(540, 400)
            buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
            form.addRow(buttons)
            buttons.accepted.connect(dlg.accept); buttons.rejected.connect(dlg.reject)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                return
            course_id = ensure_course(cname, validity)
            tdate = d_train.date().toString("yyyy-MM-dd")
            edate = "" if no_expiry.isChecked() else d_exp.date().toString("yyyy-MM-dd")
            verify_status = verification.currentText()
            assess_result = assessment.currentText()
            if verify_status == "Verified":
                verified_by = "Local User"
                verified_at = datetime.now().isoformat(timespec="seconds")
            else:
                verified_by = ""
                verified_at = ""
            if rec:
                db.execute(
                    """UPDATE training_records SET training_date=?, expiry_date=?, provider=?,
                       certificate_no=?, attachment_path=?, notes=?, verification_status=?,
                       verified_by=?, verified_at=?, assessment_result=? WHERE id=?""",
                    (tdate, edate, trainer.text().strip(), cert_no.text().strip(), attachment_path.text().strip(),
                     notes.text().strip(), verify_status, verified_by, verified_at, assess_result, rec["id"])
                )
                record_id = int(rec["id"])
                action = "Updated"
            else:
                cur = db.execute(
                    """INSERT INTO training_records
                       (employee_db_id, course_id, training_date, expiry_date, provider,
                        certificate_no, attachment_path, notes, verification_status,
                        verified_by, verified_at, assessment_result)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (emp_id, course_id, tdate, edate, trainer.text().strip(), cert_no.text().strip(),
                     attachment_path.text().strip(), notes.text().strip(), verify_status,
                     verified_by, verified_at, assess_result)
                )
                record_id = int(cur.lastrowid)
                action = "Created"
            db.execute(
                "INSERT INTO training_audit_log (entity_type, entity_id, action, details) VALUES (?,?,?,?)",
                ("training_record", record_id, action,
                 f"Employee DB ID={emp_id}; course={cname}; training_date={tdate}; verification={verify_status}; assessment={assess_result}")
            )
            load_tracker()

        def clear_selected_cell():
            row = track_table.currentRow()
            column = track_table.currentColumn()
            if row < 0 or column < 6:
                return QMessageBox.information(self, "Clear", "Select a training cell first.")
            item = track_table.item(row, column)
            if not item:
                return
            data = item.data(Qt.ItemDataRole.UserRole)
            if not isinstance(data, dict):
                return
            rec = latest_record(data["emp_id"], data["course"])
            if not rec:
                return QMessageBox.information(self, "Clear", "No training record on this cell.")
            reason, ok = QInputDialog.getText(self, "Archive Training Record", "Reason for removing this record from the active register:")
            if not ok:
                return
            if not reason.strip():
                return QMessageBox.warning(self, "Reason Required", "Enter a reason. The record will be retained in Training History.")
            try:
                db.conn.execute("BEGIN")
                db.conn.execute("""INSERT INTO training_record_history
                    (original_record_id, employee_db_id, course_id, training_date, expiry_date,
                     provider, certificate_no, attachment_path, notes, created_at,
                     archived_by, archive_reason, verification_status, verified_by, verified_at,
                     assessment_result, assessment_score)
                    SELECT id, employee_db_id, course_id, training_date, expiry_date,
                           provider, certificate_no, attachment_path, notes, created_at,
                           'Local User', ?, verification_status, verified_by, verified_at,
                           assessment_result, assessment_score
                    FROM training_records WHERE id=?""", (reason.strip(), rec["id"]))
                db.conn.execute("DELETE FROM training_records WHERE id=?", (rec["id"],))
                db.conn.execute(
                    "INSERT INTO training_audit_log (entity_type, entity_id, action, details) VALUES (?,?,?,?)",
                    ("training_record", int(rec["id"]), "Archived", reason.strip())
                )
                db.conn.commit()
            except Exception as exc:
                db.conn.rollback()
                logging.exception("Training record archive failed")
                QMessageBox.critical(self, "Archive Error", f"The training record was not changed.\n\n{exc}")
                return
            load_tracker()

        def export_tracker_excel():
            path, _ = QFileDialog.getSaveFileName(self, "Export Tracking Register", "Training_Tracking_Register.xlsx", "Excel Workbook (*.xlsx)")
            if not path:
                return
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            load_tracker()
            wb = Workbook(); ws = wb.active; ws.title = "Tracking Register"
            headers = [track_table.horizontalHeaderItem(c).text() for c in range(track_table.columnCount())]
            ws.append(["HSE TRAINING TRACKING REGISTER"])
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, len(headers)))
            ws["A1"].font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
            ws["A1"].fill = PatternFill("solid", fgColor="17365D")
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws.append(headers)
            for cell in ws[2]:
                cell.font = Font(bold=True, color="17365D"); cell.fill = PatternFill("solid", fgColor="D9EAF7")
                cell.alignment = Alignment(wrap_text=True, horizontal="center")
            for r in range(track_table.rowCount()):
                vals = [(track_table.item(r, c).text() if track_table.item(r, c) else "") for c in range(track_table.columnCount())]
                ws.append(vals)
                for c, val in enumerate(vals, start=1):
                    cell = ws.cell(ws.max_row, c)
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    low = safe(val).lower()
                    if "expired" in low or "not completed" in low or "overdue" in low:
                        cell.fill = PatternFill("solid", fgColor="FDE2E2")
                    elif "expiring" in low:
                        cell.fill = PatternFill("solid", fgColor="FFF4CE")
                    elif "valid" in low:
                        cell.fill = PatternFill("solid", fgColor="DDF5E5")
            ws.freeze_panes = "A3"
            from openpyxl.utils import get_column_letter
            for col in ws.columns:
                letter = get_column_letter(col[0].column)
                longest = max((len(str(c.value or "").split("\\n")[0]) for c in col), default=10)
                ws.column_dimensions[letter].width = min(max(longest + 2, 12), 28)
            wb.save(path)
            QMessageBox.information(self, "Export", f"Exported:\\n{path}")

        def export_tracker_word():
            path, _ = QFileDialog.getSaveFileName(self, "Export Tracking Register", "Training_Tracking_Register.docx", "Word Document (*.docx)")
            if not path:
                return
            if not path.lower().endswith(".docx"):
                path += ".docx"
            load_tracker()
            doc = Document()
            set_docx_landscape(doc)
            add_docx_report_header(doc, "HSE TRAINING TRACKING REGISTER", "TRN", f"{document_prefix()}-TRN-{datetime.now().strftime('%Y%m%d%H%M%S')}")
            table = doc.add_table(rows=1, cols=6)
            for i, h in enumerate(["S.No.", "ID", "Name", "Trade", "Company", "Mandatory %"]):
                table.rows[0].cells[i].text = h
            for r in range(track_table.rowCount()):
                cells = table.add_row().cells
                for c in range(6):
                    it = track_table.item(r, c)
                    cells[c].text = xml_safe(it.text() if it else "")
            style_docx_table(table, header=True, font_size=8)
            doc.add_paragraph("Full course-level status is available in the Excel export and on-screen register.")
            add_docx_footer(doc)
            save_docx_validated(doc, path)
            QMessageBox.information(self, "Export", f"Exported:\\n{path}")

        track_table.cellDoubleClicked.connect(log_or_edit_cell)
        track_refresh.clicked.connect(load_tracker)
        track_search.textChanged.connect(load_tracker)
        track_edit.clicked.connect(lambda: log_or_edit_cell())
        track_delete.clicked.connect(clear_selected_cell)
        track_export_xlsx.clicked.connect(export_tracker_excel)
        track_export_docx.clicked.connect(export_tracker_word)
        load_tracker()
        tabs.addTab(track_page, "6. Tracking Register")
        tabs.insertTab(0, compliance_page, "1. Compliance Dashboard")
        refresh_compliance_dashboard()

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
        self.form_fields["permit_holder_employee_id"]=QComboBox()
        self.form_fields["permit_holder_employee_id"].setMinimumWidth(240)
        self.form_fields["permit_holder_employee_id"].addItem("Select PTW-trained employee ID", "")
        self._populate_ptw_holder_ids()
        self.form_fields["permit_holder_employee_id"].currentIndexChanged.connect(self._on_ptw_holder_id_selected)
        work.addRow("Permit Holder ID *",self.form_fields["permit_holder_employee_id"])
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
        for label,fn in [("View / Load",self._load_selected),("Edit Draft",self._edit_selected),("Delete Permit",self._delete_selected_permit),("Issue Selected",self._issue_selected_from_register),("Export Excel",self._export_excel_register),("Export Selected Permit",self._export_selected),("Daily Revalidation",self._open_revalidation_selected),("History",self._show_history)]:
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
        if isinstance(w,QComboBox):
            if key == "permit_holder_employee_id": return safe(w.currentData()).strip()
            return w.currentText().strip()
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
            elif isinstance(w,QComboBox) and k == "permit_holder_employee_id": w.setCurrentIndex(0)
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

    def _populate_ptw_holder_ids(self):
        combo = self.form_fields.get("permit_holder_employee_id")
        if not isinstance(combo, QComboBox):
            return
        current_id = safe(combo.currentData()).strip()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("Select PTW-trained employee ID", "")
        rows = self.db.fetchall(
            """SELECT DISTINCT e.employee_id
               FROM employees e
               JOIN training_records tr ON tr.employee_db_id=e.id
               JOIN training_courses c ON c.id=tr.course_id
               WHERE COALESCE(trim(e.employee_id),'')<>''
                 AND (lower(COALESCE(c.course_name,'')) LIKE '%ptw%'
                      OR lower(COALESCE(c.course_name,'')) LIKE '%permit to work%')
                 AND (COALESCE(trim(tr.expiry_date),'')='' OR trim(tr.expiry_date)>=?)
               ORDER BY e.employee_id COLLATE NOCASE""",
            (today(),)
        )
        for row in rows:
            employee_id = safe(row["employee_id"]).strip()
            if employee_id:
                combo.addItem(employee_id, employee_id)
        idx = combo.findData(current_id) if current_id else 0
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        combo.blockSignals(False)
        if combo.currentData():
            self._on_ptw_holder_id_selected(combo.currentIndex())

    def _on_ptw_holder_id_selected(self, _index=0):
        combo = self.form_fields.get("permit_holder_employee_id")
        employee_id = safe(combo.currentData()).strip() if isinstance(combo, QComboBox) else ""
        if not employee_id:
            for key in ("permit_holder", "permit_holder_designation"):
                widget = self.form_fields.get(key)
                if isinstance(widget, QLineEdit): widget.clear()
            return
        self._search_ptw_holder_training(employee_id)

    def _clear_ptw_holder_details_on_id_change(self, _text=""):
        # Permit holder identity is always sourced from the Training module search.
        for key in ("permit_holder", "permit_holder_designation"):
            widget=self.form_fields.get(key)
            if isinstance(widget,QLineEdit): widget.clear()

    def _search_ptw_holder_training(self, selected_employee_id=None):
        holder_widget = self.form_fields["permit_holder_employee_id"]
        employee_id = safe(selected_employee_id).strip() if selected_employee_id is not None else (safe(holder_widget.currentData()).strip() if isinstance(holder_widget,QComboBox) else "")
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
                if k == "permit_holder_employee_id":
                    i=w.findData(_s(val)); w.setCurrentIndex(i if i>=0 else 0)
                else:
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
        """Export the complete currently filtered PTW register and daily revalidation."""
        path,_=QFileDialog.getSaveFileName(self,"Export PTW Register and Daily Revalidation","ptw_register_and_revalidation.xlsx","Excel Workbook (*.xlsx)")
        if not path: return
        if not path.lower().endswith(".xlsx"): path += ".xlsx"
        try:
            self.refresh_register()
            wb=Workbook(); ws=wb.active; ws.title="PTW Register"
            headers=[self.permit_table.horizontalHeaderItem(c).text() for c in range(self.permit_table.columnCount())]
            ws.append(["PERMIT TO WORK REGISTER"]); ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=max(1,len(headers)))
            ws["A1"].font=Font(name="Arial",size=16,bold=True,color="FFFFFF"); ws["A1"].fill=PatternFill("solid",fgColor="17365D"); ws["A1"].alignment=Alignment(horizontal="center",vertical="center"); ws.row_dimensions[1].height=28
            ws.append(headers)
            for r in range(self.permit_table.rowCount()):
                ws.append([self.permit_table.item(r,c).text() if self.permit_table.item(r,c) else "" for c in range(self.permit_table.columnCount())])
            headfill=PatternFill("solid",fgColor="D9EAF7"); thin=Side(style="thin",color="B7C9D6")
            for cell in ws[2]:
                cell.font=Font(bold=True,color="17365D"); cell.fill=headfill; cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cell.border=Border(bottom=thin)
            for row in ws.iter_rows(min_row=3):
                for cell in row: cell.alignment=Alignment(vertical="top",wrap_text=True)
            ws.freeze_panes="A3"
            if ws.max_row>=2: ws.auto_filter.ref=f"A2:{ws.cell(ws.max_row,ws.max_column).coordinate}"
            for col in ws.columns:
                from openpyxl.utils import get_column_letter
                letter=get_column_letter(col[0].column); longest=max((len(str(c.value or "")) for c in col),default=10); ws.column_dimensions[letter].width=min(max(longest+2,12),32)
            daily=wb.create_sheet("Daily Revalidation")
            dh=["Permit No.","Work Date","Day","Revalidation Mark","Daily Status","Daily PTW Holder Name","Permit Issuing Authority Name","Comments","Recorded By","Recorded At"]
            daily.append(["DAILY PTW REVALIDATION REGISTER"]); daily.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(dh))
            daily["A1"].font=Font(name="Arial",size=16,bold=True,color="FFFFFF"); daily["A1"].fill=PatternFill("solid",fgColor="17365D"); daily["A1"].alignment=Alignment(horizontal="center",vertical="center"); daily.row_dimensions[1].height=28; daily.append(dh)
            records=self.db.fetchall("SELECT p.permit_number,d.work_date,d.day_status,d.permit_holder,d.permit_authority,d.comments,d.recorded_by,d.recorded_at FROM ptw_daily_revalidation d JOIN ptw_permits p ON p.id=d.permit_id ORDER BY p.permit_number,d.work_date")
            for d in records:
                try: day_name=datetime.fromisoformat(d["work_date"]).strftime("%A")
                except Exception: day_name=""
                st=_s(d["day_status"]); mark="✓" if st=="Revalidated" else "✗" if st in ("Suspended","On Hold","Rejected","No Work") else "—"
                daily.append([d["permit_number"],d["work_date"],day_name,mark,st,d["permit_holder"],d["permit_authority"],d["comments"],d["recorded_by"],d["recorded_at"]])
                fill="E2F0D9" if mark=="✓" else "FCE4D6" if mark=="✗" else "FFF2CC"
                for cell in daily[daily.max_row]: cell.fill=PatternFill("solid",fgColor=fill)
            for cell in daily[2]:
                cell.font=Font(bold=True,color="17365D"); cell.fill=headfill; cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cell.border=Border(bottom=thin)
            for row in daily.iter_rows(min_row=3):
                for cell in row: cell.alignment=Alignment(vertical="top",wrap_text=True)
            daily.freeze_panes="A3"
            if daily.max_row>=2: daily.auto_filter.ref=f"A2:{daily.cell(daily.max_row,daily.max_column).coordinate}"
            for col in daily.columns:
                from openpyxl.utils import get_column_letter
                letter=get_column_letter(col[0].column); longest=max((len(str(c.value or "")) for c in col),default=10); daily.column_dimensions[letter].width=min(max(longest+2,12),36)
            wb.save(path)
            if not Path(path).is_file() or Path(path).stat().st_size==0: raise OSError("The workbook was not created successfully.")
            QMessageBox.information(self,"Excel Exported",f"PTW Register and Daily Revalidation were saved in separate worksheets.\n{path}")
        except Exception as exc:
            logging.exception("PTW register Excel export failed")
            QMessageBox.critical(self,"Excel Export Failed",f"Could not create the Excel file.\n\n{exc}")

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

