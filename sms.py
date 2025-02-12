import streamlit as st
import sqlite3
from datetime import datetime, timedelta
import pandas as pd
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
# Alias ReportLab’s Table to avoid conflict with openpyxl Table.
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table as RLTable,
    TableStyle,
    Image,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from PIL import Image as PILImage
import io
import zipfile
from pathlib import Path
import shutil
import plotly.express as px
import fitz
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from contextlib import contextmanager
import time
import json
from typing import Union, Dict, List  # Add this line

# Email configuration constants (replace with your actual SMTP server details)
SMTP_SERVER = "smtp.example.com"  # e.g., "smtp.gmail.com"
SMTP_PORT = 587  # Typical port for TLS
SMTP_USERNAME = "your_email@example.com"
SMTP_PASSWORD = "your_email_password"

if not os.path.exists("uploads"):
    os.makedirs("uploads")


@contextmanager
def get_db_connection(max_retries=5, retry_delay=1):
    """
    Context manager for database connections with retry mechanism
    for handling locked database situations.
    """
    attempt = 0
    while attempt < max_retries:
        try:
            conn = sqlite3.connect("student_registration.db", timeout=20)
            conn.row_factory = sqlite3.Row
            yield conn
            conn.close()
            break
        except sqlite3.OperationalError as e:
            if "database is locked" in str(e):
                attempt += 1
                if attempt == max_retries:
                    raise
                time.sleep(retry_delay)
            else:
                raise


def init_db():
    conn = sqlite3.connect("student_registration.db")
    c = conn.cursor()

    c.execute(
        """
        CREATE TABLE IF NOT EXISTS admin (
            admin_id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            password TEXT,
            email TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """
    )

    # Insert admin users if they don't exist
    admins = [
        ("EdinamSD", "prettyFLACO", "edinam.ayisadu@gmail.com"),
        ("admin2", "admin456", "admin2@school.edu"),
    ]

    for admin in admins:
        c.execute(
            """
            INSERT OR IGNORE INTO admin (username, password, email) 
            VALUES (?, ?, ?)
        """,
            admin,
        )

    c.execute(
        """
                CREATE TABLE IF NOT EXISTS student_info (
                    student_id TEXT PRIMARY KEY,
                    surname TEXT,
                    other_names TEXT,
                    date_of_birth DATE,
                    place_of_birth TEXT,
                    home_town TEXT,
                    residential_address TEXT,
                    postal_address TEXT,
                    email TEXT,
                    telephone TEXT,
                    ghana_card_id TEXT,
                    nationality TEXT,
                    marital_status TEXT,
                    gender TEXT,
                    religion TEXT,
                    denomination TEXT,
                    disability_status TEXT,
                    disability_description TEXT,
                    guardian_name TEXT,
                    guardian_relationship TEXT,
                    guardian_occupation TEXT,
                    guardian_address TEXT,
                    guardian_telephone TEXT,
                    previous_school TEXT,
                    qualification_type TEXT,
                    completion_year TEXT,
                    aggregate_score TEXT,
                    ghana_card_path TEXT,
                    passport_photo_path TEXT,
                    transcript_path TEXT,
                    certificate_path TEXT,
                    receipt_path TEXT,
                    receipt_amount REAL DEFAULT 0.0,
                    approval_status TEXT DEFAULT 'pending',
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    programme TEXT,
                    password TEXT,
                    last_login DATETIME,
                    password_reset_required BOOLEAN DEFAULT 1
                )
            """
    )

    try:
        c.execute("SELECT programme FROM student_info LIMIT 1")
    except sqlite3.OperationalError:
        c.execute("ALTER TABLE student_info ADD COLUMN programme TEXT")

    c.execute(
        """
        CREATE TABLE IF NOT EXISTS course_registration (
            registration_id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id TEXT,
            index_number TEXT,
            programme TEXT,
            specialization TEXT,
            level TEXT,
            session TEXT,
            academic_year TEXT,
            semester TEXT,
            courses TEXT,
            total_credits INTEGER,
            date_registered DATE,
            approval_status TEXT DEFAULT 'pending',
            receipt_path TEXT,
            receipt_amount REAL DEFAULT 0.0,
            FOREIGN KEY (student_id) REFERENCES student_info (student_id)
        )
    """
    )

    # Create notifications table for sending notifications
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS notifications (
            notification_id INTEGER PRIMARY KEY AUTOINCREMENT,
            recipient_id TEXT,
            recipient_type TEXT,  -- 'student', 'program', or 'all'
            title TEXT,
            message TEXT,
            notification_type TEXT,  -- 'info', 'warning', 'success', 'error'
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            read_at DATETIME,
            metadata TEXT,
            expires_at DATETIME
        )
        """
    )

    # Create notification_reads table to track read notifications for each student
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS notification_reads (
            notification_id INTEGER,
            student_id TEXT,
            read_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (notification_id, student_id)
        )
        """
    )

    conn.commit()
    conn.close()



import sqlite3
import pandas as pd
import json
from datetime import datetime
import os
import shutil
import zipfile
from typing import Dict, List, Any, Optional
import logging


class DatabaseMigrationHandler:
    """
    Handles database migrations, exports, and imports while maintaining data consistency
    """

    SCHEMA_VERSION = "1.0"

    # Define the expected schema for each table
    SCHEMAS = {
        "student_info": {
            "student_id": "TEXT PRIMARY KEY",
            "surname": "TEXT",
            "other_names": "TEXT",
            "date_of_birth": "DATE",
            "place_of_birth": "TEXT",
            "home_town": "TEXT",
            "residential_address": "TEXT",
            "postal_address": "TEXT",
            "email": "TEXT",
            "telephone": "TEXT",
            "ghana_card_id": "TEXT",
            "nationality": "TEXT",
            "marital_status": "TEXT",
            "gender": "TEXT",
            "religion": "TEXT",
            "denomination": "TEXT",
            "disability_status": "TEXT",
            "disability_description": "TEXT",
            "guardian_name": "TEXT",
            "guardian_relationship": "TEXT",
            "guardian_occupation": "TEXT",
            "guardian_address": "TEXT",
            "guardian_telephone": "TEXT",
            "previous_school": "TEXT",
            "qualification_type": "TEXT",
            "completion_year": "TEXT",
            "aggregate_score": "TEXT",
            "ghana_card_path": "TEXT",
            "passport_photo_path": "TEXT",
            "transcript_path": "TEXT",
            "certificate_path": "TEXT",
            "receipt_path": "TEXT",
            "receipt_amount": "REAL DEFAULT 0.0",
            "approval_status": 'TEXT DEFAULT "pending"',
            "created_at": "DATETIME DEFAULT CURRENT_TIMESTAMP",
            "programme": "TEXT",
            "password": "TEXT",
            "last_login": "DATETIME",
            "password_reset_required": "BOOLEAN DEFAULT 1",
        },
        "course_registration": {
            "registration_id": "INTEGER PRIMARY KEY AUTOINCREMENT",
            "student_id": "TEXT",
            "index_number": "TEXT",
            "programme": "TEXT",
            "specialization": "TEXT",
            "level": "TEXT",
            "session": "TEXT",
            "academic_year": "TEXT",
            "semester": "TEXT",
            "courses": "TEXT",
            "total_credits": "INTEGER",
            "date_registered": "DATE",
            "approval_status": 'TEXT DEFAULT "pending"',
            "receipt_path": "TEXT",
            "receipt_amount": "REAL DEFAULT 0.0",
        },
    }

    def __init__(self, db_path: str, backup_dir: str = "db_backups"):
        self.db_path = db_path
        self.backup_dir = backup_dir
        self.logger = self._setup_logger()

        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)

    def _setup_logger(self) -> logging.Logger:
        """Sets up logging configuration"""
        logger = logging.getLogger("DatabaseMigration")
        logger.setLevel(logging.INFO)
        handler = logging.FileHandler("db_migration.log")
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        return logger

    def backup_database(self) -> str:
        """Creates a backup of the current database"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(self.backup_dir, f"backup_{timestamp}.db")
        shutil.copy2(self.db_path, backup_path)
        self.logger.info(f"Database backed up to {backup_path}")
        return backup_path

    def export_database(self, export_path: str) -> str:
        """
        Exports the database to a zip file containing:
        - Excel files for each table
        - JSON schema file
        - Metadata file
        """
        try:
            # Create a temporary directory for export files
            temp_dir = "temp_export"
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)

            conn = sqlite3.connect(self.db_path)

            # Export each table to Excel
            for table_name in self.SCHEMAS.keys():
                df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
                excel_path = os.path.join(temp_dir, f"{table_name}.xlsx")
                df.to_excel(excel_path, index=False)

            # Create metadata file
            metadata = {
                "schema_version": self.SCHEMA_VERSION,
                "export_date": datetime.now().isoformat(),
                "tables": list(self.SCHEMAS.keys()),
            }
            with open(os.path.join(temp_dir, "metadata.json"), "w") as f:
                json.dump(metadata, f)

            # Create schema file
            with open(os.path.join(temp_dir, "schema.json"), "w") as f:
                json.dump(self.SCHEMAS, f)

            # Create zip file
            with zipfile.ZipFile(export_path, "w") as zipf:
                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.basename(file_path)
                        zipf.write(file_path, arcname)

            self.logger.info(f"Database exported to {export_path}")
            return export_path

        except Exception as e:
            self.logger.error(f"Export failed: {str(e)}")
            raise

        finally:
            if "conn" in locals():
                conn.close()
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

    def import_database(self, import_path: str, validate: bool = True) -> bool:
        """
        Imports database from a zip file while maintaining data consistency
        """
        try:
            # Create temporary directory for import
            temp_dir = "temp_import"
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)

            # Extract zip file
            with zipfile.ZipFile(import_path, "r") as zipf:
                zipf.extractall(temp_dir)

            # Validate import data
            if validate:
                self._validate_import(temp_dir)

            # Create backup before import
            self.backup_database()

            # Import data
            conn = sqlite3.connect(self.db_path)
            for table_name in self.SCHEMAS.keys():
                excel_path = os.path.join(temp_dir, f"{table_name}.xlsx")
                if os.path.exists(excel_path):
                    df = pd.read_excel(excel_path)
                    df.to_sql(table_name, conn, if_exists="replace", index=False)

            self.logger.info("Database import completed successfully")
            return True

        except Exception as e:
            self.logger.error(f"Import failed: {str(e)}")
            raise

        finally:
            if "conn" in locals():
                conn.close()
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

    def _validate_import(self, import_dir: str) -> bool:
        """
        Validates import data against schema
        """
        # Check metadata
        with open(os.path.join(import_dir, "metadata.json"), "r") as f:
            metadata = json.load(f)
            if metadata["schema_version"] != self.SCHEMA_VERSION:
                raise ValueError(
                    f"Schema version mismatch. Expected {self.SCHEMA_VERSION}, got {metadata['schema_version']}"
                )

        # Check schema
        with open(os.path.join(import_dir, "schema.json"), "r") as f:
            import_schema = json.load(f)
            if import_schema != self.SCHEMAS:
                raise ValueError("Schema mismatch between import and current database")

        # Validate data in Excel files
        for table_name in self.SCHEMAS.keys():
            excel_path = os.path.join(import_dir, f"{table_name}.xlsx")
            if not os.path.exists(excel_path):
                raise FileNotFoundError(f"Missing table data: {table_name}")

            df = pd.read_excel(excel_path)
            expected_columns = set(self.SCHEMAS[table_name].keys())
            actual_columns = set(df.columns)

            if not expected_columns.issubset(actual_columns):
                missing_columns = expected_columns - actual_columns
                raise ValueError(f"Missing columns in {table_name}: {missing_columns}")

        return True

    def get_schema_info(self) -> Dict[str, Any]:
        """
        Returns current schema information
        """
        return {
            "version": self.SCHEMA_VERSION,
            "tables": self.SCHEMAS,
            "backup_location": self.backup_dir,
        }


def reset_db():
    conn = sqlite3.connect("student_registration.db")
    c = conn.cursor()

    # Drop existing tables
    c.execute("DROP TABLE IF EXISTS admin")

    conn.commit()
    conn.close()

    # Reinitialize the database
    init_db()


def check_db_locked():
    """
    Utility function to check if database is currently locked
    Returns True if locked, False otherwise
    """
    try:
        with get_db_connection(max_retries=1, retry_delay=0) as conn:
            conn.cursor().execute("SELECT 1")
            return False
    except sqlite3.OperationalError as e:
        if "database is locked" in str(e):
            return True
        raise


def get_program_courses(program):
    courses = {
        "CIMG": {
            "Pathway 1": [
                "PCM 101|FUNDAMENTALS OF MARKETING|3",
                "PCM 103|BUYER BEHAVIOUR|3",
                "PCM 102|BUSINESS LAW AND ETHICS|3",
            ],
            "Pathway 2": [
                "PAC 202|MANAGEMENT IN PRACTICE|3",
                "PCM 203|DIGITAL MARKETING TECHNIQUES|3",
                "PAC 201|DECISION-MAKING TECHNIQUES|3",
            ],
            "Pathway 3": [
                "PDM 301|BRANDS MANAGEMENT|3",
                "PDM 302|MARKETING RESEARCH AND INSIGHTS|3",
                "PDM 304|DIGITAL OPTIMISATION AND STRATEGY|3",
                "PDM 303|SELLING AND SALES MANAGEMENT|3",
            ],
            "Pathway 4": [
                "PDA 407|MASTERING MARKETING METRICS|3",
                "PDA 408|MANAGING CORPORATE REPUTATION|3",
                "PDA 404|DIGITAL CUSTOMER EXPERIENCE|3",
                "PDA 405|PRODUCT MANAGEMENT|3",
                "PDA 403|MANAGING MARKETING PROJECTS|3",
                "PDA 406|CUSTOMER RELATIONSHIP MANAGEMENT|3",
                "PDA 402|FINANCIAL MANAGEMENT FOR MARKETERS|3",
                "PDA 401|INTERNATIONAL MARKETING|3",
            ],
            "Pathway 5": [
                "PGD 502|STRATEGIC MARKETING PRACTICE- CASE STUDY|3",
                "PGD 503|STRATEGIC MARKETING MANAGEMENT|3",
                "PGD 501|INTEGRATED MARKETING COMMUNICATIONS|3",
                "PGD 504|ADVANCED DIGITAL MARKETING|3",
            ],
            "Pathway 6": [
                "PMS 613|SPECIALISED COMMODITIES MARKETING|3",
                "PMS 607|TRANSPORT AND LOGISTICS MARKETING|3",
                "PMS 606|NGO MARKETING|3",
                "PMS 608|AGRI-BUSINESS MARKETING|3",
                "PMS 604|PUBLIC SECTOR MARKETING|3",
                "PMS 601|FINANCIAL SERVICES MARKETING|3",
                "PMS 611|EDUCATION, HEALTHCARE AND HOSPITALITY MARKETING|3",
                "PMS 602|ENERGY MARKETING|3",
                "PMS 610|PRINTING, COMMUNICATIONS AGENCY AND PUBLISHING MARKETING|3",
                "PMS 609|TELECOMMUNICATIONS AND DIGITAL PLATFORM MARKETING|3",
                "PMS 605|POLITICAL MARKETING|3",
                "PMS 612|SPORTS AND ENTERTAINMENT MARKETING|3",
                "PMS 603|FAST MOVING CONSUMER GOOD MARKETING|3",
            ],
            "Pathway 7": [
                "PMD 701|MARKETING CONSULTANCY PRACTICE|3",
                "PMD 703|PROFESSIONAL SERVICES MARKETING|3",
                "PMD 702|CHANGE AND TRANSFORMATION MARKETING|3",
            ],
        },
        "CIM-UK": {
            "Level 4": [
                "CIM101|Marketing Principles|6",
                "CIM102|Communications in Practice|6",
                "CIM103|Customer Communications|6",
            ],
            "Level 5": [
                "CIM201|Applied Marketing|6",
                "CIM202|Planning Campaigns|6",
                "CIM203|Customer Insights|6",
            ],
            "Level 6": [
                "CIM301|Marketing & Digital Strategy|6",
                "CIM302|Innovation in Marketing|6",
                "CIM303|Resource Management|6",
            ],
            "Level 7": [
                "CIM401|Global Marketing Decisions|6",
                "CIM402|Corporate Digital Communications|6",
                "CIM403|Creating Entrepreneurial Change|6",
            ],
        },
        "ICAG": {
            "Level 1": [
                "ICAG101|Financial Accounting|3",
                "ICAG102|Business Management & Information Systems|3",
                "ICAG103|Business Law|3",
                "ICAG104|Introduction to Management Accounting|3",
            ],
            "Level 2": [
                "ICAG201|Financial Reporting|3",
                "ICAG202|Management Accounting|3",
                "ICAG203|Audit & Assurance|3",
                "ICAG204|Financial Management|3",
                "ICAG205|Corporate Law|3",
                "ICAG206|Public Sector Accounting|3",
            ],
            "Level 3": [
                "ICAG301|Corporate Reporting|3",
                "ICAG302|Advanced Management Accounting|3",
                "ICAG303|Advanced Audit & Assurance|3",
                "ICAG304|Advanced Financial Management|3",
                "ICAG305|Strategy & Governance|3",
                "ICAG306|Advanced Taxation|3",
            ],
        },
        "ACCA": {
            "Level 1 (Applied Knowledge)": [
                "AB101|Accountant in Business|3",
                "MA101|Management Accounting|3",
                "FA101|Financial Accounting|3",
            ],
            "Level 2 (Applied Skills)": [
                "LW201|Corporate and Business Law|3",
                "PM201|Performance Management|3",
                "TX201|Taxation|3",
                "FR201|Financial Reporting|3",
                "AA201|Audit and Assurance|3",
                "FM201|Financial Management|3",
            ],
            "Level 3 Strategic Professional (Essentials)": [
                "SBL301|Strategic Business Leader|6",
                "SBR301|Strategic Business Reporting|6",
            ],
            "Strategic Professional (Options)": [
                "AFM401|Advanced Financial Management|6",
                "APM401|Advanced Performance Management|6",
                "ATX401|Advanced Taxation|6",
                "AAA401|Advanced Audit and Assurance|6",
            ],
        },
    }
    return courses.get(program, {})


def generate_student_info_pdf(data):
    filename = f"student_info_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
    )

    # Styles
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="CustomTitle",
            parent=styles["Heading1"],
            fontSize=16,
            alignment=TA_CENTER,
            spaceAfter=30,
            textColor=colors.HexColor("#003366"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="SectionHeader",
            parent=styles["Heading2"],
            fontSize=12,
            textColor=colors.HexColor("#003366"),
            spaceBefore=15,
            spaceAfter=10,
        )
    )

    elements = []

    # Header with Logo
    header_data = [
        [
            Image("upsa_logo.jpg", width=1.2 * inch, height=1.2 * inch),
            Paragraph("UNIVERSITY OF PROFESSIONAL STUDIES, ACCRA<br/>IPS DIRECTORATE", styles["CustomTitle"]),
            Image("upsa_logo.jpg", width=1.2 * inch, height=1.2 * inch),
        ]
    ]
    # Use RLTable instead of Table
    header_table = RLTable(header_data, [2 * inch, 4 * inch, 2 * inch])
    header_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    elements.append(header_table)
    elements.append(Spacer(1, 20))

    # Document Title
    elements.append(Paragraph("PROFESSIONAL STUDENT'S INFORMATION DOCUMENT", styles["CustomTitle"]))
    elements.append(Spacer(1, 20))

    # Add passport photo if available
    if data["passport_photo_path"]:
        try:
            photo_data = [[Image(data["passport_photo_path"], width=1.5 * inch, height=1.5 * inch)]]
            photo_table = RLTable(photo_data, [1.5 * inch])
            photo_table.setStyle(
                TableStyle(
                    [
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("BOX", (0, 0), (-1, -1), 1, colors.black),
                    ]
                )
            )
            elements.append(photo_table)
            elements.append(Spacer(1, 20))
        except:
            pass

    # Personal Information Section
    elements.append(Paragraph("Personal Information", styles["SectionHeader"]))
    personal_info = [
        ["Student ID:", data["student_id"]],
        ["Surname:", data["surname"]],
        ["Other Names:", data["other_names"]],
        ["Date of Birth:", str(data["date_of_birth"])],
        ["Place of Birth:", data["place_of_birth"]],
        ["Home Town:", data["home_town"]],
        ["Nationality:", data["nationality"]],
        ["Gender:", data["gender"]],
        ["Marital Status:", data["marital_status"]],
        ["Religion:", data["religion"]],
        ["Denomination:", data["denomination"]],
    ]
    t = RLTable(personal_info, [2.5 * inch, 4 * inch])
    t.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f0f0f0")),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#003366")),
                ("PADDING", (0, 0), (-1, -1), 6),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ]
        )
    )
    elements.append(t)
    elements.append(Spacer(1, 20))

    # Contact Information Section
    elements.append(Paragraph("Contact Information", styles["SectionHeader"]))
    contact_info = [
        ["Residential Address:", data["residential_address"]],
        ["Postal Address:", data["postal_address"]],
        ["Email:", data["email"]],
        ["Telephone:", data["telephone"]],
        ["Ghana Card No:", data["ghana_card_id"]],
    ]
    t = RLTable(contact_info, [2.5 * inch, 4 * inch])
    t.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f0f0f0")),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#003366")),
                ("PADDING", (0, 0), (-1, -1), 6),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ]
        )
    )
    elements.append(t)
    elements.append(Spacer(1, 20))

    # Guardian Information Section
    elements.append(Paragraph("Guardian Information", styles["SectionHeader"]))
    guardian_info = [
        ["Name:", data["guardian_name"]],
        ["Relationship:", data["guardian_relationship"]],
        ["Occupation:", data["guardian_occupation"]],
        ["Address:", data["guardian_address"]],
        ["Telephone:", data["guardian_telephone"]],
    ]
    t = RLTable(guardian_info, [2.5 * inch, 4 * inch])
    t.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f0f0f0")),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#003366")),
                ("PADDING", (0, 0), (-1, -1), 6),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ]
        )
    )
    elements.append(t)
    elements.append(Spacer(1, 20))

    if data["receipt_path"]:
        elements.append(Paragraph("Payment Information", styles["SectionHeader"]))
        payment_info = [
            ["Receipt Status:", "Uploaded"],
            ["Receipt Amount:", f"GHS {data['receipt_amount']:.2f}"],
        ]
        t = RLTable(payment_info, [2.5 * inch, 4 * inch])
        t.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f0f0f0")),
                    ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#003366")),
                    ("PADDING", (0, 0), (-1, -1), 6),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ]
            )
        )
        elements.append(t)
        elements.append(Spacer(1, 20))

    # Footer
    footer_style = ParagraphStyle(
        "Footer",
        parent=styles["Normal"],
        fontSize=8,
        textColor=colors.grey,
        alignment=TA_CENTER,
    )
    elements.append(
        Paragraph(
            f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | UPSA Student Information System",
            footer_style,
        )
    )

    doc.build(elements)
    return filename


def generate_course_registration_pdf(data):
    filename = f"course_registration_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
    )

    # Styles
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="CustomTitle",
            parent=styles["Heading1"],
            fontSize=16,
            alignment=TA_CENTER,
            spaceAfter=30,
            textColor=colors.HexColor("#003366"),
        )
    )
    styles.add(
        ParagraphStyle(
            name="SectionHeader",
            parent=styles["Heading2"],
            fontSize=12,
            textColor=colors.HexColor("#003366"),
            spaceBefore=15,
            spaceAfter=10,
        )
    )

    elements = []
    # Get student info from database
    conn = sqlite3.connect("student_registration.db")
    c = conn.cursor()
    c.execute(
        """
        SELECT passport_photo_path, surname, other_names, email 
        FROM student_info 
        WHERE student_id = ?
        """,
        (data["student_id"],),
    )
    student_info = c.fetchone()
    conn.close()

    header_elements = []
    if student_info and student_info[0] and os.path.exists(student_info[0]):
        try:
            with PILImage.open(student_info[0]) as img:
                img.thumbnail((100, 100))
                img_buffer = io.BytesIO()
                img.save(img_buffer, format="JPEG")
                img_buffer.seek(0)
                header_elements.append(Image(img_buffer))
        except Exception as e:
            header_elements.append(Image("upsa_logo.jpg", width=1.2 * inch, height=1.2 * inch))
    else:
        header_elements.append(Image("upsa_logo.jpg", width=1.2 * inch, height=1.2 * inch))

    header_elements.extend(
        [
            Paragraph("UNIVERSITY OF PROFESSIONAL STUDIES, ACCRA<br/>Proof of Registration", styles["CustomTitle"]),
            Image("upsa_logo.jpg", width=1.2 * inch, height=1.2 * inch),
        ]
    )
    header_table = RLTable([header_elements], [2 * inch, 4 * inch, 2 * inch])
    header_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    elements.append(header_table)
    elements.append(Spacer(1, 20))

    if student_info:
        student_details = [
            [Paragraph(f"<b>Name:</b> {student_info[1]} {student_info[2]}", styles["Normal"])],
            [Paragraph(f"<b>Email:</b> {student_info[3]}", styles["Normal"])],
        ]
        student_table = RLTable(student_details, [7 * inch])
        student_table.setStyle(
            TableStyle(
                [
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("PADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        elements.append(student_table)
        elements.append(Spacer(1, 20))

    elements.append(Paragraph("Registration Details", styles["SectionHeader"]))
    reg_info = [
        ["Student ID:", data["student_id"]],
        ["Index Number:", data["index_number"]],
        ["Programme:", data["programme"]],
        ["Specialization:", data["specialization"]],
        ["Level:", data["level"]],
        ["Session:", data["session"]],
        ["Academic Year:", data["academic_year"]],
        ["Semester:", data["semester"]],
    ]
    t = RLTable(reg_info, [2.5 * inch, 4 * inch])
    t.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f0f0f0")),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#003366")),
                ("PADDING", (0, 0), (-1, -1), 6),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ]
        )
    )
    elements.append(t)
    elements.append(Spacer(1, 20))

    elements.append(Paragraph("Selected Courses", styles["SectionHeader"]))
    courses_list = data["courses"].split("\n")
    courses_data = [["Course Code", "Course Title", "Credit Hours"]]
    for course in courses_list:
        if "|" in course:
            courses_data.append(course.split("|"))
    t = RLTable(courses_data, [2 * inch, 3.5 * inch, 1 * inch])
    t.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#003366")),
                ("TEXTCOLOR", (0, 0), (0, 0), colors.white),
                ("PADDING", (0, 0), (-1, -1), 6),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (-1, 0), (-1, -1), "CENTER"),
            ]
        )
    )
    elements.append(t)
    elements.append(Paragraph(f"<b>Total Credit Hours:</b> {data['total_credits']}", ParagraphStyle("TotalCredits", parent=styles["Normal"], fontSize=10, textColor=colors.HexColor("#003366"), alignment=TA_LEFT, spaceBefore=10)))
    elements.append(Spacer(1, 30))

    if data.get("receipt_path"):
        elements.append(Paragraph("Payment Information", styles["SectionHeader"]))
        payment_info = [
            ["Receipt Status:", "Uploaded"],
            ["Receipt Amount:", f"GHS {data.get('receipt_amount', 0.0):.2f}"],
        ]
        t = RLTable(payment_info, [2.5 * inch, 4 * inch])
        t.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f0f0f0")),
                    ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#003366")),
                    ("PADDING", (0, 0), (-1, -1), 6),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ]
            )
        )
        elements.append(t)
        elements.append(Spacer(1, 20))

    signature_data = [
        ["_______________________", "_______________________"],
        ["Student's Signature", "IPS Directorate Officer"],
        ["Date: ________________", "Date: ________________"],
    ]
    sig_table = RLTable(signature_data, [4 * inch, 4 * inch])
    sig_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
                ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor("#003366")),
                ("FONTSIZE", (0, 1), (-1, 1), 8),
                ("TOPPADDING", (0, 2), (-1, 2), 20),
            ]
        )
    )
    elements.append(sig_table)
    elements.append(Spacer(1, 30))
    elements.append(
        Paragraph(
            f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | UPSA Course Registration System",
            ParagraphStyle("Footer", parent=styles["Normal"], fontSize=8, textColor=colors.grey, alignment=TA_CENTER),
        )
    )

    doc.build(elements)
    return filename


def review_student_info(form_data, uploaded_files):
    st.subheader("Review Student Information")

    cols = st.columns(2)
    with cols[0]:
        st.write("**Personal Information**")
        st.write(f"Student ID: {form_data['student_id']}")
        st.write(f"Surname: {form_data['surname']}")
        st.write(f"Other Names: {form_data['other_names']}")
        st.write(f"Date of Birth: {form_data['date_of_birth']}")
        st.write(f"Place of Birth: {form_data['place_of_birth']}")
        st.write(f"Home Town: {form_data['home_town']}")
        st.write(f"Nationality: {form_data['nationality']}")
        st.write(f"Gender: {form_data['gender']}")
        st.write(f"Marital Status: {form_data['marital_status']}")
        st.write(f"Religion: {form_data['religion']}")
        st.write(f"Denomination: {form_data['denomination']}")

    with cols[1]:
        st.write("**Contact Information**")
        st.write(f"Residential Address: {form_data['residential_address']}")
        st.write(f"Postal Address: {form_data['postal_address']}")
        st.write(f"Email: {form_data['email']}")
        st.write(f"Telephone: {form_data['telephone']}")
        st.write(f"Ghana Card No: {form_data['ghana_card_id']}")

    st.write("**Guardian Information**")
    st.write(f"Name: {form_data['guardian_name']}")
    st.write(f"Relationship: {form_data['guardian_relationship']}")
    st.write(f"Occupation: {form_data['guardian_occupation']}")
    st.write(f"Address: {form_data['guardian_address']}")
    st.write(f"Telephone: {form_data['guardian_telephone']}")

    st.write("**Educational Background**")
    st.write(f"Previous School: {form_data['previous_school']}")
    st.write(f"Qualification: {form_data['qualification_type']}")
    st.write(f"Completion Year: {form_data['completion_year']}")
    st.write(f"Aggregate Score: {form_data['aggregate_score']}")

    st.write("**Uploaded Documents**")
    for doc_name, file in uploaded_files.items():
        if doc_name == "Receipt":
            if file:
                st.write(f"✅ {doc_name} uploaded (Optional)")
            else:
                st.write(f"⚪ {doc_name} not uploaded (Optional)")
        else:
            if file:
                st.write(f"✅ {doc_name} uploaded")
            else:
                st.write(f"❌ {doc_name} not uploaded")


def review_course_registration(form_data):
    col1, col2 = st.columns(2)

    with col1:
        st.write("**Student Information**")
        st.write(f"Student ID: {form_data['student_id']}")
        st.write(f"Programme: {form_data['programme']}")
        st.write(f"Level: {form_data['level']}")
        st.write(f"Specialization: {form_data['specialization']}")

    with col2:
        st.write("**Registration Details**")
        st.write(f"Session: {form_data['session']}")
        st.write(f"Academic Year: {form_data['academic_year']}")
        st.write(f"Semester: {form_data['semester']}")

    st.write("**Selected Courses**")
    if form_data["courses"]:
        courses_list = form_data["courses"].split("\n")

        table_data = []
        for course in courses_list:
            if "|" in course:
                code, title, credits = course.split("|")
                table_data.append([code, title, f"{credits} credits"])

        if table_data:
            df = pd.DataFrame(
                table_data, columns=["Course Code", "Course Title", "Credit Hours"]
            )
            st.table(df)

            st.write(f"**Total Credit Hours:** {form_data['total_credits']}")
    else:
        st.warning("No courses selected")


def save_student_info(form_data):
    with sqlite3.connect("student_registration.db") as conn:
        try:
            cursor = conn.cursor()
            # ... database operations ...
            conn.commit()
        except sqlite3.Error as e:
            conn.rollback()
            raise


def student_info_form():
    st.header("📝 Student Information Form")

    form_data = {}

    st.subheader("Personal Information")
    col1, col2 = st.columns(2)

    with col1:
        form_data["student_id"] = st.text_input("Student ID")
        form_data["surname"] = st.text_input("Surname")
        form_data["other_names"] = st.text_input("First & Middle Names")
        form_data["date_of_birth"] = st.date_input("Date of Birth")
        form_data["place_of_birth"] = st.text_input("Place of Birth")
        form_data["home_town"] = st.text_input("Home Town")
        form_data["nationality"] = st.text_input("Nationality")

    with col2:
        form_data["gender"] = st.selectbox("Gender", ["Male", "Female", "Other"])
        form_data["marital_status"] = st.selectbox(
            "Marital Status", ["Single", "Married", "Divorced", "Widowed"]
        )
        form_data["religion"] = st.text_input("Religion")
        form_data["denomination"] = st.text_input("Denomination")
        disability_status = st.selectbox("Disability Status", ["None", "Yes"])
        form_data["disability_status"] = disability_status
        if disability_status == "Yes":
            form_data["disability_description"] = st.text_area("Disability Description")
        else:
            form_data["disability_description"] = "None"

    st.subheader("Contact Information")
    col3, col4 = st.columns(2)

    with col3:
        form_data["residential_address"] = st.text_area("Residential Address")
        form_data["postal_address"] = st.text_area("Postal Address")
        form_data["email"] = st.text_input("Email Address")

    with col4:
        form_data["telephone"] = st.text_input("Telephone Number")
        form_data["ghana_card_id"] = st.text_input("Ghana Card ID Number")

    st.subheader("Guardian Information")
    col5, col6 = st.columns(2)

    with col5:
        form_data["guardian_name"] = st.text_input("Guardian's Name")
        form_data["guardian_relationship"] = st.text_input("Relationship to Guardian")
        form_data["guardian_occupation"] = st.text_input("Guardian's Occupation")

    with col6:
        form_data["guardian_address"] = st.text_area("Guardian's Address")
        form_data["guardian_telephone"] = st.text_input("Guardian's Telephone")

    st.subheader("Educational Background")
    col7, col8 = st.columns(2)

    with col7:
        form_data["previous_school"] = st.text_input("Previous School")
        form_data["qualification_type"] = st.text_input("Qualification Type")

    with col8:
        form_data["completion_year"] = st.text_input("Year of Completion")
        form_data["aggregate_score"] = st.text_input("Aggregate Score")

    st.subheader("📎 Required Documents")
    col9, col10 = st.columns(2)

    with col9:
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        ghana_card = st.file_uploader("Upload Ghana Card", type=["pdf", "jpg", "png"])
        passport_photo = st.file_uploader("Upload Passport Photo", type=["jpg", "png"])
        transcript = st.file_uploader("Upload Transcript", type=["pdf"])
        st.markdown("</div>", unsafe_allow_html=True)

    with col10:
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        certificate = st.file_uploader("Upload Certificate", type=["pdf"])
        st.write("Optional Payment Receipt")
        receipt = st.file_uploader(
            "Upload Payment Receipt (Optional)", type=["pdf", "jpg", "png"]
        )
        st.markdown("</div>", unsafe_allow_html=True)

        if receipt:
            receipt_amount = st.number_input(
                "Receipt Amount (GHS)", min_value=0.0, format="%.2f"
            )
            if receipt_amount < 100.0:
                st.warning(
                    "Receipt amount seems low. Please verify the payment amount."
                )

    uploaded_files = {
        "Ghana Card": ghana_card,
        "Passport Photo": passport_photo,
        "Transcript/ Results": transcript,
        "Certificate": certificate,
        "Receipt": receipt,
    }

    col_buttons = st.columns([2, 2, 1])

    with col_buttons[0]:
        if st.button("Review Information", use_container_width=True):
            st.session_state.review_mode = True
            st.session_state.form_data = form_data
            st.session_state.uploaded_files = uploaded_files
            st.rerun()

    if "review_mode" in st.session_state and st.session_state.review_mode:
        review_student_info(st.session_state.form_data, st.session_state.uploaded_files)

        with col_buttons[1]:
            if st.button("Edit Information", use_container_width=True):
                st.session_state.review_mode = False
                st.rerun()

        with col_buttons[2]:
            if st.button("Confirm and Submit", use_container_width=True):
                ghana_card_path = save_uploaded_file(ghana_card, "uploads")
                passport_photo_path = save_uploaded_file(passport_photo, "uploads")
                transcript_path = save_uploaded_file(transcript, "uploads")
                certificate_path = save_uploaded_file(certificate, "uploads")
                receipt_path = (
                    save_uploaded_file(receipt, "uploads") if receipt else None
                )

                file_paths = {
                    "ghana_card_path": ghana_card_path,
                    "passport_photo_path": passport_photo_path,
                    "transcript_path": transcript_path,
                    "certificate_path": certificate_path,
                    "receipt_path": receipt_path,
                }

                if receipt and "receipt_amount" in locals():
                    form_data["receipt_amount"] = receipt_amount

                try:
                    conn = sqlite3.connect("student_registration.db")
                    c = conn.cursor()
                    insert_student_info(c, form_data, file_paths)
                    conn.commit()
                    st.success(
                        "Information submitted successfully! Pending admin approval."
                    )
                    st.session_state.review_mode = False

                except sqlite3.IntegrityError:
                    st.error("Student ID already exists!")
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                finally:
                    conn.close()


def validate_file(uploaded_file, max_size_mb=5):
    if uploaded_file.size > max_size_mb * 1024 * 1024:
        raise ValueError(f"File size exceeds {max_size_mb}MB limit")


def download_all_documents():
    temp_dir = "temp_downloads"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    try:
        conn = sqlite3.connect("student_registration.db")
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT student_id, surname, other_names, 
                ghana_card_path, passport_photo_path, 
                transcript_path, certificate_path, receipt_path,
                receipt_amount
            FROM student_info
        """
        )
        students = cursor.fetchall()

        cursor.execute(
            """
            SELECT cr.registration_id, cr.student_id, si.surname, si.other_names,
                   cr.receipt_path, cr.receipt_amount
            FROM course_registration cr
            LEFT JOIN student_info si ON cr.student_id = si.student_id
            WHERE cr.receipt_path IS NOT NULL
        """
        )
        registrations = cursor.fetchall()

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"all_documents_{timestamp}.zip"

        with zipfile.ZipFile(zip_filename, "w") as zipf:
            for student in students:
                student_id, surname, other_names = student[:3]
                documents = student[3:8]
                doc_names = [
                    "ghana_card",
                    "passport_photo",
                    "transcript",
                    "certificate",
                    "receipt",
                ]
                student_dir = f"student_documents/{student_id}_{surname}_{other_names}"
                for doc_path, doc_name in zip(documents, doc_names):
                    if doc_path and os.path.exists(doc_path):
                        _, ext = os.path.splitext(doc_path)
                        archive_path = f"{student_dir}/{doc_name}{ext}"
                        zipf.write(doc_path, archive_path)

            for registration in registrations:
                (
                    reg_id,
                    student_id,
                    surname,
                    other_names,
                    receipt_path,
                    receipt_amount,
                ) = registration
                if receipt_path and os.path.exists(receipt_path):
                    reg_dir = f"course_registration_receipts/{student_id}_{surname}_{other_names}"
                    _, ext = os.path.splitext(receipt_path)
                    archive_path = f"{reg_dir}/registration_{reg_id}_receipt{ext}"
                    zipf.write(receipt_path, archive_path)

        return zip_filename

    except Exception as e:
        st.error(f"Error creating zip file: {str(e)}")
        return None

    finally:
        conn.close()
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


def course_registration_form():
    st.header("📚 Course Registration Form (A7)")

    form_data = {}

    form_data["student_id"] = st.text_input("Student ID")

    if form_data["student_id"]:
        conn = sqlite3.connect("student_registration.db")
        c = conn.cursor()

        c.execute(
            "SELECT * FROM student_info WHERE student_id = ?",
            (form_data["student_id"],),
        )
        student_info = c.fetchone()

        if student_info:
            st.markdown("---")
            col_photo, col_info = st.columns([1, 3])

            with col_photo:
                if student_info[28]:
                    try:
                        image = PILImage.open(student_info[28])
                        st.image(image, caption="Student Photo", width=150)
                    except Exception as e:
                        st.error(f"Error loading passport photo: {str(e)}")
                else:
                    st.warning("No passport photo available")

            with col_info:
                st.markdown(f"### {student_info[1]} {student_info[2]}")
                st.write(f"**Student ID:** {student_info[0]}")
                st.write(f"**Email:** {student_info[8]}")
                st.write(f"**Phone:** {student_info[9]}")

            col3, col4 = st.columns(2)
            with col3:
                form_data["programme"] = st.selectbox(
                    "Programme", ["CIMG", "CIM-UK", "ICAG", "ACCA"]
                )
                program_levels = list(
                    get_program_courses(form_data["programme"]).keys()
                )
                form_data["level"] = st.selectbox("Level/Part", program_levels)
                form_data["specialization"] = st.text_input("Specialization (Optional)")
            with col4:
                form_data["session"] = st.selectbox(
                    "Session", ["Morning", "Evening", "Weekend"]
                )
                form_data["academic_year"] = st.selectbox(
                    "Academic Year", [f"{year}-{year+1}" for year in range(2025, 2035)]
                )
                form_data["semester"] = st.selectbox(
                    "Semester", ["First", "Second", "Third"]
                )

            st.subheader("Course Selection")
            available_courses = get_program_courses(form_data["programme"]).get(
                form_data["level"], []
            )
            selected_courses = st.multiselect(
                "Select Courses",
                available_courses,
                format_func=lambda x: f"{x.split('|')[0]} - {x.split('|')[1]} ({x.split('|')[2]} credits)",
            )

            total_credits = sum(
                [int(course.split("|")[2]) for course in selected_courses]
            )
            form_data["courses"] = "\n".join(selected_courses)
            form_data["total_credits"] = total_credits

            st.text_area(
                "Selected Courses", form_data["courses"], height=150, disabled=True
            )
            st.number_input(
                "Total Credit Hours",
                value=total_credits,
                min_value=0,
                max_value=24,
                disabled=True,
            )

            if total_credits > 24:
                st.error("Total credits cannot exceed 24 hours!")
                return

            st.subheader("📎 Payment Information (Optional)")
            col5, col6 = st.columns(2)
            with col5:
                receipt = st.file_uploader(
                    "Upload Payment Receipt (Optional)", type=["pdf", "jpg", "png"]
                )
                form_data["receipt_path"] = (
                    save_uploaded_file(receipt, "uploads") if receipt else None
                )
            with col6:
                form_data["receipt_amount"] = (
                    st.number_input(
                        "Receipt Amount (GHS)", min_value=0.0, format="%.2f"
                    )
                    if receipt
                    else 0.0
                )

            if st.button("Review Registration"):
                review_course_registration(form_data)

            if st.button("Confirm and Submit"):
                try:
                    conn = sqlite3.connect("student_registration.db")
                    c = conn.cursor()
                    c.execute(
                        """
                        INSERT INTO course_registration 
                        (student_id, programme, specialization, level, 
                        session, academic_year, semester, courses, total_credits, 
                        date_registered, approval_status, receipt_path, receipt_amount)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                        (
                            form_data["student_id"],
                            form_data["programme"],
                            form_data["specialization"],
                            form_data["level"],
                            form_data["session"],
                            form_data["academic_year"],
                            form_data["semester"],
                            form_data["courses"],
                            form_data["total_credits"],
                            datetime.now().date(),
                            "pending",
                            form_data["receipt_path"],
                            form_data["receipt_amount"],
                        ),
                    )
                    conn.commit()
                    st.success("Course registration submitted! Pending admin approval.")
                except sqlite3.IntegrityError:
                    st.error(
                        "Error in registration. Please check if student ID exists."
                    )
                finally:
                    conn.close()
        else:
            st.warning(
                "No matching student record found. Please verify the Student ID."
            )
            return

        conn.close()


def admin_dashboard():
    st.title("Admin Dashboard")

    menu = st.sidebar.selectbox(
        "Menu",
        [
            "Upload Data",
            "Student Records",
            "Course Registrations",
            "Programs",
            "Database Management",
            "Pending Approvals",
            "Generate Reports",
            "Send Emails",
            "Notifications",
        ],  # Added Notifications
    )

    if menu == "Upload Data":
        upload_data_from_excel_and_docs()
    elif menu == "Student Records":
        manage_student_records()
    elif menu == "Course Registrations":
        manage_course_registrations()
    elif menu == "Programs":
        manage_programs()
    elif menu == "Database Management":
        manage_database()
    elif menu == "Pending Approvals":
        show_pending_approvals()
    elif menu == "Generate Reports":
        generate_reports()
    elif menu == "Send Emails":
        send_emails()
    elif menu == "Notifications":
        admin_notification_interface()



def zip_uploads_folder():
    """
    Zips the entire 'uploads' folder preserving its structure exactly as is.
    
    Returns:
        The filename of the generated zip file, or None if the uploads folder does not exist.
    """
    uploads_dir = "uploads"
    if not os.path.exists(uploads_dir):
        return None
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"uploads_folder_{timestamp}"
    zip_filename = f"{base_name}.zip"
    # This will create a zip archive with the base directory inside it as it is.
    shutil.make_archive(base_name=base_name, format="zip", root_dir=uploads_dir)
    return zip_filename


def manage_database():
    st.subheader("Database Management")

    col1, col2, col3, col4 = st.columns(4)  # Added a new column for the new button

    with col1:
        st.write("### Export Complete Database")
        if st.button("Download Complete Database (Excel)"):
            try:
                export_dir = "temp_export"
                if not os.path.exists(export_dir):
                    os.makedirs(export_dir)

                conn = sqlite3.connect("student_registration.db")

                tables = {
                    "student_info": pd.read_sql_query(
                        """
                        SELECT *, 
                            CASE WHEN receipt_path IS NOT NULL THEN 'Yes' ELSE 'No' END as has_receipt,
                            receipt_amount
                        FROM student_info
                    """,
                        conn,
                    ),
                    "course_registration": pd.read_sql_query(
                        """
                        SELECT *,
                            CASE WHEN receipt_path IS NOT NULL THEN 'Yes' ELSE 'No' END as has_receipt,
                            receipt_amount
                        FROM course_registration
                    """,
                        conn,
                    ),
                }

                excel_files = []
                for table_name, df in tables.items():
                    excel_filename = os.path.join(export_dir, f"{table_name}.xlsx")
                    with pd.ExcelWriter(excel_filename, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name=table_name)
                        workbook = writer.book
                        worksheet = writer.sheets[table_name]

                        for column_cells in worksheet.columns:
                            max_length = 0
                            column = column_cells[0].column_letter
                            for cell in column_cells:
                                try:
                                    cell_length = len(str(cell.value))
                                    if cell_length > max_length:
                                        max_length = cell_length
                                except Exception:
                                    pass
                            adjusted_width = max_length + 2
                            worksheet.column_dimensions[column].width = adjusted_width
                    excel_files.append(excel_filename)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                zip_filename = f"complete_database_{timestamp}.zip"
                with zipfile.ZipFile(zip_filename, "w") as zipf:
                    for file in excel_files:
                        arcname = os.path.basename(file)
                        zipf.write(file, arcname)

                for file in excel_files:
                    os.remove(file)
                os.rmdir(export_dir)

                conn.close()

                with open(zip_filename, "rb") as f:
                    st.download_button(
                        label="Download Database ZIP",
                        data=f,
                        file_name=zip_filename,
                        mime="application/zip",
                    )
                os.remove(zip_filename)
            except Exception as e:
                st.error(f"Error exporting database: {str(e)}")

    with col2:
        st.write("### Download All Documents")
        if st.button("Download All Documents"):
            with st.spinner("Creating zip file of all documents..."):
                zip_file = download_all_documents()
                if zip_file and os.path.exists(zip_file):
                    with open(zip_file, "rb") as f:
                        st.download_button(
                            label="Download Documents ZIP",
                            data=f,
                            file_name=zip_file,
                            mime="application/zip",
                        )
                    os.remove(zip_file)
                else:
                    st.error("Error creating zip file or no documents found")

    with col3:
        st.write("### Download All Receipts")
        if st.button("Download All Receipts"):
            with st.spinner("Creating zip file of all receipts..."):
                zip_file = download_receipts()
                if zip_file and os.path.exists(zip_file):
                    with open(zip_file, "rb") as f:
                        st.download_button(
                            label="Download Receipts ZIP",
                            data=f,
                            file_name=zip_file,
                            mime="application/zip",
                        )
                    os.remove(zip_file)
                else:
                    st.error("Error creating zip file or no receipts found")

    with col4:
        st.write("### Download Uploads Folder")
        if st.button("Download Uploads Folder"):
            with st.spinner("Creating zip file of the uploads folder..."):
                zip_file = zip_uploads_folder()
                if zip_file and os.path.exists(zip_file):
                    with open(zip_file, "rb") as f:
                        st.download_button(
                            label="Download Uploads ZIP",
                            data=f,
                            file_name=zip_file,
                            mime="application/zip",
                        )
                    os.remove(zip_file)
                else:
                    st.error("Uploads folder not found or error creating zip")

    st.write("### Generate All PDFs")
    col_pdfs1, col_pdfs2 = st.columns(2)
    with col_pdfs1:
        if st.button("Generate All Student Info PDFs"):
            with st.spinner("Generating student information PDFs..."):
                zip_file = generate_batch_pdfs("student_info")
                if zip_file and os.path.exists(zip_file):
                    with open(zip_file, "rb") as f:
                        st.download_button(
                            label="Download Student Info PDFs",
                            data=f,
                            file_name=zip_file,
                            mime="application/zip",
                        )
                    os.remove(zip_file)
                else:
                    st.error("Error generating PDFs")
    with col_pdfs2:
        if st.button("Generate All Course Registration PDFs"):
            with st.spinner("Generating course registration PDFs..."):
                zip_file = generate_batch_pdfs("course_registration")
                if zip_file and os.path.exists(zip_file):
                    with open(zip_file, "rb") as f:
                        st.download_button(
                            label="Download Course Registration PDFs",
                            data=f,
                            file_name=zip_file,
                            mime="application/zip",
                        )
                    os.remove(zip_file)
                else:
                    st.error("Error generating PDFs")


def upload_data_from_excel_and_docs():
    """
    Bulk upload function updated to accept two Excel files:
      - One for student information (now preserving document path values if provided)
      - One for course registration data (authoritative for programme and other registration data)
    Optionally uploads a zip file for related documents. 
    The zip file is saved to the uploads folder as-is.
    """
    import pandas as pd
    import sqlite3
    import os
    import shutil
    import streamlit as st
    from datetime import datetime

    st.header("Bulk Upload Data from Excel & Documents")

    # Use two file uploaders: one for student info and one for course registration data.
    st.markdown("### Excel Files Upload")
    col1, col2 = st.columns(2)
    with col1:
        student_excel = st.file_uploader(
            "Upload Excel File with Student Data",
            type=["xlsx", "xls"],
            key="student_excel",
        )
    with col2:
        reg_excel = st.file_uploader(
            "Upload Excel File with Course Registration Data",
            type=["xlsx", "xls"],
            key="reg_excel",
        )

    docs_zip = st.file_uploader("Upload Zip File of Documents (Optional)", type=["zip"])

    if st.button("Process Bulk Upload"):
        if student_excel is None or reg_excel is None:
            st.error(
                "Please upload both Excel files: one for student information and one for course registration data."
            )
            return

        try:
            # Read the two Excel files.
            student_df = pd.read_excel(student_excel)
            reg_df = pd.read_excel(reg_excel)

            conn = sqlite3.connect("student_registration.db")
            c = conn.cursor()

            # Updated: Insert student data including document path fields.
            # Expected Excel columns: student_id, surname, other_names, date_of_birth, place_of_birth,
            # home_town, residential_address, postal_address, email, telephone, ghana_card_id, nationality,
            # marital_status, gender, religion, denomination, disability_status, disability_description,
            # guardian_name, guardian_relationship, guardian_occupation, guardian_address, guardian_telephone,
            # previous_school, qualification_type, completion_year, aggregate_score,
            # ghana_card_path, passport_photo_path, transcript_path, certificate_path, receipt_path.
            insert_student_query = """
                INSERT OR IGNORE INTO student_info (
                    student_id, surname, other_names, date_of_birth, place_of_birth,
                    home_town, residential_address, postal_address, email, telephone,
                    ghana_card_id, nationality, marital_status, gender, religion,
                    denomination, disability_status, disability_description,
                    guardian_name, guardian_relationship, guardian_occupation,
                    guardian_address, guardian_telephone, previous_school,
                    qualification_type, completion_year, aggregate_score,
                    ghana_card_path, passport_photo_path, transcript_path, certificate_path, receipt_path,
                    programme
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """
            for _, row in student_df.iterrows():
                params = (
                    row.get("student_id"),
                    row.get("surname"),
                    row.get("other_names"),
                    row.get("date_of_birth"),
                    row.get("place_of_birth"),
                    row.get("home_town"),
                    row.get("residential_address"),
                    row.get("postal_address"),
                    row.get("email"),
                    row.get("telephone"),
                    row.get("ghana_card_id"),
                    row.get("nationality"),
                    row.get("marital_status"),
                    row.get("gender"),
                    row.get("religion"),
                    row.get("denomination"),
                    row.get("disability_status", "None"),
                    row.get("disability_description", "None"),
                    row.get("guardian_name"),
                    row.get("guardian_relationship"),
                    row.get("guardian_occupation"),
                    row.get("guardian_address"),
                    row.get("guardian_telephone"),
                    row.get("previous_school"),
                    row.get("qualification_type"),
                    row.get("completion_year"),
                    row.get("aggregate_score"),
                    row.get("ghana_card_path"),
                    row.get("passport_photo_path"),
                    row.get("transcript_path"),
                    row.get("certificate_path"),
                    row.get("receipt_path"),
                    ""  # Programme will be updated from course registration data
                )
                c.execute(insert_student_query, params)

            # Insert course registration data.
            insert_reg_query = """
                INSERT INTO course_registration (
                    student_id, index_number, programme, specialization, level,
                    session, academic_year, semester, courses, total_credits,
                    date_registered, approval_status, receipt_path, receipt_amount
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """
            for _, row in reg_df.iterrows():
                params = (
                    row.get("student_id"),
                    row.get("index_number"),
                    row.get("programme"),  # authoritative programme data
                    row.get("specialization"),
                    row.get("level"),
                    row.get("session"),
                    row.get("academic_year"),
                    row.get("semester"),
                    row.get("courses"),
                    row.get("total_credits"),
                    row.get("date_registered"),
                    row.get("approval_status", "pending"),
                    row.get("receipt_path"),
                    row.get("receipt_amount", 0.0),
                )
                c.execute(insert_reg_query, params)
                # Update the student's programme field based on the registration Excel.
                update_query = "UPDATE student_info SET programme = ? WHERE student_id = ?"
                c.execute(update_query, (row.get("programme"), row.get("student_id")))

            conn.commit()
            conn.close()

            st.success("Excel data uploaded successfully!")
        except Exception as e:
            st.error(f"Error processing Excel files: {e}")

        # Process the documents zip file if provided.
        if docs_zip:
            # Save the zip file as-is in the uploads folder.
            saved_zip_path = save_uploaded_file(docs_zip, "uploads")
            st.success("Documents zip uploaded successfully!")



def generate_reports():
    """
    Generate various types of reports including:
    - Student Statistics
    - Course Registration Summary
    - Approval Status Summary
    - Payment Statistics
    """
    st.subheader("Generate Reports")

    # Report type selection
    report_type = st.selectbox(
        "Select Report Type",
        [
            "Student Statistics",
            "Course Registration Summary",
            "Approval Status Summary",
            "Payment Statistics",
        ],
    )

    # Connect to the database
    conn = sqlite3.connect("student_registration.db")

    if report_type == "Student Statistics":
        # Gender distribution
        gender_dist = pd.read_sql_query(
            "SELECT gender, COUNT(*) as count FROM student_info GROUP BY gender", conn
        )
        if not gender_dist.empty:
            st.write("**Gender Distribution**")
            fig = px.pie(
                gender_dist, names="gender", values="count", title="Gender Distribution"
            )
            st.plotly_chart(fig)
        else:
            st.info("No gender distribution data available")

        # Programme distribution
        prog_dist = pd.read_sql_query(
            "SELECT programme, COUNT(*) as count FROM course_registration GROUP BY programme",
            conn,
        )
        if not prog_dist.empty:
            st.write("**Programme Distribution**")
            fig = px.pie(
                prog_dist,
                names="programme",
                values="count",
                title="Programme Distribution",
            )
            st.plotly_chart(fig)
        else:
            st.info("No programme distribution data available")

    elif report_type == "Course Registration Summary":
        summary = pd.read_sql_query(
            """
            SELECT cr.programme, cr.level, cr.semester, 
                   COUNT(*) as registrations,
                   AVG(cr.total_credits) as avg_credits
            FROM course_registration cr
            GROUP BY cr.programme, cr.level, cr.semester
            """,
            conn,
        )
        if not summary.empty:
            st.write("**Course Registration Summary**")
            st.dataframe(summary)
        else:
            st.info("No course registration data available")

    elif report_type == "Approval Status Summary":
        # Student approval status
        student_status = pd.read_sql_query(
            "SELECT approval_status, COUNT(*) as count FROM student_info GROUP BY approval_status",
            conn,
        )
        # Course approval status
        course_status = pd.read_sql_query(
            "SELECT approval_status, COUNT(*) as count FROM course_registration GROUP BY approval_status",
            conn,
        )

        col1, col2 = st.columns(2)
        with col1:
            if not student_status.empty:
                st.write("**Student Info Approval Status**")
                fig = px.pie(
                    student_status,
                    names="approval_status",
                    values="count",
                    title="Student Info Approval Status",
                )
                st.plotly_chart(fig)
            else:
                st.info("No student approval status data available")

        with col2:
            if not course_status.empty:
                st.write("**Course Registration Approval Status**")
                fig = px.pie(
                    course_status,
                    names="approval_status",
                    values="count",
                    title="Course Registration Approval Status",
                )
                st.plotly_chart(fig)
            else:
                st.info("No course registration approval status data available")

    elif report_type == "Payment Statistics":
        # Receipt statistics
        receipt_stats = pd.read_sql_query(
            """
            SELECT 
                CASE 
                    WHEN receipt_path IS NOT NULL THEN 'With Receipt'
                    ELSE 'Without Receipt'
                END as receipt_status,
                COUNT(*) as count,
                COALESCE(AVG(CASE WHEN receipt_amount IS NOT NULL THEN receipt_amount ELSE 0 END), 0) as avg_amount
            FROM student_info
            GROUP BY CASE 
                        WHEN receipt_path IS NOT NULL THEN 'With Receipt'
                        ELSE 'Without Receipt'
                    END
            """,
            conn,
        )

        st.write("**Receipt Statistics**")
        if not receipt_stats.empty:
            col1, col2 = st.columns(2)
            with col1:
                fig = px.pie(
                    receipt_stats,
                    names="receipt_status",
                    values="count",
                    title="Receipt Upload Distribution",
                )
                st.plotly_chart(fig)

            with col2:
                with_receipt_data = receipt_stats[
                    receipt_stats["receipt_status"] == "With Receipt"
                ]
                if not with_receipt_data.empty:
                    avg_amount = with_receipt_data["avg_amount"].iloc[0]
                    st.write(f"Average Receipt Amount: GHS {avg_amount:.2f}")
                else:
                    st.write("No receipt data available")

                # Additional payment statistics
                total_payments = pd.read_sql_query(
                    """
                    SELECT 
                        COUNT(*) as total_receipts,
                        COALESCE(SUM(receipt_amount), 0) as total_amount
                    FROM student_info 
                    WHERE receipt_path IS NOT NULL
                    """,
                    conn,
                )
                st.write("**Payment Summary**")
                st.write(f"Total Receipts: {total_payments['total_receipts'].iloc[0]}")
                st.write(
                    f"Total Amount: GHS {total_payments['total_amount'].iloc[0]:.2f}"
                )
        else:
            st.info("No payment statistics available")

    # Close the database connection
    conn.close()


def show_pending_approvals():
    st.subheader("Pending Approvals")

    tabs = st.tabs(["Student Information", "Course Registrations"])

    conn = sqlite3.connect("student_registration.db")

    try:
        with tabs[0]:
            pending_students = pd.read_sql_query(
                "SELECT * FROM student_info WHERE approval_status='pending'", conn
            )

            if pending_students.empty:
                st.info("No pending student applications")
            else:
                for _, student in pending_students.iterrows():
                    with st.expander(
                        f"Student: {student['surname']} {student['other_names']}"
                    ):
                        col1, col2 = st.columns(2)

                        with col1:
                            st.write("**Personal Information**")
                            st.write(f"Student ID: {student['student_id']}")
                            st.write(
                                f"Name: {student['surname']} {student['other_names']}"
                            )
                            st.write(f"Gender: {student['gender']}")
                            st.write(f"Email: {student['email']}")
                            st.write(f"Phone: {student['telephone']}")

                        with col2:
                            st.write("**Educational Background**")
                            st.write(f"Previous School: {student['previous_school']}")
                            st.write(f"Qualification: {student['qualification_type']}")
                            st.write(f"Completion Year: {student['completion_year']}")

                            st.write("**Payment Information**")
                            if student["receipt_path"]:
                                st.write(
                                    f"Receipt Amount: GHS {student['receipt_amount']:.2f}"
                                )
                                if os.path.exists(student["receipt_path"]):
                                    st.write(
                                        f"[View Receipt]({student['receipt_path']})"
                                    )
                            else:
                                st.write("No receipt uploaded (Optional)")

                        if student["passport_photo_path"] and os.path.exists(
                            student["passport_photo_path"]
                        ):
                            try:
                                image = PILImage.open(student["passport_photo_path"])
                                st.image(image, width=150, caption="Passport Photo")
                            except Exception as e:
                                st.error(f"Error loading passport photo: {str(e)}")

                        col1, col2 = st.columns(2)
                        with col1:
                            if st.button(
                                "Approve", key=f"approve_{student['student_id']}"
                            ):
                                c = conn.cursor()
                                c.execute(
                                    "UPDATE student_info SET approval_status='approved' WHERE student_id=?",
                                    (student["student_id"],),
                                )
                                conn.commit()
                                st.success("Application Approved!")
                                st.rerun()
                        with col2:
                            if st.button(
                                "Reject", key=f"reject_{student['student_id']}"
                            ):
                                c = conn.cursor()
                                c.execute(
                                    "UPDATE student_info SET approval_status='rejected' WHERE student_id=?",
                                    (student["student_id"],),
                                )
                                conn.commit()
                                st.error("Application Rejected!")
                                st.rerun()

        with tabs[1]:
            pending_registrations = pd.read_sql_query(
                """
                SELECT cr.*, si.surname, si.other_names 
                FROM course_registration cr 
                LEFT JOIN student_info si ON cr.student_id = si.student_id 
                WHERE cr.approval_status='pending'
                """,
                conn,
            )

            if pending_registrations.empty:
                st.info("No pending course registrations")
            else:
                for _, registration in pending_registrations.iterrows():
                    with st.expander(
                        f"Registration ID: {registration['registration_id']} - {registration['surname']} {registration['other_names']}"
                    ):
                        col1, col2 = st.columns(2)

                        with col1:
                            st.write("**Registration Details**")
                            st.write(f"Student ID: {registration['student_id']}")
                            st.write(f"Programme: {registration['programme']}")
                            st.write(f"Level: {registration['level']}")
                            st.write(f"Session: {registration['session']}")

                        with col2:
                            st.write("**Academic Information**")
                            st.write(f"Academic Year: {registration['academic_year']}")
                            st.write(f"Semester: {registration['semester']}")
                            st.write(f"Total Credits: {registration['total_credits']}")

                        st.write("**Selected Courses**")
                        if registration["courses"]:
                            courses_list = registration["courses"].split("\n")
                            table_data = []
                            for course in courses_list:
                                if "|" in course:
                                    code, title, credits = course.split("|")
                                    table_data.append(
                                        [code, title, f"{credits} credits"]
                                    )
                            if table_data:
                                df = pd.DataFrame(
                                    table_data,
                                    columns=[
                                        "Course Code",
                                        "Course Title",
                                        "Credit Hours",
                                    ],
                                )
                                st.table(df)

                        st.write("**Payment Information**")
                        if registration["receipt_path"]:
                            st.write(
                                f"Receipt Amount: GHS {registration['receipt_amount']:.2f}"
                            )
                            if os.path.exists(registration["receipt_path"]):
                                st.write(
                                    f"[View Receipt]({registration['receipt_path']})"
                                )
                        else:
                            st.write("No receipt uploaded (Optional)")

                        col1, col2 = st.columns(2)
                        with col1:
                            if st.button(
                                "Approve",
                                key=f"approve_reg_{registration['registration_id']}",
                            ):
                                c = conn.cursor()
                                c.execute(
                                    "UPDATE course_registration SET approval_status='approved' WHERE registration_id=?",
                                    (registration["registration_id"],),
                                )
                                conn.commit()
                                st.success("Registration Approved!")
                                st.rerun()
                        with col2:
                            if st.button(
                                "Reject",
                                key=f"reject_reg_{registration['registration_id']}",
                            ):
                                c = conn.cursor()
                                c.execute(
                                    "UPDATE course_registration SET approval_status='rejected' WHERE registration_id=?",
                                    (registration["registration_id"],),
                                )
                                conn.commit()
                                st.error("Registration Rejected!")
                                st.rerun()

    finally:
        conn.close()

# Helper functions for disability status handling
def get_disability_status_display(status):
    """
    Convert database disability status to display format
    Returns "None" or "Yes" for display in selectbox
    """
    if status is None or str(status).lower() in ['none', 'no', '']:
        return "None"
    return "Yes"

def get_disability_status_index(status):
    """
    Get the correct index for disability status selectbox
    Returns 0 for None/No/empty and 1 for Yes
    """
    if status is None or str(status).lower() in ['none', 'no', '']:
        return 0  # Index for "None"
    return 1  # Index for "Yes"

# Update the disability status section in manage_student_records()
def update_disability_fields(student):
    """
    Handle disability status fields in the student record form
    Returns dictionary with updated disability fields
    """
    edited_data = {}
    
    # Convert database value to display format
    current_status = get_disability_status_display(student["disability_status"])
    
    # Create selectbox with correct initial value
    edited_data["disability_status"] = st.selectbox(
        "Disability Status",
        ["None", "Yes"],
        index=get_disability_status_index(current_status),
        key=f"edit_disability_{student['student_id']}"
    )
    
    # Show description field if status is "Yes"
    if edited_data["disability_status"] == "Yes":
        edited_data["disability_description"] = st.text_area(
            "Disability Description",
            value=student["disability_description"] or "",
            key=f"edit_disability_desc_{student['student_id']}"
        )
    else:
        edited_data["disability_description"] = "None"
    
    return edited_data

def manage_student_records():
    st.subheader("Student Records Management")

    # Add search phrase input for filtering records
    search_phrase = st.text_input("Search by phrase (ID, surname, or other names)", "")

    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        sort_by = st.selectbox(
            "Sort by", ["Student ID", "Surname", "Date Added", "Programme"]
        )
    with col2:
        sort_order = st.selectbox("Order", ["Ascending", "Descending"])
    with col3:
        status_filter = st.selectbox(
            "Status", ["All", "Pending", "Approved", "Rejected"]
        )

    conn = sqlite3.connect("student_registration.db")
    sort_field = {
        "Student ID": "student_id",
        "Surname": "surname",
        "Date Added": "created_at",
        "Programme": "programme",
    }[sort_by]

    order = "ASC" if sort_order == "Ascending" else "DESC"
    # Base query and parameters list
    query = """
        SELECT 
            student_id,
            surname,
            other_names,
            date_of_birth,
            place_of_birth,
            home_town,
            residential_address,
            postal_address,
            email,
            telephone,
            ghana_card_id,
            nationality,
            marital_status,
            gender,
            religion,
            denomination,
            disability_status,
            disability_description,
            guardian_name,
            guardian_relationship,
            guardian_occupation,
            guardian_address,
            guardian_telephone,
            previous_school,
            qualification_type,
            completion_year,
            aggregate_score,
            ghana_card_path,
            passport_photo_path,
            transcript_path,
            certificate_path,
            receipt_path,
            CAST(receipt_amount AS FLOAT) as receipt_amount,
            approval_status,
            created_at,
            programme
        FROM student_info 
        WHERE 1=1
    """
    params = []

    # Add status filter if not "All"
    if status_filter != "All":
        query += " AND approval_status = ?"
        params.append(status_filter.lower())
    
    # Add search by phrase filter if present
    if search_phrase.strip():
        query += " AND (student_id LIKE ? OR surname LIKE ? OR other_names LIKE ?)"
        like_phrase = f"%{search_phrase.strip()}%"
        params.extend([like_phrase, like_phrase, like_phrase])
    
    # Add ordering
    query += f" ORDER BY {sort_field} {order}"
    
    df = pd.read_sql_query(query, conn, params=params)

    if not df.empty:
        for _, student in df.iterrows():
            with st.expander(
                f"{student['surname']}, {student['other_names']} ({student['student_id']})"
            ):
                tab1, tab2, tab3, tab4 = st.tabs(
                    ["Details", "Edit Form", "Documents", "Actions"]
                )

                with tab1:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**Personal Information**")
                        st.write(f"Student ID: {student['student_id']}")
                        st.write(f"Name: {student['surname']} {student['other_names']}")
                        st.write(f"Date of Birth: {student['date_of_birth']}")
                        st.write(f"Gender: {student['gender']}")
                        st.write(f"Nationality: {student['nationality']}")
                        st.write(f"Religion: {student['religion']}")
                        st.write(f"Denomination: {student['denomination']}")

                    with col2:
                        st.write("**Contact Information**")
                        st.write(f"Email: {student['email']}")
                        st.write(f"Phone: {student['telephone']}")
                        st.write(f"Ghana Card: {student['ghana_card_id']}")
                        st.write(f"Address: {student['residential_address']}")

                        st.write("**Payment Information**")
                        if (
                            pd.notna(student["receipt_path"])
                            and student["receipt_path"]
                        ):
                            st.write("✅ Receipt Uploaded")
                            receipt_amount = (
                                student["receipt_amount"].iloc[0]
                                if isinstance(student["receipt_amount"], pd.Series)
                                else student["receipt_amount"]
                            )
                            st.write(f"Receipt Amount: GHS {float(receipt_amount):.2f}")
                        else:
                            st.write("⚪ No Receipt (Optional)")

                with tab2:
                    edited_data = {}

                    st.subheader("Personal Information")
                    col1, col2 = st.columns(2)

                    with col1:
                        edited_data["student_id"] = st.text_input(
                            "Student ID",
                            student["student_id"],
                            key=f"edit_id_{student['student_id']}",
                        )
                        edited_data["surname"] = st.text_input(
                            "Surname",
                            student["surname"],
                            key=f"edit_surname_{student['student_id']}",
                        )
                        edited_data["other_names"] = st.text_input(
                            "Other Names",
                            student["other_names"],
                            key=f"edit_other_{student['student_id']}",
                        )
                        edited_data["date_of_birth"] = st.date_input(
                            "Date of Birth",
                            datetime.strptime(
                                student["date_of_birth"], "%Y-%m-%d"
                            ).date(),
                            key=f"edit_dob_{student['student_id']}",
                        )
                        edited_data["place_of_birth"] = st.text_input(
                            "Place of Birth",
                            student["place_of_birth"],
                            key=f"edit_pob_{student['student_id']}",
                        )
                        edited_data["home_town"] = st.text_input(
                            "Home Town",
                            student["home_town"],
                            key=f"edit_hometown_{student['student_id']}",
                        )
                        edited_data["nationality"] = st.text_input(
                            "Nationality",
                            student["nationality"],
                            key=f"edit_nationality_{student['student_id']}",
                        )

                    with col2:
                        edited_data["gender"] = st.selectbox(
                            "Gender",
                            ["Male", "Female", "Other"],
                            index=["Male", "Female", "Other"].index(student["gender"]),
                            key=f"edit_gender_{student['student_id']}",
                        )
                        edited_data["marital_status"] = st.selectbox(
                            "Marital Status",
                            ["Single", "Married", "Divorced", "Widowed"],
                            index=["Single", "Married", "Divorced", "Widowed"].index(
                                student["marital_status"]
                            ),
                            key=f"edit_marital_{student['student_id']}",
                        )
                        edited_data["religion"] = st.text_input(
                            "Religion",
                            student["religion"],
                            key=f"edit_religion_{student['student_id']}",
                        )
                        edited_data["denomination"] = st.text_input(
                            "Denomination",
                            student["denomination"],
                            key=f"edit_denom_{student['student_id']}",
                        )



                        # In the manage_student_records function, replace the disability status selectbox code with:
                        edited_data["disability_status"] = st.selectbox(
                            "Disability Status",
                            ["None", "Yes"],
                            index=get_disability_status_index(student["disability_status"]),
                            key=f"edit_disability_{student['student_id']}"
                        )

                        # If disability status is "Yes", show description field
                        if edited_data["disability_status"] == "Yes":
                            edited_data["disability_description"] = st.text_area(
                                "Disability Description",
                                student["disability_description"],
                                key=f"edit_disability_desc_{student['student_id']}"
                        )
                    st.subheader("Contact Information")
                    col3, col4 = st.columns(2)

                    with col3:
                        edited_data["residential_address"] = st.text_area(
                            "Residential Address",
                            student["residential_address"],
                            key=f"edit_res_{student['student_id']}",
                        )
                        edited_data["postal_address"] = st.text_area(
                            "Postal Address",
                            student["postal_address"],
                            key=f"edit_postal_{student['student_id']}",
                        )
                        edited_data["email"] = st.text_input(
                            "Email",
                            student["email"],
                            key=f"edit_email_{student['student_id']}",
                        )

                    with col4:
                        edited_data["telephone"] = st.text_input(
                            "Telephone",
                            student["telephone"],
                            key=f"edit_tel_{student['student_id']}",
                        )
                        edited_data["ghana_card_id"] = st.text_input(
                            "Ghana Card ID",
                            student["ghana_card_id"],
                            key=f"edit_ghana_{student['student_id']}",
                        )

                    st.subheader("Guardian Information")
                    col5, col6 = st.columns(2)

                    with col5:
                        edited_data["guardian_name"] = st.text_input(
                            "Guardian's Name",
                            student["guardian_name"],
                            key=f"edit_guard_name_{student['student_id']}",
                        )
                        edited_data["guardian_relationship"] = st.text_input(
                            "Relationship",
                            student["guardian_relationship"],
                            key=f"edit_guard_rel_{student['student_id']}",
                        )
                        edited_data["guardian_occupation"] = st.text_input(
                            "Occupation",
                            student["guardian_occupation"],
                            key=f"edit_guard_occ_{student['student_id']}",
                        )

                    with col6:
                        edited_data["guardian_address"] = st.text_area(
                            "Address",
                            student["guardian_address"],
                            key=f"edit_guard_addr_{student['student_id']}",
                        )
                        edited_data["guardian_telephone"] = st.text_input(
                            "Telephone",
                            student["guardian_telephone"],
                            key=f"edit_guard_tel_{student['student_id']}",
                        )

                    st.subheader("Educational Background")
                    col7, col8 = st.columns(2)

                    with col7:
                        edited_data["previous_school"] = st.text_input(
                            "Previous School",
                            student["previous_school"],
                            key=f"edit_prev_sch_{student['student_id']}",
                        )
                        edited_data["qualification_type"] = st.text_input(
                            "Qualification",
                            student["qualification_type"],
                            key=f"edit_qual_{student['student_id']}",
                        )

                    with col8:
                        edited_data["completion_year"] = st.text_input(
                            "Completion Year",
                            student["completion_year"],
                            key=f"edit_comp_year_{student['student_id']}",
                        )
                        edited_data["aggregate_score"] = st.text_input(
                            "Aggregate Score",
                            student["aggregate_score"],
                            key=f"edit_agg_{student['student_id']}",
                        )

                    edited_data["approval_status"] = st.selectbox(
                        "Approval Status",
                        ["pending", "approved", "rejected"],
                        index=["pending", "approved", "rejected"].index(
                            student["approval_status"]
                        ),
                        key=f"edit_status_{student['student_id']}",
                    )

                    st.subheader("Payment Information")
                    if pd.notna(student["receipt_path"]) and student["receipt_path"]:
                        current_amount = (
                            float(student["receipt_amount"])
                            if pd.notna(student["receipt_amount"])
                            else 0.0
                        )
                        edited_data["receipt_amount"] = st.number_input(
                            "Receipt Amount (GHS)",
                            value=current_amount,
                            min_value=0.0,
                            format="%.2f",
                            key=f"edit_receipt_amount_{student['student_id']}",
                        )
                        if edited_data["receipt_amount"] < 100.0:
                            st.warning(
                                "Receipt amount seems low. Please verify the payment amount."
                            )

                    if st.button(
                        "Save Changes", key=f"save_changes_{student['student_id']}"
                    ):
                        try:
                            c = conn.cursor()
                            update_query = """
                                UPDATE student_info 
                                SET student_id=?, surname=?, other_names=?, date_of_birth=?, 
                                    place_of_birth=?, home_town=?, nationality=?, gender=?,
                                    marital_status=?, religion=?, denomination=?, 
                                    disability_status=?, disability_description=?,
                                    residential_address=?, postal_address=?, email=?,
                                    telephone=?, ghana_card_id=?, guardian_name=?,
                                    guardian_relationship=?, guardian_occupation=?,
                                    guardian_address=?, guardian_telephone=?,
                                    previous_school=?, qualification_type=?,
                                    completion_year=?, aggregate_score=?,
                                    approval_status=?, receipt_amount=?
                                WHERE student_id=?
                            """

                            receipt_amount = float(
                                edited_data.get("receipt_amount", 0.0)
                            )

                            c.execute(
                                update_query,
                                (
                                    edited_data["student_id"],
                                    edited_data["surname"],
                                    edited_data["other_names"],
                                    edited_data["date_of_birth"],
                                    edited_data["place_of_birth"],
                                    edited_data["home_town"],
                                    edited_data["nationality"],
                                    edited_data["gender"],
                                    edited_data["marital_status"],
                                    edited_data["religion"],
                                    edited_data["denomination"],
                                    edited_data["disability_status"],
                                    edited_data.get("disability_description", "None"),
                                    edited_data["residential_address"],
                                    edited_data["postal_address"],
                                    edited_data["email"],
                                    edited_data["telephone"],
                                    edited_data["ghana_card_id"],
                                    edited_data["guardian_name"],
                                    edited_data["guardian_relationship"],
                                    edited_data["guardian_occupation"],
                                    edited_data["guardian_address"],
                                    edited_data["guardian_telephone"],
                                    edited_data["previous_school"],
                                    edited_data["qualification_type"],
                                    edited_data["completion_year"],
                                    edited_data["aggregate_score"],
                                    edited_data["approval_status"],
                                    receipt_amount,
                                    student["student_id"],
                                ),
                            )
                            conn.commit()
                            st.success("Changes saved successfully!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error saving changes: {str(e)}")

                with tab3:
                    st.write("**Document Management**")
                    documents = {
                        "Ghana Card": student["ghana_card_path"],
                        "Passport Photo": student["passport_photo_path"],
                        "Transcript": student["transcript_path"],
                        "Certificate": student["certificate_path"],
                        "Receipt": student["receipt_path"],
                    }

                    for doc_name, doc_path in documents.items():
                        col1, col2, col3 = st.columns([3, 1, 1])

                        with col1:
                            if pd.notna(doc_path) and doc_path:  # Check for both not null and not empty string
                                st.write(f"✅ {doc_name} uploaded")
                                if os.path.exists(doc_path):  # Verify file exists on disk
                                    if doc_name == "Passport Photo" or doc_path.lower().endswith(('.jpg', '.jpeg', '.png')):
                                        try:
                                            image = PILImage.open(doc_path)
                                            st.image(image, width=150, caption=doc_name)
                                        except Exception as e:
                                            st.error(f"Error loading {doc_name}: {str(e)}")
                                    elif doc_path.lower().endswith('.pdf'):
                                        st.write(f"[View {doc_name}]({doc_path})")
                                else:
                                    st.error(f"{doc_name} file not found on disk: {doc_path}")
                            else:
                                st.write(f"❌ {doc_name} not uploaded")

                        with col2:
                            new_file = st.file_uploader(
                                f"Upload new {doc_name}",
                                type=(["pdf", "jpg", "jpeg", "png"] if doc_name == "Passport Photo" 
                                    else ["pdf"] if doc_name in ["Transcript", "Certificate"]
                                    else ["pdf", "jpg", "jpeg", "png"]),
                                key=f"upload_{doc_name}_{student['student_id']}"
                            )

                            if new_file:
                                if st.button(f"Save {doc_name}", key=f"save_{doc_name}_{student['student_id']}"):
                                    try:
                                        # Remove old file if it exists
                                        if doc_path and os.path.exists(doc_path):
                                            os.remove(doc_path)

                                        # Save new file
                                        new_path = save_uploaded_file(new_file, "uploads")
                                        if new_path:
                                            # Update database with new file path
                                            c = conn.cursor()
                                            c.execute(
                                                f"""
                                                UPDATE student_info 
                                                SET {doc_name.lower().replace(' ', '_')}_path = ? 
                                                WHERE student_id = ?
                                                """,
                                                (new_path, student["student_id"])
                                            )
                                            conn.commit()
                                            st.success(f"{doc_name} uploaded successfully!")
                                            st.rerun()
                                        else:
                                            st.error(f"Failed to save {doc_name}")
                                    except Exception as e:
                                        st.error(f"Error uploading {doc_name}: {str(e)}")

                        with col3:
                            if pd.notna(doc_path) and doc_path and os.path.exists(doc_path):
                                if st.button(f"Delete {doc_name}", key=f"del_{doc_name}_{student['student_id']}"):
                                    try:
                                        # Remove file from disk
                                        os.remove(doc_path)
                                        
                                        # Update database
                                        c = conn.cursor()
                                        c.execute(
                                            f"""
                                            UPDATE student_info 
                                            SET {doc_name.lower().replace(' ', '_')}_path = NULL 
                                            WHERE student_id = ?
                                            """,
                                            (student["student_id"],)
                                        )
                                        conn.commit()
                                        st.success(f"{doc_name} deleted successfully!")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Error deleting {doc_name}: {str(e)}")

                with tab4:
                    st.write("**Student Actions**")
                    col1, col2 = st.columns(2)

                    with col1:
                        if st.button(
                            "Generate PDF", key=f"pdf_{student['student_id']}"
                        ):
                            pdf_file = generate_student_info_pdf(student)
                            with open(pdf_file, "rb") as file:
                                st.download_button(
                                    label="Download Student Info",
                                    data=file,
                                    file_name=pdf_file,
                                    mime="application/pdf",
                                )

                    with col2:
                        if st.button(
                            "Delete Student Record",
                            key=f"del_student_{student['student_id']}",
                            type="primary",
                        ):
                            try:
                                for doc_path in documents.values():
                                    if doc_path and os.path.exists(doc_path):
                                        os.remove(doc_path)

                                c = conn.cursor()
                                c.execute(
                                    "DELETE FROM student_info WHERE student_id = ?",
                                    (student["student_id"],),
                                )
                                conn.commit()
                                st.success("Student record deleted successfully!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error deleting student record: {str(e)}")
    else:
        st.info("No records found")

    conn.close()


def manage_course_registrations():
    st.subheader("Course Registration Management")
    
    # Add search phrase input for filtering registration records
    search_phrase = st.text_input("Search registrations (by student ID, surname, or registration ID)", "")

    col1, col2, col3 = st.columns([2, 2, 1])

    with col1:
        sort_by = st.selectbox(
            "Sort by", ["Registration ID", "Student ID", "Programme", "Date Registered"]
        )

    with col2:
        sort_order = st.selectbox("Order", ["Ascending", "Descending"], key="reg_order")

    with col3:
        status_filter = st.selectbox(
            "Status", ["All", "Pending", "Approved", "Rejected"], key="reg_status"
        )

    conn = sqlite3.connect("student_registration.db")

    sort_field = {
        "Registration ID": "cr.registration_id",
        "Student ID": "cr.student_id",
        "Programme": "cr.programme",
        "Date Registered": "cr.date_registered",
    }[sort_by]

    order = "ASC" if sort_order == "Ascending" else "DESC"

    # Base query and parameters list for registrations
    query = """
        SELECT cr.*, si.surname, si.other_names 
        FROM course_registration cr 
        LEFT JOIN student_info si ON cr.student_id = si.student_id 
        WHERE 1=1
    """
    params = []
    
    # Add status filter if not All
    if status_filter != "All":
        query += " AND cr.approval_status = ?"
        params.append(status_filter.lower())
    
    # Add search by phrase filter if provided, search in registration_id, student_id, surname, or other_names
    if search_phrase.strip():
        query += " AND (cr.registration_id LIKE ? OR cr.student_id LIKE ? OR si.surname LIKE ? OR si.other_names LIKE ?)"
        like_phrase = f"%{search_phrase.strip()}%"
        params.extend([like_phrase, like_phrase, like_phrase, like_phrase])
    
    query += f" ORDER BY {sort_field} {order}"
    
    df = pd.read_sql_query(query, conn, params=params)

    if not df.empty:
        for _, registration in df.iterrows():
            with st.expander(
                f"Registration ID: {registration['registration_id']} - {registration['surname']} {registration['other_names']}"
            ):
                tab1, tab2, tab3, tab4 = st.tabs(
                    ["Details", "Edit Registration", "Documents", "Actions"]
                )

                with tab1:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**Registration Details**")
                        st.write(f"Student ID: {registration['student_id']}")
                        st.write(f"Programme: {registration['programme']}")
                        st.write(f"Level: {registration['level']}")
                        st.write(f"Session: {registration['session']}")

                    with col2:
                        st.write("**Academic Information**")
                        st.write(f"Academic Year: {registration['academic_year']}")
                        st.write(f"Semester: {registration['semester']}")
                        st.write(f"Total Credits: {registration['total_credits']}")

                    st.write("**Selected Courses**")
                    if registration["courses"]:
                        courses_list = registration["courses"].split("\n")
                        table_data = []
                        for course in courses_list:
                            if "|" in course:
                                code, title, credits = course.split("|")
                                table_data.append([code, title, f"{credits} credits"])
                        if table_data:
                            df = pd.DataFrame(
                                table_data,
                                columns=["Course Code", "Course Title", "Credit Hours"],
                            )
                            st.table(df)

                    st.write("**Payment Information**")
                    if registration["receipt_path"]:
                        st.write("✅ Receipt Uploaded")
                        st.write(
                            f"Receipt Amount: GHS {registration['receipt_amount']:.2f}"
                        )
                    else:
                        st.write("⚪ No Receipt (Optional)")

                with tab2:
                    edited_reg = {}
                    col1, col2 = st.columns(2)

                    with col1:
                        edited_reg["programme"] = st.selectbox(
                            "Programme",
                            ["CIMG", "CIM-UK", "ICAG", "ACCA"],
                            index=["CIMG", "CIM-UK", "ICAG", "ACCA"].index(
                                registration["programme"]
                            ),
                            key=f"prog_{registration['registration_id']}",
                        )
                        program_levels = list(
                            get_program_courses(edited_reg["programme"]).keys()
                        )
                        edited_reg["level"] = st.selectbox(
                            "Level",
                            program_levels,
                            index=(
                                program_levels.index(registration["level"])
                                if registration["level"] in program_levels
                                else 0
                            ),
                            key=f"level_{registration['registration_id']}",
                        )
                        edited_reg["session"] = st.selectbox(
                            "Session",
                            ["Morning", "Evening", "Weekend"],
                            index=["Morning", "Evening", "Weekend"].index(
                                registration["session"]
                            ),
                            key=f"session_{registration['registration_id']}",
                        )

                    with col2:
                        edited_reg["academic_year"] = st.selectbox(
                            "Academic Year",
                            [f"{year}-{year+1}" for year in range(2025, 2035)],
                            index=[
                                f"{year}-{year+1}" for year in range(2025, 2035)
                            ].index(registration["academic_year"]),
                            key=f"year_{registration['registration_id']}",
                        )
                        edited_reg["semester"] = st.selectbox(
                            "Semester",
                            ["First", "Second", "Third"],
                            index=["First", "Second", "Third"].index(
                                registration["semester"]
                            ),
                            key=f"sem_{registration['registration_id']}",
                        )
                        edited_reg["approval_status"] = st.selectbox(
                            "Status",
                            ["pending", "approved", "rejected"],
                            index=["pending", "approved", "rejected"].index(
                                registration["approval_status"]
                            ),
                            key=f"status_{registration['registration_id']}",
                        )

                    st.write("**Course Selection**")
                    available_courses = get_program_courses(
                        edited_reg["programme"]
                    ).get(edited_reg["level"], [])
                    current_courses = (
                        registration["courses"].split("\n")
                        if registration["courses"]
                        else []
                    )
                    selected_courses = st.multiselect(
                        "Select Courses",
                        available_courses,
                        default=current_courses,
                        format_func=lambda x: f"{x.split('|')[0]} - {x.split('|')[1]} ({x.split('|')[2]} credits)",
                        key=f"courses_{registration['registration_id']}",
                    )

                    edited_reg["courses"] = "\n".join(selected_courses)
                    edited_reg["total_credits"] = sum(
                        [int(course.split("|")[2]) for course in selected_courses]
                    )

                    st.write(f"Total Credits: {edited_reg['total_credits']}")
                    if edited_reg["total_credits"] > 24:
                        st.error("Total credits cannot exceed 24 hours!")

                    if st.button(
                        "Save Changes", key=f"save_{registration['registration_id']}"
                    ):
                        if edited_reg["total_credits"] <= 24:
                            try:
                                c = conn.cursor()
                                update_query = """
                                    UPDATE course_registration 
                                    SET programme=?, level=?, session=?, 
                                        academic_year=?, semester=?, approval_status=?,
                                        courses=?, total_credits=?
                                    WHERE registration_id=?
                                """
                                c.execute(
                                    update_query,
                                    (
                                        edited_reg["programme"],
                                        edited_reg["level"],
                                        edited_reg["session"],
                                        edited_reg["academic_year"],
                                        edited_reg["semester"],
                                        edited_reg["approval_status"],
                                        edited_reg["courses"],
                                        edited_reg["total_credits"],
                                        registration["registration_id"],
                                    ),
                                )
                                conn.commit()
                                st.success("Changes saved successfully!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error saving changes: {str(e)}")
                        else:
                            st.error(
                                "Cannot save changes. Total credits exceed 24 hours limit."
                            )

                with tab3:
                    st.write("**Document Management**")
                    if registration["receipt_path"]:
                        st.write("✅ Receipt uploaded")
                        if registration["receipt_path"].lower().endswith((".pdf")):
                            st.write(f"[View Receipt]({registration['receipt_path']})")
                        elif (
                            registration["receipt_path"]
                            .lower()
                            .endswith((".jpg", ".jpeg", ".png"))
                        ):
                            try:
                                image = PILImage.open(registration["receipt_path"])
                                st.image(
                                    image, caption="Receipt", use_container_width=True
                                )
                            except Exception as e:
                                st.error(f"Error loading receipt image: {str(e)}")

                        new_amount = st.number_input(
                            "Update Receipt Amount (GHS)",
                            value=float(registration["receipt_amount"]),
                            min_value=0.0,
                            format="%.2f",
                            key=f"receipt_amount_{registration['registration_id']}",
                        )

                        if new_amount != registration["receipt_amount"]:
                            if st.button(
                                "Update Amount",
                                key=f"update_amount_{registration['registration_id']}",
                            ):
                                try:
                                    c = conn.cursor()
                                    c.execute(
                                        """
                                        UPDATE course_registration 
                                        SET receipt_amount = ? 
                                        WHERE registration_id = ?
                                    """,
                                        (new_amount, registration["registration_id"]),
                                    )
                                    conn.commit()
                                    st.success("Receipt amount updated successfully!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error updating receipt amount: {str(e)}")

                        if st.button(
                            "Delete Receipt",
                            key=f"del_receipt_{registration['registration_id']}",
                        ):
                            try:
                                if os.path.exists(registration["receipt_path"]):
                                    os.remove(registration["receipt_path"])

                                c = conn.cursor()
                                c.execute(
                                    """
                                    UPDATE course_registration 
                                    SET receipt_path = NULL, receipt_amount = 0 
                                    WHERE registration_id = ?
                                """,
                                    (registration["registration_id"],),
                                )
                                conn.commit()
                                st.success("Receipt deleted successfully!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error deleting receipt: {str(e)}")
                    else:
                        st.write("❌ No receipt uploaded")
                        new_receipt = st.file_uploader(
                            "Upload Receipt",
                            type=["pdf", "jpg", "jpeg", "png"],
                            key=f"new_receipt_{registration['registration_id']}",
                        )
                        if new_receipt:
                            receipt_amount = st.number_input(
                                "Receipt Amount (GHS)",
                                min_value=0.0,
                                format="%.2f",
                                key=f"new_amount_{registration['registration_id']}",
                            )
                            if st.button(
                                "Save Receipt",
                                key=f"save_receipt_{registration['registration_id']}",
                            ):
                                try:
                                    receipt_path = save_uploaded_file(
                                        new_receipt, "uploads"
                                    )
                                    c = conn.cursor()
                                    c.execute(
                                        """
                                        UPDATE course_registration 
                                        SET receipt_path = ?, receipt_amount = ? 
                                        WHERE registration_id = ?
                                    """,
                                        (
                                            receipt_path,
                                            receipt_amount,
                                            registration["registration_id"],
                                        ),
                                    )
                                    conn.commit()
                                    st.success("Receipt uploaded successfully!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error uploading receipt: {str(e)}")

                with tab4:
                    st.write("**Registration Actions**")
                    col1, col2 = st.columns(2)

                    with col1:
                        if st.button(
                            "Generate PDF", key=f"pdf_{registration['registration_id']}"
                        ):
                            pdf_file = generate_course_registration_pdf(registration)
                            with open(pdf_file, "rb") as file:
                                st.download_button(
                                    label="Download Registration Form",
                                    data=file,
                                    file_name=pdf_file,
                                    mime="application/pdf",
                                )

                    with col2:
                        if st.button(
                            "Delete Registration",
                            key=f"del_reg_{registration['registration_id']}",
                            type="primary",
                        ):
                            try:
                                if registration["receipt_path"] and os.path.exists(
                                    registration["receipt_path"]
                                ):
                                    os.remove(registration["receipt_path"])

                                c = conn.cursor()
                                c.execute(
                                    "DELETE FROM course_registration WHERE registration_id = ?",
                                    (registration["registration_id"],),
                                )
                                conn.commit()
                                st.success("Registration deleted successfully!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error deleting registration: {str(e)}")
    else:
        st.info("No records found")

    conn.close()


def manage_programs():
    st.title("Programs Management")

    conn = sqlite3.connect("student_registration.db")
    programs_df = pd.read_sql_query(
        """
        SELECT DISTINCT programme 
        FROM course_registration 
        WHERE programme IS NOT NULL
    """,
        conn,
    )

    if not programs_df.empty:
        for program in programs_df["programme"]:
            with st.expander(f"📚 {program}"):
                levels_df = pd.read_sql_query(
                    """
                    SELECT DISTINCT level 
                    FROM course_registration 
                    WHERE programme = ? 
                    ORDER BY level
                """,
                    conn,
                    params=(program,),
                )

                for _, level_row in levels_df.iterrows():
                    level = level_row["level"]
                    st.subheader(f"{level}")

                    students_df = pd.read_sql_query(
                        """
                        SELECT DISTINCT 
                            cr.student_id,
                            si.surname,
                            si.other_names,
                            si.passport_photo_path,
                            cr.academic_year,
                            cr.semester
                        FROM course_registration cr
                        JOIN student_info si ON cr.student_id = si.student_id
                        WHERE cr.programme = ? AND cr.level = ?
                        ORDER BY si.surname, si.other_names
                    """,
                        conn,
                        params=(program, level),
                    )

                    if not students_df.empty:
                        st.write(f"Total Students: {len(students_df)}")

                        cols = st.columns(4)
                        for idx, student in students_df.iterrows():
                            with cols[idx % 4]:
                                st.write("---")
                                if student["passport_photo_path"] and os.path.exists(
                                    student["passport_photo_path"]
                                ):
                                    try:
                                        image = PILImage.open(
                                            student["passport_photo_path"]
                                        )
                                        st.image(image, width=100)
                                    except Exception as e:
                                        st.error("Error loading photo")
                                st.write(
                                    f"**{student['surname']}, {student['other_names']}**"
                                )
                                st.write(f"ID: {student['student_id']}")
                                st.write(f"Year: {student['academic_year']}")
                                st.write(f"Semester: {student['semester']}")

                        if st.button(
                            f"Download {program} - {level} Student List",
                            key=f"btn_{program}_{level}",
                        ):
                            pdf_file = generate_program_student_list(
                                program, level, students_df
                            )
                            with open(pdf_file, "rb") as file:
                                st.download_button(
                                    label=f"Download {program} - {level} PDF",
                                    data=file,
                                    file_name=pdf_file,
                                    mime="application/pdf",
                                )
                    else:
                        st.info(f"No students registered for {level}")
    else:
        st.info("No programs found in the database")

    conn.close()


# Student Portal Authentication and Views
def ensure_student_password_column():
    """
    Ensure that the student_info table has a 'password' column.
    If not, add it. This column will store the student's custom password.
    """
    conn = sqlite3.connect("student_registration.db")
    c = conn.cursor()
    try:
        # Try to query the 'password' column
        c.execute("SELECT password FROM student_info LIMIT 1")
    except sqlite3.OperationalError:
        # Add password column if it doesn't exist
        c.execute("ALTER TABLE student_info ADD COLUMN password TEXT")
    conn.commit()
    conn.close()


def student_login_form():
    """
    Display the login form for students.
    Uses student ID and birthday (default password as YYYY-MM-DD).
    On first successful login (if password is still default), prompt for a password reset.
    """
    st.header("👩‍🎓 Student Portal Login")

    # Store password reset form state
    if "show_password_reset" not in st.session_state:
        st.session_state.show_password_reset = False

    col1, col2 = st.columns(2)
    with col1:
        student_id = st.text_input("Student ID")
    with col2:
        password = st.text_input(
            "Password",
            type="password",
            help="Default password is your date of birth (YYYY-MM-DD)",
        )

    if (
        st.button("Login", use_container_width=True)
        or st.session_state.show_password_reset
    ):
        conn = sqlite3.connect("student_registration.db")
        c = conn.cursor()
        c.execute(
            """
            SELECT student_id, date_of_birth, password, approval_status 
            FROM student_info 
            WHERE student_id = ?
        """,
            (student_id,),
        )
        student = c.fetchone()
        conn.close()

        if not student:
            st.error("Student record not found. Please check your Student ID.")
            return None

        db_student_id, db_dob, db_password, approval_status = student

        if approval_status != "approved":
            st.error("Your account is pending approval. Please contact administration.")
            return None

        # Default password is the student's date_of_birth formatted as YYYY-MM-DD
        default_password = (
            db_dob if isinstance(db_dob, str) else db_dob.strftime("%Y-%m-%d")
        )

        # If a custom password is set, use it; otherwise use default
        expected_password = (
            db_password
            if db_password and db_password.strip() != ""
            else default_password
        )

        if password != expected_password and not st.session_state.show_password_reset:
            st.error("Incorrect password. Please check your credentials.")
            return None

        # Successful login: if the password is still default, force reset
        if (
            expected_password == default_password
            or st.session_state.show_password_reset
        ):
            st.session_state.show_password_reset = True
            st.warning(
                "You are using your default password. Please reset your password now."
            )
            new_password = st.text_input("Enter New Password", type="password")
            confirm_password = st.text_input("Confirm New Password", type="password")

            if st.button("Reset Password"):
                if new_password != confirm_password:
                    st.error("Passwords do not match. Try again.")
                    return None
                elif len(new_password) < 8:
                    st.error("Password must be at least 8 characters long.")
                    return None
                else:
                    conn = sqlite3.connect("student_registration.db")
                    c = conn.cursor()
                    c.execute(
                        """
                        UPDATE student_info 
                        SET password = ? 
                        WHERE student_id = ?
                    """,
                        (new_password, student_id),
                    )
                    conn.commit()
                    conn.close()
                    st.success("Password reset successfully! You are now logged in.")
                    st.session_state.show_password_reset = False
                    st.session_state.student_logged_in = student_id
                    st.rerun()
        else:
            st.success("Login successful!")
            st.session_state.student_logged_in = student_id
            st.rerun()


def student_portal():
    st.header("🎓 Welcome to Your Student Portal")
    student_id = st.session_state.get("student_logged_in")
    if not student_id:
        st.error("You must be logged in to view this page.")
        return

    # Fetch student information and course registrations from the database.
    conn = sqlite3.connect("student_registration.db")
    c = conn.cursor()
    c.execute("SELECT * FROM student_info WHERE student_id = ?", (student_id,))
    student = c.fetchone()

    c.execute("SELECT * FROM course_registration WHERE student_id = ? ORDER BY date_registered DESC", (student_id,))
    registrations = c.fetchall()
    conn.close()

    # Create portal tabs with an additional "Proof of Registration" tab.
    tabs = st.tabs(["Profile", "Course Registrations", "Documents", "Proof of Registration", "Settings", "Notifications"])

    # Profile Tab (existing code: abbreviated for brevity)
    with tabs[0]:
        st.subheader("Personal Information")
        col1, col2 = st.columns(2)
        with col1:
            if student[28] and os.path.exists(student[28]):  # passport_photo_path
                try:
                    image = PILImage.open(student[28])
                    st.image(image, width=200, caption="Student Photo")
                except Exception as e:
                    st.error(f"Error loading passport photo: {str(e)}")
            st.write("**Basic Information**")
            st.write(f"Student ID: {student[0]}")
            st.write(f"Name: {student[1]} {student[2]}")
            st.write(f"Date of Birth: {student[3]}")
            st.write(f"Gender: {student[13]}")
            st.write(f"Nationality: {student[11]}")
        with col2:
            st.write("**Contact Information**")
            st.write(f"Email: {student[8]}")
            st.write(f"Phone: {student[9]}")
            st.write(f"Residential Address: {student[6]}")
            st.write(f"Postal Address: {student[7]}")
            st.write("**Academic Information**")
            st.write(f"Programme: {student[-1]}")  # programme is the last column
            st.write(f"Previous School: {student[23]}")
            st.write(f"Qualification: {student[24]}")

    # Course Registrations Tab (existing code snippet)
    with tabs[1]:
        st.subheader("Course Registrations")
        if registrations:
            for reg in registrations:
                with st.expander(f"Registration ID: {reg[0]} - {reg[11]}"):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write(f"**Programme:** {reg[3]}")
                        st.write(f"**Level:** {reg[5]}")
                        st.write(f"**Session:** {reg[6]}")
                    with col2:
                        st.write(f"**Academic Year:** {reg[7]}")
                        st.write(f"**Semester:** {reg[8]}")
                        st.write(f"**Total Credits:** {reg[10]}")
                    if reg[9]:
                        st.write("**Selected Courses:**")
                        courses_list = reg[9].split("\n")
                        for course in courses_list:
                            if "|" in course:
                                code, title, credits = course.split("|")
                                st.write(f"- {code}: {title} ({credits} credits)")
        else:
            st.info("No course registrations found.")

    # Documents Tab (existing code snippet)
    with tabs[2]:
        st.subheader("Documents")
        documents = {
            "Ghana Card": student[27],
            "Passport Photo": student[28],
            "Transcript": student[29],
            "Certificate": student[30],
            "Receipt": student[31],
        }
        for doc_name, doc_path in documents.items():
            if doc_path and os.path.exists(doc_path):
                st.write(f"**{doc_name}**")
                if doc_path.lower().endswith((".jpg", ".jpeg", ".png")):
                    st.image(doc_path, width=200, caption=doc_name)
                else:
                    st.write(f"[View {doc_name}]({doc_path})")
            else:
                st.write(f"**{doc_name}:** Not uploaded")

    # NEW: Proof of Registration Tab
    with tabs[3]:
        st.subheader("Proof of Registration")
        if not registrations:
            st.info("No course registration records found.")
        else:
            # List all course registrations with an option to download the registration PDF.
            for reg in registrations:
                # Map tuple to dictionary required by generate_course_registration_pdf.
                reg_data = {
                    "student_id": reg[1],
                    "index_number": reg[2],
                    "programme": reg[3],
                    "specialization": reg[4],
                    "level": reg[5],
                    "session": reg[6],
                    "academic_year": reg[7],
                    "semester": reg[8],
                    "courses": reg[9],
                    "total_credits": reg[10],
                    "receipt_path": reg[13],
                    "receipt_amount": reg[14]
                }
                st.markdown(f"**Registration Record (ID: {reg[0]}) - Date Registered: {reg[11]}**")
                # When the student clicks the button, generate the PDF and provide a download.
                if st.button(f"Download Proof of Registration (ID: {reg[0]})", key=f"download_{reg[0]}"):
                    pdf_file = generate_course_registration_pdf(reg_data)
                    if os.path.exists(pdf_file):
                        with open(pdf_file, "rb") as f:
                            pdf_bytes = f.read()
                        st.download_button(
                            label=f"Download Registration PDF (ID: {reg[0]})",
                            data=pdf_bytes,
                            file_name=pdf_file,
                            mime="application/pdf"
                        )
                        # Clean up the generated PDF file.
                        os.remove(pdf_file)
                    else:
                        st.error("Error generating PDF. Please try again.")

    # Settings Tab (existing code snippet)
    with tabs[4]:
        st.subheader("Account Settings")
        if st.button("Change Password"):
            current_password = st.text_input("Current Password", type="password")
            new_password = st.text_input("New Password", type="password")
            confirm_password = st.text_input("Confirm New Password", type="password")
            if st.button("Update Password"):
                if new_password != confirm_password:
                    st.error("New passwords do not match.")
                elif len(new_password) < 8:
                    st.error("New password must be at least 8 characters long.")
                else:
                    conn = sqlite3.connect("student_registration.db")
                    c = conn.cursor()
                    c.execute("SELECT password FROM student_info WHERE student_id = ?", (student_id,))
                    stored_password = c.fetchone()[0]
                    if current_password == stored_password:
                        c.execute("UPDATE student_info SET password = ? WHERE student_id = ?", (new_password, student_id))
                        conn.commit()
                        st.success("Password updated successfully!")
                    else:
                        st.error("Current password is incorrect.")
                    conn.close()
        if st.button("Logout"):
            st.session_state.student_logged_in = None
            st.rerun()

    with tabs[5]:  # Notifications tab
        st.subheader("📢 Notifications")

        notification_system = NotificationSystem()
        show_read = st.checkbox("Show read notifications")

        notifications = notification_system.get_notifications(
            student_id=student_id, include_read=show_read
        )

        col1, col2 = st.columns([4, 1])
        with col1:
            st.write(
                f"You have {len([n for n in notifications if not n['is_read']])} unread notifications"
            )
        with col2:
            if st.button("Mark All as Read"):
                notification_system.mark_all_as_read(student_id)
                st.rerun()

        display_notifications(notifications)


def generate_program_student_list(program, level, students_df):
    filename = f"{program}_{level}_students_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
    )

    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="CustomTitle",
            parent=styles["Heading1"],
            fontSize=16,
            alignment=TA_CENTER,
            spaceAfter=30,
            textColor=colors.HexColor("#003366"),
        )
    )

    elements = []
    header_data = [
        [
            Image("upsa_logo.jpg", width=1.2 * inch, height=1.2 * inch),
            Paragraph("UNIVERSITY OF PROFESSIONAL STUDIES, ACCRA", styles["CustomTitle"]),
            Image("upsa_logo.jpg", width=1.2 * inch, height=1.2 * inch),
        ]
    ]
    header_table = RLTable(header_data, [2 * inch, 4 * inch, 2 * inch])
    header_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    elements.append(header_table)
    elements.append(Spacer(1, 20))

    elements.append(Paragraph(f"Program: {program}", styles["Heading2"]))
    elements.append(Paragraph(f"Level: {level}", styles["Heading2"]))
    elements.append(Paragraph(f"Total Students: {len(students_df)}", styles["Normal"]))
    elements.append(Spacer(1, 20))

    for _, student in students_df.iterrows():
        student_data = []
        if student["passport_photo_path"] and os.path.exists(student["passport_photo_path"]):
            try:
                photo = Image(student["passport_photo_path"], width=1 * inch, height=1 * inch)
                student_data.append(
                    [
                        photo,
                        Paragraph(
                            f"<b>Name:</b> {student['surname']}, {student['other_names']}<br/>"
                            f"<b>Student ID:</b> {student['student_id']}<br/>"
                            f"<b>Academic Year:</b> {student['academic_year']}<br/>"
                            f"<b>Semester:</b> {student['semester']}",
                            styles["Normal"],
                        ),
                    ]
                )
            except:
                student_data.append(
                    [
                        Paragraph("No Photo", styles["Normal"]),
                        Paragraph(
                            f"<b>Name:</b> {student['surname']}, {student['other_names']}<br/>"
                            f"<b>Student ID:</b> {student['student_id']}<br/>"
                            f"<b>Academic Year:</b> {student['academic_year']}<br/>"
                            f"<b>Semester:</b> {student['semester']}",
                            styles["Normal"],
                        ),
                    ]
                )
        else:
            student_data.append(
                [
                    Paragraph("No Photo", styles["Normal"]),
                    Paragraph(
                        f"<b>Name:</b> {student['surname']}, {student['other_names']}<br/>"
                        f"<b>Student ID:</b> {student['student_id']}<br/>"
                        f"<b>Academic Year:</b> {student['academic_year']}<br/>"
                        f"<b>Semester:</b> {student['semester']}",
                        styles["Normal"],
                    ),
                ]
            )
        student_table = RLTable(student_data, [1.5 * inch, 5 * inch])
        student_table.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (0, 0), (0, -1), "CENTER"),
                    ("PADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        elements.append(student_table)
        elements.append(Spacer(1, 10))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles["Normal"]))
    doc.build(elements)
    return filename


def download_receipts():
    temp_dir = "temp_receipts"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    try:
        conn = sqlite3.connect("student_registration.db")
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT student_id, surname, other_names, 
                   receipt_path, receipt_amount
            FROM student_info
            WHERE receipt_path IS NOT NULL
        """
        )
        student_receipts = cursor.fetchall()

        cursor.execute(
            """
            SELECT cr.registration_id, cr.student_id, si.surname, si.other_names,
                   cr.receipt_path, cr.receipt_amount
            FROM course_registration cr
            LEFT JOIN student_info si ON cr.student_id = si.student_id
            WHERE cr.receipt_path IS NOT NULL
        """
        )
        registration_receipts = cursor.fetchall()

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"all_receipts_{timestamp}.zip"

        with zipfile.ZipFile(zip_filename, "w") as zipf:
            for receipt in student_receipts:
                student_id, surname, other_names, receipt_path, amount = receipt
                if receipt_path and os.path.exists(receipt_path):
                    _, ext = os.path.splitext(receipt_path)
                    archive_path = f"student_receipts/{student_id}_{surname}_{other_names}_amount_{amount}{ext}"
                    zipf.write(receipt_path, archive_path)

            for receipt in registration_receipts:
                reg_id, student_id, surname, other_names, receipt_path, amount = receipt
                if receipt_path and os.path.exists(receipt_path):
                    _, ext = os.path.splitext(receipt_path)
                    archive_path = f"registration_receipts/reg_{reg_id}_{student_id}_{surname}_{other_names}_amount_{amount}{ext}"
                    zipf.write(receipt_path, archive_path)

        return zip_filename

    except Exception as e:
        print(f"Error creating receipts zip file: {str(e)}")
        return None

    finally:
        conn.close()
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


def save_uploaded_file(uploaded_file, directory):
    """
    Saves an uploaded file to the specified directory using its original filename.
    If a file with the same name already exists, appends a timestamp (before the extension)
    to avoid overwriting the existing file.
    
    This ensures that the file path stored in the database matches exactly the file on disk,
    thus avoiding errors like "No such file or directory" when accessing uploaded documents.
    
    Args:
        uploaded_file: The uploaded file object from Streamlit.
        directory: The directory where the file should be saved.
        
    Returns:
        The file path of the saved file, or None if no file was uploaded.
    """
    if uploaded_file is not None:
        # Ensure the target directory exists.
        if not os.path.exists(directory):
            os.makedirs(directory)
        
        # Use the original filename.
        file_path = os.path.join(directory, uploaded_file.name)
        
        # If there's a filename conflict, append a timestamp before the file extension.
        if os.path.exists(file_path):
            base, ext = os.path.splitext(uploaded_file.name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = os.path.join(directory, f"{base}_{timestamp}{ext}")
        
        # Save the file.
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    return None


def insert_student_info(c, form_data, file_paths):
    insert_query = """
        INSERT INTO student_info (
            student_id, surname, other_names, date_of_birth, place_of_birth,
            home_town, residential_address, postal_address, email, telephone,
            ghana_card_id, nationality, marital_status, gender, religion,
            denomination, disability_status, disability_description,
            guardian_name, guardian_relationship, guardian_occupation,
            guardian_address, guardian_telephone, previous_school,
            qualification_type, completion_year, aggregate_score,
            ghana_card_path, passport_photo_path, transcript_path,
            certificate_path, receipt_path, receipt_amount,
            approval_status, created_at, programme
        ) VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
        )
    """

    params = (
        form_data["student_id"],
        form_data["surname"],
        form_data["other_names"],
        form_data["date_of_birth"],
        form_data["place_of_birth"],
        form_data["home_town"],
        form_data["residential_address"],
        form_data["postal_address"],
        form_data["email"],
        form_data["telephone"],
        form_data["ghana_card_id"],
        form_data["nationality"],
        form_data["marital_status"],
        form_data["gender"],
        form_data["religion"],
        form_data["denomination"],
        form_data["disability_status"],
        form_data["disability_description"],
        form_data["guardian_name"],
        form_data["guardian_relationship"],
        form_data["guardian_occupation"],
        form_data["guardian_address"],
        form_data["guardian_telephone"],
        form_data["previous_school"],
        form_data["qualification_type"],
        form_data["completion_year"],
        form_data["aggregate_score"],
        file_paths.get("ghana_card_path"),
        file_paths.get("passport_photo_path"),
        file_paths.get("transcript_path"),
        file_paths.get("certificate_path"),
        file_paths.get("receipt_path"),
        form_data.get("receipt_amount", 0.0),
        "pending",
        datetime.now(),
        form_data.get("programme", ""),
    )

    c.execute(insert_query, params)


##############################
# New Email Sending Function #
##############################


def send_emails():
    st.header("Send Emails to Students")
    recipient_type = st.selectbox(
        "Select Recipient Group", ["All Students", "By Programme", "Individual Student"]
    )
    subject = st.text_input("Email Subject")
    message_body = st.text_area("Email Message")

    # For individual student, allow entering student ID.
    individual_id = None
    selected_programmes = []
    if recipient_type == "By Programme":
        selected_programmes = st.multiselect(
            "Select Programme(s)", options=["CIMG", "CIM-UK", "ICAG", "ACCA"]
        )
    elif recipient_type == "Individual Student":
        individual_id = st.text_input("Enter Student ID")

    if st.button("Send Email"):
        # Build recipients list based on selection.
        conn = sqlite3.connect("student_registration.db")
        cur = conn.cursor()
        recipients = []
        if recipient_type == "All Students":
            cur.execute("SELECT email FROM student_info")
            results = cur.fetchall()
            recipients = [email for (email,) in results if email]
        elif recipient_type == "By Programme":
            if selected_programmes:
                placeholders = ",".join("?" for _ in selected_programmes)
                query = f"SELECT email FROM student_info WHERE programme IN ({placeholders})"
                cur.execute(query, selected_programmes)
                results = cur.fetchall()
                recipients = [email for (email,) in results if email]
        elif recipient_type == "Individual Student":
            if individual_id:
                cur.execute(
                    "SELECT email FROM student_info WHERE student_id = ?",
                    (individual_id,),
                )
                result = cur.fetchone()
                if result and result[0]:
                    recipients = [result[0]]
        conn.close()

        st.write(f"Total recipients: {len(recipients)}")
        if not recipients:
            st.error("No email addresses found for the selected criteria.")
            return

        # Send email using SMTP.
        try:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            for recipient in recipients:
                msg = MIMEMultipart()
                msg["From"] = SMTP_USERNAME
                msg["To"] = recipient
                msg["Subject"] = subject
                msg.attach(MIMEText(message_body, "plain"))
                server.sendmail(SMTP_USERNAME, recipient, msg.as_string())
            server.quit()
            st.success("Emails sent successfully!")
        except Exception as e:
            st.error(f"Error sending emails: {e}")


##############################
# End Email Sending Function #
##############################


def generate_batch_pdfs(document_type="student_info"):
    temp_dir = "temp_pdfs"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    try:
        conn = sqlite3.connect("student_registration.db")

        if document_type == "student_info":
            students_df = pd.read_sql_query("SELECT * FROM student_info", conn)
            for _, student in students_df.iterrows():
                try:
                    pdf_file = generate_student_info_pdf(student)
                    new_path = os.path.join(temp_dir, os.path.basename(pdf_file))
                    shutil.move(pdf_file, new_path)
                except Exception as e:
                    print(
                        f"Error generating PDF for student {student['student_id']}: {str(e)}"
                    )
                    continue

        else:
            registrations_df = pd.read_sql_query(
                """
                SELECT cr.*, si.surname, si.other_names, si.passport_photo_path
                FROM course_registration cr
                LEFT JOIN student_info si ON cr.student_id = si.student_id
            """,
                conn,
            )

            for _, registration in registrations_df.iterrows():
                try:
                    pdf_file = generate_course_registration_pdf(registration)
                    new_path = os.path.join(temp_dir, os.path.basename(pdf_file))
                    shutil.move(pdf_file, new_path)
                except Exception as e:
                    print(
                        f"Error generating PDF for registration {registration['registration_id']}: {str(e)}"
                    )
                    continue

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"all_{document_type}_pdfs_{timestamp}.zip"

        with zipfile.ZipFile(zip_filename, "w") as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.basename(file_path)
                    zipf.write(file_path, arcname)

        return zip_filename

    except Exception as e:
        print(f"Error in batch PDF generation: {str(e)}")
        return None

    finally:
        conn.close()
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


def download_forms():
    st.subheader("Download Forms")

    conn = sqlite3.connect("student_registration.db")

    col1, col2 = st.columns(2)

    with col1:
        st.write("**Student Information Database**")
        student_df = pd.read_sql_query(
            """
            SELECT 
                student_id, surname, other_names, date_of_birth, 
                place_of_birth, home_town, residential_address, 
                postal_address, email, telephone, ghana_card_id, 
                nationality, marital_status, gender, religion, 
                denomination, disability_status, disability_description,
                guardian_name, guardian_relationship, guardian_occupation,
                guardian_address, guardian_telephone, previous_school,
                qualification_type, completion_year, aggregate_score,
                ghana_card_path, passport_photo_path, transcript_path,
                certificate_path, receipt_path, receipt_amount,
                approval_status, created_at, programme
            FROM student_info
        """,
            conn,
        )

        if not student_df.empty:
            csv = student_df.to_csv(index=False)
            st.download_button(
                label="Download Student Database (CSV)",
                data=csv,
                file_name=f"student_database_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
            )
        else:
            st.info("No student records available")

    with col2:
        st.write("**Course Registration Database**")
        registration_df = pd.read_sql_query("SELECT * FROM course_registration", conn)

        if not registration_df.empty:
            csv = registration_df.to_csv(index=False)
            st.download_button(
                label="Download Course Registrations (CSV)",
                data=csv,
                file_name=f"course_registrations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
            )
        else:
            st.info("No registration records available")

    st.markdown("---")

    st.subheader("Download Filtered Data")

    filter_col1, filter_col2 = st.columns(2)

    with filter_col1:
        status_filter = st.selectbox(
            "Filter by Approval Status", ["All", "Pending", "Approved", "Rejected"]
        )

    with filter_col2:
        date_range = st.date_input(
            "Select Date Range",
            value=(datetime.now() - timedelta(days=30), datetime.now()),
            max_value=datetime.now(),
        )

    if len(date_range) == 2:
        start_date, end_date = date_range

        query = """
            SELECT * FROM student_info 
            WHERE date_of_birth BETWEEN ? AND ?
        """

        if status_filter != "All":
            query += f" AND approval_status = '{status_filter.lower()}'"

        try:
            filtered_df = pd.read_sql_query(query, conn, params=(start_date, end_date))

            if not filtered_df.empty:
                csv = filtered_df.to_csv(index=False)
                st.download_button(
                    label="Download Filtered Data (CSV)",
                    data=csv,
                    file_name=f"filtered_student_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                )
            else:
                st.info("No records found for the selected filters")

        except pd.errors.DatabaseError:
            st.error("Error filtering data. Please try different filter criteria.")

    conn.close()



import zipfile
import shutil
from datetime import datetime
import sqlite3
from typing import Dict, List, Optional, Tuple
import logging
import json

class DocumentUploadHandler:
    """
    Handles the upload and organization of documents from zip files for the student registration system.
    Ensures proper file structure and database updates for both student information and course registration.
    No file size limit enforced.
    """

    # Allowed file extensions for each document type
    ALLOWED_EXTENSIONS = {
        "ghana_card": [".pdf", ".jpg", ".jpeg", ".png"],
        "passport_photo": [".jpg", ".jpeg", ".png"],
        "transcript": [".pdf", ".jpeg", ".jpg", ".png"],
        "certificate": [".pdf", ".jpg", ".jpeg", ".png" ],
        "receipt": [".pdf", ".jpg", ".jpeg", ".png"],
    }

    # Expected folders in the zip file
    EXPECTED_FOLDERS = {
        "student_documents": [
            "ghana_cards",
            "passport_photos",
            "transcripts",
            "certificates",
            "receipts",
        ],
        "course_registration_receipts": ["receipts"],
    }

    def __init__(self, upload_base_dir: str = "uploads"):
        """
        Initialize the document upload handler.

        Args:
            upload_base_dir: Base directory for all uploaded files.
        """
        self.upload_base_dir = upload_base_dir
        self.logger = self._setup_logger()
        self._ensure_directories()

    def _ensure_directories(self):
        """
        Ensure that the base uploads directory exists and create subdirectories for expected folders.
        """
        self._ensure_base_directory()
        # Create subdirectories for each expected main folder
        for main_folder in self.EXPECTED_FOLDERS.keys():
            dir_path = os.path.join(self.upload_base_dir, main_folder)
            os.makedirs(dir_path, exist_ok=True)

    def _setup_logger(self) -> logging.Logger:
        """Configure logging for the document handler."""
        logger = logging.getLogger("DocumentUploadHandler")
        logger.setLevel(logging.INFO)
        if not logger.handlers:
            handler = logging.FileHandler("document_uploads.log")
            formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        return logger

    def _ensure_base_directory(self):
        """Ensure the base uploads directory exists."""
        os.makedirs(self.upload_base_dir, exist_ok=True)

    def process_zip_file(self, zip_file_path: str) -> Tuple[bool, str]:
        """
        Process the uploaded zip file, extract its contents, and save all files directly in the uploads folder.
        No file size limit enforced.

        Args:
            zip_file_path: Path to the uploaded zip file.

        Returns:
            Tuple of (success: bool, message: str).
        """
        temp_dir = "temp_extract"
        try:
            # Create temporary directory for extraction
            os.makedirs(temp_dir, exist_ok=True)

            # Extract zip file without size restrictions
            with zipfile.ZipFile(zip_file_path, "r") as zip_ref:
                zip_ref.extractall(temp_dir)

            # Process student documents
            student_documents = self._process_student_documents(temp_dir)

            # Process registration documents
            reg_docs = self._process_registration_documents(temp_dir)

            # Update database with new file paths
            self._update_database(student_documents, reg_docs)

            return True, "Documents processed successfully"

        except Exception as e:
            self.logger.error(f"Error processing zip file: {str(e)}")
            return False, f"Error processing documents: {str(e)}"

        finally:
            # Clean up temporary directory
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

    def _process_student_documents(self, temp_dir: str) -> Dict[str, Dict[str, str]]:
        """Process student documents from the extracted zip file."""
        student_documents = {}
        student_dir = os.path.join(temp_dir, "student_documents")

        if not os.path.exists(student_dir):
            return student_documents

        for student_id in os.listdir(student_dir):
            student_folder = os.path.join(student_dir, student_id)
            if not os.path.isdir(student_folder):
                continue

            student_documents[student_id] = {}

            # Process each document type
            for doc_type, extensions in self.ALLOWED_EXTENSIONS.items():
                doc_file = self._find_document(student_folder, doc_type, extensions)
                if doc_file:
                    new_path = self._save_document(doc_file, student_id, doc_type)
                    student_documents[student_id][f"{doc_type}_path"] = new_path

        return student_documents

    def _process_registration_documents(self, temp_dir: str) -> Dict[str, Dict[str, str]]:
        """Process course registration documents from the extracted zip file."""
        reg_docs = {}
        reg_dir = os.path.join(temp_dir, "course_registration_receipts")

        if not os.path.exists(reg_dir):
            return reg_docs

        for reg_id in os.listdir(reg_dir):
            reg_folder = os.path.join(reg_dir, reg_id)
            if not os.path.isdir(reg_folder):
                continue

            reg_docs[reg_id] = {}

            # Process registration receipt
            receipt_file = self._find_document(reg_folder, "receipt", self.ALLOWED_EXTENSIONS["receipt"])
            if receipt_file:
                new_path = self._save_document(receipt_file, reg_id, "receipt")
                reg_docs[reg_id]["receipt_path"] = new_path

        return reg_docs

    def _find_document(self, directory: str, doc_type: str, allowed_extensions: List[str]) -> Optional[str]:
        """Find a document of the specified type in the directory."""
        for file in os.listdir(directory):
            ext = os.path.splitext(file)[1].lower()
            if ext in allowed_extensions:
                return os.path.join(directory, file)
        return None

    def _save_document(self, source_path: str, identifier: str, doc_type: str) -> str:
        """Save a document to the uploads directory."""
        ext = os.path.splitext(source_path)[1]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{identifier}_{doc_type}_{timestamp}{ext}"
        dest_path = os.path.join(self.upload_base_dir, filename)
        shutil.copy2(source_path, dest_path)
        return dest_path

    def _update_database(self, student_documents: Dict[str, Dict[str, str]], reg_docs: Dict[str, Dict[str, str]]):
        """Update database with new document paths."""
        conn = sqlite3.connect("student_registration.db")
        c = conn.cursor()

        try:
            # Update student documents
            for student_id, docs in student_documents.items():
                if docs:
                    update_fields = ", ".join(f"{key} = ?" for key in docs.keys())
                    query = f"UPDATE student_info SET {update_fields} WHERE student_id = ?"
                    c.execute(query, list(docs.values()) + [student_id])

            # Update registration documents
            for reg_id, docs in reg_docs.items():
                if docs:
                    update_fields = ", ".join(f"{key} = ?" for key in docs.keys())
                    query = f"UPDATE course_registration SET {update_fields} WHERE registration_id = ?"
                    c.execute(query, list(docs.values()) + [reg_id])

            conn.commit()

        except Exception as e:
            conn.rollback()
            self.logger.error(f"Database update error: {str(e)}")
            raise

        finally:
            conn.close()

    def validate_zip_structure(self, zip_file_path: str) -> Tuple[bool, str]:
        """Validate the structure of the uploaded zip file."""
        try:
            with zipfile.ZipFile(zip_file_path, "r") as zip_ref:
                files = zip_ref.namelist()

                # Check for main directories
                if (not any(f.startswith("student_documents/") for f in files) and 
                    not any(f.startswith("course_registration_receipts/") for f in files)):
                    return (False, "Missing required directories: student_documents or course_registration_receipts")

                # Validate file extensions
                for file in files:
                    if file.endswith("/"):  # Skip directories
                        continue
                    ext = os.path.splitext(file)[1].lower()
                    if not any(ext in exts for exts in self.ALLOWED_EXTENSIONS.values()):
                        return False, f"Invalid file extension in: {file}"

                return True, "Zip file structure is valid"

        except zipfile.BadZipFile:
            return False, "Invalid zip file"
        except Exception as e:
            return False, f"Error validating zip file: {str(e)}"
  
# End of DocumentUploadHandler


class NotificationSystem:
    def __init__(self):
        self.setup_notification_table()

    def setup_notification_table(self):
        """Initialize the notifications table in the database"""
        conn = sqlite3.connect("student_registration.db")
        c = conn.cursor()

        # Create notifications table
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS notifications (
                notification_id INTEGER PRIMARY KEY AUTOINCREMENT,
                recipient_id TEXT,
                recipient_type TEXT,  -- 'student', 'program', or 'all'
                title TEXT,
                message TEXT,
                notification_type TEXT,  -- 'info', 'warning', 'success', 'error'
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                read_at DATETIME,
                metadata TEXT,  -- JSON field for additional data
                expires_at DATETIME,
                FOREIGN KEY (recipient_id) REFERENCES student_info (student_id)
            )
        """
        )

        # Create notification_reads table for tracking read status
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS notification_reads (
                notification_id INTEGER,
                student_id TEXT,
                read_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (notification_id, student_id),
                FOREIGN KEY (notification_id) REFERENCES notifications (notification_id),
                FOREIGN KEY (student_id) REFERENCES student_info (student_id)
            )
        """
        )

        conn.commit()
        conn.close()

    def create_notification(
        self,
        title: str,
        message: str,
        recipient_type: str,
        recipient_id: Union[str, None] = None,
        notification_type: str = "info",
        metadata: Dict = None,
        expires_at: Union[datetime, None] = None,
    ) -> int:
        """
        Create a new notification
        Returns the notification ID
        """
        conn = sqlite3.connect("student_registration.db")
        c = conn.cursor()

        try:
            c.execute(
                """
                INSERT INTO notifications (
                    recipient_id, recipient_type, title, message,
                    notification_type, metadata, expires_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    recipient_id,
                    recipient_type,
                    title,
                    message,
                    notification_type,
                    json.dumps(metadata) if metadata else None,
                    expires_at.isoformat() if expires_at else None,
                ),
            )
            notification_id = c.lastrowid
            conn.commit()
            return notification_id
        finally:
            conn.close()

    def get_notifications(
        self, student_id: str, include_read: bool = False, limit: int = 50
    ) -> List[Dict]:
        """Get notifications for a specific student"""
        conn = sqlite3.connect("student_registration.db")
        c = conn.cursor()

        try:
            query = """
                SELECT 
                    n.notification_id,
                    n.title,
                    n.message,
                    n.notification_type,
                    n.created_at,
                    n.metadata,
                    CASE 
                        WHEN nr.read_at IS NOT NULL THEN 1 
                        ELSE 0 
                    END as is_read
                FROM notifications n
                LEFT JOIN notification_reads nr 
                    ON n.notification_id = nr.notification_id 
                    AND nr.student_id = ?
                WHERE (
                    n.recipient_id = ? 
                    OR n.recipient_type = 'all'
                    OR (
                        n.recipient_type = 'program' 
                        AND n.recipient_id = (
                            SELECT programme 
                            FROM student_info 
                            WHERE student_id = ?
                        )
                    )
                )
                AND (n.expires_at IS NULL OR n.expires_at > datetime('now'))
            """

            if not include_read:
                query += " AND nr.read_at IS NULL"

            query += " ORDER BY n.created_at DESC LIMIT ?"

            c.execute(query, (student_id, student_id, student_id, limit))

            notifications = []
            for row in c.fetchall():
                notification = {
                    "id": row[0],
                    "title": row[1],
                    "message": row[2],
                    "type": row[3],
                    "created_at": row[4],
                    "metadata": json.loads(row[5]) if row[5] else None,
                    "is_read": bool(row[6]),
                }
                notifications.append(notification)

            return notifications
        finally:
            conn.close()

    def mark_as_read(self, notification_id: int, student_id: str):
        """Mark a notification as read for a specific student"""
        conn = sqlite3.connect("student_registration.db")
        c = conn.cursor()

        try:
            c.execute(
                """
                INSERT OR IGNORE INTO notification_reads (
                    notification_id, student_id
                ) VALUES (?, ?)
            """,
                (notification_id, student_id),
            )
            conn.commit()
        finally:
            conn.close()

    def mark_all_as_read(self, student_id: str):
        """Mark all notifications as read for a specific student"""
        conn = sqlite3.connect("student_registration.db")
        c = conn.cursor()

        try:
            # Get all unread notifications for the student
            c.execute(
                """
                INSERT INTO notification_reads (notification_id, student_id)
                SELECT n.notification_id, ?
                FROM notifications n
                LEFT JOIN notification_reads nr 
                    ON n.notification_id = nr.notification_id 
                    AND nr.student_id = ?
                WHERE nr.read_at IS NULL
                AND (
                    n.recipient_id = ? 
                    OR n.recipient_type = 'all'
                    OR (
                        n.recipient_type = 'program' 
                        AND n.recipient_id = (
                            SELECT programme 
                            FROM student_info 
                            WHERE student_id = ?
                        )
                    )
                )
            """,
                (student_id, student_id, student_id, student_id),
            )
            conn.commit()
        finally:
            conn.close()

    def delete_notification(self, notification_id: int):
        """Delete a notification"""
        conn = sqlite3.connect("student_registration.db")
        c = conn.cursor()

        try:
            c.execute(
                "DELETE FROM notification_reads WHERE notification_id = ?",
                (notification_id,),
            )
            c.execute(
                "DELETE FROM notifications WHERE notification_id = ?",
                (notification_id,),
            )
            conn.commit()
        finally:
            conn.close()


def display_notifications(notifications: List[Dict]):
    """Display notifications in the Streamlit UI"""
    if not notifications:
        st.info("No new notifications")
        return

    for notification in notifications:
        with st.expander(
            f"📢 {notification['title']} - {notification['created_at']}",
            expanded=not notification["is_read"],
        ):
            st.write(notification["message"])

            # Display metadata if available
            if notification["metadata"]:
                st.json(notification["metadata"])

            col1, col2 = st.columns([4, 1])
            with col2:
                if not notification["is_read"]:
                    if st.button("Mark as Read", key=f"read_{notification['id']}"):
                        notification_system = NotificationSystem()
                        notification_system.mark_as_read(
                            notification["id"], st.session_state.student_logged_in
                        )
                        st.rerun()


def admin_notification_interface():
    """Admin interface for creating and managing notifications"""
    st.subheader("📢 Notification Management")

    tab1, tab2 = st.tabs(["Create Notification", "View/Manage Notifications"])

    with tab1:
        recipient_type = st.selectbox(
            "Recipient Type",
            ["all", "program", "student"],
            help="Select who should receive this notification",
        )

        recipient_id = None
        if recipient_type == "program":
            recipient_id = st.selectbox(
                "Select Program", ["CIMG", "CIM-UK", "ICAG", "ACCA"]
            )
        elif recipient_type == "student":
            # Get list of student IDs from database
            conn = sqlite3.connect("student_registration.db")
            c = conn.cursor()
            c.execute(
                """
                SELECT student_id, surname, other_names 
                FROM student_info 
                ORDER BY surname, other_names
            """
            )
            students = c.fetchall()
            conn.close()

            if students:
                options = [
                    f"{id} - {surname} {other_names}"
                    for id, surname, other_names in students
                ]
                selected = st.selectbox("Select Student", options)
                recipient_id = selected.split(" - ")[0]
            else:
                st.warning("No students found in database")
                return

        notification_type = st.selectbox(
            "Notification Type", ["info", "warning", "success", "error"]
        )

        title = st.text_input("Notification Title")
        message = st.text_area("Notification Message")

        col1, col2 = st.columns(2)
        with col1:
            include_metadata = st.checkbox("Include Additional Data")
        with col2:
            include_expiry = st.checkbox("Set Expiration")

        metadata = None
        if include_metadata:
            metadata_str = st.text_area(
                "Additional Data (JSON format)", help="Enter valid JSON data"
            )
            try:
                metadata = json.loads(metadata_str) if metadata_str else None
            except json.JSONDecodeError:
                st.error("Invalid JSON format")
                return

        expires_at = None
        if include_expiry:
            expires_at = st.date_input("Expiration Date")
            if expires_at:
                expires_at = datetime.combine(expires_at, datetime.min.time())

        if st.button("Send Notification"):
            if not title or not message:
                st.error("Title and message are required")
                return

            try:
                notification_system = NotificationSystem()
                notification_id = notification_system.create_notification(
                    title=title,
                    message=message,
                    recipient_type=recipient_type,
                    recipient_id=recipient_id,
                    notification_type=notification_type,
                    metadata=metadata,
                    expires_at=expires_at,
                )
                st.success(f"Notification created successfully! ID: {notification_id}")
            except Exception as e:
                st.error(f"Error creating notification: {str(e)}")

    with tab2:
        conn = sqlite3.connect("student_registration.db")
        c = conn.cursor()

        c.execute(
            """
            SELECT 
                n.notification_id,
                n.title,
                n.message,
                n.recipient_type,
                n.recipient_id,
                n.created_at,
                COUNT(nr.student_id) as read_count
            FROM notifications n
            LEFT JOIN notification_reads nr 
                ON n.notification_id = nr.notification_id
            GROUP BY n.notification_id
            ORDER BY n.created_at DESC
        """
        )

        notifications = c.fetchall()
        conn.close()

        if notifications:
            for notification in notifications:
                with st.expander(f"📢 {notification[1]} - {notification[5]}"):
                    st.write(f"**Message:** {notification[2]}")
                    st.write(f"**Type:** {notification[3]}")
                    st.write(f"**Recipient:** {notification[4] or 'All'}")
                    st.write(f"**Read by:** {notification[6]} students")

                    if st.button("Delete", key=f"del_{notification[0]}"):
                        try:
                            notification_system = NotificationSystem()
                            notification_system.delete_notification(notification[0])
                            st.success("Notification deleted successfully!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error deleting notification: {str(e)}")
        else:
            st.info("No notifications found")


def main():
    st.set_page_config(
        page_title="Student Registration System",
        page_icon="🎓",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown(
        """
        <style>
        .main {
            padding: 2rem;
        }
        .stButton>button {
            width: 100%;
            background-color: #4CAF50;
            color: white;
        }
        .upload-section {
            padding: 1rem;
            border-radius: 5px;
            background-color: #f0f2f6;
            margin: 1rem 0;
        }
        .login-form {
            max-width: 400px;
            margin: 0 auto;
            padding: 2rem;
            background-color: #f8f9fa;
            border-radius: 10px;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 24px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 50px;
            white-space: pre-wrap;
            background-color: #f0f2f6;
            border-radius: 5px;
            padding: 10px;
        }
        </style>
    """,
        unsafe_allow_html=True,
    )

    if "db_initialized" not in st.session_state:
        reset_db()
        st.session_state.db_initialized = True

    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False

    if not st.session_state.admin_logged_in:
        with st.sidebar:
            st.subheader("Admin Login")
            username = st.text_input("Username", key="login_username")
            password = st.text_input("Password", type="password", key="login_password")

            if st.button("Login"):
                conn = sqlite3.connect("student_registration.db")
                c = conn.cursor()
                c.execute(
                    "SELECT * FROM admin WHERE username=? AND password=?",
                    (username, password),
                )
                admin = c.fetchone()
                conn.close()

                if admin:
                    st.session_state.admin_logged_in = True
                    st.rerun()
                else:
                    st.error("Invalid credentials")

    if st.session_state.admin_logged_in:
        admin_dashboard()
    else:
        # Add this to the navigation options in main()
        page = st.sidebar.radio(
            "Navigation",
            ["Student Portal", "Student Information", "Course Registration"],
        )

        if page == "Student Portal":
            if "student_logged_in" not in st.session_state:
                student_login_form()
            else:
                student_portal()
        elif page == "Student Information":
            student_info_form()
        else:
            course_registration_form()


if __name__ == "__main__":
    init_db()
    main()
