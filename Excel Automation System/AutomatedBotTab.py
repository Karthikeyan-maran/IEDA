# AutomatedBotTab.py
import pydrive2
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import schedule
import time

import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import uuid
from datetime import datetime

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QMessageBox, QGroupBox,
    QGridLayout, QCheckBox
)
from PyQt5.QtCore import pyqtSignal

# The ReportGenerator class is now defined here to avoid circular imports.
class ReportGenerator:
    def __init__(self, dataframes, analysis_results, dashboard_chart_path=None):
        self.dataframes = dataframes
        self.analysis_results = analysis_results
        self.dashboard_chart_path = dashboard_chart_path
        self.email_worker = None

    def generate_excel(self, df, file_path):
        try:
            df.to_excel(file_path, index=False)
            return True
        except Exception as e:
            QMessageBox.critical(None, "Excel Generation Error", f"Failed to generate Excel file: {e}")
            return False

    def generate_pdf(self, file_path):
        try:
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, "Data Analysis Report", 0, 1, 'C')
            pdf.set_font("Arial", "", 12)
            
            for name, df in self.dataframes.items():
                pdf.add_page()
                pdf.set_font("Arial", "B", 14)
                pdf.cell(0, 10, f"DataFrame: {name}", 0, 1, 'L')
                pdf.set_font("Arial", "", 10)
                
                pdf.ln(5)
                headings = list(df.columns)
                table_data = [headings] + df.head(10).values.tolist()
                
                col_width = pdf.w / len(headings) - 2 * pdf.l_margin
                for row in table_data:
                    for item in row:
                        pdf.cell(col_width, 6, str(item), 1)
                    pdf.ln()

            # Add Analysis Results
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "Analysis Summary", 0, 1, 'L')
            pdf.set_font("Arial", "", 10)
            pdf.multi_cell(0, 5, self.analysis_results)
            
            # Add Dashboard Chart if available
            if self.dashboard_chart_path and os.path.exists(self.dashboard_chart_path):
                pdf.add_page()
                pdf.set_font("Arial", "B", 14)
                pdf.cell(0, 10, "Dashboard Chart", 0, 1, 'L')
                pdf.image(self.dashboard_chart_path, x=10, y=30, w=180)
            
            pdf.output(file_path, 'F')
            return True
        except Exception as e:
            QMessageBox.critical(None, "PDF Generation Error", f"Failed to generate PDF file: {e}")
            return False


class AutomatedBotTab(QWidget):
    report_generated = pyqtSignal(str, str, str) # file_name, file_type, file_path

    def __init__(self, dataframes, report_generator, parent=None):
        super().__init__(parent)
        self.dataframes = dataframes
        self.report_generator = report_generator
        self.drive = None
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)

        setup_group = QGroupBox("1. Google Drive & Operations Setup")
        setup_layout = QGridLayout()

        setup_layout.addWidget(QLabel("Google Drive Folder ID:"), 0, 0)
        self.drive_folder_id_input = QLineEdit()
        self.drive_folder_id_input.setPlaceholderText("Enter Google Drive Folder ID")
        setup_layout.addWidget(self.drive_folder_id_input, 0, 1)

        self.authenticate_btn = QPushButton("Authenticate Drive")
        self.authenticate_btn.clicked.connect(self.authenticate_drive)
        setup_layout.addWidget(self.authenticate_btn, 1, 0, 1, 2)
        
        setup_layout.addWidget(QLabel("Operation to Automate:"), 2, 0)
        self.operation_combo = QComboBox()
        self.operation_combo.addItems(["Clear Missing Values", "Create Dashboard Report", "Create Analysis Report"])
        setup_layout.addWidget(self.operation_combo, 2, 1)

        self.run_automation_btn = QPushButton("Run Automation Now")
        self.run_automation_btn.clicked.connect(self.run_automation_flow)
        setup_layout.addWidget(self.run_automation_btn, 3, 0, 1, 2)
        
        setup_group.setLayout(setup_layout)
        main_layout.addWidget(setup_group)

        schedule_group = QGroupBox("2. Scheduling Options (Optional)")
        schedule_layout = QGridLayout()
        
        self.schedule_checkbox = QCheckBox("Enable Scheduling")
        schedule_layout.addWidget(self.schedule_checkbox, 0, 0, 1, 2)
        
        schedule_layout.addWidget(QLabel("Schedule Time (24h):"), 1, 0)
        self.schedule_time_input = QLineEdit()
        self.schedule_time_input.setPlaceholderText("HH:MM")
        schedule_layout.addWidget(self.schedule_time_input, 1, 1)

        self.start_scheduler_btn = QPushButton("Start Scheduler")
        self.start_scheduler_btn.clicked.connect(self.start_scheduler)
        schedule_layout.addWidget(self.start_scheduler_btn, 2, 0)

        self.stop_scheduler_btn = QPushButton("Stop Scheduler")
        self.stop_scheduler_btn.setEnabled(False)
        self.stop_scheduler_btn.clicked.connect(self.stop_scheduler)
        schedule_layout.addWidget(self.stop_scheduler_btn, 2, 1)

        schedule_group.setLayout(schedule_layout)
        main_layout.addWidget(schedule_group)

        main_layout.addStretch(1)

    def authenticate_drive(self):
        """
        Authenticates with Google Drive API using a local client_secrets.json file.
        You must first set up a Google Cloud Project, enable the Drive API,
        and download the client_secrets.json file.
        See: https://pygdrive2.readthedocs.io/en/latest/quickstart.html#authentication
        """
        try:
            gauth = GoogleAuth()
            gauth.LocalWebserverAuth() # Creates local webserver and handles authorization
            self.drive = GoogleDrive(gauth)
            QMessageBox.information(self, "Authentication Success", "Successfully authenticated with Google Drive!")
        except Exception as e:
            QMessageBox.critical(self, "Authentication Failed", f"Failed to authenticate with Google Drive: {e}")
            self.drive = None

    def run_automation_flow(self):
        folder_id = self.drive_folder_id_input.text().strip()
        if not self.drive or not folder_id:
            QMessageBox.warning(self, "Error", "Please authenticate with Google Drive and enter a folder ID.")
            return

        try:
            # Step 1: Find the first Excel file in the folder
            file_list = self.drive.ListFile({'q': f"'{folder_id}' in parents and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel') and trashed=false"}).GetList()
            if not file_list:
                QMessageBox.warning(self, "File Error", "No Excel files found in the specified folder.")
                return
            
            first_excel_file = file_list[0]
            file_name = first_excel_file['title']
            file_path = f"temp_{file_name}"
            first_excel_file.GetContentFile(file_path)

            # Step 2: Load the file into a DataFrame
            df = pd.read_excel(file_path)
            
            # Step 3: Perform the selected operation
            operation = self.operation_combo.currentText()
            if operation == "Clear Missing Values":
                df.dropna(inplace=True)
                QMessageBox.information(self, "Operation Complete", "Missing values have been cleared.")
            
            # Step 4: Generate reports
            self.generate_reports(df)
            
            os.remove(file_path) # Clean up temp file
            
        except Exception as e:
            QMessageBox.critical(self, "Automation Error", f"An error occurred during automation: {e}")

    def generate_reports(self, df):
        # You would need to pass a valid `ReportGenerator` instance to the class constructor
        # This is just a placeholder to show the flow
        report_text = df.describe().to_string() # Placeholder analysis
        pdf_file_name = f"Auto_Analysis_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        pdf_file_path = f"temp_{uuid.uuid4().hex}.pdf"
        
        self.report_generator.analysis_results = report_text
        if self.report_generator.generate_pdf(pdf_file_path):
            self.report_generated.emit(pdf_file_name, "pdf", pdf_file_path)
            QMessageBox.information(self, "Reports Generated", "Analysis report saved to the 'Files' tab.")
        
    def start_scheduler(self):
        schedule_time = self.schedule_time_input.text().strip()
        if not self.schedule_checkbox.isChecked() or not schedule_time:
            QMessageBox.warning(self, "Scheduler Error", "Please enable scheduling and enter a valid time.")
            return

        try:
            schedule.every().day.at(schedule_time).do(self.run_automation_flow)
            
            self.start_scheduler_btn.setEnabled(False)
            self.stop_scheduler_btn.setEnabled(True)
            QMessageBox.information(self, "Scheduler", f"Automation scheduled to run daily at {schedule_time}.")

            # In a real app, this would run in a separate thread to not block the UI
            # For this example, it's just a simple start/stop
            while self.start_scheduler_btn.isEnabled():
                schedule.run_pending()
                time.sleep(1)

        except Exception as e:
            QMessageBox.critical(self, "Scheduler Error", f"Failed to start scheduler: {e}")
            self.start_scheduler_btn.setEnabled(True)
            self.stop_scheduler_btn.setEnabled(False)
    
    def stop_scheduler(self):
        schedule.clear()
        self.start_scheduler_btn.setEnabled(True)
        self.stop_scheduler_btn.setEnabled(False)
        QMessageBox.information(self, "Scheduler", "Scheduler has been stopped.")