import sys
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QComboBox, QLineEdit, QTabWidget, QTextEdit, QMessageBox, QGroupBox,
    QHeaderView, QRadioButton, QListWidget, QStyleFactory, QDialog,
    QGridLayout, QAbstractItemView, QFrame, QInputDialog, QAction, QCheckBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QFont, QDoubleValidator, QIntValidator, QIcon

from sklearn.preprocessing import StandardScaler
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import classification_report, accuracy_score, mean_absolute_error
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, r2_score
import numpy as np
import ollama
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import os
import io
import uuid
from datetime import datetime

# Import both the AutomatedBotTab and ReportGenerator from the separate file
from AutomatedBotTab import AutomatedBotTab, ReportGenerator

# --- QSS for Modern Dark Theme ---
MODERN_DARK_QSS = """
QWidget {
    background-color: #1a1a2e;
    color: #f0f0f0;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 14px;
}
QMainWindow, QTabWidget {
    background-color: #16162a;
}
QGroupBox {
    border: 1px solid #4a4a6e;
    border-radius: 8px;
    margin-top: 20px;
    padding-top: 15px;
    padding-bottom: 10px;
    font-weight: bold;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top center;
    padding: 0 10px;
    color: #94a3b8;
}
QPushButton {
    background-color: #3e2f5b;
    border: 1px solid #5a4b7a;
    border-radius: 8px;
    padding: 12px 24px;
    min-width: 120px;
    max-width: 200px;
    font-weight: bold;
    color: #ffffff;
}
QPushButton:hover {
    background-color: #5a4b7a;
}
QPushButton:pressed {
    background-color: #2b2241;
}
QPushButton:disabled {
    background-color: #2a2a3e;
    border: 1px solid #3e3e5a;
    color: #6a6a8a;
}
QLineEdit, QTextEdit, QComboBox, QListWidget, QTableWidget {
    background-color: #2b2b4d;
    border: 1px solid #4a4a6e;
    border-radius: 6px;
    padding: 8px;
    color: #e0e0e0;
}
QComboBox::drop-down {
    border: 0px;
}
QComboBox::down-arrow {
    image: url(down_arrow.png);
    width: 16px;
    height: 16px;
}
QTableWidget {
    background-color: #2b2b4d;
    gridline-color: #4a4a6e;
    border: 1px solid #4a4a6e;
}
QTableWidget QHeaderView::section {
    background-color: #3e2f5b;
    color: #ffffff;
    border-bottom: 1px solid #4a4a6e;
    padding: 8px;
    font-weight: bold;
}
QTableWidget::item {
    padding: 6px;
}
QTableWidget::item:selected {
    background-color: #475a85;
    color: #ffffff;
}
QTabWidget::pane {
    border: 1px solid #4a4a6e;
    background-color: #1a1a2e;
    border-radius: 8px;
}
QTabBar::tab {
    background: #2b2b4d;
    border: 1px solid #4a4a6e;
    border-bottom-color: #1a1a2e;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    min-width: 150px;
    padding: 10px 20px;
    margin-right: 4px;
    font-weight: bold;
}
QTabBar::tab:selected {
    background: #1a1a2e;
    border-bottom-color: #1a1a2e;
    color: #a0d9ff;
}
QListWidget {
    background-color: #2b2b4d;
    border: 1px solid #4a4a6e;
    border-radius: 6px;
    padding: 5px;
}
QListWidget::item:selected {
    background-color: #475a85;
    color: #ffffff;
}
QCheckBox {
    spacing: 5px;
    color: #f0f0f0;
}
QCheckBox::indicator {
    width: 18px;
    height: 18px;
}
QCheckBox::indicator:unchecked {
    image: url(unchecked.png);
}
QCheckBox::indicator:checked {
    image: url(checked.png);
}
"""

class FullSheetPreviewDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Full Excel Sheet Preview")
        self.setGeometry(200, 200, 1000, 800)
        self.setStyleSheet(MODERN_DARK_QSS)

        layout = QVBoxLayout(self)
        
        self.table = QTableWidget()
        self.table.setColumnCount(len(df.columns))
        self.table.setRowCount(len(df))
        self.table.setHorizontalHeaderLabels(df.columns.tolist())
        
        for i in range(len(df)):
            for j, col in enumerate(df.columns):
                item = QTableWidgetItem(str(df.iloc[i, j]))
                self.table.setItem(i, j, item)
        
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addWidget(self.table)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.accept)
        layout.addWidget(cancel_btn)

class DashboardApp(QWidget):
    dataframes_updated = pyqtSignal(dict)
    report_generated = pyqtSignal(str, str, str) # file_name, file_type, file_path

    def __init__(self, initial_dataframes=None):
        super().__init__()
        self.dataframes = initial_dataframes if initial_dataframes is not None else {}
        self.current_df_name = None
        self.current_df = None
        self.filtered_df = None
        self.last_chart_image_path = None

        self.figure = plt.figure(facecolor='#1a1a2e')
        self.canvas = FigureCanvas(self.figure)
        
        self.filter_column_combo = QComboBox()
        self.filter_value_combo = QComboBox()
        self.apply_filter_btn = QPushButton("Apply Filter")
        self.clear_filter_btn = QPushButton("Clear Filter")
        self.histogram_bins_input = QLineEdit()
        self.histogram_bins_input.setValidator(QIntValidator(1, 1000))
        self.histogram_bins_label = QLabel("Number of Bins:")
        
        self.init_ui()
        self.update_file_dropdown()

    def init_ui(self):
        main_layout = QHBoxLayout(self)
        
        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.StyledPanel)
        left_panel.setStyleSheet("QFrame { background-color: #1a1a2e; border: 1px solid #4a4a6e; border-radius: 8px; }")
        left_layout = QVBoxLayout(left_panel)
        
        select_data_group = QGroupBox("1. Select Data")
        select_data_layout = QGridLayout()
        select_data_layout.addWidget(QLabel("Select Excel File:"), 0, 0)
        self.file_combo = QComboBox()
        self.file_combo.currentIndexChanged.connect(self.load_selected_file_for_dashboard)
        self.file_combo.setEnabled(False)
        select_data_layout.addWidget(self.file_combo, 0, 1)
        select_data_group.setLayout(select_data_layout)
        left_layout.addWidget(select_data_group)
        
        preview_group = QGroupBox("2. Data Preview (First 5 Rows)")
        preview_layout = QVBoxLayout()
        self.preview_table = QTableWidget()
        self.preview_table.setMinimumHeight(120)
        preview_layout.addWidget(self.preview_table)
        preview_group.setLayout(preview_layout)
        left_layout.addWidget(preview_group)

        columns_group = QGroupBox("3. Select Columns")
        columns_layout = QVBoxLayout()
        self.column_list = QListWidget()
        self.column_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.column_list.setMinimumHeight(100)
        self.column_list.setEnabled(False)
        self.column_list.itemSelectionChanged.connect(self.update_chart_options)
        columns_layout.addWidget(self.column_list)
        columns_group.setLayout(columns_layout)
        left_layout.addWidget(columns_group)

        filter_group = QGroupBox("4. Filter Data")
        filter_layout = QGridLayout()
        filter_layout.addWidget(QLabel("Column:"), 0, 0)
        self.filter_column_combo.currentIndexChanged.connect(self.update_filter_values)
        self.filter_column_combo.setEnabled(False)
        filter_layout.addWidget(self.filter_column_combo, 0, 1)
        filter_layout.addWidget(QLabel("Value:"), 1, 0)
        self.filter_value_combo.setEnabled(False)
        filter_layout.addWidget(self.filter_value_combo, 1, 1)
        
        filter_btn_layout = QHBoxLayout()
        self.apply_filter_btn.clicked.connect(self.apply_filter)
        self.apply_filter_btn.setEnabled(False)
        filter_btn_layout.addWidget(self.apply_filter_btn)
        self.clear_filter_btn.clicked.connect(self.clear_filter)
        self.clear_filter_btn.setEnabled(False)
        filter_btn_layout.addWidget(self.clear_filter_btn)
        
        filter_layout.addLayout(filter_btn_layout, 2, 0, 1, 2)
        filter_group.setLayout(filter_layout)
        left_layout.addWidget(filter_group)
        
        left_layout.addStretch(1)

        right_panel = QFrame()
        right_panel.setFrameShape(QFrame.StyledPanel)
        right_panel.setStyleSheet("QFrame { background-color: #1a1a2e; border: 1px solid #4a4a6e; border-radius: 8px; }")
        right_layout = QVBoxLayout(right_panel)
        
        chart_controls_group = QGroupBox("5. Chart Options")
        chart_controls_layout = QGridLayout()
        
        chart_controls_layout.addWidget(QLabel("Chart Type:"), 0, 0)
        self.chart_type_combo = QComboBox()
        self.chart_type_combo.addItems(["", "Line Chart", "Bar Chart", "Pie Chart", "Scatter Plot", "Histogram"])
        self.chart_type_combo.currentIndexChanged.connect(self.update_chart_specific_options)
        self.chart_type_combo.setEnabled(False)
        chart_controls_layout.addWidget(self.chart_type_combo, 0, 1)
        
        chart_controls_layout.addWidget(QLabel("Chart Title:"), 1, 0)
        self.chart_title_input = QLineEdit()
        self.chart_title_input.setPlaceholderText("Enter Chart Title")
        chart_controls_layout.addWidget(self.chart_title_input, 1, 1)
        
        self.histogram_bins_label.hide()
        self.histogram_bins_input.setText("10")
        self.histogram_bins_input.hide()
        chart_controls_layout.addWidget(self.histogram_bins_label, 2, 0)
        chart_controls_layout.addWidget(self.histogram_bins_input, 2, 1)
        
        self.generate_chart_btn = QPushButton("Generate Chart")
        self.generate_chart_btn.clicked.connect(self.generate_chart)
        self.generate_chart_btn.setEnabled(False)
        chart_controls_layout.addWidget(self.generate_chart_btn, 3, 0, 1, 2)
        
        chart_controls_group.setLayout(chart_controls_layout)
        right_layout.addWidget(chart_controls_group)
        
        right_layout.addWidget(self.canvas, 1)

        main_layout.addWidget(left_panel, 1)
        main_layout.addWidget(right_panel, 3)

    def update_dataframes(self, new_dataframes):
        self.dataframes = new_dataframes
        self.update_file_dropdown()
        self.filtered_df = None

    def update_file_dropdown(self):
        self.file_combo.clear()
        if self.dataframes:
            self.file_combo.addItems(list(self.dataframes.keys()))
            if self.file_combo.count() > 0:
                self.file_combo.setCurrentIndex(0)
                self.file_combo.setEnabled(True)
        else:
            self.file_combo.setPlaceholderText("No Excel files loaded in main app.")
            self.file_combo.setEnabled(False)
    
    def load_selected_file_for_dashboard(self):
        self.current_df_name = self.file_combo.currentText()
        self.column_list.clear()
        self.preview_table.clearContents()
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)
        self.current_df = None
        self.filtered_df = None
        self.clear_filter_ui()

        self.column_list.setEnabled(False)
        self.chart_type_combo.setEnabled(False)
        self.generate_chart_btn.setEnabled(False)
        self.update_chart_specific_options()

        if self.current_df_name and self.current_df_name in self.dataframes:
            self.current_df = self.dataframes[self.current_df_name]
            self.filtered_df = self.current_df.copy()
            self.display_dataframe_in_table(self.filtered_df.head(5), self.preview_table)

            for col in self.filtered_df.columns:
                self.column_list.addItem(col)
            self.column_list.setEnabled(True)
            self.update_chart_options()
            
            self.filter_column_combo.clear()
            self.filter_column_combo.addItems([""] + self.filtered_df.columns.tolist())
            self.filter_column_combo.setEnabled(True)
            self.apply_filter_btn.setEnabled(True)
            self.clear_filter_btn.setEnabled(True)
        else:
            pass

    def update_filter_values(self):
        selected_col = self.filter_column_combo.currentText()
        self.filter_value_combo.clear()
        if selected_col and self.current_df is not None:
            unique_values = self.current_df[selected_col].dropna().unique()
            self.filter_value_combo.addItems([str(v) for v in unique_values])
            self.filter_value_combo.setEnabled(True)
        else:
            self.filter_value_combo.setEnabled(False)

    def display_dataframe_in_table(self, df, table_widget):
        table_widget.clearContents()
        table_widget.setRowCount(len(df))
        table_widget.setColumnCount(len(df.columns))
        table_widget.setHorizontalHeaderLabels(df.columns.tolist())
        for i in range(len(df)):
            for j, col in enumerate(df.columns):
                item = QTableWidgetItem(str(df.iloc[i, j]))
                table_widget.setItem(i, j, item)
        table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def update_chart_options(self):
        selected_columns = [item.text() for item in self.column_list.selectedItems()]
        self.chart_type_combo.setEnabled(bool(selected_columns))
        self.generate_chart_btn.setEnabled(bool(selected_columns))
        self.figure.clear()
        self.canvas.draw_idle()
        self.update_chart_specific_options()
    
    def update_chart_specific_options(self):
        chart_type = self.chart_type_combo.currentText()
        if chart_type == "Histogram":
            self.histogram_bins_label.show()
            self.histogram_bins_input.show()
        else:
            self.histogram_bins_label.hide()
            self.histogram_bins_input.hide()
    
    def apply_filter(self):
        filter_col = self.filter_column_combo.currentText()
        filter_val = self.filter_value_combo.currentText()
        if not filter_col or not filter_val:
            QMessageBox.warning(self, "Filter Error", "Please select a column and a value to filter.")
            return
        if self.current_df is None:
            QMessageBox.warning(self, "Filter Error", "No data loaded.")
            return

        try:
            if pd.api.types.is_numeric_dtype(self.current_df[filter_col]):
                self.filtered_df = self.current_df[self.current_df[filter_col] == float(filter_val)].copy()
            else:
                self.filtered_df = self.current_df[self.current_df[filter_col] == filter_val].copy()

            if self.filtered_df.empty:
                QMessageBox.warning(self, "Filter Result", "The filter returned no data.")
                self.filtered_df = self.current_df.copy()
            else:
                QMessageBox.information(self, "Filter Applied", f"Filtered data for '{filter_col}' equal to '{filter_val}'.")
            
            self.display_dataframe_in_table(self.filtered_df.head(5), self.preview_table)
            self.generate_chart()
        except ValueError:
            QMessageBox.critical(self, "Input Error", f"The filter value '{filter_val}' is not valid for the selected numeric column.")
        except Exception as e:
            QMessageBox.critical(self, "Filter Error", f"An error occurred during filtering: {e}")

    def clear_filter(self):
        self.filtered_df = self.current_df.copy() if self.current_df is not None else None
        self.clear_filter_ui()
        if self.filtered_df is not None:
            self.display_dataframe_in_table(self.filtered_df.head(5), self.preview_table)
        self.generate_chart()
        QMessageBox.information(self, "Filter Cleared", "The data filter has been cleared.")

    def clear_filter_ui(self):
        self.filter_column_combo.setCurrentIndex(0)
        self.filter_value_combo.clear()

    def generate_chart(self):
        selected_columns = [item.text() for item in self.column_list.selectedItems()]
        chart_type = self.chart_type_combo.currentText()
        chart_title = self.chart_title_input.text().strip()
        df_to_plot = self.filtered_df if self.filtered_df is not None else self.current_df
        
        if df_to_plot is None:
            QMessageBox.warning(self, "Chart Error", "No data loaded. Please select a file and sheet.")
            return
        if not selected_columns:
            QMessageBox.warning(self, "Chart Error", "Please select at least one column for charting.")
            return
        if not chart_type:
            QMessageBox.warning(self, "Chart Error", "Please select a chart type.")
            return
            
        self.figure.clear()
        plt.style.use('dark_background')
        
        try:
            if chart_type == "Line Chart":
                ax = self.figure.add_subplot(111)
                for col in selected_columns:
                    if not pd.api.types.is_numeric_dtype(df_to_plot[col]):
                        QMessageBox.warning(self, "Chart Error", f"Column '{col}' is not numeric. Line Chart requires numeric data.")
                        self.figure.clear()
                        self.canvas.draw_idle()
                        return
                    ax.plot(df_to_plot.index, df_to_plot[col], label=col)
                ax.set_xlabel("Index", color='#f0f0f0')
                ax.set_ylabel("Value", color='#f0f0f0')
                ax.legend(facecolor='#2b2b4d', edgecolor='#4a4a6e', labelcolor='#f0f0f0')
                ax.set_title(chart_title if chart_title else "Line Chart of Selected Columns", color='#f0f0f0')
                ax.tick_params(axis='x', colors='#e0e0e0')
                ax.tick_params(axis='y', colors='#e0e0e0')
                ax.spines['bottom'].set_color('#4a4a6e')
                ax.spines['top'].set_color('#4a4a6e')
                ax.spines['right'].set_color('#4a4a6e')
                ax.spines['left'].set_color('#4a4a6e')

            elif chart_type == "Bar Chart":
                ax = self.figure.add_subplot(111)
                if len(selected_columns) > 1:
                    df_to_plot[selected_columns].plot(kind='bar', ax=ax, rot=45)
                    ax.set_ylabel("Value", color='#f0f0f0')
                else:
                    col = selected_columns[0]
                    counts = df_to_plot[col].value_counts().sort_index()
                    counts.plot(kind='bar', ax=ax, rot=45)
                    ax.set_ylabel("Count", color='#f0f0f0')
                
                ax.set_title(chart_title if chart_title else "Bar Chart of Selected Columns", color='#f0f0f0')
                ax.set_xlabel("Category" if len(selected_columns) > 1 else selected_columns[0], color='#f0f0f0')
                ax.tick_params(axis='x', colors='#e0e0e0')
                ax.tick_params(axis='y', colors='#e0e0e0')
                ax.legend(facecolor='#2b2b4d', edgecolor='#4a4a6e', labelcolor='#f0f0f0')

            elif chart_type == "Pie Chart":
                if len(selected_columns) != 1:
                    QMessageBox.warning(self, "Chart Error", "Pie chart requires exactly one categorical column.")
                    return
                col = selected_columns[0]
                if pd.api.types.is_numeric_dtype(df_to_plot[col]) and len(df_to_plot[col].unique()) > 20:
                    QMessageBox.warning(self, "Chart Error", "Pie chart is suitable for categorical or discrete numeric data with few unique values.")
                    return
                counts = df_to_plot[col].value_counts()
                if len(counts) > 15:
                    QMessageBox.warning(self, "Chart Error", f"Pie chart may be cluttered with too many unique values ({len(counts)}). Consider a bar chart instead.")
                    return
                ax = self.figure.add_subplot(111)
                ax.pie(counts, labels=counts.index, autopct='%1.1f%%', startangle=90)
                ax.axis('equal')
                ax.set_title(chart_title if chart_title else f"Pie Chart of {col}", color='#f0f0f0')
                
            elif chart_type == "Scatter Plot":
                if len(selected_columns) != 2:
                    QMessageBox.warning(self, "Chart Error", "Scatter plot requires exactly two numeric columns (X and Y).")
                    return
                x_col = selected_columns[0]
                y_col = selected_columns[1]
                if not pd.api.types.is_numeric_dtype(df_to_plot[x_col]) or not pd.api.types.is_numeric_dtype(df_to_plot[y_col]):
                    QMessageBox.warning(self, "Chart Error", "Scatter plot requires two numeric columns.")
                    return
                ax = self.figure.add_subplot(111)
                ax.scatter(df_to_plot[x_col], df_to_plot[y_col], color='#7a9d7d')
                ax.set_xlabel(x_col, color='#f0f0f0')
                ax.set_ylabel(y_col, color='#f0f0f0')
                ax.set_title(chart_title if chart_title else f"Scatter Plot of {x_col} vs {y_col}", color='#f0f0f0')
                ax.tick_params(axis='x', colors='#e0e0e0')
                ax.tick_params(axis='y', colors='#e0e0e0')

            elif chart_type == "Histogram":
                num_cols = len(selected_columns)
                fig, axes = plt.subplots(num_cols, 1, figsize=(8, 4 * num_cols), sharex=False, facecolor='#1a1a2e')
                if num_cols == 1:
                    axes = [axes]
                
                try:
                    num_bins = int(self.histogram_bins_input.text())
                    if num_bins <= 0:
                        QMessageBox.warning(self, "Input Error", "Number of bins must be a positive integer. Using default of 10.")
                        num_bins = 10
                except ValueError:
                    QMessageBox.warning(self, "Input Error", "Invalid value for number of bins. Using default of 10.")
                    num_bins = 10

                for i, col in enumerate(selected_columns):
                    if not pd.api.types.is_numeric_dtype(df_to_plot[col]):
                        QMessageBox.warning(self, "Chart Error", f"Column '{col}' is not numeric. Histogram requires numeric data.")
                        continue
                    axes[i].hist(df_to_plot[col].dropna(), bins=num_bins, color='#c999d9')
                    axes[i].set_xlabel(col, color='#f0f0f0')
                    axes[i].set_ylabel("Frequency", color='#f0f0f0')
                    axes[i].set_title(f"Histogram of {col}", color='#f0f0f0')
                    axes[i].tick_params(axis='x', colors='#e0e0e0')
                    axes[i].tick_params(axis='y', colors='#e0e0e0')
                
                fig.suptitle(chart_title if chart_title else "Histograms of Selected Columns", color='#f0f0f0')
                self.figure = fig
                self.canvas.figure = fig
                fig.tight_layout(rect=[0, 0.03, 1, 0.95])
            
            # Save the chart to a temporary file for reporting
            chart_file_path = f"temp_chart_{uuid.uuid4().hex}.png"
            self.figure.savefig(chart_file_path, facecolor=self.figure.get_facecolor(), bbox_inches='tight')
            self.last_chart_image_path = chart_file_path
            
            self.canvas.draw()
            
            # Auto-generate PDF report and signal to main app
            report_name = f"Dashboard_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            self.report_generated.emit(report_name, "pdf", self.last_chart_image_path)
            
        except Exception as e:
            QMessageBox.critical(self, "Chart Generation Error", f"An error occurred: {e}")
            self.figure.clear()
            self.canvas.draw_idle()

class LLMWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    
    def __init__(self, prompt, dataframes_info):
        super().__init__()
        self.prompt = prompt
        self.dataframes_info = dataframes_info

    def run(self):
        try:
            context = f"Here's information about the loaded Excel dataframes:\n{self.dataframes_info}\n"
            full_prompt = f"{context}\nUser Request: {self.prompt}"
            response = ollama.chat(model='phi3', messages=[{'role': 'user', 'content': full_prompt}])
            generated_text = response['message']['content']
            self.finished.emit(generated_text)
        except Exception as e:
            self.error.emit(f"Error during LLM call: {str(e)}\n"
                            "Please ensure Ollama is running and the 'phi3' model is downloaded.")


class EmailSenderWorker(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(self, subject, body, to_email, file_path, smtp_server, port, from_email, password):
        super().__init__()
        self.subject = subject
        self.body = body
        self.to_email = to_email
        self.file_path = file_path
        self.smtp_server = smtp_server
        self.port = port
        self.from_email = from_email
        self.password = password

    def run(self):
        try:
            msg = MIMEMultipart()
            msg['From'] = self.from_email
            msg['To'] = self.to_email
            msg['Subject'] = self.subject
            msg.attach(MIMEText(self.body, 'plain'))

            if self.file_path and os.path.exists(self.file_path):
                with open(self.file_path, "rb") as attachment:
                    part = MIMEMultipart()
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {os.path.basename(self.file_path)}",
                )
                msg.attach(part)

            server = smtplib.SMTP(self.smtp_server, self.port)
            server.starttls()
            server.login(self.from_email, self.password)
            text = msg.as_string()
            server.sendmail(self.from_email, self.to_email, text)
            server.quit()
            
            self.finished.emit(True, f"Email successfully sent to {self.to_email}.")
        except Exception as e:
            self.finished.emit(False, f"Failed to send email: {e}\n"
                                     "Please check your email configuration (sender email and app password) and network connection.")

class OperationResultDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Operation Result")
        self.setGeometry(200, 200, 800, 600)
        self.setStyleSheet(MODERN_DARK_QSS)

        layout = QVBoxLayout(self)
        
        self.result_table = QTableWidget()
        self.display_dataframe_in_table(df, self.result_table)
        layout.addWidget(self.result_table)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

    def display_dataframe_in_table(self, df, table_widget):
        table_widget.clearContents()
        if df.empty:
            table_widget.setRowCount(0)
            table_widget.setColumnCount(0)
            return

        table_widget.setRowCount(len(df))
        table_widget.setColumnCount(len(df.columns))
        table_widget.setHorizontalHeaderLabels(df.columns.tolist())
        for i in range(len(df)):
            for j, col in enumerate(df.columns):
                item = QTableWidgetItem(str(df.iloc[i, j]))
                table_widget.setItem(i, j, item)
        table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)


class ExcelOperations(QWidget):
    report_generated = pyqtSignal(str, str, str) # file_name, file_type, file_path

    def __init__(self, dataframes, parent=None):
        super().__init__(parent)
        self.dataframes = dataframes
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)

        # Merge Section
        merge_group = QGroupBox("1. Merge DataFrames")
        merge_layout = QGridLayout()
        
        self.merge_df1_combo = QComboBox()
        self.merge_df2_combo = QComboBox()
        self.merge_on_col_input = QLineEdit()
        self.merge_on_col_input.setPlaceholderText("Common column name")
        self.merge_type_combo = QComboBox()
        self.merge_type_combo.addItems(["inner", "left", "right", "outer"])
        self.perform_merge_btn = QPushButton("Perform Merge")
        self.perform_merge_btn.clicked.connect(self.perform_merge)
        
        merge_layout.addWidget(QLabel("DataFrame 1:"), 0, 0)
        merge_layout.addWidget(self.merge_df1_combo, 0, 1)
        merge_layout.addWidget(QLabel("DataFrame 2:"), 1, 0)
        merge_layout.addWidget(self.merge_df2_combo, 1, 1)
        merge_layout.addWidget(QLabel("Merge On:"), 2, 0)
        merge_layout.addWidget(self.merge_on_col_input, 2, 1)
        merge_layout.addWidget(QLabel("Merge Type:"), 3, 0)
        merge_layout.addWidget(self.merge_type_combo, 3, 1)
        merge_layout.addWidget(self.perform_merge_btn, 4, 0, 1, 2)
        
        merge_group.setLayout(merge_layout)
        main_layout.addWidget(merge_group)
        main_layout.addSpacing(20)

        # Filter Section
        filter_group = QGroupBox("2. Filter DataFrame")
        filter_layout = QGridLayout()
        
        self.filter_df_combo = QComboBox()
        self.filter_by_combo = QComboBox()
        self.filter_by_combo.addItems(["Column", "Row Index"])
        self.filter_by_combo.currentIndexChanged.connect(self.update_filter_ui)
        
        self.filter_col_label = QLabel("Column:")
        self.filter_col_combo = QComboBox()
        
        self.filter_op_label = QLabel("Operator:")
        self.filter_op_combo = QComboBox()
        self.filter_op_combo.addItems(["==", "!=", ">", "<", ">=", "<=", "contains"])
        
        self.filter_val_label = QLabel("Value:")
        self.filter_val_input = QLineEdit()
        
        self.apply_filter_btn = QPushButton("Apply Filter")
        self.apply_filter_btn.clicked.connect(self.apply_filter_operation)
        self.filter_df_combo.currentIndexChanged.connect(self.update_filter_ui)

        filter_layout.addWidget(QLabel("DataFrame:"), 0, 0)
        filter_layout.addWidget(self.filter_df_combo, 0, 1)
        filter_layout.addWidget(QLabel("Filter By:"), 1, 0)
        filter_layout.addWidget(self.filter_by_combo, 1, 1)
        
        filter_layout.addWidget(self.filter_col_label, 2, 0)
        filter_layout.addWidget(self.filter_col_combo, 2, 1)
        filter_layout.addWidget(self.filter_op_label, 3, 0)
        filter_layout.addWidget(self.filter_op_combo, 3, 1)
        filter_layout.addWidget(self.filter_val_label, 4, 0)
        filter_layout.addWidget(self.filter_val_input, 4, 1)
        filter_layout.addWidget(self.apply_filter_btn, 5, 0, 1, 2)
        
        filter_group.setLayout(filter_layout)
        main_layout.addWidget(filter_group)
        main_layout.addSpacing(20)

        # Search Section
        search_group = QGroupBox("3. Search & Highlight")
        search_layout = QHBoxLayout()
        self.search_df_combo = QComboBox()
        self.search_val_input = QLineEdit()
        self.search_val_input.setPlaceholderText("Enter value to search")
        self.search_btn = QPushButton("Search")
        self.search_btn.clicked.connect(self.search_dataframe)
        
        search_layout.addWidget(QLabel("DataFrame:"))
        search_layout.addWidget(self.search_df_combo)
        search_layout.addWidget(QLabel("Value:"))
        search_layout.addWidget(self.search_val_input)
        search_layout.addWidget(self.search_btn)
        
        search_group.setLayout(search_layout)
        main_layout.addWidget(search_group)

        main_layout.addStretch(1)
        self.update_dropdowns()
    
    def update_dropdowns(self, new_dataframes=None):
        if new_dataframes:
            self.dataframes = new_dataframes
        
        df_names = list(self.dataframes.keys())
        for combo in [self.merge_df1_combo, self.merge_df2_combo, self.filter_df_combo, self.search_df_combo]:
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(df_names)
            combo.blockSignals(False)

        if df_names:
            self.update_filter_ui()

    def update_filter_ui(self):
        filter_by = self.filter_by_combo.currentText()
        df_name = self.filter_df_combo.currentText()
        
        if filter_by == "Column":
            self.filter_col_label.show()
            self.filter_col_combo.show()
            self.filter_op_label.show()
            self.filter_op_combo.show()
            self.filter_val_label.show()
            self.filter_val_input.show()

            self.filter_col_combo.clear()
            if df_name in self.dataframes:
                self.filter_col_combo.addItems(self.dataframes[df_name].columns.tolist())
        elif filter_by == "Row Index":
            self.filter_col_label.hide()
            self.filter_col_combo.hide()
            self.filter_op_label.hide()
            self.filter_op_combo.hide()
            self.filter_val_label.show()
            self.filter_val_input.show()
            self.filter_val_label.setText("Row Index:")
            self.filter_val_input.setPlaceholderText("Enter row index (e.g., 5)")

    def show_result_dialog(self, df, title):
        if df.empty:
            QMessageBox.warning(self, "Operation Result", "The operation returned no data.")
            return
        
        result_dialog = OperationResultDialog(df, self)
        result_dialog.setWindowTitle(title)
        result_dialog.exec_()
    
    def perform_merge(self):
        df1_name = self.merge_df1_combo.currentText()
        df2_name = self.merge_df2_combo.currentText()
        merge_col = self.merge_on_col_input.text().strip()
        merge_type = self.merge_type_combo.currentText()
        if not df1_name or not df2_name or not merge_col:
            QMessageBox.warning(self, "Input Error", "Please select both DataFrames and enter a merge column.")
            return
        if df1_name == df2_name:
            QMessageBox.warning(self, "Input Error", "Cannot merge a DataFrame with itself.")
            return
        df1 = self.dataframes.get(df1_name)
        df2 = self.dataframes.get(df2_name)
        if merge_col not in df1.columns or merge_col not in df2.columns:
            QMessageBox.critical(self, "Merge Error", f"Merge column '{merge_col}' not found in one or both DataFrames.")
            return
        try:
            merged_df = pd.merge(df1, df2, on=merge_col, how=merge_type, suffixes=('_df1', '_df2'))
            
            file_name = f"Merged_{df1_name}_{df2_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            file_path = f"temp_{uuid.uuid4().hex}.xlsx"
            
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            merged_df.to_excel(writer, sheet_name='Merged_Data', index=False)
            writer.close()

            self.report_generated.emit(file_name, "xlsx", file_path)
            QMessageBox.information(self, "Merge Successful", "DataFrames merged successfully! Result shown in a pop-up window and saved as an Excel file.")
            self.show_result_dialog(merged_df, "Merged DataFrame Result")
        except Exception as e:
            QMessageBox.critical(self, "Merge Error", f"An error occurred during merge: {e}")

    def apply_filter_operation(self):
        df_name = self.filter_df_combo.currentText()
        filter_by = self.filter_by_combo.currentText()
        df = self.dataframes.get(df_name)
        
        if not df_name:
            QMessageBox.warning(self, "Filter Error", "Please select a DataFrame.")
            return

        try:
            if filter_by == "Column":
                col = self.filter_col_combo.currentText()
                op = self.filter_op_combo.currentText()
                val = self.filter_val_input.text().strip()
                
                if not col or not val:
                    QMessageBox.warning(self, "Filter Error", "Please fill all column filter fields.")
                    return
                
                if op == "contains":
                    filtered_df = df[df[col].astype(str).str.contains(val, case=False, na=False)]
                else:
                    val_typed = type(df[col].iloc[0])(val) if pd.api.types.is_numeric_dtype(df[col]) else val
                    if op == "==":
                        filtered_df = df[df[col] == val_typed]
                    elif op == "!=":
                        filtered_df = df[df[col] != val_typed]
                    elif op == ">":
                        filtered_df = df[df[col] > val_typed]
                    elif op == "<":
                        filtered_df = df[df[col] < val_typed]
                    elif op == ">=":
                        filtered_df = df[df[col] >= val_typed]
                    elif op == "<=":
                        filtered_df = df[df[col] <= val_typed]
            
            elif filter_by == "Row Index":
                row_index_str = self.filter_val_input.text().strip()
                if not row_index_str:
                    QMessageBox.warning(self, "Filter Error", "Please enter a row index.")
                    return
                row_index = int(row_index_str)
                if row_index < 0 or row_index >= len(df):
                    QMessageBox.warning(self, "Filter Error", f"Row index out of range. Must be between 0 and {len(df) - 1}.")
                    return
                
                filtered_df = df.iloc[[row_index]]

            file_name = f"{df_name}_Filtered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            file_path = f"temp_{uuid.uuid4().hex}.xlsx"
            
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            filtered_df.to_excel(writer, sheet_name='Filtered_Data', index=False)
            writer.close()

            self.report_generated.emit(file_name, "xlsx", file_path)
            QMessageBox.information(self, "Filter Successful", "DataFrame filtered successfully! Result shown in a pop-up window and saved as an Excel file.")
            self.show_result_dialog(filtered_df, "Filtered DataFrame Result")
            
        except ValueError:
            QMessageBox.critical(self, "Input Error", f"The filter value is not valid for the selected column/index.")
        except Exception as e:
            QMessageBox.critical(self, "Filter Error", f"An error occurred during filtering: {e}")

    def search_dataframe(self):
        df_name = self.search_df_combo.currentText()
        search_val = self.search_val_input.text().strip()
        
        if not df_name or not search_val:
            QMessageBox.warning(self, "Search Error", "Please select a DataFrame and enter a value to search.")
            return

        df = self.dataframes.get(df_name)
        if df is None:
            QMessageBox.warning(self, "Search Error", "Selected DataFrame not found.")
            return
            
        found = False
        highlighted_df = df.copy().astype(str)

        for col in highlighted_df.columns:
            matches = highlighted_df[col].str.contains(search_val, case=False, na=False)
            if matches.any():
                found = True

        if not found:
            QMessageBox.information(self, "Search Result", f"'{search_val}' not found in the DataFrame.")
        else:
            QMessageBox.information(self, "Search Result", f"Found occurrences of '{search_val}'. See the pop-up for the full DataFrame.")
        
        self.show_result_dialog(highlighted_df, f"Search Results for '{search_val}'")

class FilesTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.generated_files = {} # Stores {file_name: file_path, file_type}
        self.email_worker = None
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        
        main_layout.addWidget(QLabel("Generated Files & Reports:"))
        self.files_list = QListWidget()
        self.files_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.files_list.setMinimumHeight(200)
        main_layout.addWidget(self.files_list)
        
        buttons_layout = QHBoxLayout()
        self.download_btn = QPushButton("Download File")
        self.download_btn.clicked.connect(self.download_file)
        self.download_btn.setEnabled(False)
        buttons_layout.addWidget(self.download_btn)
        
        self.email_btn = QPushButton("Email File")
        self.email_btn.clicked.connect(self.email_file_dialog)
        self.email_btn.setEnabled(False)
        buttons_layout.addWidget(self.email_btn)
        
        main_layout.addLayout(buttons_layout)
        main_layout.addStretch(1)

        self.files_list.itemSelectionChanged.connect(self.update_buttons)

    def add_generated_file(self, file_name, file_type, file_path):
        unique_id = uuid.uuid4().hex
        display_name = f"[{file_type.upper()}] {file_name}"
        self.generated_files[unique_id] = {'display_name': display_name, 'file_path': file_path, 'file_type': file_type}
        self.files_list.addItem(display_name)
        self.files_list.item(self.files_list.count()-1).setData(Qt.UserRole, unique_id)
    
    def update_buttons(self):
        self.download_btn.setEnabled(bool(self.files_list.selectedItems()))
        self.email_btn.setEnabled(bool(self.files_list.selectedItems()))

    def get_selected_file_info(self):
        selected_items = self.files_list.selectedItems()
        if not selected_items:
            return None
        unique_id = selected_items[0].data(Qt.UserRole)
        return self.generated_files.get(unique_id)

    def download_file(self):
        file_info = self.get_selected_file_info()
        if not file_info:
            QMessageBox.warning(self, "Download Error", "No file selected.")
            return

        file_path = file_info['file_path']
        file_type = file_info['file_type']
        display_name = file_info['display_name']

        if file_type == 'pdf':
            file_filter = "PDF Files (*.pdf)"
        elif file_type == 'xlsx':
            file_filter = "Excel Files (*.xlsx)"
        else:
            file_filter = "All Files (*.*)"

        initial_name = os.path.splitext(display_name.split('] ')[-1])[0]
        save_path, _ = QFileDialog.getSaveFileName(self, "Save File", initial_name, file_filter)
        
        if save_path:
            try:
                # Add the correct extension if not already present
                if not save_path.lower().endswith(f".{file_type}"):
                    save_path += f".{file_type}"
                
                # Check if the temporary file exists before trying to move it
                if os.path.exists(file_path):
                    os.rename(file_path, save_path)
                    
                    # Update the file path in our dictionary and the list widget
                    selected_item = self.files_list.selectedItems()[0]
                    unique_id = selected_item.data(Qt.UserRole)
                    self.generated_files[unique_id]['file_path'] = save_path
                    self.generated_files[unique_id]['display_name'] = os.path.basename(save_path)
                    selected_item.setText(os.path.basename(save_path))
                    
                    QMessageBox.information(self, "Download Complete", f"File saved to {save_path}")
                else:
                    QMessageBox.critical(self, "Download Error", "Source file does not exist.")
            except Exception as e:
                QMessageBox.critical(self, "Download Error", f"Failed to save file: {e}")

    def email_file_dialog(self):
        file_info = self.get_selected_file_info()
        if not file_info:
            QMessageBox.warning(self, "Email Error", "No file selected.")
            return
        
        file_path = file_info['file_path']
        
        email, ok = QInputDialog.getText(self, "Send Report via Email", "Enter recipient's email address:")
        if ok and email:
            self.email_file(email, file_path)

    def email_file(self, recipient_email, file_path):
        subject = f"Generated File from Intelligent Excel Analyst: {os.path.basename(file_path)}"
        body = "Please find the attached file."
        
        # NOTE: You need to configure your email credentials here.
        # This is a sample for a Gmail account. You need to enable "App Passwords"
        # in your Google Account security settings.
        sender_email = "projectklaites01@gmail.com"
        sender_password = "rvxpgajoesexfiqp"
        
        if sender_email == "your_email@gmail.com" or sender_password == "your_app_password":
            QMessageBox.critical(self, "Email Configuration Error", 
                                 "Please update the 'sender_email' and 'sender_password' variables "
                                 "in the 'email_file' method with your own credentials.")
            return

        self.email_worker = EmailSenderWorker(subject, body, recipient_email, file_path, 'smtp.gmail.com', 587, sender_email, sender_password)
        self.email_worker.finished.connect(self.handle_email_result)
        self.email_worker.start()
        QMessageBox.information(self, "Sending Email", "Email is being sent in the background.")

    def handle_email_result(self, success, message):
        if success:
            QMessageBox.information(self, "Email Status", message)
        else:
            QMessageBox.critical(self, "Email Status", message)


class ExcelAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Intelligent Excel Data Analyst")
        self.setGeometry(100, 100, 1400, 900)
        self.dataframes = {}
        self.llm_prediction_explanation_text_added = False
        self.current_llm_worker = None
        self.model = None
        self.scaler = None
        self.prediction_target_col = None
        self.prediction_feature_cols = None
        self.analysis_results_text = "No analysis performed yet."

        self.chat_output_text = QTextEdit()
        self.chat_input_text = QLineEdit()
        self.data_context_checkbox = QRadioButton("Include DataFrame Info")
        self.send_chat_btn = QPushButton("Send to LLM")
        self.chat_clear_btn = QPushButton("Clear Chat")
        self.data_preview_table = QTableWidget()
        self.loaded_files_list = QListWidget()
        self.loaded_dataframes_info_text = QTextEdit("No Excel files loaded yet.")
        self.preview_full_sheet_btn = QPushButton("Preview Full Sheet")
        self.preview_full_sheet_btn.setEnabled(False)

        self.analysis_df_combo = QComboBox()
        self.column_list_analysis = QListWidget()
        self.show_descriptive_analysis_btn = QPushButton("Descriptive Statistics")
        self.show_missing_values_btn = QPushButton("Missing Values")
        self.show_value_counts_btn = QPushButton("Value Counts")
        self.analysis_output_text = QTextEdit()
        self.generate_analysis_report_btn = QPushButton("Generate PDF Report")

        self.predict_df_combo = QComboBox()
        self.predict_target_combo = QComboBox()
        self.predict_features_list = QListWidget()
        self.predict_model_combo = QComboBox()
        self.predict_model_combo.addItems(["Linear Regression", "Decision Tree Classifier"])
        self.scale_data_checkbox = QRadioButton("Scale Features")
        self.perform_prediction_btn = QPushButton("Train Model & Predict")
        self.prediction_results_label = QLabel("Prediction Results:")
        self.prediction_results_text = QTextEdit()
        self.prediction_input_layout = QVBoxLayout()
        self.predict_new_btn = QPushButton("Predict New Value(s)")
        self.new_value_inputs = {}
        self.llm_prediction_explanation_btn = QPushButton("Ask LLM to Explain Prediction")
        self.llm_prediction_explanation_text = QTextEdit()
        
        self.init_ui()
        self.update_dataframe_dropdowns()

    def init_ui(self):
        self.setStyleSheet(MODERN_DARK_QSS)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        
        left_panel = QVBoxLayout()
        left_panel.addWidget(QLabel("1. Load Excel Files:"))
        
        file_buttons_layout = QHBoxLayout()
        self.add_file_btn = QPushButton("Add Excel File(s)")
        self.add_file_btn.clicked.connect(self.load_excel_file)
        file_buttons_layout.addWidget(self.add_file_btn)
        self.preview_full_sheet_btn.clicked.connect(self.preview_full_sheet)
        file_buttons_layout.addWidget(self.preview_full_sheet_btn)
        left_panel.addLayout(file_buttons_layout)

        left_panel.addWidget(QLabel("Loaded DataFrames:"))
        self.loaded_files_list.setMinimumHeight(150)
        self.loaded_files_list.itemSelectionChanged.connect(self.display_selected_file_preview)
        left_panel.addWidget(self.loaded_files_list)
        self.loaded_dataframes_info_text.setReadOnly(True)
        self.loaded_dataframes_info_text.setMinimumHeight(100)
        left_panel.addWidget(self.loaded_dataframes_info_text)
        left_panel.addWidget(QLabel("Data Preview (First 10 Rows of Selected File):"))
        self.data_preview_table.setMinimumHeight(200)
        left_panel.addWidget(self.data_preview_table)
        
        left_panel.addStretch(1)
        main_layout.addLayout(left_panel, 1)

        right_panel = QVBoxLayout()
        self.tabs = QTabWidget()
        
        self.files_tab = FilesTab()
        
        self.dashboard_tab = DashboardApp(initial_dataframes=self.dataframes)
        self.dashboard_tab.report_generated.connect(self.files_tab.add_generated_file)
        self.tabs.addTab(self.dashboard_tab, "Dashboard")
        
        self.excel_operations_tab = ExcelOperations(self.dataframes, parent=self)
        self.excel_operations_tab.report_generated.connect(self.files_tab.add_generated_file)
        self.tabs.addTab(self.excel_operations_tab, "Excel Operations")
        
        self.tabs.addTab(self.create_analysis_tab(), "Analysis")
        self.tabs.addTab(self.create_prediction_tab(), "Prediction")
        self.tabs.addTab(self.create_llm_chat_tab(), "LLM Chat")
        
        # --- NEW TAB INTEGRATION START ---
        # Initialize a ReportGenerator instance from the imported class
        report_generator_instance = ReportGenerator(self.dataframes, self.analysis_results_text)
        self.automated_bot_tab = AutomatedBotTab(self.dataframes, report_generator_instance)
        # Connect the bot's report signal to the files tab
        self.automated_bot_tab.report_generated.connect(self.files_tab.add_generated_file)
        self.tabs.addTab(self.automated_bot_tab, "Automation Bot")
        # --- NEW TAB INTEGRATION END ---

        self.tabs.addTab(self.files_tab, "Files")

        right_panel.addWidget(self.tabs)
        main_layout.addLayout(right_panel, 3)

    def load_excel_file(self):
        file_dialog = QFileDialog()
        file_paths, _ = file_dialog.getOpenFileNames(
            self, "Open Excel File(s)", "", "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_paths:
            for file_path in file_paths:
                file_name = os.path.basename(file_path)
                try:
                    df = pd.read_excel(file_path)
                    df_name = file_name.replace(".xlsx", "").replace(".xls", "")
                    original_df_name = df_name
                    counter = 1
                    while df_name in self.dataframes:
                        df_name = f"{original_df_name}_{counter}"
                        counter += 1
                    self.dataframes[df_name] = df
                    if df_name not in [self.loaded_files_list.item(i).text() for i in range(self.loaded_files_list.count())]:
                        self.loaded_files_list.addItem(df_name)
                    QMessageBox.information(self, "File Loaded", f"'{file_name}' loaded successfully as '{df_name}'.")
                    self.update_loaded_dataframes_info()
                    self.update_dataframe_dropdowns()
                    self.dashboard_tab.update_dataframes(self.dataframes)
                    self.excel_operations_tab.update_dropdowns(self.dataframes)
                    self.loaded_files_list.setCurrentRow(self.loaded_files_list.count() - 1)
                    self.preview_full_sheet_btn.setEnabled(True)
                except Exception as e:
                    QMessageBox.critical(self, "Error Loading File", f"Could not load file '{file_name}': {e}")
    
    def preview_full_sheet(self):
        selected_items = self.loaded_files_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Preview Error", "Please select a DataFrame to preview.")
            return
        
        df_name = selected_items[0].text()
        if df_name in self.dataframes:
            preview_dialog = FullSheetPreviewDialog(self.dataframes[df_name], self)
            preview_dialog.exec_()
    
    def display_selected_file_preview(self):
        selected_items = self.loaded_files_list.selectedItems()
        if selected_items:
            file_name = selected_items[0].text()
            if file_name in self.dataframes:
                df = self.dataframes[file_name]
                self.display_dataframe_in_table(df.head(10), self.data_preview_table)
                self.preview_full_sheet_btn.setEnabled(True)
            else:
                self.data_preview_table.clearContents()
                self.data_preview_table.setRowCount(0)
                self.data_preview_table.setColumnCount(0)
                self.preview_full_sheet_btn.setEnabled(False)
        else:
            self.data_preview_table.clearContents()
            self.data_preview_table.setRowCount(0)
            self.data_preview_table.setColumnCount(0)
            self.preview_full_sheet_btn.setEnabled(False)

    def update_loaded_dataframes_info(self):
        info_text = ""
        if not self.dataframes:
            info_text = "No Excel files loaded yet."
        else:
            for name, df in self.dataframes.items():
                info_text += f"DataFrame: {name}\n"
                info_text += f"  Shape: {df.shape[0]} rows, {df.shape[1]} columns\n"
                info_text += f"  Columns: {', '.join(df.columns.tolist())}\n\n"
        self.loaded_dataframes_info_text.setText(info_text)

    def update_dataframe_dropdowns(self):
        df_names = list(self.dataframes.keys())
        self.analysis_df_combo.clear()
        self.predict_df_combo.clear()
        
        self.analysis_df_combo.addItems(df_names)
        self.predict_df_combo.addItems(df_names)

        if df_names:
            self.update_analysis_columns()
            self.update_predict_columns()

    def display_dataframe_in_table(self, df, table_widget):
        table_widget.clearContents()
        table_widget.setRowCount(len(df))
        table_widget.setColumnCount(len(df.columns))
        table_widget.setHorizontalHeaderLabels(df.columns.tolist())
        for i in range(len(df)):
            for j, col in enumerate(df.columns):
                item = QTableWidgetItem(str(df.iloc[i, j]))
                table_widget.setItem(i, j, item)
        table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def create_analysis_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        analysis_group = QGroupBox("Data Analysis")
        analysis_layout = QVBoxLayout()
        
        analysis_layout.addWidget(QLabel("<b>1. Select DataFrame for Analysis:</b>"))
        self.analysis_df_combo.currentIndexChanged.connect(self.update_analysis_columns)
        analysis_layout.addWidget(self.analysis_df_combo)

        analysis_layout.addWidget(QLabel("<b>2. Select Columns:</b>"))
        self.column_list_analysis = QListWidget()
        self.column_list_analysis.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        analysis_layout.addWidget(self.column_list_analysis)

        analysis_layout.addWidget(QLabel("<b>3. Choose Analysis Type:</b>"))
        analysis_buttons_layout = QHBoxLayout()
        self.show_descriptive_analysis_btn.clicked.connect(self.show_descriptive_analysis)
        analysis_buttons_layout.addWidget(self.show_descriptive_analysis_btn)
        self.show_missing_values_btn.clicked.connect(self.show_missing_values)
        analysis_buttons_layout.addWidget(self.show_missing_values_btn)
        self.show_value_counts_btn.clicked.connect(self.show_value_counts)
        analysis_buttons_layout.addWidget(self.show_value_counts_btn)
        analysis_layout.addLayout(analysis_buttons_layout)
        
        analysis_group.setLayout(analysis_layout)
        layout.addWidget(analysis_group)
        
        layout.addWidget(QLabel("Analysis Results:"))
        self.analysis_output_text = QTextEdit()
        self.analysis_output_text.setReadOnly(True)
        layout.addWidget(self.analysis_output_text)
        
        self.generate_analysis_report_btn.clicked.connect(self.generate_analysis_report)
        layout.addWidget(self.generate_analysis_report_btn)
        
        layout.addStretch(1)
        return tab

    def update_analysis_columns(self):
        df_name = self.analysis_df_combo.currentText()
        self.column_list_analysis.clear()
        if df_name in self.dataframes:
            columns = self.dataframes[df_name].columns.tolist()
            self.column_list_analysis.addItems(columns)
        self.analysis_output_text.clear()

    def show_descriptive_analysis(self):
        df_name = self.analysis_df_combo.currentText()
        if not df_name:
            QMessageBox.warning(self, "Analysis Error", "Please select a DataFrame.")
            return

        selected_items = self.column_list_analysis.selectedItems()
        selected_columns = [item.text() for item in selected_items]
        
        if not selected_columns:
            QMessageBox.warning(self, "Analysis Error", "Please select one or more columns.")
            return
        
        df = self.dataframes[df_name][selected_columns]
        summary = df.describe(include='all').T
        output = "--- Descriptive Statistics ---\n"
        output += summary.to_string()
        self.analysis_output_text.setText(output)
        self.analysis_results_text = output

    def show_missing_values(self):
        df_name = self.analysis_df_combo.currentText()
        if not df_name:
            QMessageBox.warning(self, "Analysis Error", "Please select a DataFrame.")
            return
        
        df = self.dataframes[df_name]
        missing_data = df.isnull().sum()
        missing_percentage = (df.isnull().sum() / len(df)) * 100
        missing_table = pd.DataFrame({'Total Missing': missing_data, 'Percentage': missing_percentage})
        
        output = "--- Missing Values Analysis ---\n"
        output += missing_table.to_string()
        self.analysis_output_text.setText(output)
        self.analysis_results_text = output

    def show_value_counts(self):
        df_name = self.analysis_df_combo.currentText()
        if not df_name:
            QMessageBox.warning(self, "Analysis Error", "Please select a DataFrame.")
            return
        
        selected_items = self.column_list_analysis.selectedItems()
        selected_columns = [item.text() for item in selected_items]
        if not selected_columns:
            QMessageBox.warning(self, "Analysis Error", "Please select at least one column for value counts.")
            return
        
        output = "--- Value Counts ---\n"
        for col in selected_columns:
            output += f"\nColumn: '{col}'\n"
            value_counts = self.dataframes[df_name][col].value_counts(dropna=False)
            output += value_counts.to_string()
            output += "\n"
        
        self.analysis_output_text.setText(output)
        self.analysis_results_text = output
    
    def generate_analysis_report(self):
        if not self.analysis_output_text.toPlainText():
            QMessageBox.warning(self, "Report Error", "No analysis results to generate a report.")
            return
        
        report_generator = ReportGenerator(self.dataframes, self.analysis_results_text)
        file_name = f"Analysis_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        file_path = f"temp_{uuid.uuid4().hex}.pdf"
        
        if report_generator.generate_pdf(file_path):
            self.files_tab.add_generated_file(file_name, "pdf", file_path)
            QMessageBox.information(self, "Report Generated", "Analysis report saved to the 'Files' tab.")
        else:
            QMessageBox.critical(self, "Report Error", "Failed to generate PDF report.")

    def create_prediction_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        predict_group = QGroupBox("Prediction Model")
        predict_layout = QVBoxLayout()
        
        predict_df_layout = QHBoxLayout()
        predict_df_layout.addWidget(QLabel("Select DataFrame:"))
        self.predict_df_combo.currentIndexChanged.connect(self.update_predict_columns)
        predict_layout.addLayout(predict_df_layout)
        predict_layout.addWidget(self.predict_df_combo)
        
        predict_model_layout = QHBoxLayout()
        predict_model_layout.addWidget(QLabel("Select Model:"))
        self.predict_model_combo.currentIndexChanged.connect(self.update_predict_columns)
        predict_layout.addLayout(predict_model_layout)
        predict_layout.addWidget(self.predict_model_combo)
        
        predict_target_layout = QHBoxLayout()
        predict_target_layout.addWidget(QLabel("Select Target Column:"))
        predict_layout.addLayout(predict_target_layout)
        predict_layout.addWidget(self.predict_target_combo)
        
        predict_features_layout = QHBoxLayout()
        predict_features_layout.addWidget(QLabel("Select Feature Columns (Multi-select):"))
        self.predict_features_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        predict_layout.addLayout(predict_features_layout)
        predict_layout.addWidget(self.predict_features_list)
        
        predict_layout.addWidget(self.scale_data_checkbox)
        
        self.perform_prediction_btn.clicked.connect(self.perform_prediction)
        predict_layout.addWidget(self.perform_prediction_btn)
        predict_group.setLayout(predict_layout)
        layout.addWidget(predict_group)
        layout.addStretch(1)

        self.prediction_results_label.setFont(QFont('Segoe UI', 12, QFont.Bold))
        self.prediction_results_label.setStyleSheet("color: #a0d9ff;")
        self.prediction_results_label.hide()
        layout.addWidget(self.prediction_results_label)

        self.prediction_results_text.setReadOnly(True)
        self.prediction_results_text.hide()
        layout.addWidget(self.prediction_results_text)
        
        layout.addLayout(self.prediction_input_layout)
        
        self.predict_new_btn.clicked.connect(self.predict_new_value)
        self.predict_new_btn.hide()
        layout.addWidget(self.predict_new_btn)
        
        self.llm_prediction_explanation_btn.clicked.connect(self.ask_llm_for_explanation)
        self.llm_prediction_explanation_btn.hide()
        layout.addWidget(self.llm_prediction_explanation_btn)

        self.llm_prediction_explanation_text.setReadOnly(True)
        self.llm_prediction_explanation_text.hide()
        layout.addWidget(self.llm_prediction_explanation_text)

        return tab

    def update_predict_columns(self):
        df_name = self.predict_df_combo.currentText()
        model_type = self.predict_model_combo.currentText()
        self.predict_target_combo.clear()
        self.predict_features_list.clear()

        if df_name in self.dataframes:
            df = self.dataframes[df_name]
            if model_type == "Linear Regression":
                columns = df.select_dtypes(include=np.number).columns.tolist()
            else:
                columns = df.columns.tolist()
            
            self.predict_target_combo.addItems(columns)
            self.predict_features_list.addItems(columns)
        self.clear_prediction_ui()

    def clear_prediction_ui(self):
        self.prediction_results_label.hide()
        self.prediction_results_text.hide()
        self.predict_new_btn.hide()
        self.llm_prediction_explanation_btn.hide()
        self.llm_prediction_explanation_text.hide()
        
        for i in reversed(range(self.prediction_input_layout.count())):
            item = self.prediction_input_layout.itemAt(i)
            if item.widget():
                item.widget().setParent(None)
            elif item.layout():
                while item.layout().count() > 0:
                    child = item.layout().takeAt(0)
                    if child.widget():
                        child.widget().setParent(None)
        self.new_value_inputs.clear()

    def perform_prediction(self):
        df_name = self.predict_df_combo.currentText()
        target_col = self.predict_target_combo.currentText()
        selected_items = self.predict_features_list.selectedItems()
        feature_cols = [item.text() for item in selected_items]
        model_type = self.predict_model_combo.currentText()
        scale_data = self.scale_data_checkbox.isChecked()

        if not df_name or not target_col or not feature_cols:
            QMessageBox.warning(self, "Prediction Error", "Please select DataFrame, Target, and at least one Feature column.")
            return

        df = self.dataframes[df_name].dropna(subset=feature_cols + [target_col]).copy()
        if df.empty:
            QMessageBox.warning(self, "Prediction Error", "Selected columns contain no valid data. Cannot train model.")
            return

        try:
            X = df[feature_cols]
            y = df[target_col]

            if scale_data:
                self.scaler = StandardScaler()
                X = self.scaler.fit_transform(X)
            else:
                self.scaler = None
            
            self.prediction_target_col = target_col
            self.prediction_feature_cols = feature_cols

            if model_type == "Linear Regression":
                self.model = LinearRegression()
                self.model.fit(X, y)
                y_pred = self.model.predict(X)
                mse = mean_squared_error(y, y_pred)
                mae = mean_absolute_error(y, y_pred)
                r2 = r2_score(y, y_pred)
                
                output = f"--- Model: Linear Regression ---\n"
                output += f"R-squared (R²): {r2:.4f}\n"
                output += f"Mean Squared Error (MSE): {mse:.2f}\n"
                output += f"Mean Absolute Error (MAE): {mae:.2f}\n"
                output += f"Coefficients:\n"
                for i, col in enumerate(feature_cols):
                    output += f"  - {col}: {self.model.coef_[i]:.4f}\n"
                output += f"Intercept: {self.model.intercept_:.4f}\n"
            
            elif model_type == "Decision Tree Classifier":
                if y.dtype.kind not in 'i':
                    QMessageBox.warning(self, "Prediction Error", "Decision Tree Classifier requires the target column to be categorical or integer.")
                    return
                self.model = DecisionTreeClassifier()
                self.model.fit(X, y)
                y_pred = self.model.predict(X)
                accuracy = accuracy_score(y, y_pred)
                
                output = f"--- Model: Decision Tree Classifier ---\n"
                output += f"Accuracy: {accuracy:.4f}\n"
                output += f"Classification Report:\n{classification_report(y, y_pred)}\n"
            
            self.prediction_results_text.setText(output)
            self.prediction_results_label.show()
            self.prediction_results_text.show()
            self.predict_new_btn.show()
            self.llm_prediction_explanation_btn.show()
            self.llm_prediction_explanation_text.hide()
            
            self.clear_prediction_input_fields()
            
            for col in feature_cols:
                input_layout = QHBoxLayout()
                label = QLabel(f"{col}:")
                line_edit = QLineEdit()
                line_edit.setPlaceholderText(f"Enter value for {col}")
                line_edit.setValidator(QDoubleValidator())
                self.new_value_inputs[col] = line_edit
                input_layout.addWidget(label)
                input_layout.addWidget(line_edit)
                self.prediction_input_layout.addLayout(input_layout)

        except Exception as e:
            QMessageBox.critical(self, "Prediction Error", f"An error occurred during training: {e}")
            self.clear_prediction_ui()

    def predict_new_value(self):
        new_values = []
        for col in self.prediction_feature_cols:
            if col in self.new_value_inputs:
                try:
                    new_values.append(float(self.new_value_inputs[col].text()))
                except ValueError:
                    QMessageBox.warning(self, "Input Error", f"Invalid value for '{col}'. Please enter a number.")
                    return
        
        if len(new_values) != len(self.prediction_feature_cols):
            QMessageBox.warning(self, "Input Error", "Please enter a value for all feature columns.")
            return

        new_data = np.array(new_values).reshape(1, -1)
        if self.scaler:
            new_data = self.scaler.transform(new_data)

        predicted_value = self.model.predict(new_data)[0]
        
        output = f"--- New Prediction ---\n"
        output += f"For values: {dict(zip(self.prediction_feature_cols, new_values))}\n"
        output += f"The predicted value is: {predicted_value:.4f}"
        
        QMessageBox.information(self, "Prediction Result", output)

    def clear_prediction_input_fields(self):
        for widget in self.new_value_inputs.values():
            widget.setText("")

    def ask_llm_for_explanation(self):
        prompt = f"Explain the following prediction model results:\n\n{self.prediction_results_text.toPlainText()}\n\n"
        prompt += f"The target variable is '{self.prediction_target_col}' and the feature variables are {self.prediction_feature_cols}. "
        prompt += f"The model used is {self.predict_model_combo.currentText()}. "
        prompt += "Explain what the metrics mean and how the model works in simple terms."
        
        self.llm_prediction_explanation_text.setHtml("<i>Requesting explanation from LLM... Please wait.</i>")
        self.llm_prediction_explanation_text.show()
        
        dataframes_info = self.loaded_dataframes_info_text.toPlainText() if self.data_context_checkbox.isChecked() else ""
        
        self.current_llm_worker = LLMWorker(prompt, dataframes_info)
        self.current_llm_worker.finished.connect(self.handle_llm_response)
        self.current_llm_worker.error.connect(self.handle_llm_error)
        self.current_llm_worker.start()

    def handle_llm_response(self, response):
        self.llm_prediction_explanation_text.setText(response)
        self.llm_prediction_explanation_text.show()

    def handle_llm_error(self, error_message):
        QMessageBox.critical(self, "LLM Error", error_message)
        self.llm_prediction_explanation_text.setText(error_message)

    def create_llm_chat_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        chat_group = QGroupBox("LLM Chat")
        chat_layout = QVBoxLayout()
        self.chat_output_text.setReadOnly(True)
        self.chat_output_text.setStyleSheet("border: 1px solid #4a4a6e; border-radius: 6px; padding: 10px;")
        chat_layout.addWidget(self.chat_output_text, 1)

        input_layout = QHBoxLayout()
        self.chat_input_text.setPlaceholderText("Type your question for the LLM here...")
        input_layout.addWidget(self.chat_input_text, 1)
        self.send_chat_btn.clicked.connect(self.send_chat_to_llm)
        input_layout.addWidget(self.send_chat_btn)
        chat_layout.addLayout(input_layout)
        
        options_layout = QHBoxLayout()
        self.data_context_checkbox.setChecked(True)
        options_layout.addWidget(self.data_context_checkbox)
        self.chat_clear_btn.clicked.connect(self.chat_output_text.clear)
        options_layout.addWidget(self.chat_clear_btn)
        chat_layout.addLayout(options_layout)
        
        chat_group.setLayout(chat_layout)
        layout.addWidget(chat_group)
        
        return tab

    def send_chat_to_llm(self):
        user_prompt = self.chat_input_text.text().strip()
        if not user_prompt:
            return

        self.chat_output_text.append(f"<b style='color:#a0d9ff;'>You:</b> {user_prompt}\n")
        self.chat_input_text.clear()

        dataframes_info = self.loaded_dataframes_info_text.toPlainText() if self.data_context_checkbox.isChecked() else ""
        
        self.current_llm_worker = LLMWorker(user_prompt, dataframes_info)
        self.current_llm_worker.finished.connect(self.handle_llm_chat_response)
        self.current_llm_worker.error.connect(self.handle_llm_chat_error)
        self.current_llm_worker.start()

    def handle_llm_chat_response(self, response):
        self.chat_output_text.append(f"<b style='color:#a0d9ff;'>LLM:</b> {response}\n")

    def handle_llm_chat_error(self, error_message):
        self.chat_output_text.append(f"<b style='color:red;'>Error:</b> {error_message}\n")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    app.setStyle(QStyleFactory.create("Fusion"))

    main_app = ExcelAnalyzerApp()
    main_app.show()
    sys.exit(app.exec_())