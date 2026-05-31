# merge_tab.py

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog
import pandas as pd

def create_merge_tab(self):
    self.merge_tab = QWidget()
    layout = QVBoxLayout(self.merge_tab)

    layout.addWidget(QLabel("🔗 Merge Excel DataFrames", self))

    self.df1_combo = QComboBox()
    self.df2_combo = QComboBox()
    self.merge_column_input = QLineEdit()
    self.merge_type_combo = QComboBox()
    self.merge_type_combo.addItems(["inner", "left", "right", "outer"])

    self.merge_button = QPushButton("Perform Merge")
    self.merge_button.clicked.connect(self.perform_merge)

    self.download_merged_button = QPushButton("Download Merged File")
    self.download_merged_button.clicked.connect(self.download_merged_excel)
    self.download_merged_button.setEnabled(False)

    input_layout = QHBoxLayout()
    input_layout.addWidget(QLabel("DataFrame 1:"))
    input_layout.addWidget(self.df1_combo)
    input_layout.addWidget(QLabel("DataFrame 2:"))
    input_layout.addWidget(self.df2_combo)
    input_layout.addWidget(QLabel("Merge Column:"))
    input_layout.addWidget(self.merge_column_input)
    input_layout.addWidget(QLabel("Merge Type:"))
    input_layout.addWidget(self.merge_type_combo)
    input_layout.addWidget(self.merge_button)
    input_layout.addWidget(self.download_merged_button)

    layout.addLayout(input_layout)

    self.merge_table = QTableWidget()
    layout.addWidget(self.merge_table)

    return self.merge_tab
