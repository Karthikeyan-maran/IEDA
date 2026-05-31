# analysis_tab.py

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QComboBox, QRadioButton, QPushButton, QTextEdit, QButtonGroup

def create_analysis_tab(self):
    self.analysis_tab = QWidget()
    layout = QVBoxLayout(self.analysis_tab)

    layout.addWidget(QLabel("📈 Analyze Excel Data", self))

    self.analysis_df_combo = QComboBox()
    self.analysis_df_combo.currentIndexChanged.connect(self.update_analysis_columns)
    layout.addWidget(self.analysis_df_combo)

    self.column_combo = QComboBox()
    layout.addWidget(self.column_combo)

    self.stats_radio = QRadioButton("Descriptive Statistics")
    self.value_counts_radio = QRadioButton("Value Counts")
    self.missing_radio = QRadioButton("Missing Values")

    self.analysis_radio_group = QButtonGroup()
    self.analysis_radio_group.addButton(self.stats_radio)
    self.analysis_radio_group.addButton(self.value_counts_radio)
    self.analysis_radio_group.addButton(self.missing_radio)

    layout.addWidget(self.stats_radio)
    layout.addWidget(self.value_counts_radio)
    layout.addWidget(self.missing_radio)

    self.analyze_button = QPushButton("Analyze")
    self.analyze_button.clicked.connect(self.perform_analysis)
    layout.addWidget(self.analyze_button)

    self.analysis_output = QTextEdit()
    layout.addWidget(self.analysis_output)

    return self.analysis_tab
