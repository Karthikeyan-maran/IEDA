# prediction_tab.py

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QComboBox, QPushButton, QHBoxLayout, QLineEdit, QTextEdit, QScrollArea, QFormLayout, QGroupBox

def create_prediction_tab(self):
    self.prediction_tab = QWidget()
    layout = QVBoxLayout(self.prediction_tab)

    layout.addWidget(QLabel("🤖 Train & Predict with ML", self))

    self.predict_df_combo = QComboBox()
    self.predict_df_combo.currentIndexChanged.connect(self.update_predict_columns)
    layout.addWidget(self.predict_df_combo)

    self.target_combo = QComboBox()
    layout.addWidget(QLabel("Target Column:"))
    layout.addWidget(self.target_combo)

    self.feature_combo = QComboBox()
    self.feature_combo.setEditable(True)
    layout.addWidget(QLabel("Feature Columns (comma-separated):"))
    layout.addWidget(self.feature_combo)

    self.train_button = QPushButton("Train Model")
    self.train_button.clicked.connect(self.train_prediction_model)
    layout.addWidget(self.train_button)

    self.prediction_input_group = QGroupBox("Enter Values for Prediction")
    self.prediction_input_layout = QFormLayout()
    self.prediction_input_group.setLayout(self.prediction_input_layout)

    self.prediction_scroll = QScrollArea()
    self.prediction_scroll.setWidget(self.prediction_input_group)
    self.prediction_scroll.setWidgetResizable(True)
    layout.addWidget(self.prediction_scroll)

    self.predict_button = QPushButton("Predict New Value")
    self.predict_button.clicked.connect(self.predict_new_value)
    layout.addWidget(self.predict_button)

    self.prediction_result = QTextEdit()
    layout.addWidget(self.prediction_result)

    return self.prediction_tab
