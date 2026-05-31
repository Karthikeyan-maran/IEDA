# llm_chat_tab.py

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QTextEdit, QLineEdit, QCheckBox, QPushButton, QLabel

def create_llm_chat_tab(self):
    self.llm_tab = QWidget()
    layout = QVBoxLayout(self.llm_tab)

    layout.addWidget(QLabel("💬 Ask LLM (Ollama)"))

    self.chat_history = QTextEdit()
    self.chat_history.setReadOnly(True)
    layout.addWidget(self.chat_history)

    self.chat_input = QLineEdit()
    layout.addWidget(self.chat_input)

    self.include_data_checkbox = QCheckBox("Include Excel Data Context")
    layout.addWidget(self.include_data_checkbox)

    self.send_chat_button = QPushButton("Send to LLM")
    self.send_chat_button.clicked.connect(self.send_chat_message)
    layout.addWidget(self.send_chat_button)

    return self.llm_tab
