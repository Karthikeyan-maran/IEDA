# dashboard_tab.py

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

class DashboardApp(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.dataframes = None
        self.current_dataframe_name = None

        self.layout = QVBoxLayout(self)
        self.label = QLabel("📊 Excel Dashboard", self)
        self.label.setStyleSheet("font-size: 20px; font-weight: bold;")

        self.df_selector = QComboBox(self)
        self.df_selector.currentTextChanged.connect(self.update_columns)

        self.x_selector = QComboBox(self)
        self.y_selector = QComboBox(self)
        self.chart_type_selector = QComboBox(self)
        self.chart_type_selector.addItems(["Bar", "Line", "Scatter"])

        self.generate_button = QPushButton("Generate Chart", self)
        self.generate_button.clicked.connect(self.generate_chart)

        self.layout.addWidget(self.label)
        self.layout.addWidget(self.df_selector)
        self.layout.addWidget(QLabel("X Axis"))
        self.layout.addWidget(self.x_selector)
        self.layout.addWidget(QLabel("Y Axis"))
        self.layout.addWidget(self.y_selector)
        self.layout.addWidget(QLabel("Chart Type"))
        self.layout.addWidget(self.chart_type_selector)
        self.layout.addWidget(self.generate_button)

        self.canvas = FigureCanvas(Figure(figsize=(5, 3)))
        self.layout.addWidget(self.canvas)
        self.ax = self.canvas.figure.subplots()

    def set_dataframes(self, dataframes):
        self.dataframes = dataframes
        self.df_selector.clear()
        self.df_selector.addItems(dataframes.keys())

    def update_columns(self, df_name):
        self.current_dataframe_name = df_name
        if df_name and df_name in self.dataframes:
            df = self.dataframes[df_name]
            self.x_selector.clear()
            self.y_selector.clear()
            self.x_selector.addItems(df.columns.astype(str).tolist())
            self.y_selector.addItems(df.columns.astype(str).tolist())

    def generate_chart(self):
        df_name = self.current_dataframe_name
        if not df_name or df_name not in self.dataframes:
            return

        df = self.dataframes[df_name]
        x_col = self.x_selector.currentText()
        y_col = self.y_selector.currentText()
        chart_type = self.chart_type_selector.currentText()

        if x_col not in df.columns or y_col not in df.columns:
            return

        self.ax.clear()
        if chart_type == "Bar":
            self.ax.bar(df[x_col], df[y_col])
        elif chart_type == "Line":
            self.ax.plot(df[x_col], df[y_col])
        elif chart_type == "Scatter":
            self.ax.scatter(df[x_col], df[y_col])
        self.ax.set_xlabel(x_col)
        self.ax.set_ylabel(y_col)
        self.ax.set_title(f"{chart_type} Chart of {y_col} vs {x_col}")
        self.canvas.draw()
