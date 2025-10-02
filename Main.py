import sys
import os
import re
import pandas as pd
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QFileDialog, QMessageBox, QProgressDialog, QCompleter
)
from PyQt6.QtCore import Qt, QTimer, QStringListModel
from PyQt6.QtGui import QColor, QFont


class RawMaterialApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Raw Material Inventory Dashboard")
        self.setGeometry(100, 100, 1200, 800)

        self._current_df = None
        self._pivot_df = None
        self.month_years_for_headers = []
        self._search_text = ""

        self.init_ui()
        self.apply_styles()

    # ---------------- UI SETUP ----------------
    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # --- Header ---
        header_label = QLabel("Raw Material Consumption")
        header_label.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(header_label)

        # --- Controls ---
        input_action_layout = QHBoxLayout()
        input_action_layout.setSpacing(10)
        input_action_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        # Month input
        input_action_layout.addWidget(QLabel("<b>Start Month:</b>"))
        self.month_combo = QComboBox()
        self.month_combo.setFont(QFont("Segoe UI", 10))
        self.month_combo.addItems([
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ])
        self.month_combo.setCurrentIndex(0)
        self.month_combo.setMinimumWidth(130)
        self.month_combo.currentIndexChanged.connect(self.generate_table)
        input_action_layout.addWidget(self.month_combo)

        # Year input
        input_action_layout.addWidget(QLabel("<b>Start Year:</b>"))
        self.year_edit = QLineEdit()
        self.year_edit.setFont(QFont("Segoe UI", 10))
        self.year_edit.setPlaceholderText("ex. 2023")
        self.year_edit.setText(str(datetime.now().year - 1))
        self.year_edit.setFixedWidth(70)
        self.year_edit.returnPressed.connect(self.generate_table)
        input_action_layout.addWidget(self.year_edit)

        # year completer
        current_year = datetime.now().year
        years = [str(year) for year in range(2000, current_year + 1)]
        completer = QCompleter()
        model = QStringListModel(years)
        completer.setModel(model)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.year_edit.setCompleter(completer)

        # Filter dropdown
        input_action_layout.addWidget(QLabel("<b>Display Filter:</b>"))
        self.filter_combo = QComboBox()
        self.filter_combo.setFont(QFont("Segoe UI", 10))
        self.filter_combo.addItems(["All Data", "Set 1", "Set 2", "Set 3"])
        self.filter_combo.setCurrentIndex(0)
        self.filter_combo.setMinimumWidth(140)
        self.filter_combo.currentIndexChanged.connect(self.apply_filter)
        input_action_layout.addWidget(self.filter_combo)
        input_action_layout.addStretch()
        # Search input
        input_action_layout.addWidget(QLabel("<b>Search Code:</b>"))
        self.search_edit = QLineEdit()
        self.search_edit.setFont(QFont("Segoe UI", 10))
        self.search_edit.setPlaceholderText("Enter material code...")
        self.search_edit.setFixedWidth(180)
        self.setup_finished_typing(self.search_edit, self.search_table)
        self.search_edit.returnPressed.connect(self.search_table)
        input_action_layout.addWidget(self.search_edit)

        # Buttons
        self.load_excel_button = QPushButton("🗂️ Load Excel")
        self.load_excel_button.clicked.connect(self.load_data_from_excel)
        self.load_excel_button.setCursor(Qt.CursorShape.PointingHandCursor)
        input_action_layout.addWidget(self.load_excel_button)

        self.export_excel_button = QPushButton("📊 Export Excel")
        self.export_excel_button.clicked.connect(self.export_to_excel)
        self.export_excel_button.setCursor(Qt.CursorShape.PointingHandCursor)
        input_action_layout.addWidget(self.export_excel_button)

        main_layout.addLayout(input_action_layout)

        # --- Table ---
        self.table_widget = QTableWidget()
        self.table_widget.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table_widget.setFont(QFont("Segoe UI", 9))
        self.table_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        main_layout.addWidget(self.table_widget)

        self.setLayout(main_layout)
        self.generate_table()

    def apply_styles(self):

        self.setStyleSheet("""
            QPushButton {
                background-color: #5C7CFA; /* Primary blue */
                color: white;
                border: none;
                border-radius: 8px; /* More rounded corners */
                padding: 10px 20px; /* More padding */
                font-size: 10pt;
                font-weight: bold;
                transition: background-color 0.3s ease; /* Smooth transition */
            }
            QPushButton:hover {
                background-color: #6C7BEC; /* Lighter blue on hover */
                box-shadow: 2px 2px 8px rgba(0, 0, 0, 0.2); /* Subtle shadow on hover */
            }
            QPushButton:pressed {
                background-color: #4C6CDA; /* Darker blue on press */
            }
            QTableWidget {
                content: "code";
                border: 1px solid #D1D9E0;
                gridline-color: #E0E0E0;
                background-color: #FFFFFF;
                selection-background-color: #6C7BEC;
                selection-color: #FFFFFF;
            }
            QTableCornerButton::section {
                background-color: #DDE4ED;
                font-weight: bold;
                font-size: 10pt;
                padding: 6px;
            }
            QHeaderView::section {
                background-color: #DDE4ED;
                font-weight: bold;
                font-size: 10pt;
                padding: 6px;
            }
            QLineEdit, QComboBox {
                border: 1px solid #D1D9E0;
                border-radius: 5px;
                padding: 5px;
            }
            QComboBox::drop-down {
                border: none;
                background-color: #E0E0E0;
                width: 20px;
                subcontrol-origin: padding;
                subcontrol-position: top right;
                border-top-right-radius: 5px;
                border-bottom-right-radius: 5px;
            }
            QMessageBox {
                background-color: #F8F9FA;
                color: #343A40;
            }
            
        """)

    # ---------------- TABLE GENERATION ----------------
    def generate_table(self):
        try:
            start_month_index = self.month_combo.currentIndex() + 1
            start_year = int(self.year_edit.text())
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid year.")
            return

        current_date = datetime.now()
        current_month = current_date.month
        current_year = current_date.year

        headers = ["Raw Material Code"]
        self.month_years_for_headers = []
        start_date = pd.to_datetime(f"{start_month_index:02d}/01/{start_year}")
        end_date = pd.to_datetime(f"{current_month:02d}/01/{current_year}")

        current_date = start_date
        while current_date <= end_date:
            month_name = current_date.strftime("%b %Y")
            headers.append(month_name)
            self.month_years_for_headers.append(month_name)
            current_date = (current_date + pd.offsets.MonthBegin(1)).replace(day=1)

        self.table_widget.setColumnCount(len(headers))
        self.table_widget.setHorizontalHeaderLabels(headers)

        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table_widget.setColumnWidth(0, 150)
        for col in range(1, len(headers)):
            self.table_widget.setColumnWidth(col, 80)

        self.table_widget.setRowCount(0)
        if self._current_df is not None:
            self.populate_table(self._current_df)

    # ---------------- LOAD EXCEL ----------------
    def load_data_from_excel(self):
        documents_path = os.path.expanduser("~/Documents")
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", documents_path,
                                                   "Excel Files (*.xlsx *.xls)")
        if not file_name:
            return

        try:
            progress = QProgressDialog("Loading Excel file...", None, 0, 0, self)
            progress.setWindowModality(Qt.WindowModality.WindowModal)
            progress.setStyleSheet("QProgressBar { margin-left: 36px; }")

            progress.show()
            QApplication.processEvents()

            df = pd.read_excel(file_name, sheet_name=0)
            self._current_df = df
            self._search_text = ""  # Reset search text
            self.search_edit.clear()  # Clear search input
            self.table_widget.clearContents()

            self.generate_table()

            progress.close()
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Error", f"Error loading Excel file: {e}")

    # ---------------- SEARCH TABLE ----------------
    def search_table(self):
        self._search_text = self.search_edit.text().strip().upper()
        if self._current_df is not None:
            self.populate_table(self._current_df)

    # ---------------- FILTER TABLE ----------------
    def apply_filter(self):
        if self._current_df is not None:
            self.populate_table(self._current_df)

    # ---------------- POPULATE TABLE ----------------
    def populate_table(self, df):
        if not self.month_years_for_headers:
            QMessageBox.warning(self, "Error", "Headers not generated.")
            return

        df.columns = df.columns.str.strip().str.lower()
        required_cols = {"prod_date", "raw material", "qty used"}
        if not required_cols.issubset(df.columns):
            QMessageBox.warning(self, "Error", "Excel must contain: prod_date, raw material, qty used")
            return
        df = df[list(required_cols)]

        df["prod_date"] = pd.to_datetime(df["prod_date"], errors="coerce")
        df = df.dropna(subset=["prod_date"])

        start_date = pd.to_datetime(f"{self.month_combo.currentText()} {self.year_edit.text()}")
        end_date = pd.to_datetime(datetime.now().strftime("%b %Y"))
        df = df[(df["prod_date"] >= start_date) & (df["prod_date"] <= end_date)]

        df["qty used"] = pd.to_numeric(
            df["qty used"].astype(str).str.replace(",", "", regex=False).str.strip(),
            errors="coerce"
        ).fillna(0)

        df["raw material"] = df["raw material"].astype(str).str.strip().str.upper()
        df = df[df["raw material"].str.len() > 0]

        if self._search_text:
            df = df[df["raw material"].str.contains(self._search_text, case=False, na=False)]

        df["month_year"] = df["prod_date"].dt.strftime("%b %Y")

        grouped = df.groupby(["raw material", "month_year"])["qty used"].sum().reset_index()
        pivot_df = grouped.pivot_table(
            index="raw material", columns="month_year",
            values="qty used", aggfunc="sum", fill_value=0
        )

        def natural_keys(text):
            return tuple(int(s) if s.isdigit() else s.lower() for s in re.split(r'(\d+)', str(text)))

        for m in self.month_years_for_headers:
            if m not in pivot_df.columns:
                pivot_df[m] = 0

        pivot_df = pivot_df.reset_index()[["raw material"] + self.month_years_for_headers]

        filter_choice = self.filter_combo.currentText()
        pivot_df["len_category"] = pd.cut(
            pivot_df["raw material"].str.len(),
            bins=[-1, 5, 10, float("inf")],
            labels=[1, 2, 3]
        )
        if filter_choice == "Set 1":
            pivot_df = pivot_df[pivot_df["len_category"] == 1]
        elif filter_choice == "Set 2":
            pivot_df = pivot_df[pivot_df["len_category"] == 2]
        elif filter_choice == "Set 3":
            pivot_df = pivot_df[pivot_df["len_category"] == 3]

        pivot_df = pivot_df.sort_values(
            by=["len_category", "raw material"],
            key=lambda col: col.map(natural_keys) if col.name == "raw material" else col
        )
        pivot_df = pivot_df.drop(columns=["len_category"])

        self.table_widget.setRowCount(len(pivot_df))
        self.table_widget.setUpdatesEnabled(False)
        try:
            for row_idx, row in enumerate(pivot_df.itertuples(index=False)):
                item_code = QTableWidgetItem(str(row[0]))
                item_code.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                font = QFont("Segoe UI", 9, QFont.Weight.Bold)
                item_code.setFont(font)
                self.table_widget.setItem(row_idx, 0, item_code)

                for col_offset, qty in enumerate(row[1:], 1):
                    item = QTableWidgetItem(f"{qty:.2f}")
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    if qty > 0:
                        item.setData(Qt.ItemDataRole.BackgroundRole, QColor("#C9FFFA"))
                        bold_font = QFont("Segoe UI", 9, QFont.Weight.Bold)
                        item.setFont(bold_font)
                        item.setForeground(QColor("#003366"))
                    self.table_widget.setItem(row_idx, col_offset, item)
        finally:
            self.table_widget.setUpdatesEnabled(True)
        self.table_widget.verticalScrollBar().setValue(0)
        self._pivot_df = pivot_df

    # ---------------- EXPORT ----------------
    def export_to_excel(self):
        if self._pivot_df is None or self._pivot_df.empty:
            QMessageBox.warning(self, "No Data", "Load data before exporting.")
            return

        file_name, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
        if not file_name:
            return

        try:
            if not file_name.endswith(".xlsx"):
                file_name += ".xlsx"
            self._pivot_df.to_excel(file_name, index=False)
            QMessageBox.information(self, "Success", f"Exported to {file_name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error exporting: {e}")

    # ---------------- DEBOUNCING METHOD ----------------
    def setup_finished_typing(self, line_edit, callback, delay=800):
        timer = QTimer()
        timer.setSingleShot(True)
        timer.timeout.connect(callback)
        line_edit.textChanged.connect(lambda: timer.start(delay))
        return timer


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RawMaterialApp()
    window.show()
    sys.exit(app.exec())
