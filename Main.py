import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QMessageBox, QProgressDialog
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPalette, QColor, QFont
from datetime import datetime

class RawMaterialApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Raw Material Inventory Dashboard")
        self.setGeometry(100, 100, 1200, 800)
        self.init_ui()
        self.apply_styles()

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Header Section (Title)
        header_label = QLabel("Raw Material Consumption")
        header_label.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(header_label)

        # Input & Action Section
        input_action_layout = QHBoxLayout()
        input_action_layout.setSpacing(10)
        input_action_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        # Month Input
        month_label = QLabel("Start Month:")
        month_label.setFont(QFont("Segoe UI", 10))
        input_action_layout.addWidget(month_label)
        self.month_combo = QComboBox()
        self.month_combo.setFont(QFont("Segoe UI", 10))
        self.month_combo.addItems([
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ])
        self.month_combo.setCurrentIndex(0)
        self.month_combo.setMinimumWidth(120)
        input_action_layout.addWidget(self.month_combo)

        # Year Input
        year_label = QLabel("Start Year:")
        year_label.setFont(QFont("Segoe UI", 10))
        input_action_layout.addWidget(year_label)
        self.year_edit = QLineEdit()
        self.year_edit.setFont(QFont("Segoe UI", 10))
        self.year_edit.setPlaceholderText("e.g., 2023")
        self.year_edit.setText(str(datetime.now().year - 2))
        self.year_edit.setFixedWidth(80)
        input_action_layout.addWidget(self.year_edit)

        # Generate Table Button
        self.generate_button = QPushButton("Refresh View")
        self.generate_button.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.generate_button.setFixedSize(120, 30)
        self.generate_button.clicked.connect(self.generate_table)
        input_action_layout.addWidget(self.generate_button)

        # Load Excel Button
        self.load_excel_button = QPushButton("Load from Excel")
        self.load_excel_button.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.load_excel_button.setFixedSize(140, 30)
        self.load_excel_button.clicked.connect(self.load_data_from_excel)
        input_action_layout.addWidget(self.load_excel_button)

        # Export to Excel Button
        self.export_excel_button = QPushButton("Export to Excel")
        self.export_excel_button.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.export_excel_button.setFixedSize(140, 30)
        self.export_excel_button.clicked.connect(self.export_to_excel)
        input_action_layout.addWidget(self.export_excel_button)

        main_layout.addLayout(input_action_layout)

        # Table Section
        self.table_widget = QTableWidget()
        self.table_widget.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table_widget.setAlternatingRowColors(True)
        self.table_widget.setFont(QFont("Segoe UI", 9))
        main_layout.addWidget(self.table_widget)

        self.setLayout(main_layout)
        self.generate_table()  # Initial table setup

    def apply_styles(self):
        palette = self.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor("#F0F2F5"))
        palette.setColor(QPalette.ColorRole.WindowText, QColor("#333333"))
        palette.setColor(QPalette.ColorRole.Base, QColor("#FFFFFF"))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#E8ECF1"))
        palette.setColor(QPalette.ColorRole.Text, QColor("#333333"))
        palette.setColor(QPalette.ColorRole.Highlight, QColor("#4A90E2"))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
        self.setPalette(palette)

        self.setStyleSheet("""
            QWidget {
                background-color: #F0F2F5;
                color: #333333;
            }
            QLabel#header_label {
                color: #2C3E50;
                margin-bottom: 10px;
            }
            QComboBox, QLineEdit {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                padding: 5px;
                background-color: #FFFFFF;
                selection-background-color: #4A90E2;
                selection-color: #FFFFFF;
            }
            QPushButton {
                background-color: #5C6BC0;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 15px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #7986CB;
            }
            QPushButton:pressed {
                background-color: #3F51B5;
            }
            QTableWidget {
                border: 1px solid #D1D9E0;
                gridline-color: #E0E0E0;
                background-color: #FFFFFF;
                selection-background-color: #A7C7ED;
                selection-color: #333333;
                border-radius: 5px;
            }
            QHeaderView::section {
                background-color: #DDE4ED;
                color: #2C3E50;
                padding: 6px;
                border: 1px solid #CCD2D9;
                font-weight: bold;
                font-size: 10pt;
            }
            QHeaderView::section:horizontal {
                border-bottom: 1px solid #B0BACC;
            }
        """)

    def generate_table(self):
        try:
            start_month_index = self.month_combo.currentIndex() + 1
            start_year = int(self.year_edit.text())
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid year.")
            return

        # Generate the 12-month range for headers
        headers = ["Raw Material Code"]
        self.month_years_for_headers = []
        current_month = start_month_index
        current_year = start_year

        for i in range(12):
            month_name = datetime(current_year, current_month, 1).strftime("%b %Y")
            headers.append(month_name)
            self.month_years_for_headers.append(month_name)
            current_month += 1
            if current_month > 12:
                current_month = 1
                current_year += 1

        self.table_widget.setColumnCount(len(headers))
        self.table_widget.setHorizontalHeaderLabels(headers)

        self.table_widget.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        self.table_widget.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.ResizeToContents
        )

        self.table_widget.setRowCount(0)

    def load_data_from_excel(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_name:
            try:
                # Show progress dialog
                progress = QProgressDialog("Loading Excel file...", "Cancel", 0, 0, self)
                progress.setWindowModality(Qt.WindowModality.WindowModal)
                progress.setCancelButton(None)
                progress.show()
                QApplication.processEvents()

                # Load Excel file
                df = pd.read_excel(file_name, sheet_name=0)
                self._current_df = df
                self.table_widget.clearContents()

                # Update headers and populate table
                self.generate_table()
                self.populate_table(df)

                progress.close()
            except Exception as e:
                progress.close()
                QMessageBox.critical(self, "Error", f"Error loading Excel file: {e}")

    def populate_table(self, df):
        if not hasattr(self, 'month_years_for_headers') or not self.month_years_for_headers:
            QMessageBox.warning(self, "Error", "Headers not generated. Please click 'Refresh View' first.")
            return

        # Standardize column names
        df.columns = df.columns.str.strip().str.lower()
        required_cols = {"prod_date", "raw material", "qty used"}
        available_cols = [col for col in df.columns if col in required_cols]
        if not available_cols:
            QMessageBox.warning(self, "Error", "No required columns found in Excel file.")
            return
        df = df[available_cols]

        # Convert prod_date early and filter invalid dates
        df["prod_date"] = pd.to_datetime(df["prod_date"], errors="coerce")
        df = df.dropna(subset=["prod_date"])

        # Filter rows for relevant months
        start_date = pd.to_datetime(f"{self.month_combo.currentText()} {self.year_edit.text()}")
        end_date = start_date + pd.offsets.MonthEnd(12)
        df = df[(df["prod_date"] >= start_date) & (df["prod_date"] <= end_date)]

        # Convert qty used to numeric
        df["qty used"] = pd.to_numeric(
            df["qty used"].astype(str).str.replace(",", "", regex=False).str.strip(),
            errors="coerce"
        ).fillna(0)

        # Normalize material codes
        df["raw material"] = df["raw material"].astype(str).str.strip().str.upper()

        # Extract month-year
        df["month_year"] = df["prod_date"].dt.strftime("%b %Y")

        # Group and pivot
        grouped = df.groupby(["raw material", "month_year"])["qty used"].sum().reset_index()
        pivot_df = grouped.pivot_table(
            index="raw material",
            columns="month_year",
            values="qty used",
            aggfunc="sum",
            fill_value=0
        )

        # Ensure all required months appear
        for m in self.month_years_for_headers:
            if m not in pivot_df.columns:
                pivot_df[m] = 0

        # Reorder columns
        columns_to_select = ["raw material"] + self.month_years_for_headers
        pivot_df = pivot_df.reset_index()[columns_to_select]

        # Populate QTableWidget
        self.table_widget.setRowCount(0)
        self.table_widget.setRowCount(len(pivot_df))

        # Disable updates for performance
        self.table_widget.setUpdatesEnabled(False)
        try:
            for row_idx, row in pivot_df.iterrows():
                # Material code
                item_code = QTableWidgetItem(str(row["raw material"]))
                item_code.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table_widget.setItem(row_idx, 0, item_code)

                # Monthly values with highlighting for non-zero
                for col_offset, month in enumerate(self.month_years_for_headers, 1):
                    qty = row[month]
                    item_value = QTableWidgetItem(f"{qty:.2f}")
                    item_value.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
                    if qty > 0:  # Highlight cells with non-zero values
                        item_value.setBackground(QColor("#E6F0FA"))  # Light blue highlight
                    self.table_widget.setItem(row_idx, col_offset, item_value)
        finally:
            self.table_widget.setUpdatesEnabled(True)

        self._pivot_df = pivot_df

    def export_to_excel(self):
        if not hasattr(self, '_pivot_df') or self._pivot_df.empty:
            QMessageBox.warning(self, "No Data", "No data available to export. Please load data first.")
            return

        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "", "Excel Files (*.xlsx)"
        )
        if file_name:
            try:
                if not file_name.endswith('.xlsx'):
                    file_name += '.xlsx'
                self._pivot_df.to_excel(file_name, index=False)
                QMessageBox.information(self, "Success", f"Data exported successfully to {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error exporting to Excel: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RawMaterialApp()
    window.show()
    sys.exit(app.exec())