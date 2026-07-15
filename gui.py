import os
import pandas as pd
import math
from PySide6.QtWidgets import (
    QWidget,
    QLabel,
    QPushButton,
    QLineEdit,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QProgressBar,
    QRadioButton,
    QButtonGroup,
    QMessageBox,
    QCheckBox,
    QSpinBox,
)
from PySide6.QtCore import Qt
from excel_engine import GeneratorWorker


class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("NMMA AUTOMATE")
        self.resize(750, 530)
        
        self.worker = None

        self.apply_styles()
        self.create_ui()

    def apply_styles(self):
        """
        Applies a modern dark-mode stylesheet to the interface
        to give it a premium, professional appearance.
        """
        self.setStyleSheet("""
            QWidget {
                background-color: #1a1a24;
                color: #e2e8f0;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
            }
            QLabel {
                font-weight: bold;
                color: #cbd5e1;
            }
            QLineEdit {
                background-color: #2d3748;
                color: #ffffff;
                border: 1px solid #4a5568;
                border-radius: 6px;
                padding: 8px 12px;
            }
            QLineEdit:focus {
                border: 1px solid #3182ce;
            }
            QPushButton {
                background-color: #4a5568;
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #718096;
            }
            QPushButton:pressed {
                background-color: #2d3748;
            }
            QPushButton:disabled {
                background-color: #1e2530;
                color: #4a5568;
            }
            QPushButton#generate_btn {
                background-color: #3182ce;
                font-size: 14px;
                padding: 10px 20px;
            }
            QPushButton#generate_btn:hover {
                background-color: #4299e1;
            }
            QPushButton#generate_btn:pressed {
                background-color: #2b6cb0;
            }
            QPushButton#cancel_btn {
                background-color: #e53e3e;
                font-size: 14px;
                padding: 10px 20px;
            }
            QPushButton#cancel_btn:hover {
                background-color: #f56565;
            }
            QPushButton#cancel_btn:pressed {
                background-color: #c53030;
            }
            QProgressBar {
                border: 1px solid #4a5568;
                border-radius: 6px;
                background-color: #2d3748;
                text-align: center;
                color: #ffffff;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #3182ce;
                border-radius: 5px;
            }
            QRadioButton {
                color: #cbd5e1;
                font-size: 13px;
            }
            QRadioButton::indicator {
                width: 16px;
                height: 16px;
            }
            QCheckBox {
                color: #cbd5e1;
                font-size: 13px;
            }
            QSpinBox {
                background-color: #2d3748;
                color: #ffffff;
                border: 1px solid #4a5568;
                border-radius: 6px;
                padding: 4px;
            }
        """)

    def create_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(25, 25, 25, 25)

        # Title / Header
        title_label = QLabel("National Mission on Monuments and Antiquities Form Automator")
        title_label.setStyleSheet("font-size: 18px; color: #ffffff; padding-bottom: 10px;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # -----------------------------
        # Book1 File Row
        # -----------------------------
        book_layout = QHBoxLayout()
        book_label = QLabel("Master Data File:")
        book_label.setFixedWidth(120)
        self.book_edit = QLineEdit()
        self.book_edit.setPlaceholderText("Select Master Data Excel File (e.g., Book1.xlsx)")

        self.book_browse_btn = QPushButton("Browse")
        self.book_browse_btn.clicked.connect(self.select_book)

        book_layout.addWidget(book_label)
        book_layout.addWidget(self.book_edit)
        book_layout.addWidget(self.book_browse_btn)

        # -----------------------------
        # Template File Row
        # -----------------------------
        template_layout = QHBoxLayout()
        template_label = QLabel("NMMA Template:")
        template_label.setFixedWidth(120)
        self.template_edit = QLineEdit()
        self.template_edit.setPlaceholderText("Select NMMA Form Template (e.g., NMMA Template.xlsx)")

        self.template_browse_btn = QPushButton("Browse")
        self.template_browse_btn.clicked.connect(self.select_template)

        template_layout.addWidget(template_label)
        template_layout.addWidget(self.template_edit)
        template_layout.addWidget(self.template_browse_btn)

        # -----------------------------
        # Output Folder Row
        # -----------------------------
        output_layout = QHBoxLayout()
        output_label = QLabel("Output Folder:")
        output_label.setFixedWidth(120)
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("Select Output Folder for generated forms")

        self.output_browse_btn = QPushButton("Browse")
        self.output_browse_btn.clicked.connect(self.select_output)

        output_layout.addWidget(output_label)
        output_layout.addWidget(self.output_edit)
        output_layout.addWidget(self.output_browse_btn)

        # Layout assignments
        layout.addLayout(book_layout)
        layout.addLayout(template_layout)
        layout.addLayout(output_layout)

        # -----------------------------
        # Output Format Selection
        # -----------------------------
        format_group_box = QHBoxLayout()
        format_label = QLabel("Output Format:")
        format_label.setFixedWidth(120)
        
        self.excel_radio = QRadioButton("Excel (.xlsx)")
        self.pdf_radio = QRadioButton("PDF (.pdf)")
        self.both_radio = QRadioButton("Excel + PDF")
        
        self.excel_radio.setChecked(True)

        self.format_group = QButtonGroup(self)
        self.format_group.addButton(self.excel_radio)
        self.format_group.addButton(self.pdf_radio)
        self.format_group.addButton(self.both_radio)

        format_group_box.addWidget(format_label)
        format_group_box.addWidget(self.excel_radio)
        format_group_box.addWidget(self.pdf_radio)
        format_group_box.addWidget(self.both_radio)
        format_group_box.addStretch()
        
        layout.addLayout(format_group_box)

        # -----------------------------
        # Grouping Settings Row
        # -----------------------------
        grouping_layout = QHBoxLayout()
        grouping_label = QLabel("Workbook Grouping:")
        grouping_label.setFixedWidth(120)

        self.group_checkbox = QCheckBox("Group forms into multi-sheet workbooks")
        
        self.group_size_label = QLabel("Sheets per file:")
        self.group_size_label.setStyleSheet("padding-left: 15px; color: #a0aec0;")
        self.group_size_spin = QSpinBox()
        self.group_size_spin.setRange(2, 5000)
        self.group_size_spin.setValue(50)
        self.group_size_spin.setFixedWidth(80)
        self.group_size_spin.setEnabled(False)

        # Connect checkbox to toggle SpinBox
        self.group_checkbox.toggled.connect(self.group_size_spin.setEnabled)

        grouping_layout.addWidget(grouping_label)
        grouping_layout.addWidget(self.group_checkbox)
        grouping_layout.addWidget(self.group_size_label)
        grouping_layout.addWidget(self.group_size_spin)
        grouping_layout.addStretch()

        layout.addLayout(grouping_layout)

        # -----------------------------
        # Progress Bar & Status
        # -----------------------------
        progress_layout = QVBoxLayout()
        progress_layout.setSpacing(5)
        
        progress_header = QLabel("Progress:")
        self.progress = QProgressBar()
        self.progress.setValue(0)
        
        self.status = QLabel("Ready")
        self.status.setStyleSheet("color: #a0aec0; font-style: italic;")

        progress_layout.addWidget(progress_header)
        progress_layout.addWidget(self.progress)
        progress_layout.addWidget(self.status)
        
        layout.addLayout(progress_layout)

        layout.addStretch()

        # -----------------------------
        # Action Buttons
        # -----------------------------
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(15)

        self.generate_button = QPushButton("Generate NMMA Forms")
        self.generate_button.setObjectName("generate_btn")
        self.generate_button.clicked.connect(self.generate_forms)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setObjectName("cancel_btn")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel_generation)

        buttons_layout.addWidget(self.generate_button)
        buttons_layout.addWidget(self.cancel_button)

        layout.addLayout(buttons_layout)

        self.setLayout(layout)

    # ==========================================
    # Selection Dialogs
    # ==========================================

    def select_book(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select Master Data File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if filename:
            self.book_edit.setText(filename)

    def select_template(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select NMMA Template File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if filename:
            self.template_edit.setText(filename)

    def select_output(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder"
        )
        if folder:
            self.output_edit.setText(folder)

    # ==========================================
    # State Helpers
    # ==========================================

    def set_inputs_enabled(self, enabled: bool):
        self.book_edit.setEnabled(enabled)
        self.book_browse_btn.setEnabled(enabled)
        self.template_edit.setEnabled(enabled)
        self.template_browse_btn.setEnabled(enabled)
        self.output_edit.setEnabled(enabled)
        self.output_browse_btn.setEnabled(enabled)
        
        self.excel_radio.setEnabled(enabled)
        self.pdf_radio.setEnabled(enabled)
        self.both_radio.setEnabled(enabled)
        
        self.group_checkbox.setEnabled(enabled)
        self.group_size_spin.setEnabled(enabled and self.group_checkbox.isChecked())
        
        self.generate_button.setEnabled(enabled)
        self.cancel_button.setEnabled(not enabled)

    # ==========================================
    # Processing Logic
    # ==========================================

    def generate_forms(self):
        book_path = self.book_edit.text().strip()
        template_path = self.template_edit.text().strip()
        output_folder = self.output_edit.text().strip()

        # Path Validations
        if not book_path:
            QMessageBox.warning(self, "Validation Error", "Please select the Master Data file.")
            return
        if not os.path.exists(book_path):
            QMessageBox.critical(self, "Error", f"Master Data file does not exist:\n{book_path}")
            return

        if not template_path:
            QMessageBox.warning(self, "Validation Error", "Please select the NMMA Template file.")
            return
        if not os.path.exists(template_path):
            QMessageBox.critical(self, "Error", f"Template file does not exist:\n{template_path}")
            return

        if not output_folder:
            QMessageBox.warning(self, "Validation Error", "Please select the Output Folder.")
            return

        # Output format resolution
        output_format = "Excel"
        if self.pdf_radio.isChecked():
            output_format = "PDF"
        elif self.both_radio.isChecked():
            output_format = "Excel + PDF"

        # Resolve group size
        group_size = 1
        if self.group_checkbox.isChecked():
            group_size = self.group_size_spin.value()

        # Read row count to verify size and warn user if PDF generation will take too long
        try:
            df = pd.read_excel(book_path)
            row_count = len(df)
        except Exception as e:
            QMessageBox.critical(self, "Error Reading Excel", f"Failed to read master data workbook:\n{e}")
            return

        if row_count == 0:
            QMessageBox.warning(self, "No Records", "The selected Master Data file contains no records.")
            return

        # Show performance warning for large PDF runs
        if output_format in ["PDF", "Excel + PDF"] and row_count > 1000:
            approx_minutes = int(row_count * 1.5 / 60)
            approx_hours = round(approx_minutes / 60, 1)
            time_display = f"{approx_minutes} minutes" if approx_hours < 1.0 else f"{approx_hours} hours"
            
            warning_msg = (
                f"You have selected PDF format for {row_count} records.\n\n"
                f"PDF conversion via Microsoft Excel takes about 1.5 seconds per file. "
                f"Processing {row_count} records will take approximately **{time_display}**.\n\n"
                f"Do you want to proceed anyway?"
            )
            response = QMessageBox.warning(
                self,
                "Performance Warning",
                warning_msg,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if response == QMessageBox.No:
                return

        # Determine total items to process for progress bar
        total_items = row_count
        if group_size > 1:
            total_items = math.ceil(row_count / group_size)

        # Disable UI to prevent user interactions during processing
        self.set_inputs_enabled(False)
        self.progress.setValue(0)
        self.progress.setMaximum(total_items)
        self.status.setText("Initializing worker thread...")

        # Setup and start background worker
        self.worker = GeneratorWorker(book_path, template_path, output_folder, output_format, group_size)
        self.worker.progress_updated.connect(self.on_progress)
        self.worker.status_updated.connect(self.on_status)
        self.worker.error_occurred.connect(self.on_error)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def cancel_generation(self):
        if self.worker and self.worker.isRunning():
            self.cancel_button.setEnabled(False)
            self.status.setText("Cancellation requested. Terminating processes, please wait...")
            self.worker.cancel()

    # ==========================================
    # Worker Thread Signal Slots
    # ==========================================

    def on_progress(self, completed, total):
        self.progress.setMaximum(total)
        self.progress.setValue(completed)

    def on_status(self, msg):
        self.status.setText(msg)

    def on_error(self, error_msg):
        QMessageBox.critical(self, "Generation Error", error_msg)
        self.reset_ui()

    def on_finished(self, success_count, error_count, summary):
        QMessageBox.information(self, "Process Finished", summary)
        self.reset_ui()

    def reset_ui(self):
        self.set_inputs_enabled(True)
        self.progress.setValue(0)
        self.status.setText("Ready")
        self.worker = None