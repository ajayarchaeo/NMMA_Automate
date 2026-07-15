import os
import json
import pandas as pd
import openpyxl
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
    QComboBox,
    QScrollArea,
    QFormLayout,
    QFrame,
    QInputDialog,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from excel_engine import GeneratorWorker


class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("NMMA AUTOMATE")
        self.resize(1100, 600)
        
        self.worker = None
        self.mapping_widgets = {}
        self.saved_mapping_dict = {}
        self.presets = {}
        self.block_preset_signal = False

        self.apply_styles()
        self.create_ui()
        self.load_settings()

    def apply_styles(self):
        """
        Applies a premium warm-mode stylesheet using the pairing of Rockwell
        (for headers, titles, and buttons) and Garamond (for body labels, entries,
        and dropdown text).
        """
        self.setStyleSheet("""
            QWidget {
                background-color: #FFF5EE;
                color: #2d3748;
                font-family: 'Garamond', 'Georgia', serif;
                font-size: 15px;
            }
            QLabel {
                font-weight: bold;
                color: #2d3748;
            }
            QLabel#main_title {
                font-family: 'Rockwell', 'Courier New', serif;
                font-size: 21px;
                font-weight: 900;
                color: #368F8E;
            }
            QLabel#subtitle {
                font-family: 'Rockwell', 'Courier New', serif;
                font-size: 11px;
                font-weight: bold;
                color: #E29F2C;
            }
            QLabel#mapping_header {
                font-family: 'Rockwell', 'Courier New', serif;
                font-size: 18px;
                color: #368F8E;
            }
            QLineEdit {
                background-color: #ffffff;
                color: #1a1a24;
                border: 1px solid #F2D4A0;
                border-radius: 6px;
                padding: 8px 12px;
            }
            QLineEdit:focus {
                border: 2px solid #368F8E;
            }
            QPushButton {
                background-color: #368F8E;
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
                font-family: 'Rockwell', 'Courier New', serif;
            }
            QPushButton:hover {
                background-color: #2d7978;
            }
            QPushButton:pressed {
                background-color: #246261;
            }
            QPushButton:disabled {
                background-color: #e2e8f0;
                color: #a0aec0;
            }
            QPushButton#generate_btn {
                background-color: #E29F2C;
                color: #ffffff;
                font-size: 14px;
                padding: 10px 20px;
            }
            QPushButton#generate_btn:hover {
                background-color: #c78822;
            }
            QPushButton#generate_btn:pressed {
                background-color: #a67018;
            }
            QPushButton#cancel_btn {
                background-color: #e53e3e;
                color: #ffffff;
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
                border: 1px solid #F2D4A0;
                border-radius: 6px;
                background-color: #F6E3C0;
                text-align: center;
                color: #1a1a24;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #368F8E;
                border-radius: 5px;
            }
            QRadioButton {
                color: #2d3748;
                font-size: 14px;
            }
            QCheckBox {
                color: #2d3748;
                font-size: 14px;
            }
            QSpinBox {
                background-color: #ffffff;
                color: #1a1a24;
                border: 1px solid #F2D4A0;
                border-radius: 6px;
                padding: 4px;
            }
            QComboBox {
                background-color: #ffffff;
                color: #1a1a24;
                border: 1px solid #F2D4A0;
                border-radius: 4px;
                padding: 4px 8px;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                color: #1a1a24;
                selection-background-color: #368F8E;
                selection-color: #ffffff;
            }
            QScrollArea {
                border: 1px solid #F2D4A0;
                border-radius: 6px;
                background-color: #F6E3C0;
            }
        """)

    def create_ui(self):
        main_layout = QHBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # -----------------------------
        # Left Panel (Settings & Progress)
        # -----------------------------
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(12)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Title & Logo Header Row
        header_layout = QHBoxLayout()
        header_layout.setSpacing(15)
        header_layout.setContentsMargins(5, 5, 5, 5)
        
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo.png")
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            scaled_pixmap = pixmap.scaled(60, 85, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(scaled_pixmap)
            logo_label.setFixedSize(60, 85)
            logo_label.setStyleSheet("border: 2px solid #E29F2C; border-radius: 6px; padding: 2px; background-color: #ffffff;")
        else:
            logo_label.setText("[Logo]")
            logo_label.setFixedSize(60, 85)
            logo_label.setAlignment(Qt.AlignCenter)
            logo_label.setStyleSheet("color: #E29F2C; border: 2px dashed #E29F2C; border-radius: 6px;")
            
        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(4)
        
        main_title = QLabel("NMMA Automated")
        main_title.setObjectName("main_title")
        
        subtitle = QLabel("Excel automation for Archaeological Documentation")
        subtitle.setObjectName("subtitle")
        
        title_vbox.addWidget(main_title)
        title_vbox.addWidget(subtitle)
        title_vbox.addStretch()
        
        header_layout.addWidget(logo_label)
        header_layout.addLayout(title_vbox)
        header_layout.addStretch()
        
        left_layout.addLayout(header_layout)

        # Book1 File Row
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

        # Template File Row
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

        # Output Folder Row
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

        # Photos Folder Row (Optional)
        photo_folder_layout = QHBoxLayout()
        photo_folder_label = QLabel("Photos Folder:")
        photo_folder_label.setFixedWidth(120)
        self.photo_edit = QLineEdit()
        self.photo_edit.setPlaceholderText("Select folder containing photographs (Optional)")
        self.photo_browse_btn = QPushButton("Browse")
        self.photo_browse_btn.clicked.connect(self.select_photos_folder)
        photo_folder_layout.addWidget(photo_folder_label)
        photo_folder_layout.addWidget(self.photo_edit)
        photo_folder_layout.addWidget(self.photo_browse_btn)

        # Photo Dimensions Row
        photo_dim_layout = QHBoxLayout()
        photo_dim_label = QLabel("Photo Size (px):")
        photo_dim_label.setFixedWidth(120)
        
        w_label = QLabel("Width:")
        w_label.setStyleSheet("font-weight: normal; color: #4a5568; margin-left: 10px;")
        self.photo_width_spin = QSpinBox()
        self.photo_width_spin.setRange(50, 2000)
        self.photo_width_spin.setValue(220)
        self.photo_width_spin.setFixedWidth(80)
        
        h_label = QLabel("Height:")
        h_label.setStyleSheet("font-weight: normal; color: #4a5568; margin-left: 10px;")
        self.photo_height_spin = QSpinBox()
        self.photo_height_spin.setRange(50, 2000)
        self.photo_height_spin.setValue(200)
        self.photo_height_spin.setFixedWidth(80)
        
        photo_dim_layout.addWidget(photo_dim_label)
        photo_dim_layout.addWidget(w_label)
        photo_dim_layout.addWidget(self.photo_width_spin)
        photo_dim_layout.addWidget(h_label)
        photo_dim_layout.addWidget(self.photo_height_spin)
        photo_dim_layout.addStretch()

        left_layout.addLayout(book_layout)
        left_layout.addLayout(template_layout)
        left_layout.addLayout(output_layout)
        left_layout.addLayout(photo_folder_layout)
        left_layout.addLayout(photo_dim_layout)

        # Output Format Selection
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
        left_layout.addLayout(format_group_box)

        # Grouping Settings Row
        grouping_section = QVBoxLayout()
        grouping_section.setSpacing(8)

        grouping_header_layout = QHBoxLayout()
        grouping_label = QLabel("Workbook Grouping:")
        grouping_label.setFixedWidth(120)
        self.group_checkbox = QCheckBox("Group multiple antiquities into single files")
        self.group_size_label = QLabel("Forms per file:")
        self.group_size_label.setStyleSheet("padding-left: 15px; color: #E29F2C;")
        self.group_size_spin = QSpinBox()
        self.group_size_spin.setRange(2, 5000)
        self.group_size_spin.setValue(50)
        self.group_size_spin.setFixedWidth(80)
        self.group_size_spin.setEnabled(False)

        grouping_header_layout.addWidget(grouping_label)
        grouping_header_layout.addWidget(self.group_checkbox)
        grouping_header_layout.addWidget(self.group_size_label)
        grouping_header_layout.addWidget(self.group_size_spin)
        grouping_header_layout.addStretch()

        # Layout options row (visible/enabled only when grouping is checked)
        layout_options_layout = QHBoxLayout()
        layout_options_label = QLabel("Grouping Layout:")
        layout_options_label.setFixedWidth(120)
        layout_options_label.setStyleSheet("color: #4a5568; font-weight: normal;")
        self.layout_sheets_radio = QRadioButton("Separate sheets (tabs)")
        self.layout_stacked_radio = QRadioButton("Stacked vertically")
        self.layout_horizontal_radio = QRadioButton("Stacked horizontally")
        self.layout_sheets_radio.setChecked(True)
        self.layout_sheets_radio.setEnabled(False)
        self.layout_stacked_radio.setEnabled(False)
        self.layout_horizontal_radio.setEnabled(False)

        self.layout_group = QButtonGroup(self)
        self.layout_group.addButton(self.layout_sheets_radio)
        self.layout_group.addButton(self.layout_stacked_radio)
        self.layout_group.addButton(self.layout_horizontal_radio)

        layout_options_layout.addWidget(layout_options_label)
        layout_options_layout.addWidget(self.layout_sheets_radio)
        layout_options_layout.addWidget(self.layout_stacked_radio)
        layout_options_layout.addWidget(self.layout_horizontal_radio)
        layout_options_layout.addStretch()

        # Connect checkbox toggled event to control sub-inputs
        def toggle_grouping_options(checked):
            self.group_size_spin.setEnabled(checked)
            self.layout_sheets_radio.setEnabled(checked)
            self.layout_stacked_radio.setEnabled(checked)
            self.layout_horizontal_radio.setEnabled(checked)

        self.group_checkbox.toggled.connect(toggle_grouping_options)

        grouping_section.addLayout(grouping_header_layout)
        grouping_section.addLayout(layout_options_layout)
        left_layout.addLayout(grouping_section)

        # Progress Bar & Status
        progress_layout = QVBoxLayout()
        progress_layout.setSpacing(5)
        progress_header = QLabel("Progress:")
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.status = QLabel("Ready")
        self.status.setStyleSheet("color: #368F8E; font-style: italic;")
        progress_layout.addWidget(progress_header)
        progress_layout.addWidget(self.progress)
        progress_layout.addWidget(self.status)
        left_layout.addLayout(progress_layout)

        left_layout.addStretch()

        # Action Buttons
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
        left_layout.addLayout(buttons_layout)

        # -----------------------------
        # Vertical Separator Line
        # -----------------------------
        separator = QFrame()
        separator.setFrameShape(QFrame.VLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("color: #F2D4A0;")

        # -----------------------------
        # Right Panel (Column Mapping)
        # -----------------------------
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setSpacing(10)
        right_layout.setContentsMargins(0, 0, 0, 0)

        self.mapping_header = QLabel("Column Mapping")
        self.mapping_header.setObjectName("mapping_header")
        
        self.mapping_status = QLabel("Select Master Data & NMMA Template to scan fields...")
        self.mapping_status.setStyleSheet("color: #4a5568; font-style: italic;")
        self.mapping_status.setWordWrap(True)

        # Preset Selector Bar
        preset_layout = QHBoxLayout()
        preset_label = QLabel("Preset:")
        preset_label.setStyleSheet("color: #2d3748; font-weight: bold;")
        
        self.preset_combo = QComboBox()
        self.preset_combo.addItem("(Auto Fuzzy Map)")
        self.preset_combo.currentIndexChanged.connect(self.apply_selected_preset)
        
        self.save_preset_btn = QPushButton("Save")
        self.save_preset_btn.setToolTip("Save current column mappings as a preset")
        self.save_preset_btn.clicked.connect(self.save_current_preset)
        
        self.delete_preset_btn = QPushButton("Delete")
        self.delete_preset_btn.setToolTip("Delete the selected preset")
        self.delete_preset_btn.clicked.connect(self.delete_current_preset)
        self.delete_preset_btn.setEnabled(False)
        
        preset_layout.addWidget(preset_label)
        preset_layout.addWidget(self.preset_combo, stretch=1)
        preset_layout.addWidget(self.save_preset_btn)
        preset_layout.addWidget(self.delete_preset_btn)

        # Scroll Area for the mapping table
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        scroll_content = QWidget()
        self.mapping_form_layout = QFormLayout(scroll_content)
        self.mapping_form_layout.setSpacing(10)
        self.mapping_form_layout.setContentsMargins(10, 10, 10, 10)
        scroll.setWidget(scroll_content)

        right_layout.addWidget(self.mapping_header)
        right_layout.addWidget(self.mapping_status)
        right_layout.addLayout(preset_layout)
        right_layout.addWidget(scroll)

        # Connect paths text changes to scan fields automatically
        self.book_edit.textChanged.connect(self.update_column_mappings)
        self.template_edit.textChanged.connect(self.update_column_mappings)

        # Assemble main window
        main_layout.addWidget(left_widget, stretch=5)
        main_layout.addWidget(separator)
        main_layout.addWidget(right_widget, stretch=4)
        
        self.setLayout(main_layout)

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

    def select_photos_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Photographs Folder"
        )
        if folder:
            self.photo_edit.setText(folder)

    # ==========================================
    # Mapping & Presets Logic
    # ==========================================

    def clear_mapping_ui(self):
        while self.mapping_form_layout.count():
            child = self.mapping_form_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        self.mapping_widgets = {}

    def update_column_mappings(self):
        book_path = self.book_edit.text().strip()
        template_path = self.template_edit.text().strip()
        
        if not book_path or not template_path:
            return
        if not os.path.exists(book_path) or not os.path.exists(template_path):
            return
            
        try:
            # 1. Read columns from Master Excel
            df = pd.read_excel(book_path, nrows=0)
            master_cols = df.columns.tolist()
            
            # 2. Read fields from NMMA Template Column B
            wb = openpyxl.load_workbook(template_path, data_only=True)
            ws = wb.active
            
            fields = []
            # Scan rows 2 to 24 (Column B)
            for r in range(2, 25):
                cell_val = ws[f"B{r}"].value
                if cell_val and str(cell_val).strip():
                    fields.append((r, str(cell_val).strip()))
            wb.close()
            
            # 3. Populate Right Pane
            self.clear_mapping_ui()
            
            # Helper to normalize string for fuzzy matching
            import re
            def normalize(text):
                text = str(text).lower().strip()
                text = re.sub(r'[^a-z0-9]', '', text)
                for filler in ["ofthe", "of", "in", "incom", "weightingram", "weight", "no"]:
                    text = text.replace(filler, "")
                return text

            for r, label in fields:
                cell_ref = f"C{r}"
                field_label = QLabel(f"{label} ({cell_ref}):")
                field_label.setStyleSheet("color: #2d3748; font-weight: normal;")
                
                combo = QComboBox()
                combo.addItem("(Use Template Default)")
                for col in master_cols:
                    combo.addItem(col)
                
                # Check if we have a saved mapping first
                saved_col = self.saved_mapping_dict.get(cell_ref)
                if saved_col and (saved_col in master_cols or saved_col == "(Use Template Default)"):
                    combo.setCurrentText(saved_col)
                else:
                    # Apply Fuzzy Matching
                    norm_label = normalize(label)
                    best_idx = 0
                    for idx, col in enumerate(master_cols, start=1):
                        norm_col = normalize(col)
                        if norm_label == norm_col or norm_label in norm_col or norm_col in norm_label:
                            best_idx = idx
                            break
                    
                    # Dynamic fallbacks
                    if best_idx == 0:
                        if "date" in norm_label:
                            for idx, col in enumerate(master_cols, start=1):
                                if "date" in col.lower():
                                    best_idx = idx
                                    break
                        elif "accession" in norm_label:
                            for idx, col in enumerate(master_cols, start=1):
                                if "accession" in col.lower() or "nmma" in col.lower():
                                    best_idx = idx
                                    break
                    
                    combo.setCurrentIndex(best_idx)
                
                self.mapping_widgets[cell_ref] = combo
                self.mapping_form_layout.addRow(field_label, combo)

            self.mapping_header.setText("Column Mapping (Scanned)")
            self.mapping_status.setText(f"Found {len(fields)} form fields. Adjust column mappings below if needed.")
            self.mapping_status.setStyleSheet("color: #368F8E; font-style: italic;")

        except Exception as e:
            self.mapping_status.setText(f"Error scanning template fields:\n{e}")
            self.mapping_status.setStyleSheet("color: #e53e3e; font-style: italic;")

    def save_current_preset(self):
        # Prompt for preset name
        name, ok = QInputDialog.getText(
            self,
            "Save Preset",
            "Enter a name for the column mapping preset:"
        )
        if not ok or not name.strip():
            return
            
        preset_name = name.strip()
        if preset_name == "(Auto Fuzzy Map)":
            QMessageBox.warning(self, "Invalid Name", "Cannot overwrite default auto fuzzy mapping.")
            return
            
        # Compile mapping dictionary
        mapping_dict = {}
        for cell_ref, combo in self.mapping_widgets.items():
            mapping_dict[cell_ref] = combo.currentText()
            
        self.presets[preset_name] = mapping_dict
        
        # Refresh ComboBox list
        self.block_preset_signal = True
        self.preset_combo.clear()
        self.preset_combo.addItem("(Auto Fuzzy Map)")
        for p in sorted(self.presets.keys()):
            self.preset_combo.addItem(p)
        self.preset_combo.setCurrentText(preset_name)
        self.delete_preset_btn.setEnabled(True)
        self.block_preset_signal = False
        
        # Save to configuration
        self.save_settings()
        QMessageBox.information(self, "Preset Saved", f"Preset '{preset_name}' saved successfully.")

    def delete_current_preset(self):
        preset_name = self.preset_combo.currentText()
        if preset_name == "(Auto Fuzzy Map)":
            return
            
        response = QMessageBox.question(
            self,
            "Delete Preset",
            f"Are you sure you want to delete preset '{preset_name}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if response == QMessageBox.No:
            return
            
        if preset_name in self.presets:
            del self.presets[preset_name]
            
        # Rebuild ComboBox
        self.block_preset_signal = True
        self.preset_combo.clear()
        self.preset_combo.addItem("(Auto Fuzzy Map)")
        for p in sorted(self.presets.keys()):
            self.preset_combo.addItem(p)
        self.preset_combo.setCurrentIndex(0)
        self.delete_preset_btn.setEnabled(False)
        self.block_preset_signal = False
        
        self.save_settings()
        
        # Reset and trigger auto fuzzy mapping
        self.saved_mapping_dict = {}
        self.update_column_mappings()

    def apply_selected_preset(self):
        if self.block_preset_signal:
            return
            
        preset_name = self.preset_combo.currentText()
        if preset_name == "(Auto Fuzzy Map)":
            self.delete_preset_btn.setEnabled(False)
            self.saved_mapping_dict = {}
            self.update_column_mappings()
        else:
            self.delete_preset_btn.setEnabled(True)
            mapping_dict = self.presets.get(preset_name, {})
            
            # Apply to mappings
            self.saved_mapping_dict = mapping_dict
            self.update_column_mappings()

    # ==========================================
    # Settings Saving & Loading
    # ==========================================

    def get_settings_filepath(self):
        local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")
        try:
            temp_file = local_path + ".tmp"
            with open(temp_file, "w") as f:
                f.write("")
            os.remove(temp_file)
            return local_path
        except Exception:
            return os.path.join(os.path.expanduser("~"), ".nmma_automate_settings.json")

    def save_settings(self):
        settings = {
            "book_path": self.book_edit.text().strip(),
            "template_path": self.template_edit.text().strip(),
            "output_folder": self.output_edit.text().strip(),
            "photos_dir": self.photo_edit.text().strip(),
            "photo_width": self.photo_width_spin.value(),
            "photo_height": self.photo_height_spin.value(),
            "output_format": "Excel" if self.excel_radio.isChecked() else "PDF" if self.pdf_radio.isChecked() else "Excel + PDF",
            "is_grouped": self.group_checkbox.isChecked(),
            "group_size": self.group_size_spin.value(),
            "group_layout": "stacked" if self.layout_stacked_radio.isChecked() else "horizontal" if self.layout_horizontal_radio.isChecked() else "sheets",
            "mapping_dict": {},
            "presets": getattr(self, "presets", {}),
            "selected_preset": self.preset_combo.currentText()
        }
        
        for cell_ref, combo in self.mapping_widgets.items():
            settings["mapping_dict"][cell_ref] = combo.currentText()
            
        try:
            filepath = self.get_settings_filepath()
            with open(filepath, "w") as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def load_settings(self):
        filepath = self.get_settings_filepath()
        if not os.path.exists(filepath):
            return
            
        try:
            with open(filepath, "r") as f:
                settings = json.load(f)
                
            # Load presets list first
            self.presets = settings.get("presets", {})
            
            # Rebuild preset ComboBox items
            self.block_preset_signal = True
            self.preset_combo.clear()
            self.preset_combo.addItem("(Auto Fuzzy Map)")
            for p in sorted(self.presets.keys()):
                self.preset_combo.addItem(p)
                
            selected_preset = settings.get("selected_preset", "(Auto Fuzzy Map)")
            if selected_preset in self.presets or selected_preset == "(Auto Fuzzy Map)":
                self.preset_combo.setCurrentText(selected_preset)
                self.delete_preset_btn.setEnabled(selected_preset != "(Auto Fuzzy Map)")
            self.block_preset_signal = False
            
            # Load basic settings
            self.book_edit.setText(settings.get("book_path", ""))
            self.template_edit.setText(settings.get("template_path", ""))
            self.output_edit.setText(settings.get("output_folder", ""))
            self.photo_edit.setText(settings.get("photos_dir", ""))
            self.photo_width_spin.setValue(settings.get("photo_width", 220))
            self.photo_height_spin.setValue(settings.get("photo_height", 200))
            
            output_format = settings.get("output_format", "Excel")
            if output_format == "Excel":
                self.excel_radio.setChecked(True)
            elif output_format == "PDF":
                self.pdf_radio.setChecked(True)
            else:
                self.both_radio.setChecked(True)
                
            self.group_checkbox.setChecked(settings.get("is_grouped", False))
            self.group_size_spin.setValue(settings.get("group_size", 50))
            
            group_layout = settings.get("group_layout", "sheets")
            if group_layout == "stacked":
                self.layout_stacked_radio.setChecked(True)
            elif group_layout == "horizontal":
                self.layout_horizontal_radio.setChecked(True)
            else:
                self.layout_sheets_radio.setChecked(True)
                
            is_grouped = self.group_checkbox.isChecked()
            self.group_size_spin.setEnabled(is_grouped)
            self.layout_sheets_radio.setEnabled(is_grouped)
            self.layout_stacked_radio.setEnabled(is_grouped)
            self.layout_horizontal_radio.setEnabled(is_grouped)
            
            # If a preset was active, set self.saved_mapping_dict to that preset's dict
            if selected_preset != "(Auto Fuzzy Map)":
                self.saved_mapping_dict = self.presets.get(selected_preset, {})
            else:
                self.saved_mapping_dict = settings.get("mapping_dict", {})
                
            self.update_column_mappings()
            
        except Exception as e:
            print(f"Error loading settings: {e}")

    def closeEvent(self, event):
        """
        Auto-save settings when application window is closed.
        """
        self.save_settings()
        event.accept()

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
        self.photo_edit.setEnabled(enabled)
        self.photo_browse_btn.setEnabled(enabled)
        
        self.photo_width_spin.setEnabled(enabled)
        self.photo_height_spin.setEnabled(enabled)
        
        self.excel_radio.setEnabled(enabled)
        self.pdf_radio.setEnabled(enabled)
        self.both_radio.setEnabled(enabled)
        
        self.group_checkbox.setEnabled(enabled)
        is_grouped = self.group_checkbox.isChecked()
        self.group_size_spin.setEnabled(enabled and is_grouped)
        self.layout_sheets_radio.setEnabled(enabled and is_grouped)
        self.layout_stacked_radio.setEnabled(enabled and is_grouped)
        self.layout_horizontal_radio.setEnabled(enabled and is_grouped)
        
        # Disable mapping comboboxes during run
        for combo in self.mapping_widgets.values():
            combo.setEnabled(enabled)
            
        # Disable presets UI during run
        self.preset_combo.setEnabled(enabled)
        self.save_preset_btn.setEnabled(enabled)
        self.delete_preset_btn.setEnabled(enabled and self.preset_combo.currentText() != "(Auto Fuzzy Map)")
            
        self.generate_button.setEnabled(enabled)
        self.cancel_button.setEnabled(not enabled)

    # ==========================================
    # Processing Logic
    # ==========================================

    def generate_forms(self):
        book_path = self.book_edit.text().strip()
        template_path = self.template_edit.text().strip()
        output_folder = self.output_edit.text().strip()
        photos_dir = self.photo_edit.text().strip()

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

        if not photos_dir:
            photos_dir = None
        elif not os.path.exists(photos_dir):
            QMessageBox.critical(self, "Error", f"Photos directory does not exist:\n{photos_dir}")
            return

        # Output format resolution
        output_format = "Excel"
        if self.pdf_radio.isChecked():
            output_format = "PDF"
        elif self.both_radio.isChecked():
            output_format = "Excel + PDF"

        # Resolve group size and layout
        group_size = 1
        group_layout = "sheets"
        if self.group_checkbox.isChecked():
            group_size = self.group_size_spin.value()
            if self.layout_stacked_radio.isChecked():
                group_layout = "stacked"
            elif self.layout_horizontal_radio.isChecked():
                group_layout = "horizontal"

        # Pull custom photo size
        photo_width = self.photo_width_spin.value()
        photo_height = self.photo_height_spin.value()

        # Save settings on run
        self.save_settings()

        # Compile mapping dict
        mapping_dict = {}
        for cell_ref, combo in self.mapping_widgets.items():
            mapping_dict[cell_ref] = combo.currentText()

        # Read row count to verify size
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

        # Disable UI during run
        self.set_inputs_enabled(False)
        self.progress.setValue(0)
        self.progress.setMaximum(total_items)
        self.status.setText("Initializing worker thread...")

        # Setup and start background worker
        self.worker = GeneratorWorker(book_path, template_path, output_folder, output_format, group_size, group_layout, mapping_dict, photos_dir, photo_width, photo_height)
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