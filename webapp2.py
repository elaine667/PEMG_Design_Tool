import sys
import pandas as pd
import os
import re
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QComboBox, QFrame, QMessageBox, QScrollArea)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

class CoreDimensionApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Universal Magnetic Core Explorer")
        self.setMinimumSize(700, 800)
        self.setStyleSheet("background-color: #121212; color: #e0e0e0;")

        self.df = None
        # Base font size for the whole app
        self.base_font_size = 16 
        
        self.init_ui()
        self.load_data()

    def format_latex_label(self, col_name):
        """
        Dynamically converts column names to LaTeX-style HTML.
        Example: 'a_c_e_mm_2' -> A<sub>e</sub> [mm<sup>2</sup>]
        """
        # 1. Handle Symbols
        text = col_name.lower()
        
        # Mapping common magnetic symbols
        symbol_map = {
            r'^l_e': '<i>l</i><sub>e</sub>',
            r'^a_c_e': '<i>A</i><sub>e</sub>',
            r'^a_c_min': '<i>A</i><sub>min</sub>',
            r'^v_c_e': '<i>V</i><sub>e</sub>',
            r'^l_t': '<i>l</i><sub>t</sub>',
            r'^a_w': '<i>A</i><sub>w</sub>',
            r'^l_n': '<i>l</i><sub>n</sub>',
            r'^i_t': '<i>I</i><sub>t</sub>',
            r'^i_n': '<i>I</i><sub>n</sub>',
        }
        
        found_symbol = False
        for pattern, latex in symbol_map.items():
            if re.search(pattern, text):
                text = re.sub(pattern, latex, text)
                found_symbol = True
                break
        
        if not found_symbol:
            # Fallback: Capitalize first letter and handle underscores as subscripts
            parts = text.split('_')
            text = f"<i>{parts[0].upper()}</i>" + (f"<sub>{'_'.join(parts[1:])}</sub>" if len(parts) > 1 else "")

        # 2. Handle Units (mm_2 -> [mm²], mm_3 -> [mm³])
        text = text.replace('mm_2', ' [mm<sup>2</sup>]')
        text = text.replace('mm_3', ' [mm<sup>3</sup>]')
        text = text.replace('_mm', ' [mm]')
        
        # 3. Clean up remaining underscores and prettify
        text = text.replace('_', ' ')
        return text

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(40, 40, 40, 40)

        # Header
        header = QLabel("Core Property Viewer")
        header.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header.setStyleSheet("color: #007acc; margin-bottom: 10px;")
        main_layout.addWidget(header)

        # Selector Area
        selector_label = QLabel("Select Core Type:")
        selector_label.setStyleSheet(f"font-size: {self.base_font_size}px; font-weight: bold;")
        main_layout.addWidget(selector_label)
        
        self.core_selector = QComboBox()
        self.core_selector.setStyleSheet(f"""
            QComboBox {{ 
                background-color: #252525; 
                border: 2px solid #3d3d3d; 
                padding: 10px; 
                border-radius: 6px; 
                font-size: {self.base_font_size}px;
                color: white;
            }}
            QComboBox::drop-down {{ border: none; }}
            QComboBox QAbstractItemView {{ background-color: #252525; selection-background-color: #007acc; }}
        """)
        self.core_selector.currentIndexChanged.connect(self.on_core_selected)
        main_layout.addWidget(self.core_selector)

        main_layout.addSpacing(30)

        # Scrollable Area for Details
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none; background-color: transparent;")
        
        self.details_widget = QWidget()
        self.details_container = QVBoxLayout(self.details_widget)
        self.details_container.setSpacing(10)
        scroll.setWidget(self.details_widget)
        main_layout.addWidget(scroll)
        
        self.placeholder = QLabel("Select a core to view detailed dimensions")
        self.placeholder.setStyleSheet(f"color: #666666; font-size: {self.base_font_size}px; font-style: italic;")
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.details_container.addWidget(self.placeholder)
        
        main_layout.addStretch()

    def load_data(self):
        file_path = "Cores-Info-Database.xlsx"
        if not os.path.exists(file_path):
            QMessageBox.critical(self, "Error", f"File '{file_path}' not found.")
            return

        try:
            self.df = pd.read_excel(file_path)
            
            # Robust Cleaning
            original_cols = self.df.columns.tolist()
            def clean_column_name(col):
                c = str(col).strip().lower()
                c = re.sub(r'[^a-z0-9]+', '_', c)
                return c.strip('_')

            self.df.columns = [clean_column_name(c) for c in self.df.columns]

            # Better Header Detection (Fixes 'core_type' error)
            core_col = None
            # Priority 1: Check for 'core' in name
            for c in self.df.columns:
                if 'core' in c:
                    core_col = c
                    break
            
            # Priority 2: Use the very first column if no 'core' word is found
            if not core_col and len(self.df.columns) > 0:
                core_col = self.df.columns[0]

            if core_col:
                self.df.rename(columns={core_col: 'core_type'}, inplace=True)
            else:
                raise KeyError("Spreadsheet is empty or has no readable columns.")

            core_names = self.df['core_type'].dropna().unique().tolist()
            self.core_selector.clear()
            self.core_selector.addItems(sorted([str(c) for c in core_names]))
            
        except Exception as e:
            QMessageBox.critical(self, "Data Error", f"Failed to load: {str(e)}")

    def add_stat_row(self, html_label, value):
        row = QFrame()
        row.setStyleSheet("""
            QFrame { 
                background-color: #1e1e1e; 
                border: 1px solid #333333; 
                border-radius: 8px; 
            }
            QFrame:hover { background-color: #252525; border-color: #007acc; }
        """)
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(20, 15, 20, 15)
        
        lbl = QLabel(html_label)
        lbl.setStyleSheet(f"color: #bbbbbb; font-size: {self.base_font_size + 2}px; border: none;")
        lbl.setTextFormat(Qt.TextFormat.RichText)
        
        val = QLabel(str(value))
        val.setStyleSheet(f"font-weight: bold; font-size: {self.base_font_size + 4}px; color: #4dabff; border: none;")
        
        row_layout.addWidget(lbl)
        row_layout.addStretch()
        row_layout.addWidget(val)
        self.details_container.addWidget(row)

    def on_core_selected(self):
        if self.df is None or self.core_selector.currentIndex() < 0:
            return

        selected_name = self.core_selector.currentText()
        
        # Clear previous widgets
        while self.details_container.count():
            child = self.details_container.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        row_data = self.df[self.df['core_type'] == selected_name].iloc[0]
        
        for col in self.df.columns:
            if col == 'core_type':
                continue
            
            # Use the universal LaTeX-style formatter
            html_label = self.format_latex_label(col)
            val = row_data[col]
            
            if pd.isna(val): val = "—"
            self.add_stat_row(html_label, val)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = CoreDimensionApp()
    window.show()
    sys.exit(app.exec())