import sys
import pandas as pd
import os
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QComboBox, QFrame, QMessageBox, QScrollArea)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QPixmap

class CoreDimensionApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Magnetic Core Explorer")
        self.setMinimumSize(950, 1000)
        self.setStyleSheet("background-color: #121212; color: #e0e0e0;")

        self.df = None
        self.core_col = None
        self.pic_col = None
        
        # UI Font Sizes
        self.title_font_size = 36
        self.label_font_size = 24
        self.value_font_size = 28
        
        self.init_ui()
        self.load_data()

    def format_latex_label(self, text):
        """Universal LaTeX formatter for magnetic core properties."""
        text = str(text)
        replacements = {
            r'l_e': '<i>l</i><sub>e</sub>',
            r'A_c,e': '<i>A</i><sub>e</sub>',
            r'A_c,min': '<i>A</i><sub>min</sub>',
            r'V_c,e': '<i>V</i><sub>e</sub>',
            r'l_t': '<i>l</i><sub>t</sub>',
            r'A_w': '<i>A</i><sub>w</sub>',
            r'l_N': '<i>l</i><sub>N</sub>',
            r'\[mm\^2\]': '[mm<sup>2</sup>]',
            r'\[mm\^2\|': '[mm<sup>2</sup>]', # Handles typo in screenshot
            r'\[mm\^3\]': '[mm<sup>3</sup>]',
            r'\[mm\]': '[mm]'
        }
        for pattern, sub in replacements.items():
            text = re.sub(pattern, sub, text)
        if '_' in text and '<' not in text:
            text = text.replace('_', ' ').title()
        return text

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(50, 50, 50, 50)

        title = QLabel("Core Property Browser")
        title.setFont(QFont("Arial", self.title_font_size, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #007acc; margin-bottom: 25px;")
        main_layout.addWidget(title)

        self.core_selector = QComboBox()
        self.core_selector.setStyleSheet(f"""
            QComboBox {{ 
                background-color: #252525; border: 3px solid #3d3d3d; 
                padding: 15px; border-radius: 12px; 
                font-size: {self.label_font_size}px; color: white;
            }}
            QComboBox QAbstractItemView {{ background-color: #252525; color: white; selection-background-color: #007acc; }}
        """)
        self.core_selector.currentIndexChanged.connect(self.on_core_selected)
        main_layout.addWidget(self.core_selector)

        main_layout.addSpacing(40)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none; background-color: transparent;")
        
        self.content_container = QWidget()
        self.layout = QVBoxLayout(self.content_container)
        self.layout.setSpacing(20)
        scroll.setWidget(self.content_container)
        main_layout.addWidget(scroll)

        self.image_display = QLabel()
        self.image_display.setFixedSize(850, 480)
        self.image_display.setAlignment(Qt.AlignCenter)
        self.image_display.setStyleSheet("background-color: #1e1e1e; border-radius: 20px; border: 2px solid #333;")
        self.layout.addWidget(self.image_display)

        self.stats_widget = QWidget()
        self.stats_layout = QVBoxLayout(self.stats_widget)
        self.stats_layout.setContentsMargins(0, 0, 0, 0)
        self.layout.addWidget(self.stats_widget)

    def load_data(self):
        file_path = "Cores-Info-Database.xlsx"
        try:
            self.df = pd.read_excel(file_path)
            # Find Core Type and Picture columns regardless of exact naming/case
            for col in self.df.columns:
                c_low = str(col).lower().strip()
                if ("core" in c_low or "type" in c_low) and not self.core_col:
                    self.core_col = col
                if ("picture" in c_low or "image" in c_low) and not self.pic_col:
                    self.pic_col = col

            if not self.core_col: self.core_col = self.df.columns[0]
            
            core_list = sorted(self.df[self.core_col].dropna().unique().astype(str))
            self.core_selector.addItems(core_list)
        except Exception as e:
            QMessageBox.critical(self, "Data Error", f"Could not read spreadsheet.\nCheck if 'Core Type' column exists.\nError: {str(e)}")

    def on_core_selected(self):
        if self.df is None or self.core_selector.currentIndex() < 0: return
        selected = self.core_selector.currentText()
        row = self.df[self.df[self.core_col].astype(str) == selected].iloc[0]

        # Update Image
        img_found = False
        if self.pic_col and not pd.isna(row[self.pic_col]):
            img_path = str(row[self.pic_col]).strip()
            if not os.path.exists(img_path):
                img_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), img_path)
            
            pix = QPixmap(img_path)
            if not pix.isNull():
                self.image_display.setPixmap(pix.scaled(820, 450, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                self.image_display.setText("")
                img_found = True
        
        if not img_found:
            self.image_display.setPixmap(QPixmap())
            self.image_display.setText("Image not found\nAdd filenames to 'Picture' column")
            self.image_display.setStyleSheet("background-color: #1e1e1e; color: #666; font-size: 20px; border-radius: 20px;")

        # Clear and Refill Stats
        while self.stats_layout.count():
            item = self.stats_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()

        for col in self.df.columns:
            if col in [self.core_col, self.pic_col]: continue
            val = row[col]
            frame = QFrame()
            frame.setStyleSheet("background-color: #1e1e1e; border: 1px solid #333; border-radius: 15px;")
            h = QHBoxLayout(frame)
            h.setContentsMargins(30, 20, 30, 20)
            
            name_lbl = QLabel(self.format_latex_label(col))
            name_lbl.setTextFormat(Qt.RichText)
            name_lbl.setStyleSheet(f"font-size: {self.label_font_size}px; color: #aaaaaa; border: none;")
            
            val_text = str(val) if not pd.isna(val) else "—"
            val_lbl = QLabel(val_text)
            val_lbl.setStyleSheet(f"font-size: {self.value_font_size}px; font-weight: bold; color: #4dabff; border: none;")
            
            h.addWidget(name_lbl)
            h.addStretch()
            h.addWidget(val_lbl)
            self.stats_layout.addWidget(frame)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = CoreDimensionApp()
    win.show()
    sys.exit(app.exec())