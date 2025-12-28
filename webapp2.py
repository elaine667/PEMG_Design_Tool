import sys
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QComboBox, QTextEdit, QLabel)
from PyQt5.QtCore import Qt

class MagneticCoreGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Magnetic Core Database")
        self.setGeometry(100, 100, 700, 600)
        
        self.df = pd.read_excel('Cores-Info-Database.xlsx')
        
        self.core_data = {}
        for idx, row in self.df.iterrows():
            core_name = str(row['Core Type'])
            self.core_data[core_name] = row.to_dict()
        
        self.init_ui()
    
    def init_ui(self):
        central_widget = QWidget()
        central_widget.setStyleSheet("background-color: #1a1a1a;")
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        label = QLabel("Select a Magnetic Core:")
        label.setStyleSheet("""
            font-family: 'SF Pro Display', 'Helvetica Neue', 'Arial';
            font-size: 18px;
            color: #e0e0e0;
            font-weight: 300;
        """)
        layout.addWidget(label)
        
        self.combo = QComboBox()
        self.combo.addItems(list(self.core_data.keys()))
        self.combo.setStyleSheet("""
            QComboBox {
                background-color: #2d2d2d;
                color: #ffffff;
                border: 1px solid #404040;
                border-radius: 6px;
                padding: 10px 15px;
                font-family: 'SF Pro Text', 'Helvetica Neue', 'Arial';
                font-size: 14px;
                font-weight: 300;
            }
            QComboBox:hover {
                border: 1px solid #606060;
            }
            QComboBox::drop-down {
                border: none;
                padding-right: 10px;
            }
            QComboBox QAbstractItemView {
                background-color: #2d2d2d;
                color: #ffffff;
                selection-background-color: #505050;
                selection-color: #ffffff;
                border: 1px solid #404040;
                outline: none;
            }
            QComboBox QAbstractItemView::item {
                padding: 8px 15px;
                min-height: 25px;
            }
            QComboBox QAbstractItemView::item:hover {
                background-color: #454545;
                color: #ffffff;
            }
            QComboBox QAbstractItemView::item:selected {
                background-color: #505050;
                color: #ffffff;
            }
        """)
        self.combo.currentTextChanged.connect(self.display_core_info)
        layout.addWidget(self.combo)
        
        self.info_display = QTextEdit()
        self.info_display.setReadOnly(True)
        self.info_display.setStyleSheet("""
            QTextEdit {
                background-color: #242424;
                color: #d0d0d0;
                border: 1px solid #333333;
                border-radius: 8px;
                padding: 20px;
                font-family: 'SF Mono', 'Monaco', 'Menlo', 'Courier New';
                font-size: 13px;
                line-height: 1.6;
            }
        """)
        layout.addWidget(self.info_display)
        
        central_widget.setLayout(layout)
    
    def display_core_info(self, core_name):
        if core_name in self.core_data:
            core_info = self.core_data[core_name]
            
            text = f"{core_name}\n"
            text += "─" * 60 + "\n\n"
            
            for key, value in core_info.items():
                text += f"{key:<25} {value}\n"
            
            self.info_display.setText(text)

def main():
    app = QApplication(sys.argv)
    window = MagneticCoreGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
