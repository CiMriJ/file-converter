"""界面配色与样式表。"""

ACCENT = "#3b6ef6"
ACCENT_DARK = "#2f5fe0"
TEXT = "#1c2b45"
MUTED = "#8a97ad"
GREEN = "#18a058"
RED = "#d03050"

QSS = """
QWidget {
    color: #1c2b45;
    font-family: "Microsoft YaHei UI", "Segoe UI", "PingFang SC";
    font-size: 13px;
}
QWidget#Root { background: #f5f7fb; }

QFrame#Card {
    background: #ffffff;
    border: 1px solid #e7ebf5;
    border-radius: 12px;
}

QLabel#HeaderTitle { font-size: 21px; font-weight: 600; color: #16233c; }
QLabel#HeaderSub { color: #8a97ad; font-size: 12px; }
QLabel#SectionTitle { font-size: 13px; font-weight: 600; color: #24344f; }
QLabel#Hint { color: #8a97ad; font-size: 12px; }
QLabel#Badge {
    background: #eef3ff; color: #2d5ce0; border-radius: 9px;
    padding: 3px 10px; font-size: 12px; font-weight: 600;
}

QPushButton {
    background: #ffffff;
    border: 1px solid #dde3f0;
    border-radius: 8px;
    padding: 7px 14px;
    color: #34445f;
}
QPushButton:hover { border-color: #b9c6e6; background: #f7f9ff; }
QPushButton:pressed { background: #eef2fd; }
QPushButton:disabled { color: #b3bccd; border-color: #e8ecf5; background: #fafbfe; }

QPushButton#Primary {
    background: #3b6ef6; border: none; color: #ffffff;
    font-weight: 600; padding: 10px 24px;
}
QPushButton#Primary:hover { background: #2f5fe0; }
QPushButton#Primary:pressed { background: #2754c9; }
QPushButton#Primary:disabled { background: #b9c6e2; color: #ffffff; }

QPushButton#Ghost {
    background: transparent; border: none; color: #5b6b88; padding: 6px 10px;
}
QPushButton#Ghost:hover { background: #eef2fb; color: #2d5ce0; }

QLineEdit, QSpinBox, QComboBox {
    background: #ffffff;
    border: 1px solid #dfe5f1;
    border-radius: 8px;
    padding: 7px 10px;
    min-height: 18px;
    selection-background-color: #cfe0ff;
    selection-color: #1c2b45;
}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus { border-color: #7d9ff9; }
QLineEdit:disabled, QSpinBox:disabled { background: #f7f9fd; color: #a3aec3; }
QComboBox::drop-down { border: none; width: 24px; }
QComboBox QAbstractItemView {
    background: #ffffff;
    border: 1px solid #e0e6f2;
    border-radius: 8px;
    padding: 4px;
    outline: 0;
    selection-background-color: #eef3ff;
    selection-color: #1c2b45;
}

QTableWidget {
    background: #ffffff;
    border: none;
    gridline-color: #f0f3fa;
    selection-background-color: #eef3ff;
    selection-color: #1c2b45;
}
QHeaderView::section {
    background: #fafbfe;
    border: none;
    border-bottom: 1px solid #eef1f8;
    padding: 9px 10px;
    color: #7a889f;
    font-weight: 600;
}
QTableWidget::item { padding: 6px 10px; border-bottom: 1px solid #f5f7fc; }

QProgressBar {
    border: none; background: #eaeef9; border-radius: 5px;
    height: 10px; text-align: center; color: transparent;
}
QProgressBar::chunk { background: #3b6ef6; border-radius: 5px; }

QTabWidget::pane { border: none; background: transparent; }
QTabBar::tab {
    background: transparent; padding: 9px 20px; margin-right: 6px;
    border-radius: 9px; color: #6b7a94; font-weight: 600;
}
QTabBar::tab:hover { background: #eef2fb; color: #34445f; }
QTabBar::tab:selected { background: #e8efff; color: #2d5ce0; }

QListWidget {
    background: #ffffff;
    border: 1px solid #e7ebf5;
    border-radius: 10px;
    padding: 4px;
    outline: 0;
}
QListWidget::item { padding: 7px 9px; border-radius: 6px; color: #34445f; }
QListWidget::item:hover { background: #f6f8fe; }
QListWidget::item:selected { background: #eef3ff; color: #1c2b45; }

QTextEdit {
    background: #fafbfe;
    border: 1px solid #e7ebf5;
    border-radius: 10px;
    padding: 8px;
    color: #44546f;
    font-family: "Consolas", "Microsoft YaHei UI";
    font-size: 12px;
}

QCheckBox { spacing: 7px; color: #3d4d68; }
QCheckBox::indicator {
    width: 16px; height: 16px; border-radius: 4px;
    border: 1px solid #cfd8ea; background: #ffffff;
}
QCheckBox::indicator:checked { background: #3b6ef6; border-color: #3b6ef6; }
QCheckBox::indicator:hover { border-color: #7d9ff9; }

QSlider::groove:horizontal { height: 5px; background: #e6ebf7; border-radius: 3px; }
QSlider::handle:horizontal {
    background: #3b6ef6; width: 14px; height: 14px;
    margin: -5px 0; border-radius: 7px;
}
QSlider::handle:horizontal:hover { background: #2f5fe0; }

QFrame#DropArea {
    background: #fbfcff;
    border: 2px dashed #c9d6f2;
    border-radius: 12px;
}
QFrame#DropArea[drag="true"] {
    background: #eef4ff;
    border-color: #6f93f5;
}

QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
QScrollBar::handle:vertical { background: #d5dcec; border-radius: 5px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #bcc7e0; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: transparent; height: 10px; margin: 2px; }
QScrollBar::handle:horizontal { background: #d5dcec; border-radius: 5px; min-width: 30px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }

QToolTip {
    background: #24344f; color: #ffffff; border: none;
    padding: 6px 8px; border-radius: 6px;
}
"""
