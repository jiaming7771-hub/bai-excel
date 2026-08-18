APP_STYLE = """
QWidget {
    font-family: "PingFang SC", "Microsoft YaHei", "Noto Sans CJK SC", sans-serif;
    color: #1F2937;
    font-size: 14px;
}

QMainWindow, #root {
    background: #F3F5F8;
}

QLabel#appTitle {
    font-size: 28px;
    font-weight: 700;
    color: #111827;
}

QLabel#appSubtitle {
    font-size: 15px;
    color: #6B7280;
}

QLabel#pageTitle {
    font-size: 22px;
    font-weight: 700;
    color: #111827;
}

QLabel#pageHint {
    font-size: 13px;
    color: #6B7280;
}

QLabel#muted {
    color: #6B7280;
}

QFrame#card {
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 16px;
}

QFrame#dropZone {
    background: #F8FAFC;
    border: 2px dashed #CBD5E1;
    border-radius: 14px;
}

QFrame#dropZone[active="true"] {
    background: #EFF6FF;
    border: 2px dashed #2563EB;
}

QPushButton {
    border: none;
    border-radius: 10px;
    padding: 10px 18px;
    font-size: 14px;
}

QPushButton#primary {
    background: #2563EB;
    color: white;
    font-weight: 600;
    min-height: 40px;
}

QPushButton#primary:hover {
    background: #1D4ED8;
}

QPushButton#primary:disabled {
    background: #93C5FD;
}

QPushButton#secondary {
    background: #EEF2FF;
    color: #1D4ED8;
    font-weight: 600;
}

QPushButton#ghost {
    background: transparent;
    color: #4B5563;
}

QPushButton#ghost:hover {
    background: #E5E7EB;
}

QComboBox, QLineEdit {
    border: 1px solid #D1D5DB;
    border-radius: 10px;
    padding: 8px 12px;
    background: white;
    min-height: 36px;
}

QCheckBox {
    spacing: 8px;
}

QProgressBar {
    border: none;
    background: #E5E7EB;
    border-radius: 8px;
    height: 14px;
    text-align: center;
}

QProgressBar::chunk {
    background: #2563EB;
    border-radius: 8px;
}

QListWidget {
    border: 1px solid #E5E7EB;
    border-radius: 10px;
    background: white;
    padding: 6px;
}

QFrame#errorBox {
    background: #FEF2F2;
    border: 1px solid #FECACA;
    border-radius: 12px;
}

QLabel#errorTitle {
    color: #B91C1C;
    font-weight: 700;
}

QFrame#okBox {
    background: #ECFDF5;
    border: 1px solid #A7F3D0;
    border-radius: 12px;
}

QLabel#okTitle {
    color: #047857;
    font-weight: 700;
}
"""
