APP_STYLE = """
QWidget {
    font-family: "PingFang SC", "Microsoft YaHei", "Noto Sans CJK SC", sans-serif;
    color: #1F2937;
    font-size: 14px;
}

QMainWindow, #root {
    background: #F3F5F8;
}

QScrollArea, QScrollArea::viewport {
    background: #F3F5F8;
    border: none;
}

QLabel#appTitle {
    font-size: 28px;
    font-weight: 700;
    color: #111827;
}

QLabel#appSubtitle {
    font-size: 15px;
    color: #374151;
}

QLabel#pageTitle {
    font-size: 22px;
    font-weight: 700;
    color: #111827;
}

QLabel#pageHint {
    font-size: 13px;
    color: #374151;
}

QLabel#muted {
    color: #374151;
}

QLabel#fileInfo {
    color: #111827;
    font-size: 14px;
    font-weight: 600;
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
    padding: 0 18px;
    font-size: 14px;
    margin: 0;
}

QPushButton#primary {
    background-color: #2563EB;
    color: white;
    font-weight: 600;
    border: 1px solid #2563EB;
    min-width: 108px;
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
    color: #111827;
    min-height: 36px;
}

QCheckBox {
    spacing: 8px;
    color: #111827;
}

QProgressBar {
    border: none;
    background: #E5E7EB;
    border-radius: 8px;
    height: 16px;
    text-align: center;
    color: #111827;
    font-size: 11px;
}

QProgressBar::chunk {
    background: #2563EB;
    border-radius: 8px;
}

QFrame#progressPanel {
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 12px;
}

QLabel#progressLabel {
    color: #111827;
    font-size: 13px;
    font-weight: 600;
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
