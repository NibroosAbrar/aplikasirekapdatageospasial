import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QWidget, QTableWidgetItem, QMessageBox, QMainWindow, QApplication, QVBoxLayout

excel_file_path = "D:/ipb capekk/MSIB/02. Projeck Aplikasi Rekap Data/04. Data Grafik/Data Buat Nibroos 2.xlsx"

def load_chart_data(self):
    """Membaca data dari file Excel untuk bar chart dan mengisi ComboBox Provinsi"""
    try:
        self.chart_data = pd.read_excel(self.excel_file_path, sheet_name=0)
    except Exception as e:
        QMessageBox.critical(self, "Error", f"Gagal memuat file Excel untuk chart: {e}")