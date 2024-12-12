from PyQt5.QtWidgets import QApplication, QMainWindow
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import sys

class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)
        super(MplCanvas, self).__init__(fig)

class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Matplotlib with PyQt5")
        canvas = MplCanvas(self, width=5, height=4, dpi=100)
        self.setCentralWidget(canvas)
        canvas.axes.plot([0, 1, 2, 3, 4], [10, 1, 20, 3, 40])  # Contoh plot sederhana

app = QApplication(sys.argv)
window = MainApp()
window.show()
sys.exit(app.exec_())
