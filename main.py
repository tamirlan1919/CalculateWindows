import sys
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QStackedWidget, QTableWidget, QTableWidgetItem, QComboBox


class WindowSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.loadPrices()
        self.window_type = None  # Инициализация атрибута для хранения типа окна

    def initUI(self):
        self.setWindowTitle('Калькулятор окон')
        self.setGeometry(100, 100, 500, 300)

        main_layout = QVBoxLayout(self)

        self.stacked_widget = QStackedWidget(self)

        self.selection_screen = QWidget()
        self.selection_layout = QVBoxLayout(self.selection_screen)

        self.single_button = QPushButton("Стекла")
        self.single_button.clicked.connect(self.showSize)

        self.double_button = QPushButton("Двери")
        self.double_button.clicked.connect(self.doors)

        self.st = QPushButton("Расчет стеклопакета")
        self.st.clicked.connect(self.priceSt)

        self.selection_layout.addWidget(self.single_button)
        self.selection_layout.addWidget(self.double_button)
        self.selection_layout.addWidget(self.st)

        self.stacked_widget.addWidget(self.selection_screen)

        main_layout.addWidget(self.stacked_widget)

        self.back_button = QPushButton('Назад')
        self.back_button.clicked.connect(self.goBack)
        self.back_button.hide()

        back_layout = QHBoxLayout()
        back_layout.addWidget(self.back_button)
        main_layout.addLayout(back_layout)

        self.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                border: none;
                color: white;
                padding: 10px 24px;
                text-align: center;
                text-decoration: none;
                display: inline-block;
                font-size: 16px;
                margin: 4px 2px;
                cursor: pointer;
                border-radius: 12px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QLineEdit, QComboBox {
                padding: 5px;
                font-size: 16px;
                border-radius: 6px;
                border: 1px solid #ccc;
            }
            QLabel {
                font-size: 16px;
            }
        """)

        self.show()

    def doors(self):
        selection_screen = QWidget()
        selection_layout = QVBoxLayout(selection_screen)

        single_button = QPushButton('Одностворчатая дверь')
        single_button.clicked.connect(self.singleDoor)  # Передача типа окна

        double_button = QPushButton('Двухстворчатая дверь')
        double_button.clicked.connect(self.doubleDoor)  # Передача типа окна

        selection_layout.addWidget(single_button)
        selection_layout.addWidget(double_button)

        self.stacked_widget.addWidget(selection_screen)
        self.selection_screen_priceSt = selection_screen

        self.stacked_widget.setCurrentWidget(selection_screen)

        self.back_button.show()

    def doubleDoor(self):
        if not hasattr(self, 'input_screen_priceDoor'):
            self.input_screen_priceDoor = QWidget()
            self.input_layout_priceDoor = QVBoxLayout(self.input_screen_priceDoor)

            self.height_label_priceDoor = QLabel('Высота (м):')
            self.height_lineedit_priceDoor = QLineEdit()

            self.width_label_priceDoor = QLabel('Ширина (м):')
            self.width_lineedit_priceDoor = QLineEdit()

            self.left_label_priceDoor = QLabel('Левая дверь (мм):')
            self.left_lineedit_priceDoor = QLineEdit()

            self.right_label_priceDoor = QLabel('Правая дверь (мм):')
            self.right_lineedit_priceDoor = QLineEdit()

            self.calculate_button_priceDoor = QPushButton('Рассчитать')
            self.calculate_button_priceDoor.clicked.connect(self.CalculateDoubleDoor)

            self.input_layout_priceDoor.addWidget(self.height_label_priceDoor)
            self.input_layout_priceDoor.addWidget(self.height_lineedit_priceDoor)
            self.input_layout_priceDoor.addWidget(self.width_label_priceDoor)
            self.input_layout_priceDoor.addWidget(self.width_lineedit_priceDoor)
            self.input_layout_priceDoor.addWidget(self.left_label_priceDoor)
            self.input_layout_priceDoor.addWidget(self.left_lineedit_priceDoor)
            self.input_layout_priceDoor.addWidget(self.right_label_priceDoor)
            self.input_layout_priceDoor.addWidget(self.right_lineedit_priceDoor)

            self.input_layout_priceDoor.addWidget(self.calculate_button_priceDoor)

            self.stacked_widget.addWidget(self.input_screen_priceDoor)

        self.stacked_widget.removeWidget(self.selection_screen_priceSt)
        self.stacked_widget.setCurrentWidget(self.input_screen_priceDoor)
        self.back_button.show()

    def CalculateDoubleDoor(self):
        try:
            height = float(self.height_lineedit_priceDoor.text())
            width = float(self.width_lineedit_priceDoor.text())
            left = float(self.left_lineedit_priceDoor.text())
            right = float(self.right_lineedit_priceDoor.text())

            X1 = left - 0.21
            X2 = right - 0.21
            Y1 = height - 0.268
            result_text = f"Двухстворчатая дверь\n{X1} x {Y1} * {1} пакета\n{X2} * {Y1} * 1 пакета"
            self.showResult(result_text)
        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")

    def singleDoor(self):
        if not hasattr(self, 'input_screen_priceDoor_One'):
            self.input_screen_priceDoor_One = QWidget()
            self.input_layout_priceDoor_One = QVBoxLayout(self.input_screen_priceDoor_One)

            self.height_label_priceDoor_One = QLabel('Высота (м):')
            self.height_lineedit_priceDoor_One = QLineEdit()

            self.width_label_priceDoor_One = QLabel('Ширина (м):')
            self.width_lineedit_priceDoor_One = QLineEdit()

            self.calculate_button_priceDoor_One = QPushButton('Рассчитать')
            self.calculate_button_priceDoor_One.clicked.connect(self.calculateSinleDoor)

            self.input_layout_priceDoor_One.addWidget(self.height_label_priceDoor_One)
            self.input_layout_priceDoor_One.addWidget(self.height_lineedit_priceDoor_One)
            self.input_layout_priceDoor_One.addWidget(self.width_label_priceDoor_One)
            self.input_layout_priceDoor_One.addWidget(self.width_lineedit_priceDoor_One)
            self.input_layout_priceDoor_One.addWidget(self.calculate_button_priceDoor_One)

            self.stacked_widget.addWidget(self.input_screen_priceDoor_One)

        self.stacked_widget.removeWidget(self.selection_screen_priceSt)
        self.stacked_widget.setCurrentWidget(self.input_screen_priceDoor_One)
        self.back_button.show()

    def calculateSinleDoor(self):
        try:
            height = float(self.height_lineedit_priceDoor_One.text())
            width = float(self.width_lineedit_priceDoor_One.text())
            X1 = width - 0.29
            Y1 = height - 0.268
            result_text = f"Одностворчатая дверь\n{X1} x {Y1} * {1} пакета"
            self.showResult(result_text)
        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")

    def priceSt(self):
        selection_screen = QWidget()
        selection_layout = QVBoxLayout(selection_screen)

        single_button = QPushButton('Одностворчатое окно')
        single_button.clicked.connect(lambda: self.calcSingleSize('Одностворчатое окно'))  # Передача типа окна

        double_button = QPushButton('Двухстворчатое окно')
        double_button.clicked.connect(lambda: self.showSingleSize('Двухстворчатое окно'))  # Передача типа окна

        triple_button = QPushButton('Трехстворчатое окно')
        triple_button.clicked.connect(lambda: self.showThreeSize('Трехстворчатое окно'))  # Передача типа окна

        selection_layout.addWidget(single_button)
        selection_layout.addWidget(double_button)
        selection_layout.addWidget(triple_button)

        self.stacked_widget.addWidget(selection_screen)
        self.selection_screen_priceSt = selection_screen
        self.stacked_widget.setCurrentWidget(selection_screen)

        self.back_button.show()

    def calcSingleSize(self, window_type):
        self.window_type = window_type  # Сохранение типа окна
        if not hasattr(self, 'input_screen_priceSt_One'):
            self.input_screen_priceSt_One = QWidget()
            self.input_layout_priceSt_One = QVBoxLayout(self.input_screen_priceSt_One)

            self.height_label_window_priceSt_One = QLabel('Высота (м):')
            self.height_lineedit_window_priceSt_One = QLineEdit()

            self.width_label_window_priceSt_One = QLabel('Ширина (м):')
            self.width_lineedit_window_priceSt_One = QLineEdit()

            self.calculate_button_window_priceSt_One = QPushButton('Рассчитать')
            self.calculate_button_window_priceSt_One.clicked.connect(self.calculateWindowSingle)

            self.input_layout_priceSt_One.addWidget(self.height_label_window_priceSt_One)
            self.input_layout_priceSt_One.addWidget(self.height_lineedit_window_priceSt_One)
            self.input_layout_priceSt_One.addWidget(self.width_label_window_priceSt_One)
            self.input_layout_priceSt_One.addWidget(self.width_lineedit_window_priceSt_One)
            self.input_layout_priceSt_One.addWidget(self.calculate_button_window_priceSt_One)

            self.stacked_widget.addWidget(self.input_screen_priceSt_One)

        self.stacked_widget.setCurrentWidget(self.input_screen_priceSt_One)
        self.back_button.show()

    def calculateWindowSingle(self):
        try:
            height = float(self.height_lineedit_window_priceSt_One.text())
            width = float(self.width_lineedit_window_priceSt_One.text())
            X1 = width - 0.1 - 0.124
            Y1 = height - 0.228
            result_text = f"Одностворчатое окно\n{X1} x {Y1} * {1} пакета"
            self.showResult(result_text)
        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")

    def showSingleSize(self, window_type):
        self.window_type = window_type  # Сохранение типа окна
        if not hasattr(self, 'input_screen_priceSt'):
            self.input_screen_priceSt = QWidget()
            self.input_layout_priceSt = QVBoxLayout(self.input_screen_priceSt)

            self.height_label_window_priceSt = QLabel('Высота (м):')
            self.height_lineedit_window_priceSt = QLineEdit()

            self.width_label_window_priceSt = QLabel('Ширина (м):')
            self.width_lineedit_window_priceSt = QLineEdit()

            self.spacer_size_label_window_priceSt = QLabel('Шпатик (мм):')
            self.spacer_size_lineedit_window_priceSt = QLineEdit()

            self.calculate_button_window_priceSt = QPushButton('Рассчитать')
            self.calculate_button_window_priceSt.clicked.connect(self.calculateWindow)

            self.input_layout_priceSt.addWidget(self.height_label_window_priceSt)
            self.input_layout_priceSt.addWidget(self.height_lineedit_window_priceSt)
            self.input_layout_priceSt.addWidget(self.width_label_window_priceSt)
            self.input_layout_priceSt.addWidget(self.width_lineedit_window_priceSt)
            self.input_layout_priceSt.addWidget(self.spacer_size_label_window_priceSt)
            self.input_layout_priceSt.addWidget(self.spacer_size_lineedit_window_priceSt)
            self.input_layout_priceSt.addWidget(self.calculate_button_window_priceSt)

            self.stacked_widget.addWidget(self.input_screen_priceSt)

        self.stacked_widget.removeWidget(self.selection_screen_priceSt)
        self.stacked_widget.setCurrentWidget(self.input_screen_priceSt)
        self.back_button.show()


    def showThreeSize(self, window_type):
        self.window_type = window_type  # Сохранение типа окна
        if not hasattr(self, 'input_screen_priceSt_Three'):
            self.input_screen_priceSt_Three  = QWidget()
            self.input_layout_priceSt_Three  = QVBoxLayout(self.input_screen_priceSt_Three)

            self.height_label_window_priceSt_Three = QLabel('Высота (м):')
            self.height_lineedit_window_priceSt_Three  = QLineEdit()

            self.width_label_window_priceSt_Three = QLabel('Ширина (м):')
            self.width_lineedit_window_priceSt_Three = QLineEdit()

            self.spacer_size_label_window_priceSt_Three = QLabel('Шпатик (мм):')
            self.spacer_size_lineedit_window_priceSt_Three = QLineEdit()

            self.calculate_button_window_priceSt_Three = QPushButton('Рассчитать')
            self.calculate_button_window_priceSt_Three.clicked.connect(self.calculateThreeWindow)

            self.input_layout_priceSt_Three.addWidget(self.height_label_window_priceSt_Three)
            self.input_layout_priceSt_Three.addWidget(self.height_lineedit_window_priceSt_Three)
            self.input_layout_priceSt_Three.addWidget(self.width_label_window_priceSt_Three)
            self.input_layout_priceSt_Three.addWidget(self.width_lineedit_window_priceSt_Three)
            self.input_layout_priceSt_Three.addWidget(self.spacer_size_label_window_priceSt_Three)
            self.input_layout_priceSt_Three.addWidget(self.spacer_size_lineedit_window_priceSt_Three)
            self.input_layout_priceSt_Three.addWidget(self.calculate_button_window_priceSt_Three)

            self.stacked_widget.addWidget(self.input_screen_priceSt_Three)

        self.stacked_widget.setCurrentWidget(self.input_screen_priceSt_Three)
        self.back_button.show()

    def calculateWindow(self):
        try:
            height = float(self.height_lineedit_window_priceSt.text())
            width = float(self.width_lineedit_window_priceSt.text())
            spacer_size = float(self.spacer_size_lineedit_window_priceSt.text())

            # Ширина рамы, импоста и высота рамы (в м)
            frame_width = 0.05
            mullion_width = 0.047
            count_impost = 1
            first_pack = 1
            second_pack = 1
            # В зависимости от типа окна выбираем количество импостов


            # Рассчитываем размеры другого окна
            leader_window = width - ((count_impost * mullion_width) + (frame_width+frame_width))
            leader_window -= spacer_size
            leader_window /= count_impost

            # Рассчитываем размеры пакета стекла
            X1 = leader_window - 0.01
            Y1 = height - 0.05*2
            Y1 -= 0.01

            X2 = spacer_size - 0.124
            Y2 = height - 0.228

            # Показываем результат
            result_text = f"Двухстворчатое оконо\n{X1} x {Y1} * {first_pack} пакета\n{X2} x {Y2} * {second_pack} пакета"
            self.showResult(result_text)
        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")

    def calculateThreeWindow(self):
        try:
            height = float(self.height_lineedit_window_priceSt_Three.text())
            width = float(self.width_lineedit_window_priceSt_Three.text())
            spacer_size = float(self.spacer_size_lineedit_window_priceSt_Three.text())

            # Ширина рамы, импоста и высота рамы (в м)
            frame_width = 0.05
            mullion_width = 0.047
            count_impost = 2
            first_pack = 2
            second_pack = 1


            # Рассчитываем размеры другого окна
            leader_window = width - ((count_impost * mullion_width) + (frame_width + frame_width))
            leader_window -= spacer_size
            leader_window /= count_impost

            # Рассчитываем размеры пакета стекла
            X1 = leader_window - 0.01
            Y1 = height - 0.05 * 2
            Y1 -= 0.01

            X2 = spacer_size - 0.124
            Y2 = height - 0.228

            # Показываем результат
            result_text = f"Трехстворчатое окно\n{X1} x {Y1} * {first_pack} пакета\n{X2} x {Y2} * {second_pack} пакета"
            self.showResult(result_text)
        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")

    def showResult(self, result_text):
        # Выводим результат на экран
        result_label = QLabel(result_text)
        self.stacked_widget.addWidget(result_label)
        self.stacked_widget.removeWidget(self.selection_screen_priceSt)
        self.stacked_widget.setCurrentWidget(result_label)


    def showSize(self):
        if not hasattr(self, 'input_screen'):
            self.input_screen = QWidget()
            self.input_layout = QVBoxLayout(self.input_screen)

            self.height_label = QLabel('Высота (м):')
            self.height_lineedit = QLineEdit()

            self.width_label = QLabel('Ширина (м):')
            self.width_lineedit = QLineEdit()

            self.calculate_button = QPushButton('Рассчитать')
            self.calculate_button.clicked.connect(self.calculateCost)

            self.input_layout.addWidget(self.height_label)
            self.input_layout.addWidget(self.height_lineedit)
            self.input_layout.addWidget(self.width_label)
            self.input_layout.addWidget(self.width_lineedit)

            self.input_layout.addWidget(self.calculate_button)

            self.stacked_widget.addWidget(self.input_screen)

        self.stacked_widget.setCurrentWidget(self.input_screen)
        self.back_button.show()

    def goBack(self):
        current_index = self.stacked_widget.currentIndex()
        next_index = 0  # Индекс начального экрана

        if current_index > 1:
            if hasattr(self, 'height_lineedit'):
                self.height_lineedit.clear()
            if hasattr(self, 'width_lineedit'):
                self.width_lineedit.clear()
            if hasattr(self, 'height_lineedit_window_priceSt'):
                self.height_lineedit_window_priceSt.clear()
            if hasattr(self, 'width_lineedit_window_priceSt'):
                self.width_lineedit_window_priceSt.clear()
            if hasattr(self, 'spacer_size_lineedit_window_priceSt'):
                self.spacer_size_lineedit_window_priceSt.clear()
            if hasattr(self, 'height_lineedit_priceDoor'):
                self.height_lineedit_priceDoor.clear()
            if hasattr(self, 'width_lineedit_priceDoor'):
                self.width_lineedit_priceDoor.clear()

        self.stacked_widget.setCurrentIndex(next_index)

        if next_index == 0:
            self.back_button.hide()
    def calculateCost(self):
        try:
            height = float(self.height_lineedit.text())
            width = float(self.width_lineedit.text())

            frame_cost = self.frame_price * (2 * (height + width))
            mullion_cost = self.mullion_price * (2 * (height + width))
            sash_cost = self.sash_price * height

            total_cost = frame_cost + mullion_cost + sash_cost

            self.showCostTable(frame_cost, mullion_cost, sash_cost, total_cost)
        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")

    def showCostTable(self, frame_cost, mullion_cost, sash_cost, total_cost):
        # Удаляем предыдущий экземпляр таблицы, если он существует
        if hasattr(self, 'result_table'):
            self.result_table.deleteLater()

        self.result_table = QTableWidget()
        self.result_table.setRowCount(4)
        self.result_table.setColumnCount(6)

        headers = ['N', 'Товар', 'Кол-во', 'Единица', 'Сумма', 'Цвет']
        self.result_table.setHorizontalHeaderLabels(headers)

        items = [
            ('1', 'Рама', f'{2 * (float(self.height_lineedit.text()) + float(self.width_lineedit.text()))}', 'м',
             f'{frame_cost} руб.', ''),
            ('2', 'Импост', f'{2 * (float(self.height_lineedit.text()) + float(self.width_lineedit.text()))}', 'м',
             f'{mullion_cost} руб.', ''),
            ('3', 'Створка', f'{float(self.height_lineedit.text())}', 'м', f'{sash_cost} руб.', ''),
            ('', '', '', 'Итого:', f'{total_cost} руб.', '')
        ]

        for row, item in enumerate(items):
            for col, data in enumerate(item):
                if col == 5:
                    combobox = QComboBox()
                    combobox.addItems(['Белый', 'Комбинированный', 'Цельный'])
                    self.result_table.setCellWidget(row - 1, col, combobox)
                else:
                    self.result_table.setItem(row, col, QTableWidgetItem(data))

        self.result_table.resizeColumnsToContents()
        self.result_table.setWindowTitle('Результат расчета стоимости')
        self.result_table.show()

        self.stacked_widget.addWidget(self.result_table)
        self.stacked_widget.setCurrentWidget(self.result_table)

    def loadPrices(self):

        wb = openpyxl.load_workbook('file.xlsx')
        sheet = wb.active
        print(sheet['A2'].value)
        self.frame_price = float(sheet['A2'].value)
        self.mullion_price = float(sheet['B2'].value)
        self.sash_price = float(sheet['C2'].value)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WindowSelector()
    sys.exit(app.exec_())