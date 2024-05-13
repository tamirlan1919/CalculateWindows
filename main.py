import sys
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, \
    QStackedWidget, QTableWidget, QTableWidgetItem, QComboBox, QMessageBox


class WindowSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.loadPrices()
        self.costs_dict = {}
        self.window_type = None  # Инициализация атрибута для хранения типа окна

    def initUI(self):
        self.setWindowTitle('Калькулятор окон')
        self.setGeometry(100, 100, 500, 300)

        main_layout = QVBoxLayout(self)

        self.stacked_widget = QStackedWidget(self)

        self.selection_screen = QWidget()
        self.selection_layout = QVBoxLayout(self.selection_screen)

        self.single_button = QPushButton("Себестоимость")
        self.single_button.clicked.connect(self.choose_window)

        self.single_button_st = QPushButton("Стекла")
        self.single_button_st.clicked.connect(self.showStekla)



        self.selection_layout.addWidget(self.single_button)
        self.selection_layout.addWidget(self.single_button_st)


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

            self.height_label_priceDoor = QLabel('Высота (мм):')
            self.height_lineedit_priceDoor = QLineEdit()

            self.width_label_priceDoor = QLabel('Ширина (мм):')
            self.width_lineedit_priceDoor = QLineEdit()

            self.left_label_priceDoor = QLabel('Левая дверь (мм):')
            self.left_lineedit_priceDoor = QLineEdit()

            self.right_label_priceDoor = QLabel('Правая дверь (мм):')
            self.right_lineedit_priceDoor = QLineEdit()

            self.calculate_button_priceDoor = QPushButton('Рассчитать')
            self.calculate_button_priceDoor.clicked.connect(self.CalculateDoubleDoor)
            self.input_layout_priceDoor.addWidget(self.width_label_priceDoor)
            self.input_layout_priceDoor.addWidget(self.width_lineedit_priceDoor)
            self.input_layout_priceDoor.addWidget(self.height_label_priceDoor)
            self.input_layout_priceDoor.addWidget(self.height_lineedit_priceDoor)

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

            X1 = left - 210
            X2 = right - 210
            Y1 = height - 268
            result_text = f"Двухстворчатая дверь\n{X1} x {Y1} * {1} пакета\n{X2} * {Y1} * 1 пакета"
            self.showResult(result_text)
        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")

    def singleDoor(self):
        if not hasattr(self, 'input_screen_priceDoor_One'):
            self.input_screen_priceDoor_One = QWidget()
            self.input_layout_priceDoor_One = QVBoxLayout(self.input_screen_priceDoor_One)

            self.height_label_priceDoor_One = QLabel('Высота (мм):')
            self.height_lineedit_priceDoor_One = QLineEdit()

            self.width_label_priceDoor_One = QLabel('Ширина (мм):')
            self.width_lineedit_priceDoor_One = QLineEdit()

            self.calculate_button_priceDoor_One = QPushButton('Рассчитать')
            self.calculate_button_priceDoor_One.clicked.connect(self.calculateSinleDoor)
            self.input_layout_priceDoor_One.addWidget(self.width_label_priceDoor_One)
            self.input_layout_priceDoor_One.addWidget(self.width_lineedit_priceDoor_One)
            self.input_layout_priceDoor_One.addWidget(self.height_label_priceDoor_One)
            self.input_layout_priceDoor_One.addWidget(self.height_lineedit_priceDoor_One)

            self.input_layout_priceDoor_One.addWidget(self.calculate_button_priceDoor_One)

            self.stacked_widget.addWidget(self.input_screen_priceDoor_One)

        self.stacked_widget.removeWidget(self.selection_screen_priceSt)
        self.stacked_widget.setCurrentWidget(self.input_screen_priceDoor_One)
        self.back_button.show()

    def calculateSinleDoor(self):
        try:
            height = float(self.height_lineedit_priceDoor_One.text())
            width = float(self.width_lineedit_priceDoor_One.text())
            X1 = width - 290
            Y1 = height - 268
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
        if not hasattr(self, 'uuu'):
            self.input_screen_priceSt_One = QWidget()
            self.input_layout_priceSt_One = QVBoxLayout(self.input_screen_priceSt_One)

            self.height_label_window_priceSt_One = QLabel('Высота (мм):')
            self.height_lineedit_window_priceSt_One = QLineEdit()

            self.width_label_window_priceSt_One = QLabel('Ширина (мм):')
            self.width_lineedit_window_priceSt_One = QLineEdit()

            self.calculate_button_window_priceSt_One = QPushButton('Рассчитать')
            self.calculate_button_window_priceSt_One.clicked.connect(self.calculateWindowSingle)
            self.input_layout_priceSt_One.addWidget(self.width_label_window_priceSt_One)
            self.input_layout_priceSt_One.addWidget(self.width_lineedit_window_priceSt_One)
            self.input_layout_priceSt_One.addWidget(self.height_label_window_priceSt_One)
            self.input_layout_priceSt_One.addWidget(self.height_lineedit_window_priceSt_One)

            self.input_layout_priceSt_One.addWidget(self.calculate_button_window_priceSt_One)

            self.stacked_widget.addWidget(self.input_screen_priceSt_One)

        self.stacked_widget.setCurrentWidget(self.input_screen_priceSt_One)
        self.back_button.show()

    def calculateWindowSingle(self):
        try:
            height = float(self.height_lineedit_window_priceSt_One.text())
            width = float(self.width_lineedit_window_priceSt_One.text())
            X1 = width - 224
            Y1 = height - 228
            result_text = f"Одностворчатое окно\n{X1} x {Y1} * {1} пакета"
            self.showResult(result_text)
        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")

    def showSingleSize(self, window_type):
        self.window_type = window_type  # Сохранение типа окна
        if not hasattr(self, 'input_screen_priceSt'):
            self.input_screen_priceSt = QWidget()
            self.input_layout_priceSt = QVBoxLayout(self.input_screen_priceSt)

            self.height_label_window_priceSt = QLabel('Высота (мм):')
            self.height_lineedit_window_priceSt = QLineEdit()

            self.width_label_window_priceSt = QLabel('Ширина (мм):')
            self.width_lineedit_window_priceSt = QLineEdit()

            self.spacer_size_label_window_priceSt = QLabel('Шпатик (мм):')
            self.spacer_size_lineedit_window_priceSt = QLineEdit()

            self.calculate_button_window_priceSt = QPushButton('Рассчитать')
            self.calculate_button_window_priceSt.clicked.connect(self.calculateWindow)
            self.input_layout_priceSt.addWidget(self.width_label_window_priceSt)
            self.input_layout_priceSt.addWidget(self.width_lineedit_window_priceSt)
            self.input_layout_priceSt.addWidget(self.height_label_window_priceSt)
            self.input_layout_priceSt.addWidget(self.height_lineedit_window_priceSt)

            self.input_layout_priceSt.addWidget(self.spacer_size_label_window_priceSt)
            self.input_layout_priceSt.addWidget(self.spacer_size_lineedit_window_priceSt)
            self.input_layout_priceSt.addWidget(self.calculate_button_window_priceSt)

            self.stacked_widget.addWidget(self.input_screen_priceSt)

        self.stacked_widget.removeWidget(self.selection_screen_priceSt)
        self.stacked_widget.setCurrentWidget(self.input_screen_priceSt)
        self.back_button.show()


    def showThreeSize(self, window_type):
        self.window_type = window_type  # Сохранение типа окна
        if not hasattr(self, 'mmm'):
            self.input_screen_priceSt_Three  = QWidget()
            self.input_layout_priceSt_Three  = QVBoxLayout(self.input_screen_priceSt_Three)

            self.height_label_window_priceSt_Three = QLabel('Высота (мм):')
            self.height_lineedit_window_priceSt_Three  = QLineEdit()

            self.width_label_window_priceSt_Three = QLabel('Ширина (мм):')
            self.width_lineedit_window_priceSt_Three = QLineEdit()

            self.spacer_size_label_window_priceSt_Three = QLabel('Шпатик (мм):')
            self.spacer_size_lineedit_window_priceSt_Three = QLineEdit()

            self.calculate_button_window_priceSt_Three = QPushButton('Рассчитать')
            self.calculate_button_window_priceSt_Three.clicked.connect(self.calculateThreeWindow)

            self.input_layout_priceSt_Three.addWidget(self.width_label_window_priceSt_Three)
            self.input_layout_priceSt_Three.addWidget(self.width_lineedit_window_priceSt_Three)
            self.input_layout_priceSt_Three.addWidget(self.height_label_window_priceSt_Three)
            self.input_layout_priceSt_Three.addWidget(self.height_lineedit_window_priceSt_Three)

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
            frame_width = 50
            mullion_width = 47
            count_impost = 1
            first_pack = 1
            second_pack = 1
            # В зависимости от типа окна выбираем количество импостов


            # Рассчитываем размеры другого окна
            leader_window = width - ((count_impost * mullion_width) + (frame_width+frame_width))
            leader_window -= spacer_size
            leader_window /= count_impost

            # Рассчитываем размеры пакета стекла
            X1 = leader_window - 10
            Y1 = height - 110


            X2 = spacer_size - 124
            Y2 = height - 228

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
            frame_width = 50
            mullion_width = 47
            count_impost = 2
            first_pack = 2
            second_pack = 1


            # Рассчитываем размеры другого окна
            leader_window = width - ((count_impost * mullion_width) + (frame_width + frame_width))
            leader_window -= spacer_size
            leader_window /= count_impost

            # Рассчитываем размеры пакета стекла
            X1 = leader_window - 10
            Y1 = height - 110


            X2 = spacer_size - 124
            Y2 = height - 228

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

    def showStekla(self):
        if not hasattr(self, 'input_screen'):
            self.input_screen = QWidget()
            self.input_layout = QVBoxLayout(self.input_screen)

            self.double_button = QPushButton("Дверь")
            self.double_button.clicked.connect(self.doors)

            self.st = QPushButton("Окно")
            self.st.clicked.connect(self.priceSt)

            self.input_layout.addWidget(self.double_button)
            self.input_layout.addWidget(self.st)



            self.stacked_widget.addWidget(self.input_screen)

        self.stacked_widget.setCurrentWidget(self.input_screen)
        self.back_button.show()




    def showSize(self):
        if not hasattr(self, 'input_screen_stekla'):
            self.input_screen_stekla = QWidget()
            self.input_layout_stekla = QVBoxLayout(self.input_screen_stekla)

            self.height_label = QLabel('Высота (мм):')
            self.height_lineedit = QLineEdit()

            self.width_label = QLabel('Ширина (мм):')
            self.width_lineedit = QLineEdit()

            self.btn = QLabel('Штапик (мм):')
            self.btn_edit = QLineEdit()

            self.calculate_button = QPushButton('Рассчитать')
            self.calculate_button.clicked.connect(self.calculateCost)
            self.input_layout_stekla.addWidget(self.width_label)
            self.input_layout_stekla.addWidget(self.width_lineedit)
            self.input_layout_stekla.addWidget(self.height_label)
            self.input_layout_stekla.addWidget(self.height_lineedit)

            self.input_layout_stekla.addWidget(self.btn)
            self.input_layout_stekla.addWidget(self.btn_edit)

            self.input_layout_stekla.addWidget(self.calculate_button)

            self.stacked_widget.addWidget(self.input_screen_stekla)

        self.stacked_widget.setCurrentWidget(self.input_screen_stekla)
        self.back_button.show()



    def calcThreeWindow(self):
        if not hasattr(self, 'ccc'):
            self.input_screen_stekla_Three = QWidget()
            self.input_layout_stekla_Three= QVBoxLayout(self.input_screen_stekla_Three)

            self.height_label_Three = QLabel('Высота (мм):')
            self.height_lineedit_Three = QLineEdit()

            self.width_label_Three= QLabel('Ширина (мм):')
            self.width_lineedit_Three = QLineEdit()

            self.btn_Three = QLabel('Штапик (мм):')
            self.btn_edit_Three = QLineEdit()

            self.calculate_button_Three = QPushButton('Рассчитать')
            self.calculate_button_Three.clicked.connect(self.calculateCostThree)

            self.input_layout_stekla_Three.addWidget(self.width_label_Three)
            self.input_layout_stekla_Three.addWidget(self.width_lineedit_Three)
            self.input_layout_stekla_Three.addWidget(self.height_label_Three)
            self.input_layout_stekla_Three.addWidget(self.height_lineedit_Three)

            self.input_layout_stekla_Three.addWidget(self.btn_Three)
            self.input_layout_stekla_Three.addWidget(self.btn_edit_Three)
            self.input_layout_stekla_Three.addWidget(self.calculate_button_Three)



            self.stacked_widget.addWidget(self.input_screen_stekla_Three)

        self.stacked_widget.setCurrentWidget(self.input_screen_stekla_Three)
        self.back_button.show()

    def calcDoubleWindow(self):
        if not hasattr(self, 'eee'):
            self.input_screen_stekla_Double = QWidget()
            self.input_layout_stekla_Double = QVBoxLayout(self.input_screen_stekla_Double)

            self.height_label_Double = QLabel('Высота (мм):')
            self.height_lineedit_Double = QLineEdit()

            self.width_label_Double = QLabel('Ширина (мм):')
            self.width_lineedit_Double = QLineEdit()

            self.btn_Double = QLabel('Штапик (мм):')
            self.btn_edit_Double = QLineEdit()

            self.calculate_button_Double = QPushButton('Рассчитать')
            self.calculate_button_Double.clicked.connect(self.calculateCostDouble)

            self.input_layout_stekla_Double.addWidget(self.width_label_Double)
            self.input_layout_stekla_Double.addWidget(self.width_lineedit_Double)
            self.input_layout_stekla_Double.addWidget(self.height_label_Double)
            self.input_layout_stekla_Double.addWidget(self.height_lineedit_Double)

            self.input_layout_stekla_Double.addWidget(self.btn_Double)
            self.input_layout_stekla_Double.addWidget(self.btn_edit_Double)
            self.input_layout_stekla_Double.addWidget(self.calculate_button_Double)



            self.stacked_widget.addWidget(self.input_screen_stekla_Double)

        self.stacked_widget.setCurrentWidget(self.input_screen_stekla_Double)
        self.back_button.show()

    def calcOneWindow(self):
        if not hasattr(self, 'www'):
            self.input_screen_stekla_One = QWidget()
            self.input_layout_stekla_One = QVBoxLayout(self.input_screen_stekla_One)

            self.height_label_One = QLabel('Высота (мм):')
            self.height_lineedit_One = QLineEdit()

            self.width_label_One = QLabel('Ширина (мм):')
            self.width_lineedit_One = QLineEdit()


            self.calculate_button_One = QPushButton('Рассчитать')
            self.calculate_button_One.clicked.connect(self.calculateCost_One)

            self.input_layout_stekla_One.addWidget(self.width_label_One)
            self.input_layout_stekla_One.addWidget(self.width_lineedit_One)
            self.input_layout_stekla_One.addWidget(self.height_label_One)
            self.input_layout_stekla_One.addWidget(self.height_lineedit_One)



            self.input_layout_stekla_One.addWidget(self.calculate_button_One)

            self.stacked_widget.addWidget(self.input_screen_stekla_One)
            self.stacked_widget.setCurrentWidget(self.input_screen_stekla_One)

        self.stacked_widget.setCurrentWidget(self.input_screen_stekla_One)
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
                self.left_lineedit_priceDoor.clear()
                self.right_lineedit_priceDoor.clear()
            if hasattr(self, 'width_lineedit_priceDoor_One'):
                self.height_lineedit_priceDoor_One.clear()
                self.width_lineedit_priceDoor_One.clear()

            if hasattr(self, 'height_label_Double'):
                self.height_label_Double.clear()
                self.width_label_Double.clear()
                self.btn_Double.clear()
            if hasattr(self, 'height_label_Three'):
                self.height_label_Three.clear()
                self.width_label_Three.clear()
                self.btn_Three.clear()
            if hasattr(self, 'height_label_window_priceSt_One'):
                self.height_label_window_priceSt_One.clear()
                self.width_label_window_priceSt_One.clear()





        self.stacked_widget.setCurrentIndex(next_index)

        if next_index == 0:
            self.back_button.hide()

    def choose_window(self):
        selection_screen_new = QWidget()
        selection_layout_new  = QVBoxLayout(selection_screen_new)

        single_button = QPushButton('Одностворч')
        single_button.clicked.connect(self.calcOneWindow)  # Передача типа окна

        double_button = QPushButton('Двухстворч')
        double_button.clicked.connect(self.calcDoubleWindow)  # Передача типа окна

        triple_button = QPushButton('Трехстворч')
        triple_button.clicked.connect(self.calcThreeWindow)  # Передача типа окна

        selection_layout_new.addWidget(single_button)
        selection_layout_new.addWidget(double_button)
        selection_layout_new.addWidget(triple_button)

        self.stacked_widget.addWidget(selection_screen_new)
        self.selection_screen_priceSt = selection_screen_new
        self.stacked_widget.setCurrentWidget(selection_screen_new)

        self.back_button.show()


    def calculateCostThree(self):
        try:
            height = float(self.height_lineedit_Three.text())
            width = float(self.width_lineedit_Three.text())
            shtapik = float(self.btn_edit_Three.text())

            # Здесь нужно определить значения frame_price, mullion_price и sash_price
            frame_cost = 0  # Замените на свои значения
            mullion_cost = 0  # Замените на свои значения
            sash_cost = 0  # Замените на свои значения



            self.costs_dict = self.loadPrices()

            self.showCostTable_Three(self.costs_dict)
        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")

    def calculateCost_One(self):
        try:
            height = float(self.height_lineedit_One.text())
            width = float(self.width_lineedit_One.text())

            # Здесь нужно определить значения frame_price, mullion_price и sash_price
            frame_cost = 0  # Замените на свои значения
            mullion_cost = 0  # Замените на свои значения
            sash_cost = 0  # Замените на свои значения



            self.costs_dict = self.loadPrices()

            self.showCostTable_One(self.costs_dict)

        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")

    def calculateCostDouble(self):
        try:
            height = float(self.height_lineedit_Double.text())
            width = float(self.width_lineedit_Double.text())
            shtapik = float(self.btn_edit_Double.text())

            # Здесь нужно определить значения frame_price, mullion_price и sash_price
            frame_cost = 0  # Замените на свои значения
            mullion_cost = 0  # Замените на свои значения
            sash_cost = 0  # Замените на свои значения



            self.costs_dict = self.loadPrices()



            self.showCostTable_Double(self.costs_dict)

        except ValueError:
            print("Ошибка: некорректный формат ввода данных.")


    def showCostTable_Three(self,costs_dict):
        if hasattr(self, 'result_table'):
            self.result_table.deleteLater()

        self.result_table = QTableWidget()
        self.result_table.setRowCount(len(costs_dict) + 1)  # +1 для строки с итоговой суммой
        self.result_table.setColumnCount(7)  # Добавлен один столбец для суммы

        headers = ['N', 'Наименование', 'Кол-во', 'Единица', 'Цена', 'Сумма']
        self.result_table.setHorizontalHeaderLabels(headers)

        height = float(self.height_lineedit_Three.text())  / 1000
        width = float(self.width_lineedit_Three.text()) / 1000
        shtapik = float(self.btn_edit_Three.text()) / 1000

        rama = (height + width) * 2
        impost = (height - 0.1) * 2
        stvor = ((height - 0.1) + shtapik) * 2
        shtapik = rama + impost * 2
        metall = (rama + impost + stvor)
        rezina = rama + impost* 2 + stvor * 2
        steklo = (height * width) * 0.75

        row = 0
        total_cost = 0

        for item_name, item_cost in costs_dict.items():
            total_cost_item = float(item_cost.get('Белый') ) # Цена товара из Excel
            count = 1 if item_name in ['Замок', 'Петля', 'Ручка', 'Монтаж', 'Прочее','Производство'] else 2 * (height + width)
            count_formula = None
            if item_name == "Рама":
                count_formula = rama
            elif item_name == "Импост":
                count_formula = impost
            elif item_name == "Створка":
                count_formula = stvor
            elif item_name == "Штапик":
                count_formula = shtapik
            elif item_name == "Металл":
                count_formula = metall
            elif item_name == "Резина":
                count_formula = rezina
            elif item_name == 'Производство':
                count_formula = (width * height)
            elif item_name == "Стекло":
                count_formula = steklo
            if count_formula is not None:
                count = count_formula
            print(type(count),type(total_cost))
            subtotal = round(total_cost_item * count, 2)  # Сумма товара, округленная до двух знаков

            total_cost += subtotal

            self.result_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))  # Номер строки
            self.result_table.setItem(row, 1, QTableWidgetItem(item_name))  # Название товара
            self.result_table.setItem(row, 2, QTableWidgetItem(str(round(count,3))))  # Кол-во
            self.result_table.setItem(row, 3, QTableWidgetItem(str(item_cost['Единица'])))  # Единица измерения
            self.result_table.setItem(row, 4, QTableWidgetItem(f'{total_cost_item} руб.'))  # Цена
            self.result_table.setItem(row, 5, QTableWidgetItem(f'{ subtotal} руб.'))  # Сумма


            # Добавляем QComboBox с возможностью выбора цвета
            combobox = QComboBox()
            combobox.addItems(['Белый', 'Комбинированный', 'Цельный'])
            combobox.currentIndexChanged.connect(
                lambda index, row=row, item_cost=item_cost: self.updatePrice(index, row, item_cost))
            self.result_table.setCellWidget(row, 6, combobox)

            row += 1


        # Строка с итоговой суммой
        self.result_table.setItem(row, 3, QTableWidgetItem('Итого:'))
        self.result_table.setItem(row, 5, QTableWidgetItem(f'{round(total_cost,2)} руб.'))

        self.result_table.resizeColumnsToContents()
        self.result_table.setWindowTitle('Результат расчета стоимости')
        self.result_table.show()

        self.stacked_widget.addWidget(self.result_table)
        self.stacked_widget.setCurrentWidget(self.result_table)

    def showCostTable_Double(self, costs_dict):
        if hasattr(self, 'result_table'):
            self.result_table.deleteLater()

        self.result_table = QTableWidget()
        self.result_table.setRowCount(len(costs_dict) + 1)  # +1 для строки с итоговой суммой
        self.result_table.setColumnCount(7)  # Добавлен один столбец для суммы

        headers = ['N', 'Наименование', 'Кол-во', 'Единица', 'Цена', 'Сумма']
        self.result_table.setHorizontalHeaderLabels(headers)

        height = float(self.height_lineedit_Double.text())  / 1000
        width = float(self.width_lineedit_Double.text()) / 1000
        shtapik = float(self.btn_edit_Double.text()) / 1000

        rama = (height + width) * 2
        impost = (height - 0.1) * 1
        stvor = ((height - 0.1) + shtapik) * 2
        shtapik = rama + impost * 2
        metall = (rama + impost + stvor)
        rezina = rama + impost* 2 + stvor * 2
        steklo = (height * width) * 0.75

        row = 0
        total_cost = 0

        for item_name, item_cost in costs_dict.items():
            total_cost_item = float(item_cost.get('Белый') ) # Цена товара из Excel
            count = 1 if item_name in ['Замок', 'Петля', 'Ручка', 'Монтаж', 'Прочее','Производство'] else 2 * (height + width)
            count_formula = None
            if item_name == "Рама":
                count_formula = rama
            elif item_name == "Импост":
                count_formula = impost
            elif item_name == "Створка":
                count_formula = stvor
            elif item_name == "Штапик":
                count_formula = shtapik
            elif item_name == "Металл":
                count_formula = metall
            elif item_name == "Резина":
                count_formula = rezina
            elif item_name == "Стекло":
                count_formula = steklo
            elif item_name == 'Производство':
                count_formula = (width * height)
            if count_formula is not None:
                count = count_formula
            subtotal = round(total_cost_item * count, 2)  # Сумма товара, округленная до двух знаков

            total_cost += subtotal

            self.result_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))  # Номер строки
            self.result_table.setItem(row, 1, QTableWidgetItem(item_name))  # Название товара
            self.result_table.setItem(row, 2, QTableWidgetItem(str(round(count,3))))  # Кол-во
            self.result_table.setItem(row, 3, QTableWidgetItem(str(item_cost['Единица'])))  # Единица измерения
            self.result_table.setItem(row, 4, QTableWidgetItem(f'{total_cost_item} руб.'))  # Цена
            self.result_table.setItem(row, 5, QTableWidgetItem(f'{subtotal} руб.'))  # Сумма


            # Добавляем QComboBox с возможностью выбора цвета
            combobox = QComboBox()
            combobox.addItems(['Белый', 'Комбинированный', 'Цельный'])
            combobox.currentIndexChanged.connect(
                lambda index, row=row, item_cost=item_cost: self.updatePrice(index, row, item_cost))
            self.result_table.setCellWidget(row, 6, combobox)

            row += 1

        # Строка с итоговой суммой
        self.result_table.setItem(row, 3, QTableWidgetItem('Итого:'))
        self.result_table.setItem(row, 5, QTableWidgetItem(f'{total_cost} руб.'))

        self.result_table.resizeColumnsToContents()
        self.result_table.setWindowTitle('Результат расчета стоимости')
        self.result_table.show()

        self.stacked_widget.addWidget(self.result_table)
        self.stacked_widget.setCurrentWidget(self.result_table)

    def showCostTable_One(self, costs_dict, row=None):
        if hasattr(self, 'result_table'):
            self.result_table.deleteLater()

        self.result_table = QTableWidget()
        self.result_table.setRowCount(len(costs_dict) + 1)  # +1 для строки с итоговой суммой
        self.result_table.setColumnCount(7)  # Добавлен один столбец для суммы

        headers = ['N', 'Наименование', 'Кол-во', 'Единица', 'Цена', 'Сумма', 'Цвет']
        self.result_table.setHorizontalHeaderLabels(headers)

        height = float(self.height_lineedit_One.text()) / 1000
        width = float(self.width_lineedit_One.text()) / 1000

        rama = (height + width) * 2
        stvor = ((height - 0.1) + (width - 0.1)) * 2
        shtapik = rama
        metall = (rama  + stvor)
        rezina = rama * 2 + stvor * 2
        steklo = (height * width) * 0.75

        row = 0
        total_cost = 0

        for item_name, item_cost in costs_dict.items():
            total_cost_item = float(item_cost.get('Белый'))  # Цена товара из Excel

            count = 1 if item_name in ['Замок', 'Петля', 'Ручка', 'Монтаж', 'Прочее', 'Производство'] else 2 * (
                        height + width)
            count_formula = None

            if item_name == "Рама":
                count_formula = rama
            elif item_name == "Створка":
                count_formula = stvor
            elif item_name == "Штапик":
                count_formula = shtapik
            elif item_name == "Металл":
                count_formula = metall
            elif item_name == "Резина":
                count_formula = rezina
            elif item_name == 'Производство':
                count_formula = (width * height)
            elif item_name == "Стекло":
                count_formula = steklo

            if count_formula is not None:
                count = count_formula

            subtotal = round(total_cost_item * count, 2)  # Сумма товара, округленная до двух знаков
            total_cost += subtotal

            self.result_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))  # Номер строки
            self.result_table.setItem(row, 1, QTableWidgetItem(item_name))  # Название товара
            self.result_table.setItem(row, 2, QTableWidgetItem(str(round(count, 3))))  # Кол-во
            self.result_table.setItem(row, 3, QTableWidgetItem(str(item_cost['Единица'])))  # Единица измерения
            self.result_table.setItem(row, 4, QTableWidgetItem(f'{total_cost_item} руб.'))  # Цена
            self.result_table.setItem(row, 5, QTableWidgetItem(f'{subtotal} руб.'))  # Сумма

            # Добавляем QComboBox с возможностью выбора цвета
            combobox = QComboBox()
            combobox.addItems(['Белый', 'Комбинированный', 'Цельный'])
            combobox.currentIndexChanged.connect(
                lambda index, row=row, item_cost=item_cost: self.updatePrice(index, row, item_cost))
            self.result_table.setCellWidget(row, 6, combobox)

            row += 1

        # Строка с итоговой суммой
        self.result_table.setItem(row, 3, QTableWidgetItem('Итого:'))
        self.result_table.setItem(row, 5, QTableWidgetItem(f'{total_cost} руб.'))

        self.result_table.resizeColumnsToContents()
        self.result_table.setWindowTitle('Результат расчета стоимости')
        self.result_table.show()

        self.stacked_widget.addWidget(self.result_table)
        self.stacked_widget.setCurrentWidget(self.result_table)

    # Функция для обновления цены и итоговой суммы при изменении цвета в QComboBox
    def updatePrice(self, index, row, item_cost):
        color = item_cost[['Белый', 'Комбинированный', 'Цельный'][index]]
        price_item = float(color)
        count_item = float(self.result_table.item(row, 2).text())
        subtotal = round(price_item * count_item, 2)

        self.result_table.item(row, 4).setText(f'{price_item} руб.')
        self.result_table.item(row, 5).setText(f'{subtotal} руб.')

        # Пересчитываем итоговую сумму
        total_cost = sum(
            float(self.result_table.item(i, 5).text().split()[0]) for i in range(self.result_table.rowCount() - 1))
        self.result_table.item(self.result_table.rowCount() - 1, 5).setText(f'{total_cost:.2f} руб.')

    def loadPrices(self):
        wb = openpyxl.load_workbook('file.xlsx')
        sheet = wb.active
        self.costs_dict = {}

        # Получаем значения по столбцам
        item_names = [cell.value for cell in sheet['A'][1:] if cell.value is not None]  # Названия товаров
        item_ed = [cell.value for cell in sheet['B'][1:] if cell.value is not None]  # Единица товара
        item_white_costs = [cell.value for cell in sheet['C'][1:] if
                            cell.value is not None]  # Цены товаров на белый цвет
        item_combined_costs = [cell.value for cell in sheet['D'][1:] if
                               cell.value is not None]  # Цены товаров на комбинированный цвет
        item_solid_costs = [cell.value for cell in sheet['E'][1:] if
                            cell.value is not None]  # Цены товаров на цельный цвет

        # Формируем словарь
        for name, ed, white_cost, combined_cost, solid_cost in zip(item_names, item_ed, item_white_costs,
                                                                   item_combined_costs, item_solid_costs):
            self.costs_dict[name] = {'Единица': ed, 'Белый': white_cost, 'Комбинированный': combined_cost,
                                     'Цельный': solid_cost}
        print(self.costs_dict)
        return self.costs_dict

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Message', 'Are you sure to quit?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = WindowSelector()
    mainWindow.show()
    sys.exit(app.exec_())