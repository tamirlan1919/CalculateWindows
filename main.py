import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QStackedWidget, QTableWidget, QTableWidgetItem, QComboBox


class WindowSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Калькулятор окон')
        self.setGeometry(100, 100, 500, 300)

        # Создание главного макета
        main_layout = QVBoxLayout(self)

        # Создание стек виджета
        self.stacked_widget = QStackedWidget(self)

        # Экран с кнопками выбора
        self.selection_screen = QWidget()
        self.selection_layout = QVBoxLayout(self.selection_screen)

        self.single_button = QPushButton('Одностворчатое окно')
        self.single_button.clicked.connect(self.showSingleInput)

        self.double_button = QPushButton('Двухстворчатое окно')
        self.double_button.clicked.connect(self.showDoubleInput)

        self.triple_button = QPushButton('Трехстворчатое окно')
        self.triple_button.clicked.connect(self.showTripleInput)

        self.selection_layout.addWidget(self.single_button)
        self.selection_layout.addWidget(self.double_button)
        self.selection_layout.addWidget(self.triple_button)

        self.stacked_widget.addWidget(self.selection_screen)

        # Добавляем стек виджет на главный макет
        main_layout.addWidget(self.stacked_widget)

        # Создание кнопки "Назад"
        self.back_button = QPushButton('Назад')
        self.back_button.clicked.connect(self.showSelectionScreen)
        main_layout.addWidget(self.back_button)

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

    def showSingleInput(self):
        self.createInputScreen('Одностворчатое окно')

    def showDoubleInput(self):
        self.createInputScreen('Двухстворчатое окно')

    def showTripleInput(self):
        self.createInputScreen('Трехстворчатое окно')

    def createInputScreen(self, window_type):
        # Создание экрана с полями ввода
        self.input_screen = QWidget()
        self.input_layout = QVBoxLayout(self.input_screen)

        self.height_label = QLabel('Высота (см):')
        self.height_lineedit = QLineEdit()

        self.width_label = QLabel('Ширина (см):')
        self.width_lineedit = QLineEdit()

        self.spacer_size_label = QLabel('Размер шпатика (см):')
        self.spacer_size_lineedit = QLineEdit()

        self.glass_package_label = QLabel('Стеклопакет:')
        self.glass_package_combobox = QComboBox()
        self.glass_package_combobox.addItems(['Однокамерный', 'Двухкамерный'])

        self.calculate_button = QPushButton('Рассчитать')
        self.calculate_button.clicked.connect(self.calculateCost)

        self.input_layout.addWidget(self.height_label)
        self.input_layout.addWidget(self.height_lineedit)
        self.input_layout.addWidget(self.width_label)
        self.input_layout.addWidget(self.width_lineedit)
        self.input_layout.addWidget(self.spacer_size_label)
        self.input_layout.addWidget(self.spacer_size_lineedit)
        self.input_layout.addWidget(self.glass_package_label)
        self.input_layout.addWidget(self.glass_package_combobox)
        self.input_layout.addWidget(self.calculate_button)

        # Очистка старых экранов
        self.stacked_widget.removeWidget(self.selection_screen)
        if hasattr(self, 'input_screen'):
            self.stacked_widget.removeWidget(self.input_screen)

        # Добавляем экран с полями ввода на стек виджет
        self.stacked_widget.addWidget(self.input_screen)
        self.stacked_widget.setCurrentWidget(self.input_screen)

    def showSelectionScreen(self):
        self.stacked_widget.setCurrentWidget(self.selection_screen)

    def calculateCost(self):
        height = float(self.height_lineedit.text())
        width = float(self.width_lineedit.text())
        spacer_size = float(self.spacer_size_lineedit.text())

        glass_package = self.glass_package_combobox.currentText()

        frame_cost = 182 * (2 * (height + width) / 100)  # 2 стороны
        mullion_cost = 216 * (2 * (height + width) / 100)  # 2 стороны
        sash_cost = 208 * (height / 100)  # 1 сторона

        if glass_package == 'Однокамерный':
            glass_cost = 100
        else:
            glass_cost = 150

        total_cost = frame_cost + mullion_cost + sash_cost + glass_cost

        self.showCostTable(frame_cost, mullion_cost, sash_cost, glass_cost, total_cost)

    def showCostTable(self, frame_cost, mullion_cost, sash_cost, glass_cost, total_cost):
        # Создание таблицы для отображения результата расчета
        self.result_table = QTableWidget()
        self.result_table.setRowCount(5)
        self.result_table.setColumnCount(5)

        headers = ['N', 'Товар', 'Кол-во', 'Единица', 'Сумма']
        self.result_table.setHorizontalHeaderLabels(headers)

        items = [
            ('1', 'Рама', f'{2 * ((float(self.height_lineedit.text()) + float(self.width_lineedit.text())) / 100)}', 'м', f'{frame_cost} руб.'),
            ('2', 'Импост', f'{2 * ((float(self.height_lineedit.text()) + float(self.width_lineedit.text())) / 100)}', 'м', f'{mullion_cost} руб.'),
            ('3', 'Створка', f'{float(self.height_lineedit.text()) / 100}', 'м', f'{sash_cost} руб.'),
            ('4', 'Стеклопакет', '1', 'шт', f'{glass_cost} руб.'),
            ('', '', '', 'Итого:', f'{total_cost} руб.')
        ]

        for row, item in enumerate(items):
            for col, data in enumerate(item):
                self.result_table.setItem(row, col, QTableWidgetItem(data))

        self.result_table.resizeColumnsToContents()
        self.result_table.setWindowTitle('Результат расчета стоимости')
        self.result_table.show()

        # Очистка старых экранов
        self.stacked_widget.removeWidget(self.input_screen)
        self.stacked_widget.removeWidget(self.selection_screen)

        # Добавляем таблицу на стек виджет
        self.stacked_widget.addWidget(self.result_table)
        self.stacked_widget.setCurrentWidget(self.result_table)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WindowSelector()
    sys.exit(app.exec_())
