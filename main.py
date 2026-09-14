import sys
import pandas as pd

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QComboBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMessageBox,
    QScrollArea
)

from PyQt5.QtCore import Qt


EXCEL_FILE = "ASME-Pipedimension.xlsx"

FILTER_COLUMNS = [
    "NPS",
    "DN",
    "Identification",
    "Schedule No."
]


class PipeDimensionApp(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle(
            "PipeDim - Steel Pipe Dimensions & Mass Calculator"
        )

        self.setMinimumSize(1200, 700)

        self.df = pd.DataFrame()
        self.filtered_df = pd.DataFrame()
        self.filter_boxes = {}

        self.load_excel()
        self.setup_ui()

    # ---------------------------------------------------------
    # LOAD EXCEL
    # ---------------------------------------------------------

    def load_excel(self):

        try:
            self.df = pd.read_excel(EXCEL_FILE)

            # Remove completely empty rows
            self.df.dropna(how="all", inplace=True)

            # Convert column names to strings
            self.df.columns = self.df.columns.astype(str)

            # Replace NaN values
            self.df = self.df.fillna("")

            self.filtered_df = self.df.copy()

        except Exception as e:

            QMessageBox.critical(
                self,
                "Excel Error",
                f"Unable to load Excel file.\n\n{str(e)}"
            )

            sys.exit()

    # ---------------------------------------------------------
    # USER INTERFACE
    # ---------------------------------------------------------

    def setup_ui(self):

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)

        # -----------------------------------------------------
        # TITLE
        # -----------------------------------------------------

        title = QLabel(
            "PIPE DIMENSION DATABASE"
        )

        title.setAlignment(Qt.AlignCenter)

        title.setStyleSheet("""
            QLabel {
                font-size: 22px;
                font-weight: bold;
                padding: 15px;
            }
        """)

        main_layout.addWidget(title)

        subtitle = QLabel(
            "Dimensions and Weights (Masses) of Welded and Seamless Wrought Steel Pipe"
        )

        subtitle.setAlignment(Qt.AlignCenter)

        subtitle.setStyleSheet("""
            QLabel {
                font-size: 13px;
                padding-bottom: 10px;
            }
        """)

        main_layout.addWidget(subtitle)

        # -----------------------------------------------------
        # FILTER AREA
        # -----------------------------------------------------

        filter_scroll = QScrollArea()

        filter_scroll.setWidgetResizable(True)

        filter_container = QWidget()

        filter_layout = QGridLayout(filter_container)

        columns = [
            column for column in FILTER_COLUMNS
            if column in self.df.columns
        ]

        for index, column in enumerate(columns):

            row = index // 4
            col = index % 4

            label = QLabel(column)

            combo = QComboBox()

            combo.addItem("All")

            values = (
                self.df[column]
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )

            values = sorted(
                [value for value in values if value != ""],
                key=str.lower
            )

            combo.addItems(values)

            combo.currentIndexChanged.connect(
                self.apply_filters
            )

            self.filter_boxes[column] = combo

            filter_layout.addWidget(label, row * 2, col)
            filter_layout.addWidget(combo, row * 2 + 1, col)

        filter_scroll.setWidget(filter_container)

        filter_scroll.setMaximumHeight(180)

        main_layout.addWidget(filter_scroll)

        # -----------------------------------------------------
        # BUTTONS
        # -----------------------------------------------------

        button_layout = QHBoxLayout()

        clear_button = QPushButton(
            "Clear Filters"
        )

        clear_button.clicked.connect(
            self.clear_filters
        )

        clear_button.setMinimumHeight(35)

        button_layout.addWidget(clear_button)

        button_layout.addStretch()

        self.record_label = QLabel()

        button_layout.addWidget(
            self.record_label
        )

        main_layout.addLayout(button_layout)

        # -----------------------------------------------------
        # TABLE
        # -----------------------------------------------------

        self.table = QTableWidget()

        self.table.setSortingEnabled(True)

        self.table.setAlternatingRowColors(True)

        self.table.setSelectionBehavior(
            QTableWidget.SelectRows
        )

        self.table.setEditTriggers(
            QTableWidget.NoEditTriggers
        )

        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents
        )

        main_layout.addWidget(self.table)

        # Display initial data
        self.populate_table(self.df)

    # ---------------------------------------------------------
    # APPLY FILTERS
    # ---------------------------------------------------------

    def apply_filters(self):

        filtered = self.df.copy()

        for column, combo in self.filter_boxes.items():

            selected_value = combo.currentText()

            if selected_value != "All":

                filtered = filtered[
                    filtered[column].astype(str).str.strip()
                    == selected_value
                ]

        self.filtered_df = filtered

        self.populate_table(filtered)

    # ---------------------------------------------------------
    # CLEAR FILTERS
    # ---------------------------------------------------------

    def clear_filters(self):

        for combo in self.filter_boxes.values():

            combo.blockSignals(True)

            combo.setCurrentIndex(0)

            combo.blockSignals(False)

        self.filtered_df = self.df.copy()

        self.populate_table(self.df)

    # ---------------------------------------------------------
    # POPULATE TABLE
    # ---------------------------------------------------------

    def populate_table(self, dataframe):

        self.table.setSortingEnabled(False)

        self.table.clear()

        self.table.setRowCount(len(dataframe))

        self.table.setColumnCount(len(dataframe.columns))

        self.table.setHorizontalHeaderLabels(
            list(dataframe.columns)
        )

        for row_index, row in enumerate(
            dataframe.itertuples(index=False, name=None)
        ):

            for column_index, value in enumerate(row):

                item = QTableWidgetItem(
                    str(value)
                )

                self.table.setItem(
                    row_index,
                    column_index,
                    item
                )

        self.table.setSortingEnabled(True)

        self.record_label.setText(
            f"Records: {len(dataframe)} / {len(self.df)}"
        )


# =============================================================
# APPLICATION
# =============================================================

if __name__ == "__main__":

    app = QApplication(sys.argv)

    window = PipeDimensionApp()

    window.show()

    sys.exit(app.exec_())