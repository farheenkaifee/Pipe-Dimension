import sys
import pandas as pd
import os

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
    QTextBrowser,
    QSplitter,
    QFrame,
    QListView
)

from PyQt5.QtGui import QPixmap

from PyQt5.QtCore import Qt


# =========================================================
# CONFIGURATION
# =========================================================

EXCEL_FILE = "ASME-Pipedimension.xlsx"

FILTER_COLUMNS = [
    "NPS",
    "DN",
    "Identification",
    "Schedule No."
]


# =========================================================
# MAIN APPLICATION
# =========================================================

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

    # =====================================================
    # LOAD EXCEL
    # =====================================================

    def load_excel(self):

        try:

            self.df = pd.read_excel(EXCEL_FILE)

            # Remove completely empty rows
            self.df.dropna(
                how="all",
                inplace=True
            )

            # Clean column names
            self.df.columns = (
                self.df.columns
                .astype(str)
                .str.strip()
            )

            # Replace NaN
            self.df = self.df.fillna("")

            # Initial dataframe
            self.filtered_df = self.df.copy()

        except FileNotFoundError:

            QMessageBox.critical(
                self,
                "Excel File Not Found",
                f"Could not find:\n\n"
                f"{EXCEL_FILE}\n\n"
                "Make sure the Excel file is in the same folder "
                "as this Python file."
            )

            sys.exit()

        except Exception as e:

            QMessageBox.critical(
                self,
                "Excel Error",
                f"Unable to load Excel file.\n\n{str(e)}"
            )

            sys.exit()

    # =====================================================
    # USER INTERFACE
    # =====================================================

    def setup_ui(self):

        central_widget = QWidget()

        self.setCentralWidget(
            central_widget
        )

        main_layout = QVBoxLayout(
            central_widget
        )

        main_layout.setContentsMargins(
            12, 10, 12, 10
        )

        main_layout.setSpacing(8)

        # =================================================
        # GLOBAL STYLE
        # =================================================

        self.setStyleSheet("""

            QMainWindow {
                background: #f4f6f8;
            }

            QLabel {
                color: #17202a;
            }

            QComboBox {
                background: white;
                border: 1px solid #b8c2cc;
                border-radius: 4px;
                padding: 5px 5px;
                font-size: 12px;
                min-height: 25px;
            }

            QComboBox QAbstractItemView::item {
                border-bottom: 1px solid #d3d3d3;
                padding-top: 1px;
                padding-bottom: 1px;
            }

            QComboBox QAbstractItemView::item:hover {
                background-color: #ffffff;
                color: #000000;
            }

            QComboBox:hover {
                border: 1px solid #16355d;
            }

            QComboBox:focus {
                border: 1px solid #16355d;
            }

            QPushButton {
                background: #16355d;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 7px 18px;
                font-size: 12px;
                font-weight: bold;
            }

            QPushButton:hover {
                background: #0d325c;
            }

            QPushButton:pressed {
                background: #0a2747;
            }

            QTableWidget {
                background: white;
                border: 1px solid #cbd3da;
                gridline-color: #d9dee3;
                font-size: 11px;
                selection-background-color: #dce9f7;
                selection-color: #111111;
            }

            QTableWidget::item {
                padding: 4px;
            }

            QHeaderView::section {
                background: #16355d;
                color: white;
                font-weight: bold;
                font-size: 11px;
                padding: 7px 5px;
                border: none;
                border-right: 1px solid #496784;
            }

            QTextBrowser {
                background: white;
                border: 1px solid #cbd3da;
                border-radius: 4px;
                padding: 10px;
                font-size: 12px;
            }

            QSplitter::handle {
                background: #d2d8de;
            }

            QScrollBar:vertical {
                width: 10px;
            }

            QScrollBar:horizontal {
                height: 10px;
            }

        """)

        # =================================================
        # HEADER
        # =================================================

        header = QFrame()

        header.setFixedHeight(100)

        header.setStyleSheet("""
            QFrame {
                background: white;
                border: 1px solid #d5dce3;
                border-radius: 5px;
            }
        """)

        header_layout = QHBoxLayout(header)

        header_layout.setContentsMargins(
            10, 5, 10, 5
        )

        header_layout.setSpacing(10)

        # =================================================
        # COMPANY LOGO - LEFT
        # =================================================

        logo_label = QLabel()

        logo_label.setFixedSize(
            220,
            70
        )

        logo_label.setAlignment(
            Qt.AlignLeft | Qt.AlignVCenter
        )

        logo_label.setScaledContents(False)

        logo_path = os.path.join(
            os.path.dirname(
                os.path.abspath(__file__)
            ),
            "assets",
            "company_logo.png"
        )

        logo = QPixmap(logo_path)

        if not logo.isNull():

            logo_label.setPixmap(
                logo.scaled(
                    210,
                    60,
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                )
            )

        # Add logo to LEFT
        header_layout.addWidget(
            logo_label,
            0,
            Qt.AlignLeft | Qt.AlignVCenter
        )

        # =================================================
        # TITLE CONTAINER
        # =================================================

        title_container = QWidget()

        title_layout = QVBoxLayout(
            title_container
        )

        title_layout.setContentsMargins(
            0, 0, 0, 0
        )

        title_layout.setSpacing(2)

        # =================================================
        # MAIN TITLE
        # =================================================

        title = QLabel(
            "ASME B36.10M - STEEL PIPE DIMENSIONS"
        )

        title.setAlignment(
            Qt.AlignCenter
        )

        title.setStyleSheet("""
            QLabel {
                color: #16355d;
                font-size: 20px;
                font-weight: bold;
                border: none;
            }
        """)

        # =================================================
        # SUBTITLE
        # =================================================

        subtitle = QLabel(
            "Dimensions and Weights (Masses) of "
            "Welded and Seamless Wrought Steel Pipe"
        )

        subtitle.setAlignment(
            Qt.AlignCenter
        )

        subtitle.setStyleSheet("""
            QLabel {
                color: #5b6570;
                font-size: 11px;
                border: none;
            }
        """)

        title_layout.addWidget(title)
        title_layout.addWidget(subtitle)

        # =================================================
        # TITLE - CENTER
        # =================================================

        header_layout.addWidget(
            title_container,
            1,
            Qt.AlignVCenter
        )

        # =================================================
        # OPTIONAL RIGHT SPACER
        # Keeps title truly centered relative to header
        # =================================================

        right_spacer = QWidget()

        right_spacer.setFixedWidth(
            220
        )

        header_layout.addWidget(
            right_spacer
        )


        # ADD HEADER TO MAIN LAYOUT
        main_layout.addWidget(
            header
        )

            

        # =================================================
        # FILTER FRAME
        # =================================================

        filter_frame = QFrame()

        filter_frame.setFixedHeight(88)

        filter_frame.setStyleSheet("""
            QFrame {
                background: white;
                border: 1px solid #cbd3da;
                border-radius: 5px;
            }
        """)

        filter_layout = QGridLayout(
            filter_frame
        )

        filter_layout.setContentsMargins(
            10, 7, 10, 7
        )

        filter_layout.setHorizontalSpacing(
            12
        )

        filter_layout.setVerticalSpacing(
            2
        )

        columns = [
            column
            for column in FILTER_COLUMNS
            if column in self.df.columns
        ]

        # =================================================
        # CREATE FILTER COMBO BOXES
        # =================================================

        for index, column in enumerate(columns):

            # -----------------------------
            # LABEL
            # -----------------------------

            label = QLabel(
                column
            )

            label.setStyleSheet("""
                QLabel {
                    font-weight: bold;
                    color: #16355d;
                    font-size: 11px;
                    border: none;
                }
            """)

            # -----------------------------
            # COMBO BOX
            # -----------------------------

            combo = QComboBox()

            # Dropdown height
            combo.view().setMinimumHeight(0)
            combo.view().setMaximumHeight(50)

            combo.setMinimumHeight(28)

            # Custom list view
            view = QListView()

            combo.setView(view)

            combo.setSizePolicy(
                combo.sizePolicy().Expanding,
                combo.sizePolicy().Fixed
            )

            # Add initial values
            combo.addItem("All")

            values = (
                self.df[column]
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )

            values = sorted(
                [
                    value
                    for value in values
                    if value != ""
                ],
                key=str.lower
            )

            combo.addItems(
                values
            )

            # Connect combo box
            combo.currentIndexChanged.connect(
                lambda index, col=column:
                self.filter_changed(col)
            )

            # Store combo
            self.filter_boxes[column] = combo

            # -----------------------------
            # ADD TO GRID
            # -----------------------------

            filter_layout.addWidget(
                label,
                0,
                index
            )

            filter_layout.addWidget(
                combo,
                1,
                index
            )

            filter_layout.setColumnStretch(
                index,
                1
            )

        main_layout.addWidget(
            filter_frame
        )

        # =================================================
        # CONTROL BAR
        # =================================================

        control_frame = QFrame()

        control_frame.setFixedHeight(42)

        control_frame.setStyleSheet("""
            QFrame {
                background: transparent;
                border: none;
            }
        """)

        control_layout = QHBoxLayout(
            control_frame
        )

        control_layout.setContentsMargins(
            0, 0, 0, 0
        )

        # Clear button

        clear_button = QPushButton(
            "Clear Filters"
        )

        clear_button.setFixedSize(
            120,
            32
        )

        clear_button.clicked.connect(
            self.clear_filters
        )

        control_layout.addWidget(
            clear_button
        )

        control_layout.addStretch()

        # Record count

        self.record_label = QLabel()

        self.record_label.setStyleSheet("""
            QLabel {
                color: #16355d;
                font-size: 12px;
                font-weight: bold;
                padding-right: 5px;
            }
        """)

        control_layout.addWidget(
            self.record_label
        )

        main_layout.addWidget(
            control_frame
        )

        # =================================================
        # MAIN CONTENT
        # =================================================

        splitter = QSplitter(
            Qt.Horizontal
        )

        splitter.setChildrenCollapsible(
            False
        )

        # =================================================
        # LEFT - TABLE
        # =================================================

        table_widget = QWidget()

        table_layout = QVBoxLayout(
            table_widget
        )

        table_layout.setContentsMargins(
            0, 0, 5, 0
        )

        table_layout.setSpacing(
            5
        )

        table_title = QLabel(
            "PIPE DIMENSION DATA"
        )

        table_title.setFixedHeight(
            28
        )

        table_title.setStyleSheet("""
            QLabel {
                color: #16355d;
                font-size: 15px;
                font-weight: bold;
                padding-left: 3px;
            }
        """)

        table_layout.addWidget(
            table_title
        )

        # -----------------------------
        # TABLE
        # -----------------------------

        self.table = QTableWidget()

        self.table.setSortingEnabled(
            True
        )

        self.table.setAlternatingRowColors(
            True
        )

        self.table.setSelectionBehavior(
            QTableWidget.SelectRows
        )

        self.table.setEditTriggers(
            QTableWidget.NoEditTriggers
        )

        self.table.setWordWrap(
            False
        )

        self.table.setShowGrid(
            True
        )

        self.table.verticalHeader().setDefaultSectionSize(
            25
        )

        # Row number width

        self.table.verticalHeader().setFixedWidth(
            36
        )

        # Horizontal header

        self.table.horizontalHeader().setStretchLastSection(
            False
        )

        table_layout.addWidget(
            self.table
        )

        # =================================================
        # RIGHT - WRITE UP
        # =================================================

        writeup_widget = QWidget()

        writeup_layout = QVBoxLayout(
            writeup_widget
        )

        writeup_layout.setContentsMargins(
            5, 0, 0, 0
        )

        writeup_layout.setSpacing(
            5
        )

        writeup_title = QLabel(
            "WELDED AND SEAMLESS WROUGHT STEEL PIPE"
        )

        writeup_title.setFixedHeight(
            28
        )

        writeup_title.setStyleSheet("""
            QLabel {
                color: #16355d;
                font-size: 15px;
                font-weight: bold;
                padding-left: 3px;
            }
        """)

        writeup_layout.addWidget(
            writeup_title
        )

        # -----------------------------
        # WRITE-UP
        # -----------------------------

        self.writeup = QTextBrowser()

        self.writeup.setOpenExternalLinks(
            True
        )

        self.writeup.setHtml("""

        <h2 style="color:#16355d;">
            WELDED AND SEAMLESS WROUGHT STEEL PIPE
        </h2>

        <hr>

        <h3>1 SCOPE</h3>

        <p>
        This Standard covers the standardization of dimensions
        of welded and seamless wrought steel pipe for high or low
        temperatures and pressures.
        </p>

        <p>
        The word <b>"pipe"</b> is used, as distinguished from
        <b>"tube"</b>, to apply to tubular products of dimensions
        commonly used for pipeline and piping systems.
        </p>

        <h3>2 SIZE</h3>

        <p>
        The size of pipe is identified by the dimensionless
        designator nominal pipe size (NPS) and diameter nominal
        (DN).
        </p>

        <p>
        Pipe NPS 12 (DN 300) and smaller have outside diameters
        numerically larger than their corresponding sizes.
        </p>

        <p>
        The manufacture of pipe from NPS 1/8 (DN 6) to
        NPS 12 (DN 300), inclusive, is based on a standardized
        outside diameter (O.D.).
        </p>

        <p>
        The manufacture of pipe NPS 14 (DN 350) and larger is
        based on the O.D. being the same as the nominal pipe size.
        </p>

        <h3>3 REFERENCES</h3>

        <p>
        The following publications are referenced in this
        Standard. Unless otherwise specified, the latest edition
        applies.
        </p>

        <ul>
            <li>API 5L - Specification for Line Pipe</li>
            <li>
                ASME B1.20.1 - Pipe Threads,
                General Purpose (Inch)
            </li>
        </ul>

        <h3>4 MATERIALS</h3>

        <p>
        The dimensional standards for pipe described in this
        Standard are for products covered in ASTM specifications.
        </p>

        <h3>5 WALL THICKNESS</h3>

        <p>
        The nominal wall thicknesses are given in Table 2-1.
        </p>

        <h3>6 WEIGHTS / MASSES</h3>

        <p>
        The nominal weights (masses) of steel pipe are calculated
        values and are given in Table 2-1.
        </p>

        <h4>Nominal Plain End Weight</h4>

        <p>
        The nominal plain end weight, in pounds per foot,
        is calculated using the applicable outside diameter
        and wall thickness.
        </p>

        <p>
        <b>
        W<sub>pe</sub> = 10.69(D - t)t
        </b>
        </p>

        <p><b>Where:</b></p>

        <ul>
            <li><b>D</b> = outside diameter</li>
            <li><b>t</b> = specified wall thickness</li>
            <li>
                <b>W<sub>pe</sub></b> =
                nominal plain end weight
            </li>
        </ul>

        <h4>Nominal Plain End Mass</h4>

        <p>
        The nominal plain end mass, in kilograms per meter,
        is calculated using the applicable outside diameter
        and wall thickness.
        </p>

        <p>
        <b>
        M<sub>pe</sub> = 0.024661(D - t)t
        </b>
        </p>

        <p><b>Where:</b></p>

        <ul>
            <li><b>D</b> = outside diameter</li>
            <li><b>t</b> = specified wall thickness</li>
            <li>
                <b>M<sub>pe</sub></b> =
                nominal plain end mass
            </li>
        </ul>

        <h3>7 PERMISSIBLE VARIATIONS</h3>

        <p>
        Variations in dimensions differ depending upon the method
        of manufacture employed in making the pipe to the various
        specifications available.
        </p>

        <p>
        Permissible variations for dimensions are indicated in
        the applicable specification.
        </p>

        <h3>8 PIPE THREADS</h3>

        <p>
        Unless otherwise specified, the threads of threaded pipe
        shall conform to ASME B1.20.1.
        </p>

        <p>
        Schedules 5 and 10 wall thicknesses do not permit
        threading in accordance with ASME B1.20.1.
        </p>

        <h3>9 WALL-THICKNESS DESIGNATIONS</h3>

        <p>
        The wall-thickness designations Standard (STD),
        Extra Strong (XS), and Double Extra Strong (XXS)
        have been commercially used designations for many years.
        </p>

        <p>
        Schedule numbers were subsequently added as a convenient
        designation for use in ordering pipe.
        </p>

        <h3>10 WALL-THICKNESS SELECTION</h3>

        <p>
        When the selection of wall thickness depends primarily
        upon capacity to resist internal pressure under given
        conditions, the designer shall determine the wall
        thickness suitable for the applicable design conditions
        and governing code.
        </p>

        <br>

        <hr>

        <p>
        <b>Reference:</b>
        ASME B36.10 - Welded and Seamless Wrought Steel Pipe
        </p>

        """)

        writeup_layout.addWidget(
            self.writeup
        )

        # =================================================
        # ADD TO SPLITTER
        # =================================================

        splitter.addWidget(
            table_widget
        )

        splitter.addWidget(
            writeup_widget
        )

        # 60% table / 40% document

        splitter.setStretchFactor(
            0,
            6
        )

        splitter.setStretchFactor(
            1,
            4
        )

        splitter.setSizes([
            1000,
            550
        ])

        main_layout.addWidget(
            splitter,
            1
        )

        # =================================================
        # INITIAL TABLE
        # =================================================

        self.populate_table(
            self.df
        )

        # =================================================
        # COPYRIGHT
        # =================================================

        copyright_title = QLabel(
            "Copyright@ Ashkam Energy Pvt Ltd 2026"
        )

        copyright_title.setFixedHeight(
            28
        )

        copyright_title.setStyleSheet("""
            QLabel {
                color: #16355d;
                font-size: 10px;
                font-weight: bold;
                padding-left: 3px;
            }
        """)

        main_layout.addWidget(
            copyright_title
        )

    # =====================================================
    # DEPENDENT FILTERS
    # =====================================================

    def filter_changed(self, changed_column):

        # Order of dependency
        filter_order = [
            "NPS",
            "DN",
            "Identification",
            "Schedule No."
        ]

        if changed_column not in filter_order:
            self.apply_filters()
            return

        changed_index = filter_order.index(
            changed_column
        )

        # -------------------------------------------------
        # Update filters AFTER changed filter
        # -------------------------------------------------

        for column in filter_order[changed_index + 1:]:

            if column not in self.filter_boxes:
                continue

            combo = self.filter_boxes[column]

            # Save current selection
            current_value = combo.currentText()

            # Start with complete dataframe
            temp_df = self.df.copy()

            # -------------------------------------------------
            # Apply all previous filters
            # -------------------------------------------------

            for previous_column in filter_order:

                if previous_column not in self.filter_boxes:
                    continue

                if previous_column == column:
                    break

                selected = self.filter_boxes[
                    previous_column
                ].currentText()

                if selected != "All":

                    temp_df = temp_df[
                        temp_df[previous_column]
                        .astype(str)
                        .str.strip()
                        == selected
                    ]

            # -------------------------------------------------
            # Get valid values
            # -------------------------------------------------

            values = (
                temp_df[column]
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )

            values = sorted(
                [
                    value
                    for value in values
                    if value != ""
                ],
                key=str.lower
            )

            # -------------------------------------------------
            # Rebuild combo
            # -------------------------------------------------

            combo.blockSignals(True)

            combo.clear()

            combo.addItem(
                "All"
            )

            combo.addItems(
                values
            )

            # Restore old selection if still valid
            if current_value in values:

                combo.setCurrentText(
                    current_value
                )

            else:

                combo.setCurrentIndex(
                    0
                )

            combo.blockSignals(False)

        # -------------------------------------------------
        # Apply filters to table
        # -------------------------------------------------

        self.apply_filters()

    # =====================================================
    # APPLY FILTERS
    # =====================================================

    def apply_filters(self):

        filtered = self.df.copy()

        for column, combo in self.filter_boxes.items():

            selected_value = combo.currentText()

            if selected_value != "All":

                filtered = filtered[
                    filtered[column]
                    .astype(str)
                    .str.strip()
                    == selected_value
                ]

        self.filtered_df = filtered

        self.populate_table(
            filtered
        )

    # =====================================================
    # CLEAR FILTERS
    # =====================================================

    def clear_filters(self):

        # Reset all combos
        for combo in self.filter_boxes.values():

            combo.blockSignals(True)

            combo.setCurrentIndex(
                0
            )

            combo.blockSignals(False)

        # Restore complete values
        self.update_all_filter_values()

        # Restore complete dataframe
        self.filtered_df = self.df.copy()

        self.populate_table(
            self.df
        )

    # =====================================================
    # RESTORE ALL FILTER VALUES
    # =====================================================

    def update_all_filter_values(self):

        for column, combo in self.filter_boxes.items():

            combo.blockSignals(True)

            combo.clear()

            combo.addItem(
                "All"
            )

            values = (
                self.df[column]
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )

            values = sorted(
                [
                    value
                    for value in values
                    if value != ""
                ],
                key=str.lower
            )

            combo.addItems(
                values
            )

            combo.blockSignals(False)

    # =====================================================
    # POPULATE TABLE
    # =====================================================

    def populate_table(self, dataframe):

        self.table.setSortingEnabled(
            False
        )

        self.table.clear()

        # Number of rows

        self.table.setRowCount(
            len(dataframe)
        )

        # Number of columns

        self.table.setColumnCount(
            len(dataframe.columns)
        )

        # Headers

        self.table.setHorizontalHeaderLabels(
            list(dataframe.columns)
        )

        # -------------------------------------------------
        # DATA
        # -------------------------------------------------

        for row_index, row in enumerate(
            dataframe.itertuples(
                index=False,
                name=None
            )
        ):

            for column_index, value in enumerate(
                row
            ):

                item = QTableWidgetItem(
                    str(value)
                )

                item.setTextAlignment(
                    Qt.AlignCenter
                )

                self.table.setItem(
                    row_index,
                    column_index,
                    item
                )

        # =================================================
        # COLUMN SIZING
        # =================================================

        header = self.table.horizontalHeader()

        header.setSectionResizeMode(
            QHeaderView.ResizeToContents
        )

        # Prevent extremely narrow columns

        for column_index in range(
            self.table.columnCount()
        ):

            width = self.table.columnWidth(
                column_index
            )

            if width < 70:

                self.table.setColumnWidth(
                    column_index,
                    70
                )

        self.table.setSortingEnabled(
            True
        )

        # =================================================
        # RECORD COUNT
        # =================================================

        self.record_label.setText(
            f"Records: {len(dataframe):,} / "
            f"{len(self.df):,}"
        )


# =========================================================
# APPLICATION START
# =========================================================

if __name__ == "__main__":

    app = QApplication(
        sys.argv
    )

    app.setStyle(
        "Fusion"
    )

    window = PipeDimensionApp()

    window.showMaximized()

    sys.exit(
        app.exec_()
    )