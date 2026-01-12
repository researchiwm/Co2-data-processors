import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QLineEdit, QMessageBox, QProgressBar,
    QTextEdit, QGroupBox, QFormLayout, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPalette, QColor

GAS_CONSTANT_R = 8.314
ATMOSPHERIC_PRESSURE_KPA = 101.325
MOLECULAR_WEIGHT_CO2 = 44.01
MOLECULAR_WEIGHT_C = 12.01
DEFAULT_JAR_RADIUS_CM = 9.6 / 2


def convert_ppm_to_mg_CO2(ppm, temp_C):
    T_K = temp_C + 273.15
    mol_L = (ppm * ATMOSPHERIC_PRESSURE_KPA) / (GAS_CONSTANT_R * T_K * 1e6)
    return mol_L * MOLECULAR_WEIGHT_CO2 * 1000


def convert_mg_CO2_to_mg_C(mg_CO2):
    return mg_CO2 * (MOLECULAR_WEIGHT_C / MOLECULAR_WEIGHT_CO2)


def calculate_total_porosity(bulk_density, particle_density):
    return 1 - (bulk_density / particle_density)


class FileDropLineEdit(QLineEdit):
    """QLineEdit that accepts drag and drop of files."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptDrops(True)
        self.setFont(QFont("Segoe UI", 11))
        self.setStyleSheet(
            "QLineEdit { border: 2px dashed #aaa; padding: 6px; border-radius: 10px; }"
            "QLineEdit:focus { border-color: #3a86ff; }"
        )

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) == 1 and urls[0].toLocalFile().lower().endswith('.xlsx'):
                event.acceptProposedAction()
                self.setStyleSheet(
                    "QLineEdit { border: 2px dashed #3a86ff; background-color: #e6f0ff; padding: 6px; border-radius: 10px; }"
                )
            else:
                event.ignore()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(
            "QLineEdit { border: 2px dashed #aaa; padding: 6px; border-radius: 10px; }"
        )

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) == 1:
                filepath = urls[0].toLocalFile()
                if filepath.lower().endswith('.xlsx'):
                    self.setText(filepath)
        self.setStyleSheet(
            "QLineEdit { border: 2px dashed #aaa; padding: 6px; border-radius: 10px; }"
        )


class ProcessorThread(QThread):
    progress_done = pyqtSignal()
    progress_failed = pyqtSignal(str)
    progress_status = pyqtSignal(str)
    progress_value = pyqtSignal(int)

    def __init__(self, sensor_path, metadata_path, output_path):
        super().__init__()
        self.sensor_path = sensor_path
        self.metadata_path = metadata_path
        self.output_path = output_path

    def run(self):
        try:
            self.progress_status.emit("🟡 Loading files...")
            self.progress_value.emit(10)

            metadata = pd.read_excel(self.metadata_path)
            raw = pd.read_excel(self.sensor_path, header=None)

            block_row = raw.iloc[0].ffill()
            sensor_row = raw.iloc[1].ffill()
            attr_row = raw.iloc[2]

            new_columns = ['Index', 'Date', 'Time']
            for i in range(3, raw.shape[1]):
                block = str(block_row[i]).strip()
                sensor = str(sensor_row[i]).strip()
                attr = str(attr_row[i]).strip()
                new_columns.append(f"{block}_Sensor{sensor}_{attr}")

            raw_data = raw.iloc[4:].copy()
            raw_data.columns = new_columns

            df = pd.melt(raw_data, id_vars=["Index", "Date", "Time"], var_name="Sensor_Attr", value_name="Value")
            df[['Block', 'Sensor', 'Attribute']] = df['Sensor_Attr'].str.extract(r'(Block\d+)_Sensor(\d+)_(\w+)')
            non_numeric_values = df[~df["Value"].apply(lambda x: pd.api.types.is_number(x))]
            non_numeric_count = len(non_numeric_values)
            df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
            if non_numeric_count > 0:
                self.progress_status.emit(f"⚠️ {non_numeric_count} non-numeric values found and set to NaN.")
            df = df.pivot_table(index=["Date", "Time", "Block", "Sensor"], columns="Attribute", values="Value").reset_index()
            df.columns.name = None
            df.columns = [str(c).strip() for c in df.columns]

            df['Sensor_ID'] = df['Sensor'].apply(lambda s: f"Sensor{s}")
            df = df.merge(metadata, on="Sensor_ID", how="inner")
            df = df[df["Block"] == df["Block_ID"].apply(lambda b: f"Block{int(b)}")]

            for col in ["CO2", "Temp", "Wet_Mass_g", "Bulk_Density", "Jar_Volume_L", "AFP", "OM_Mass_g",
                        "Interval_min", "MC", "Particle_Density", "Compost_Height_cm"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').round(6)

            df["Temp"] = df["Temp"].fillna(25.0)

            if "Jar_Radius_cm" in df.columns:
                df["Jar_Radius_cm"] = pd.to_numeric(df["Jar_Radius_cm"], errors="coerce").fillna(DEFAULT_JAR_RADIUS_CM).round(6)
            else:
                df["Jar_Radius_cm"] = round(DEFAULT_JAR_RADIUS_CM, 6)

            self.progress_status.emit("🧮 Calculating compost volume & gas values...")
            self.progress_value.emit(50)

            df["Compost_Vol_L"] = np.where(
                df["Compost_Height_cm"].notna(),
                np.pi * (df["Jar_Radius_cm"] ** 2) * df["Compost_Height_cm"] / 1000,
                df["Wet_Mass_g"] / df["Bulk_Density"] / 1000
            ).round(6)

            df["V_CO2_L"] = ((df["Jar_Volume_L"] - df["Compost_Vol_L"]) + (df["AFP"] * df["Compost_Vol_L"])).round(6)
            df["mg_CO2"] = df.apply(lambda row: convert_ppm_to_mg_CO2(row["CO2"], row["Temp"]) * row["V_CO2_L"], axis=1).round(6)
            df["mg_C"] = convert_mg_CO2_to_mg_C(df["mg_CO2"]).round(6)
            df["Rate_mgC_per_gOM_per_hr"] = (df["mg_C"] / (df["OM_Mass_g"] * df["Interval_min"] / 60)).round(6)

            if "Particle_Density" in df.columns:
                df["Total_Porosity"] = calculate_total_porosity(df["Bulk_Density"], df["Particle_Density"]).round(6)

            export_df = df[["Date", "Time", "Block", "Sensor", "mg_CO2", "mg_C"]].copy()
            export_df["Block"] = export_df["Block"].str.extract(r'(\d+)')
            export_df["Sensor"] = export_df["Sensor"].astype(str)

            export_df_long = pd.melt(export_df, id_vars=["Date", "Time", "Block", "Sensor"],
                                     value_vars=["mg_CO2", "mg_C"], var_name="Metric", value_name="Value")
            export_df_long["ColName"] = "Block" + export_df_long["Block"] + "_Sensor" + export_df_long["Sensor"] + "_" + export_df_long["Metric"]
            wide = export_df_long.pivot_table(index=["Date", "Time"], columns="ColName", values="Value")
            wide = wide.loc[:, ~wide.columns.duplicated()]
            wide = wide.reset_index()
            wide["Date"] = pd.to_datetime(wide["Date"], errors="coerce").dt.strftime('%Y-%m-%d')

            def sort_key(col):
                if col in ["Date", "Time"]:
                    return (0, 0, col)
                parts = col.split("_")
                return (int(parts[0].replace("Block", "")), int(parts[1].replace("Sensor", "")), parts[2])

            wide = wide[sorted(wide.columns, key=sort_key)]

            wb = Workbook()
            ws = wb.active
            ws.title = "Sensor Output"

            block_row, sensor_row, metric_row = [], [], []
            for col in wide.columns:
                if col in ["Date", "Time"]:
                    block_row.append("")
                    sensor_row.append("")
                    metric_row.append(col)
                else:
                    parts = col.split("_")
                    block_row.append(parts[0])
                    sensor_row.append(parts[1])
                    metric_row.append("_".join(parts[2:]))

            ws.append(block_row)
            ws.append(sensor_row)
            ws.append(metric_row)

            col_start = 1
            while col_start <= len(metric_row):
                if metric_row[col_start - 1] in ["Date", "Time"]:
                    col_start += 1
                    continue
                block = block_row[col_start - 1]
                block_end = col_start
                while block_end <= len(metric_row) and block_row[block_end - 1] == block:
                    block_end += 1
                block_end -= 1
                ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=block_end)
                ws.cell(row=1, column=col_start).value = block
                ws.cell(row=1, column=col_start).alignment = Alignment(horizontal='center')

                sensor_col = col_start
                while sensor_col <= block_end:
                    sensor = sensor_row[sensor_col - 1]
                    sensor_end = sensor_col
                    while sensor_end <= block_end and sensor_row[sensor_end - 1] == sensor:
                        sensor_end += 1
                    sensor_end -= 1
                    ws.merge_cells(start_row=2, start_column=sensor_col, end_row=2, end_column=sensor_end)
                    ws.cell(row=2, column=sensor_col).value = sensor
                    ws.cell(row=2, column=sensor_col).alignment = Alignment(horizontal='center')
                    sensor_col = sensor_end + 1
                col_start = block_end + 1

            for col in range(1, len(metric_row) + 1):
                ws.cell(row=3, column=col).alignment = Alignment(horizontal='center')

            for r in wide.itertuples(index=False):
                ws.append(r)

            wb.save(self.output_path)
            self.progress_value.emit(100)
            self.progress_status.emit("✅ Done! Output saved.")
            self.progress_done.emit()

        except Exception as e:
            self.progress_value.emit(0)
            self.progress_status.emit(f"❌ Error: {str(e)}")
            self.progress_failed.emit(str(e))


class CO2App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🌿 CO₂ Multi-Sensor Processor")
        self.setFixedSize(720, 560)

        self.setFont(QFont("Segoe UI", 11))

        # Input fields
        self.sensor_path_input = FileDropLineEdit()
        self.metadata_path_input = FileDropLineEdit()

        # Output log
        self.output_log = QTextEdit()
        self.output_log.setReadOnly(True)
        self.output_log.setStyleSheet(
            "background-color: #f9f9f9; border: 1px solid #ddd; border-radius: 8px; padding: 8px;"
        )
        self.output_log.setFont(QFont("Consolas", 10))

        # Progress bar
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setMinimum(0)
        self.progress.setMaximum(100)
        self.progress.setTextVisible(True)
        self.progress.setFormat("%p%")

        # Run button
        self.run_button = QPushButton("▶ Run Processor")
        self.run_button.setEnabled(False)
        self.run_button.setCursor(Qt.PointingHandCursor)
        self.run_button.setStyleSheet("""
            QPushButton {
                background-color: #3a86ff;
                color: white;
                border-radius: 14px;
                padding: 12px;
                font-weight: bold;
                font-size: 14pt;
            }
            QPushButton:disabled {
                background-color: #a0bff8;
                color: #ddd;
            }
            QPushButton:hover:!disabled {
                background-color: #265fcc;
            }
        """)

        # Instruction note
        self.instruction_label = QLabel(
            "⚠️ Please ensure your sensor and metadata files are in the specified Excel (.xlsx) format."
        )
        self.instruction_label.setStyleSheet("color: #555; font-style: italic;")
        self.instruction_label.setWordWrap(True)

        # Footer
        self.footer_label = QLabel("🌱 Innovative Waste Management | Developed by Akash")
        self.footer_label.setAlignment(Qt.AlignCenter)
        self.footer_label.setStyleSheet("color: #888; font-size: 9pt; font-style: italic;")

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        header = QLabel("Multi-Sensor CO₂ Processor")
        header.setAlignment(Qt.AlignCenter)
        header.setFont(QFont("Segoe UI", 18, QFont.Bold))
        header.setStyleSheet("color: #1a237e; margin-bottom: 12px;")
        main_layout.addWidget(header)

        # Input group
        input_group = QGroupBox("Input Files (Drag & Drop Supported)")
        input_layout = QFormLayout()
        input_layout.addRow("Sensor File:", self._build_file_row(self.sensor_path_input))
        input_layout.addRow("Metadata File:", self._build_file_row(self.metadata_path_input))
        input_group.setLayout(input_layout)
        main_layout.addWidget(input_group)

        main_layout.addWidget(self.instruction_label)
        main_layout.addSpacing(10)

        main_layout.addWidget(self.run_button)
        main_layout.addSpacing(10)
        main_layout.addWidget(self.progress)
        main_layout.addSpacing(10)

        main_layout.addWidget(QLabel("Output Log:"))
        main_layout.addWidget(self.output_log)

        main_layout.addSpacing(15)
        main_layout.addWidget(self.footer_label)

        self.setLayout(main_layout)

        # Connections
        self.sensor_path_input.textChanged.connect(self._enable_run)
        self.metadata_path_input.textChanged.connect(self._enable_run)
        self.run_button.clicked.connect(self.run_processor)

    def _build_file_row(self, input_field):
        container = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(input_field)

        browse_button = QPushButton("📂")
        browse_button.setToolTip("Browse for Excel file")
        browse_button.setFixedWidth(40)
        browse_button.setCursor(Qt.PointingHandCursor)
        browse_button.setStyleSheet("""
            QPushButton {
                background-color: #e0e0e0;
                border-radius: 10px;
                font-size: 16pt;
            }
            QPushButton:hover {
                background-color: #b0c4ff;
            }
        """)
        browse_button.clicked.connect(lambda: self._browse_file(input_field))
        layout.addWidget(browse_button)

        container.setLayout(layout)
        return container

    def _browse_file(self, target_field):
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if path:
            target_field.setText(path)

    def _enable_run(self):
        self.run_button.setEnabled(bool(self.sensor_path_input.text()) and bool(self.metadata_path_input.text()))

    def run_processor(self):
        sensor_path = self.sensor_path_input.text()
        metadata_path = self.metadata_path_input.text()

        if not sensor_path.lower().endswith('.xlsx') or not metadata_path.lower().endswith('.xlsx'):
            QMessageBox.warning(self, "Invalid Files", "Please provide valid .xlsx files for both Sensor and Metadata.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Output File As",
            f"output_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
            "Excel Files (*.xlsx)"
        )
        if not output_path:
            return

        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.output_log.clear()

        self.thread = ProcessorThread(sensor_path, metadata_path, output_path)
        self.thread.progress_status.connect(self.output_log.append)
        self.thread.progress_value.connect(self.progress.setValue)
        self.thread.progress_done.connect(self.on_success)
        self.thread.progress_failed.connect(self.on_error)
        self.thread.start()

        self.run_button.setEnabled(False)

    def on_success(self):
        self.progress.setVisible(False)
        self.output_log.append("✅ Processing complete!")
        QMessageBox.information(self, "Success", "Processing complete!")
        self.run_button.setEnabled(True)

    def on_error(self, message):
        self.progress.setVisible(False)
        self.output_log.append(f"❌ Error: {message}")
        QMessageBox.critical(self, "Error", message)
        self.run_button.setEnabled(True)


if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)

    # Optional: set Fusion style for modern look
    app.setStyle("Fusion")

    # Optional: tweak palette for subtle colors
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#fafafa"))
    palette.setColor(QPalette.Base, QColor("#ffffff"))
    palette.setColor(QPalette.AlternateBase, QColor("#f0f0f0"))
    palette.setColor(QPalette.ToolTipBase, QColor("#e0e0e0"))
    palette.setColor(QPalette.ToolTipText, Qt.black)
    palette.setColor(QPalette.Text, QColor("#212121"))
    palette.setColor(QPalette.Button, QColor("#3a86ff"))
    palette.setColor(QPalette.ButtonText, Qt.white)
    palette.setColor(QPalette.Highlight, QColor("#265fcc"))
    palette.setColor(QPalette.HighlightedText, Qt.white)
    app.setPalette(palette)

    window = CO2App()
    window.show()
    sys.exit(app.exec_())
