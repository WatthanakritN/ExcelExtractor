from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLineEdit, QLabel, QMessageBox
)
import sys
import os
from datetime import datetime
from extractor import ExcelExtractor
import pandas as pd  


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Extractor")
        self.resize(400, 250)

        layout = QVBoxLayout()

        self.file_label = QLabel("No files selected")
        layout.addWidget(self.file_label)

        btn_select = QPushButton("Select Excel Files")
        btn_select.clicked.connect(self.select_files)
        layout.addWidget(btn_select)

        # ตั้งชื่อไฟล์ output
        self.output_name = QLineEdit()
        self.output_name.setPlaceholderText("Output filename (ไม่ต้องใส่ .xlsx)")
        layout.addWidget(self.output_name)

        btn_extract = QPushButton("Extract")
        btn_extract.clicked.connect(self.extract)
        layout.addWidget(btn_extract)

        self.setLayout(layout)
        self.files = []

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Excel files",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if files:
            self.files = files
            self.file_label.setText(f"{len(files)} files selected")

    def extract(self):
        if not self.files:
            QMessageBox.warning(self, "Warning", "Please select files first")
            return

    #ตั้งชื่อไฟล์ output
        filename = self.output_name.text().strip()
        if not filename:
            filename = f"output_{datetime.now():%Y%m%d_%H%M%S}"

        output_path = os.path.join(os.getcwd(), f"{filename}.xlsx")

        extractor = ExcelExtractor("System Daily Report Entry") #เปลี่ยนชื่อ sheet ที่ต้องการข้อมูล

        df_list = []
        skipped_files = []

        for f in self.files:
            try:
                df = extractor.extract_single_file(f)
                if df.empty:
                    skipped_files.append(os.path.basename(f))
                else:
                    df_list.append(df)
            except Exception as e:
                skipped_files.append(os.path.basename(f))
                print(f"[ERROR] {f}: {e}")

        if not df_list:
            QMessageBox.warning(self, "Warning", "No valid data found")
            return

        # ใช้ concat แทน append
        result = pd.concat(df_list, ignore_index=True)
        result.to_excel(output_path, index=False)

        msg = f"Exported file:\n{output_path}"
        if skipped_files:
            msg += "\n\nSkipped files:\n" + "\n".join(skipped_files)

        QMessageBox.information(self, "Success", msg)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
