#เอารันเทส ต้องสร้างไฟล์ input_excel ไว้เก็บ excel file เพื่อทดสอบ
from extractor import ExcelExtractor
import glob

extractor = ExcelExtractor(sheet_name="System Daily Report Entry")

files = glob.glob("input_excel/*.xlsx")

df = extractor.extract_multiple_files(files)

print(df)
df.to_excel("output.xlsx", index=False)
