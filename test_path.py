import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(BASE_DIR, "class_9A_master.xlsx")

print("BASE_DIR:", BASE_DIR)
print("Expected Excel path:", excel_path)
print("Exists?", os.path.exists(excel_path))
print("Directory listing:", os.listdir(BASE_DIR))
