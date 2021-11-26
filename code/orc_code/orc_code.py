# ocr_card.py
import os
import sys
from PIL import Image
import pyocr
import pyocr.builders

# 1.インストール済みのTesseractのパスを通す
path_tesseract = "C:\\Program Files\\Tesseract-OCR"
if path_tesseract not in os.environ["PATH"].split(os.pathsep):
    os.environ["PATH"] += os.pathsep + path_tesseract

# 2.OCRエンジンの取得
tools = pyocr.get_available_tools()
tool = tools[0]

# 3.原稿画像の読み込み
img_org = Image.open("./code_image/1.png")

# 4.ＯＣＲ実行
builder = pyocr.builders.TextBuilder(tesseract_layout=7)
result = tool.image_to_string(img_org, lang="eng+jpn", builder=builder)

print(result)
print("----------------")
exe_path = os.path.join("C:","Program Files","Tesseract-OCR")
sys.path.append(exe_path)
# print(sys.path)
sys_cmd = "tesseract.exe ./code_image/1.png ./result3 -l eng -c tessedit_char_whitelist=0123456789"
eng_result=os.system(sys_cmd)

print(eng_result)

sys_cmd2 = "tesseract.exe ./code_image/1.png ./result3 -l jpn -c tessedit_char_whitelist=0123456789"
eng_result2=os.system(sys_cmd2)

print(eng_result2)