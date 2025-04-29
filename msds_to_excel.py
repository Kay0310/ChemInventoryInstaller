
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import pandas as pd
import re
import os
import shutil

# === 1. 파일 경로 설정 ===
image_folder = "/mnt/data/"  # 이미지가 저장된 폴더
backup_folder = os.path.join(image_folder, "backup")  # 백업 폴더 생성

# 파일 이름 규칙:
# 각 제품: <제품명>.product.png (예: product1.product.png)
# 그에 해당하는 화학물질 표: <제품명>.table.png (예: product1.table.png)

# === 2. OCR용 이미지 전처리 ===
def preprocess_image(image):
    image = image.convert('L')  # 흑백 변환
    image = image.filter(ImageFilter.MedianFilter())
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)
    return image

# === 3. 제품명 추출 ===
def extract_product_name(image_path):
    image = Image.open(image_path)
    image = preprocess_image(image)
    text = pytesseract.image_to_string(image)
    match = re.search(r"Product identifier.*?:\s*(.*)", text)
    return match.group(1).strip() if match else "UNKNOWN"

# === 4. 화학물질 데이터 추출 ===
def extract_chemical_data(image_path):
    image = Image.open(image_path)
    image = preprocess_image(image)
    text = pytesseract.image_to_string(image)
    lines = text.split("\n")
    data = []
    for line in lines:
        if re.search(r"\d{2,5}-\d{2}-\d", line):  # CAS 번호 포함
            parts = re.split(r"\s{2,}|\t+", line.strip())
            if len(parts) >= 4:
                data.append(parts[:4])
            elif len(parts) == 3:
                data.append([parts[0], "None", parts[1], parts[2]])
    return data

# === 5. 엑셀 저장 ===
def save_to_excel(product_name, chemical_data, output_path):
    df = pd.DataFrame(chemical_data, columns=[
        "3.화학물질명(필수)", "4.관용명및이명", "5.CAS번호", "21.함유량"
    ])
    df.insert(0, "2.제품명", product_name)
    df.to_excel(output_path, index=False)
    print(f"[INFO] 개별 파일 저장 완료: {output_path}")

# === 6. 전체 처리 ===
def process_all_images(folder):
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)

    combined_data = []
    for file in os.listdir(folder):
        if file.endswith(".product.png"):
            basename = file.replace(".product.png", "")
            product_path = os.path.join(folder, f"{basename}.product.png")
            table_path = os.path.join(folder, f"{basename}.table.png")

            if not os.path.exists(table_path):
                print(f"[WARN] 테이블 이미지 없음: {table_path}")
                continue

            # 파일 백업
            shutil.copy(product_path, os.path.join(backup_folder, os.path.basename(product_path)))
            shutil.copy(table_path, os.path.join(backup_folder, os.path.basename(table_path)))

            product_name = extract_product_name(product_path)
            chemical_data = extract_chemical_data(table_path)

            # 개별 엑셀 저장
            output_file = os.path.join(folder, f"{basename}_inventory.xlsx")
            save_to_excel(product_name, chemical_data, output_file)

            # 병합용 데이터 추가
            for row in chemical_data:
                combined_data.append([product_name] + row)

    # 전체 병합된 엑셀 저장
    final_df = pd.DataFrame(combined_data, columns=[
        "2.제품명", "3.화학물질명(필수)", "4.관용명및이명", "5.CAS번호", "21.함유량"
    ])
    merged_file = os.path.join(folder, "merged_inventory.xlsx")
    final_df.to_excel(merged_file, index=False)
    print(f"[DONE] 모든 데이터를 병합하여 저장 완료: {merged_file}")

# 실행
if __name__ == "__main__":
    process_all_images(image_folder)
