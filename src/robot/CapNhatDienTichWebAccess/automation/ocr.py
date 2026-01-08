import os
import re
import shutil
import tempfile

import cv2
import easyocr
import numpy as np
import pytesseract
import torch
from pdf2image import convert_from_path
from PIL import Image
from transformers import AutoImageProcessor, TableTransformerForObjectDetection


class OCR:
    def __init__(
        self,
        logger=None,
        tesseract_path: str = "C:/Program Files/Tesseract-OCR/tesseract.exe",
        poppler_path: str = "bin",
    ):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
        self.processor = AutoImageProcessor.from_pretrained("microsoft/table-transformer-detection")
        self.model = TableTransformerForObjectDetection.from_pretrained("microsoft/table-transformer-detection")
        self.poppler_path = poppler_path
        self.reader = easyocr.Reader(["ja", "en"], gpu=False)
        self.logger = logger

    def preprocess_image(self, img_path: str) -> None:
        img = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)

        # 1️⃣ Giảm nhiễu nhẹ (Gaussian blur)
        img = cv2.GaussianBlur(img, (3, 3), 0)

        # 2️⃣ Tăng tương phản (adaptive threshold)
        img = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 11)

        # 3️⃣ Morphological opening để loại bỏ nhiễu nhỏ
        kernel = np.ones((1, 1), np.uint8)
        img = cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel)

        # 4️⃣ Lưu đè ảnh đã xử lý
        cv2.imwrite(img_path, img)

    def pdf_2_png(self, pdf_path: str, output_folder: str | None = None) -> list[str]:
        """Convert PDF sang PNG, sau đó preprocess từng ảnh"""
        if output_folder:
            os.makedirs(output_folder, exist_ok=True)
        else:
            output_folder = os.path.abspath(".")

        # Chuyển PDF -> PNG
        images = convert_from_path(
            pdf_path,
            poppler_path=self.poppler_path,
            output_folder=output_folder,
            fmt="png",
            dpi=600,
        )

        processed_paths = []
        for i, img in enumerate(images):
            img_name = f"page_{i+1}.png"
            img_path = os.path.join(output_folder, img_name)

            img.save(img_path, "PNG")  # Lưu ảnh gốc
            self.preprocess_image(img_path)  # Tiền xử lý ảnh
            processed_paths.append(img_path)

        return processed_paths

    def get_area(self, pdf_path: str) -> float | None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output_folder = temp_dir
            self.logger.info(f"Processing {os.path.basename(pdf_path)}")
            images = self.pdf_2_png(pdf_path=pdf_path, output_folder=output_folder)
            area = []
            for image_path in images:
                image = Image.open(image_path).convert("RGB")

                # 2️⃣ Tiền xử lý ảnh cho mô hình TableTransformer
                inputs = self.processor(images=image, return_tensors="pt")
                outputs = self.model(**inputs)

                # 3️⃣ Xử lý kết quả (bảng được phát hiện)
                target_sizes = torch.tensor([image.size[::-1]])  # (height, width)
                results = self.processor.post_process_object_detection(
                    outputs, threshold=0.7, target_sizes=target_sizes
                )[0]

                # 🟩 Lọc ra các vùng có nhãn là "table"
                tables = []
                for box, score, label in zip(results["boxes"], results["scores"], results["labels"]):
                    if self.model.config.id2label[label.item()] == "table":
                        tables.append((score.item(), box))

                if not tables:
                    results = self.reader.readtext(image_path)
                    text = " ".join([text for _, text, _ in results]).replace("O", "0")
                    pattern = r"(?<!\d)(\d+(?:\s*[., ]\s*\d+)?)[ ]*m(?:\?2|z2|²|2|\?|z|e)(?![A-Za-z0-9])"
                    if match := re.findall(pattern, text):
                        value = match[-1]
                        value = value.replace(",", ".")
                        value = value.replace(" ", ".")
                        value = re.sub(r"\.{2,}", ".", value)
                        try:
                            if float(re.sub(r"[^\d.]", "", value)) and float(re.sub(r"[^\d.]", "", value)) < 1000:
                                area.append(float(re.sub(r"[^\d.]", "", value)))
                                self.logger.info(f"EasyOCR: Extract {value} from {repr(text)}")
                            else:
                                return None
                        except ValueError:
                            self.logger.warning(f"EasyOCR: Failed to convert {value}")
                    continue  # không có bảng nào

                # 🟦 Chỉ chọn bảng có score cao nhất
                _, best_box = max(tables, key=lambda x: x[0])
                x_min, y_min, x_max, y_max = best_box.int().tolist()

                # 🟧 Cắt vùng bảng có độ tin cậy cao nhất
                table_crop = image.crop((x_min, y_min, x_max, y_max))

                # Image to String
                text = pytesseract.image_to_string(table_crop, lang="eng+jpn").replace("O", "0")
                pattern = r"(?<!\d)(\d+(?:\s*[., ]\s*\d+)?)[ ]*m(?:\?2|z2|²|2|\?|z|e)(?![A-Za-z0-9])"
                if match := re.findall(pattern, text):
                    value = match[-1]
                    value = value.replace(",", ".")
                    value = value.replace(" ", ".")
                    value = re.sub(r"\.{2,}", ".", value)
                    try:
                        if float(re.sub(r"[^\d.]", "", value)) and float(re.sub(r"[^\d.]", "", value)) < 1000:
                            area.append(float(re.sub(r"[^\d.]", "", value)))
                            self.logger.info(f"Pytesseract: Extract {value} from {repr(text)}")
                        else:
                            return None
                    except ValueError:
                        self.logger.warning(f"Pytesseract: Failed to convert {value}")
                else:
                    result = self.reader.readtext(np.array(table_crop), detail=1)
                    text = " ".join([t for _, t, _ in result]).replace("O", "0")
                    if match := re.findall(pattern, text):
                        value = match[-1]
                        value = value.replace(",", ".")
                        value = value.replace(" ", ".")
                        value = re.sub(r"\.{2,}", ".", value)
                        try:
                            if float(re.sub(r"[^\d.]", "", value)) and float(re.sub(r"[^\d.]", "", value)) < 1000:
                                area.append(float(re.sub(r"[^\d.]", "", value)))
                                self.logger.info(f"EasyOCR: Extract {value} from {repr(text)}")
                            else:
                                return None
                        except ValueError:
                            self.logger.warning(f"Failed to convert {value}")
            shutil.rmtree(output_folder)
            if len(area) != len(images):
                return None
            return round(sum(area), 2)
