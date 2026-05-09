import os
import io
import re
import traceback
from urllib.parse import urlparse, unquote

import numpy as np
import torch
import requests
from PIL import Image, ImageOps

try:
    import pandas as pd
except ImportError:
    print("\n错误：QC.LoadImageFromExcelURL 节点缺少 Pandas 库。请安装: pip install pandas openpyxl\n")
    pd = None


class WorkflowStopRequested(Exception):
    """指示工作流应在此节点后停止的特殊异常。"""
    pass


def _col_letter_to_index(letter):
    letter = letter.strip().upper()
    if not letter:
        raise ValueError("列字母不能为空。")
    index = 0
    for char in letter:
        if not 'A' <= char <= 'Z':
            raise ValueError(f"列字母 '{letter}' 中包含无效字符 '{char}'。")
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1


def _sanitize_filename(name):
    name = re.sub(r'[\\/:*?"<>|]+', '_', name).strip(' .')
    return name or "image"


def _name_from_url(url, ext):
    try:
        path = urlparse(url).path
        base = os.path.basename(unquote(path)) or "image"
        stem, _ = os.path.splitext(base)
    except Exception:
        stem = "image"
    return f"{_sanitize_filename(stem)}.{ext}"


def _pil_to_tensor(image, output_format):
    """按输出格式规范化 PIL 图像，再转为 ComfyUI 的 IMAGE/MASK 张量。"""
    try:
        image = ImageOps.exif_transpose(image)
    except Exception:
        pass

    if image.mode == 'I':
        image = image.point(lambda i: i * (1 / 255))

    has_alpha = 'A' in image.getbands()
    alpha_channel = image.getchannel('A') if has_alpha else None

    # jpg 不支持透明通道：拼合到白底
    if output_format == "jpg" and has_alpha:
        bg = Image.new("RGB", image.size, (255, 255, 255))
        bg.paste(image.convert("RGBA"), mask=alpha_channel)
        image = bg
        has_alpha = False
        alpha_channel = None
    elif image.mode != 'RGB':
        image = image.convert('RGB')

    image_np = np.array(image).astype(np.float32) / 255.0
    image_tensor = torch.from_numpy(image_np).unsqueeze(0)

    if has_alpha:
        mask_np = np.array(alpha_channel).astype(np.float32) / 255.0
        mask_tensor = 1.0 - torch.from_numpy(mask_np)
    else:
        h, w = image_np.shape[:2]
        mask_tensor = torch.zeros((h, w), dtype=torch.float32)
    mask_tensor = mask_tensor.unsqueeze(0)

    return image_tensor, mask_tensor


class LoadImageFromExcelURL:
    """
    QC.LoadImageFromExcelURL

    从 Excel 指定列读取 URL，下载该图片并输出 (image, image_name, url)。
    start_row_number 会在每次运行后自动增加，超出 end_row_number 则停止工作流。
    """

    def __init__(self):
        pass

    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "excel_file_path": ("STRING", {"multiline": False, "default": "input/example.xlsx"}),
                "url_column": ("STRING", {"multiline": False, "default": "A", "description": "存放图片URL的列字母（单列，如 A / B / AA）"}),
                "start_row_number": ("INT", {"default": 1, "min": 1, "max": 100000, "step": 1, "control_after_generate": True}),
                "end_row_number": ("INT", {"default": -1, "min": -1, "max": 100000, "step": 1, "description": "-1 表示读到最后一行"}),
                "output_format": (["png", "webp", "jpg"], {"default": "png"}),
                "timeout": ("INT", {"default": 30, "min": 5, "max": 300, "step": 5}),
                "on_error": (["raise", "black", "white"], {"default": "black"}),
                "fallback_size": ("INT", {"default": 512, "min": 16, "max": 8192, "step": 16}),
            },
            "optional": {
                "sheet_name": ("STRING", {"multiline": False, "default": "0"}),
            }
        }

    RETURN_TYPES = ("IMAGE", "MASK", "STRING", "STRING")
    RETURN_NAMES = ("image", "mask", "image_name", "url")
    FUNCTION = "execute"
    CATEGORY = "ExcelUtils"

    def execute(self, excel_file_path, url_column, start_row_number, end_row_number,
                output_format, timeout, on_error="black", fallback_size=512, sheet_name="0"):
        def _make_fallback(url_for_name="", row_for_name=0):
            color = 255 if on_error == "white" else 0
            img = Image.new("RGB", (fallback_size, fallback_size), (color, color, color))
            it, mt = _pil_to_tensor(img, output_format)
            name = _name_from_url(url_for_name, output_format) if url_for_name else f"fallback_row{row_for_name}.{output_format}"
            return it, mt, name
        if pd is None:
            raise RuntimeError("错误：运行此节点需要 Pandas 库，请先安装 pandas openpyxl。")

        if not excel_file_path or not os.path.isfile(excel_file_path):
            raise RuntimeError(f"错误：文件未找到 '{excel_file_path}'")

        try:
            col_index = _col_letter_to_index(url_column)
            sheet_identifier = int(sheet_name) if str(sheet_name).isdigit() else sheet_name

            df = pd.read_excel(
                excel_file_path, header=None, index_col=None, engine='openpyxl',
                sheet_name=sheet_identifier, usecols=[col_index],
            )
            total_rows = len(df)

            actual_end_row = int(end_row_number) if end_row_number != -1 else total_rows
            current_row = int(start_row_number)

            if total_rows == 0:
                raise RuntimeError("Excel 工作表为空或指定列无数据。")

            if current_row > actual_end_row:
                msg = f"已到达结束行 ({actual_end_row})，当前起始行 {current_row}。工作流停止。"
                print(f"[QC.LoadImageFromExcelURL] {msg}")
                raise WorkflowStopRequested(msg)

            if current_row < 1:
                current_row = 1

            cell_value = df.iloc[current_row - 1, 0]
            if pd.isna(cell_value):
                raise RuntimeError(f"第 {current_row} 行的URL单元格为空。")

            url = str(cell_value).strip()
            if not url.startswith(('http://', 'https://')):
                raise RuntimeError(f"第 {current_row} 行内容不是有效URL：'{url}'")

            print(f"[QC.LoadImageFromExcelURL] 行 {current_row} URL: {url}")

            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
            resp = requests.get(url, timeout=timeout, headers=headers, stream=True)
            resp.raise_for_status()

            image = Image.open(io.BytesIO(resp.content))
            image_tensor, mask_tensor = _pil_to_tensor(image, output_format)

            image_name = _name_from_url(url, output_format)
            print(f"[QC.LoadImageFromExcelURL] 加载成功: {image_name} ({image_tensor.shape[2]}x{image_tensor.shape[1]})")

            next_start = current_row + 1
            if next_start > actual_end_row and end_row_number != -1:
                next_start = actual_end_row + 1

            return {
                "ui": {"start_row_number": [next_start]},
                "result": (image_tensor, mask_tensor, image_name, url),
            }

        except WorkflowStopRequested:
            raise
        except Exception as e:
            traceback.print_exc()
            if on_error == "raise":
                raise RuntimeError(f"加载失败: {e}")

            row_for_fallback = locals().get("current_row", int(start_row_number))
            url_for_fallback = locals().get("url", "")
            print(f"[QC.LoadImageFromExcelURL] 行 {row_for_fallback} 加载失败，返回 {on_error} 占位图: {e}")
            it, mt, name = _make_fallback(url_for_fallback, row_for_fallback)

            try:
                actual_end = int(end_row_number) if end_row_number != -1 else row_for_fallback
                next_start = row_for_fallback + 1
                if end_row_number != -1 and next_start > actual_end:
                    next_start = actual_end + 1
            except Exception:
                next_start = row_for_fallback + 1

            return {
                "ui": {"start_row_number": [next_start]},
                "result": (it, mt, name, url_for_fallback),
            }


class LoadImageFromURL:
    """
    QC.LoadImageFromURL (保留的单URL版本)

    从单个URL加载图片，输出 (image, mask, image_name, url)。
    """

    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "url": ("STRING", {"multiline": False, "default": "https://example.com/image.jpg"}),
                "output_format": (["png", "webp", "jpg"], {"default": "png"}),
            },
            "optional": {
                "timeout": ("INT", {"default": 30, "min": 5, "max": 300, "step": 5}),
            }
        }

    RETURN_TYPES = ("IMAGE", "MASK", "STRING", "STRING")
    RETURN_NAMES = ("image", "mask", "image_name", "url")
    FUNCTION = "load_image"
    CATEGORY = "ExcelUtils"

    def load_image(self, url, output_format, timeout=30):
        try:
            if not url or not isinstance(url, str):
                raise ValueError("URL不能为空且必须是字符串")
            url = url.strip()
            if not url.startswith(('http://', 'https://')):
                raise ValueError("URL必须以 http:// 或 https:// 开头")

            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
            resp = requests.get(url, timeout=timeout, headers=headers, stream=True)
            resp.raise_for_status()

            image = Image.open(io.BytesIO(resp.content))
            image_tensor, mask_tensor = _pil_to_tensor(image, output_format)
            image_name = _name_from_url(url, output_format)

            print(f"[QC.LoadImageFromURL] 加载成功: {image_name}")
            return (image_tensor, mask_tensor, image_name, url)

        except requests.exceptions.Timeout:
            raise RuntimeError(f"请求超时: URL '{url}' 在 {timeout} 秒内未响应")
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"网络请求失败: {e}")
        except Image.UnidentifiedImageError:
            raise RuntimeError(f"无法识别的图片格式: {url}")
        except Exception as e:
            traceback.print_exc()
            raise RuntimeError(f"加载图片失败: {e}")


NODE_CLASS_MAPPINGS = {
    "QC.LoadImageFromURL": LoadImageFromURL,
    "QC.LoadImageFromExcelURL": LoadImageFromExcelURL,
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "QC.LoadImageFromURL": "QC.LoadImageFromURL",
    "QC.LoadImageFromExcelURL": "QC.LoadImageFromExcelURL",
}

print("--- 加载自定义节点: QC.LoadImageFromURL / QC.LoadImageFromExcelURL v2.0 ---")
