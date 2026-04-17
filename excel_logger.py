import openpyxl
from openpyxl import load_workbook, Workbook
import os
from datetime import datetime

class ExcelAutoLogger:
    def __init__(self):
        pass
    
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "prompt": ("STRING", {"multiline": True, "forceInput": True}),
                "excel_path": ("STRING", {"default": "prompt_log.xlsx"}),
                "mode": (["manual", "auto_increment"], {"default": "auto_increment"}),
                "start_row": ("INT", {"default": 2, "min": 1}),
                "start_column": ("INT", {"default": 1, "min": 1}),
                "sheet_name": ("STRING", {"default": "Prompts"}),
            },
            "optional": {
                "negative_prompt": ("STRING", {"multiline": True, "forceInput": True}),
                "seed": ("INT", {"forceInput": True}),
                "model_name": ("STRING", {"forceInput": True}),
                "steps": ("INT", {"forceInput": True}),
                "cfg": ("FLOAT", {"forceInput": True}),
                "image_path": ("STRING", {"forceInput": True}),
            }
        }
    
    RETURN_TYPES = ("STRING", "INT")
    RETURN_NAMES = ("prompt", "row_number")
    FUNCTION = "log_to_excel"
    CATEGORY = "utils"
    OUTPUT_NODE = True

    def log_to_excel(self, prompt, excel_path, mode, start_row, start_column, 
                     sheet_name, negative_prompt=None, seed=None, model_name=None,
                     steps=None, cfg=None, image_path=None):
        try:
            # 检查文件是否存在
            if os.path.exists(excel_path):
                wb = load_workbook(excel_path)
            else:
                wb = Workbook()
                if "Sheet" in wb.sheetnames:
                    wb.remove(wb["Sheet"])
            
            # 获取或创建工作表
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                ws = wb.create_sheet(sheet_name)
                # 创建表头
                headers = ["ID", "Timestamp", "Prompt", "Negative Prompt", 
                          "Seed", "Model", "Steps", "CFG", "Image Path"]
                for idx, header in enumerate(headers, start=1):
                    ws.cell(row=1, column=idx, value=header)
            
            # 确定写入行号
            if mode == "auto_increment":
                # 找到第一个空行
                row = start_row
                while ws.cell(row=row, column=start_column).value is not None:
                    row += 1
            else:
                row = start_row
            
            col = start_column
            
            # 写入数据
            ws.cell(row=row, column=col, value=row - 1)  # ID
            ws.cell(row=row, column=col+1, 
                   value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))  # 时间戳
            ws.cell(row=row, column=col+2, value=prompt)  # 提示词
            ws.cell(row=row, column=col+3, value=negative_prompt or "")  # 负面提示词
            ws.cell(row=row, column=col+4, value=seed)  # 种子
            ws.cell(row=row, column=col+5, value=model_name or "")  # 模型名
            ws.cell(row=row, column=col+6, value=steps)  # 步数
            ws.cell(row=row, column=col+7, value=cfg)  # CFG
            ws.cell(row=row, column=col+8, value=image_path or "")  # 图片路径
            
            # 保存文件
            wb.save(excel_path)
            
            print(f"✓ Logged to Excel: {excel_path} - Row {row}")
            
            return (prompt, row)
            
        except Exception as e:
            print(f"✗ Excel logging error: {str(e)}")
            return (prompt, -1)

NODE_CLASS_MAPPINGS = {
    "ExcelAutoLogger": ExcelAutoLogger
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "ExcelAutoLogger": "Excel Auto Logger"
}