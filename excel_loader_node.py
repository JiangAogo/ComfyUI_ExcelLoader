# /ComfyUI/custom_nodes/ComfyUI_ExcelLoader/excel_loader_node.py

import random
import os
import traceback

# 尝试导入pandas
try:
    import pandas as pd
except ImportError:
    print("\n##################################################################")
    print("错误：QC.LoadExcelContent 节点缺少 Pandas 库。")
    print("请安装: pip install pandas openpyxl")
    print("##################################################################\n")
    pd = None

# 自定义异常用于停止工作流
class WorkflowStopRequested(Exception):
    """指示工作流应在此节点后停止的特殊异常。"""
    pass

class LoadExcelContentAdvanced:
    """
    QC.LoadExcelContent Node for ComfyUI (v1.4 - Multi-Column & Tidy Tags Support)

    从Excel的一列或多列按顺序递增加载数据。
    '起始行号' (start_row_number) 会在每次运行后自动增加。
    当起始行号超过结束行号时，将停止整个工作流。
    新增多列读取和 tidy_tags 功能，用于格式化输出。
    """
    def __init__(self):
        pass

    @classmethod
    def INPUT_TYPES(cls):
        """
        定义节点的输入参数。
        """
        return {
            "required": {
                "excel_file_path": ("STRING", {"multiline": False, "default": "input/example.xlsx", "description": "Excel 文件路径 (.xlsx 或 .xls)"}),
                "column_letter": ("STRING", {"multiline": False, "default": "A", "description": "要读取的列字母，支持多列，用逗号分隔 (例如 'A', 'B,C', 'A,C,E')"}),
                "tidy_tags": (["yes", "no"], {"default": "yes", "description": "是否将多列内容用逗号连接（适合Tags）。'no'则用空格连接。"}),
                "read_count": ("INT", {"default": 1, "min": 1, "step": 1, "description": "每次执行读取的行数"}),
                "start_row_number": ("INT", {"default": 1, "min": 1, "step": 1, "control_after_generate": True, "description": "从哪一行开始读取（1基索引）。运行后自动增加。"}),
                "end_row_number": ("INT", {"default": -1, "min": -1, "step": 1, "description": "在哪一行停止读取（包含该行，1基索引）。-1 表示读到最后一行。"}),
                "exclude_text": ("STRING", {"multiline": True, "default": "", "description": "要从输出中排除的文本内容（区分大小写，每行一个）"})
            },
            "optional": {
                 "sheet_name": ("STRING", {"multiline": False, "default": "0", "description": "要读取的工作表名称或0基索引。'0' 代表第一个表。"})
            }
        }

    RETURN_TYPES = ("STRING", "STRING")
    RETURN_NAMES = ("current_row_str", "output_text")
    FUNCTION = "execute"
    CATEGORY = "ExcelUtils"

    @staticmethod
    def _col_letter_to_index(letter):
        """将单个Excel列字母（如 'A', 'B', 'AA'）转换为0基索引。"""
        index = 0
        power = 1
        for char in reversed(letter.upper()):
            if not 'A' <= char <= 'Z':
                raise ValueError(f"列字母 '{letter}' 中包含无效字符 '{char}'。")
            index += (ord(char) - ord('A') + 1) * power
            power *= 26
        return index - 1
    
    @classmethod
    def _col_letters_to_indices(cls, letters_str):
        """将逗号分隔的列字母字符串（如 'A, C, AA'）转换为0基索引列表。"""
        if not letters_str or not isinstance(letters_str, str):
            raise ValueError("列字母输入无效。")
        
        letters = [l.strip() for l in letters_str.split(',')]
        indices = []
        for letter in letters:
            if letter: # 忽略空的条目 (例如 'A,,B')
                indices.append(cls._col_letter_to_index(letter))
        
        if not indices:
            raise ValueError("未指定任何有效的列字母。")
            
        return indices

    def execute(self, excel_file_path, column_letter, tidy_tags, read_count, start_row_number, end_row_number, exclude_text, sheet_name="0", **kwargs):
        if pd is None:
            return {"result": ("Error", "错误：运行此节点需要 Pandas 库，请先安装。")}

        # --- [0. 文件路径和参数预检查] ---
        if not excel_file_path or not isinstance(excel_file_path, str): return {"result": ("Error", "错误：Excel 文件路径无效。")}
        if not os.path.exists(excel_file_path): return {"result": ("Error", f"错误：文件未找到 '{excel_file_path}'")}
        if not os.path.isfile(excel_file_path): return {"result": ("Error", f"错误：路径 '{excel_file_path}' 不是一个文件。")}

        current_exec_start_row = int(start_row_number)
        current_row_str = str(current_exec_start_row)
        output_text = ""
        next_start_row = current_exec_start_row

        try:
            # --- [1. 转换列字母并初步加载以检查行数] ---
            try:
                col_indices = self._col_letters_to_indices(column_letter)
                sheet_identifier = int(sheet_name) if sheet_name.isdigit() else sheet_name
                # 只读取第一列来确定总行数，这样更快
                temp_df_for_rows = pd.read_excel(excel_file_path, header=None, index_col=None, engine='openpyxl', sheet_name=sheet_identifier, usecols=[col_indices[0]])
                total_rows = len(temp_df_for_rows)
                del temp_df_for_rows
            except Exception as e:
                 print(f"\n--- 错误：无法初步加载Excel以检查行数或解析列 '{column_letter}' ---"); traceback.print_exc(); print("-----\n")
                 raise e

            actual_end_row = int(end_row_number) if end_row_number != -1 else total_rows

            # --- [2. 核心停止逻辑] ---
            if total_rows > 0 and current_exec_start_row > actual_end_row:
                msg = f"已到达或超过结束行 ({actual_end_row})，当前请求起始行 {current_exec_start_row}。工作流停止。"
                print(f"[QC.LoadExcelContent] INFO: {msg}")
                raise WorkflowStopRequested(msg)

            # --- [3. 如果没停止，正式读取数据] ---
            df = pd.read_excel(excel_file_path, header=None, index_col=None, engine='openpyxl', sheet_name=sheet_identifier, usecols=col_indices)
            
            if total_rows == 0:
                return {"ui": {"start_row_number": [1]}, "result": ("0", "警告：Excel工作表为空或指定列无数据。")}
            
            actual_end_row = max(1, actual_end_row)

            # [修正起始行号]
            original_start_row = current_exec_start_row
            clamped = False
            if current_exec_start_row < 1:
                current_exec_start_row = 1
                clamped = True
            if current_exec_start_row > actual_end_row:
                current_exec_start_row = max(1, actual_end_row - read_count + 1)
                clamped = True
            
            if clamped:
                current_row_str = f"{current_exec_start_row} (修正自 {original_start_row})"
            else:
                current_row_str = str(current_exec_start_row)

            # --- [4. 处理数据] ---
            read_start_index = current_exec_start_row - 1
            read_end_index = min(read_start_index + read_count, total_rows)
            
            processed_rows = []
            if read_start_index < read_end_index:
                # 获取所需行的数据
                target_df_slice = df.iloc[read_start_index:read_end_index]
                
                # 定义连接符
                separator = "," if tidy_tags == "yes" else " "
                
                # 处理每一行
                for _, row in target_df_slice.iterrows():
                    # 将行内所有单元格转为字符串，并处理空值
                    string_cells = [str(cell).strip() for cell in row if pd.notna(cell) and str(cell).strip()]
                    
                    # 使用指定分隔符连接非空单元格
                    combined_string = separator.join(string_cells)
                    processed_rows.append(combined_string)

            # [处理排除项]
            exclusions = [line.strip() for line in exclude_text.splitlines() if line.strip()]
            final_contents = []
            if exclusions:
                for content in processed_rows:
                    temp_content = content
                    for ex in exclusions:
                        temp_content = temp_content.replace(ex, "")
                    final_contents.append(temp_content)
            else:
                final_contents = processed_rows
            
            output_text = "\n".join(final_contents)

            # --- [5. 计算下一个起始行号] ---
            next_start_row = current_exec_start_row + read_count
            if next_start_row > actual_end_row and end_row_number != -1:
                 # 如果下次将超过结尾，让它停在结尾，这样下次运行会触发停止条件
                 next_start_row = actual_end_row + 1

            if next_start_row < 1: next_start_row = 1

            # --- [6. 返回结果和 UI 更新信息] ---
            return {
                "ui": {"start_row_number": [next_start_row]},
                "result": (current_row_str, output_text)
            }

        # --- 异常处理 ---
        except WorkflowStopRequested as e:
            raise e
        except FileNotFoundError:
            return {"result": ("Error", f"错误：文件未找到于 '{excel_file_path}'")}
        except ImportError:
            return {"result": ("Error", "错误：Pandas 库未安装。请执行 'pip install pandas openpyxl'")}
        except pd.errors.EmptyDataError:
            return {"result": ("Error", f"错误：Excel 文件 '{excel_file_path}' 是空的或格式不正确。")}
        except ValueError as e:
            print(f"[QC.LoadExcelContent] 错误: {e}"); traceback.print_exc()
            return {"result": ("Error", f"错误：值错误 - {e}")}
        except KeyError as e:
            print(f"[QC.LoadExcelContent] 错误: {e}"); traceback.print_exc()
            return {"result": ("Error", f"错误：找不到工作表 '{sheet_name}' 或列 '{column_letter}'。")}
        except Exception as e:
            print(f"[QC.LoadExcelContent] 未知错误: {e}"); traceback.print_exc()
            return {"result": ("Error", f"发生未知错误: {e}")}


# --- ComfyUI 节点注册 (保持不变) ---
NODE_CLASS_MAPPINGS = { "QC.LoadExcelContent": LoadExcelContentAdvanced }
NODE_DISPLAY_NAME_MAPPINGS = { "QC.LoadExcelContent": "QC.LoadExcelContent-v2" }

print("--- 加载自定义节点: QC.LoadExcelContent (ExcelUtils/QC.LoadExcelContent-v2) v1.4 [Multi-Column & Tidy-Tags Support] ---")