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
    QC.LoadExcelContent Node for ComfyUI (v1.3 - Fixed Increment Mode)

    从Excel列按顺序递增加载数据。
    '起始行号' (start_row_number) 会在每次运行后自动增加。
    当起始行号超过结束行号时，将停止整个工作流。
    (读取模式固定为递增，不再提供选项)
    """
    def __init__(self):
        pass

    @classmethod
    def INPUT_TYPES(cls):
        """
        定义节点的输入参数。移除了 read_mode 选项。
        """
        return {
            "required": {
                "excel_file_path": ("STRING", {"multiline": False,"default": "input/example.xlsx","description": "Excel 文件路径 (.xlsx 或 .xls)"}),
                "column_letter": ("STRING", {"multiline": False,"default": "A","description": "要读取的列字母 (例如 'A', 'B', 'AA')"}),
                "read_count": ("INT", {"default": 1,"min": 1,"step": 1,"description": "每次执行读取的单元格（行）数"}),
                "start_row_number": ("INT", {"default": 1,"min": 1,"step": 1,"control_after_generate": True,"description": "从哪一行开始读取（1基索引）。运行后自动增加。"}),
                "end_row_number": ("INT", {"default": -1,"min": -1,"step": 1,"description": "在哪一行停止读取（包含该行，1基索引）。-1 表示读到最后一行。"}),
                # --- read_mode 已被移除 ---
                "exclude_text": ("STRING", {"multiline": True,"default": "","description": "要从输出中排除的文本内容（区分大小写，每行一个）"})
            },
            "optional": {
                 "sheet_name": ("STRING", {"multiline": False,"default": "0","description": "要读取的工作表名称或0基索引。'0' 代表第一个表。"})
            }
        }

    RETURN_TYPES = ("STRING", "STRING")
    RETURN_NAMES = ("current_row_str", "output_text")
    FUNCTION = "execute"
    CATEGORY = "ExcelUtils"

    @staticmethod
    def _col_letter_to_index(letter):
        # ... (此函数保持不变) ...
        index = 0; power = 1
        for char in reversed(letter.upper()):
            if not 'A' <= char <= 'Z': raise ValueError(f"列字母 '{letter}' 中包含无效字符 '{char}'。")
            index += (ord(char) - ord('A') + 1) * power; power *= 26
        return index - 1

    # --- 修改：execute 函数签名移除 read_mode ---
    def execute(self, excel_file_path, column_letter, read_count, start_row_number, end_row_number, exclude_text, sheet_name="0", **kwargs):
        if pd is None:
            return {"result": ("Error", "错误：运行此节点需要 Pandas 库，请先安装。")}

        # --- [0. 文件路径预检查] ---
        if not excel_file_path or not isinstance(excel_file_path, str): return {"result": ("Error", "错误：Excel 文件路径无效。")}
        if not os.path.exists(excel_file_path): return {"result": ("Error", f"错误：文件未找到 '{excel_file_path}'")}
        if not os.path.isfile(excel_file_path): return {"result": ("Error", f"错误：路径 '{excel_file_path}' 不是一个文件。")}

        current_exec_start_row = int(start_row_number)
        current_row_str = str(current_exec_start_row)
        output_text = ""
        next_start_row = current_exec_start_row # 默认

        try:
            # --- [提前加载部分信息以检查停止条件] ---
            try:
                temp_df_for_rows = pd.read_excel(excel_file_path, header=None, index_col=None, engine='openpyxl', sheet_name=(int(sheet_name) if sheet_name.isdigit() else sheet_name), usecols=[self._col_letter_to_index(column_letter)])
                total_rows = len(temp_df_for_rows); del temp_df_for_rows
            except Exception as e:
                 print(f"\n--- 错误：无法初步加载Excel以检查行数 ---"); traceback.print_exc(); print("-----\n")
                 raise e

            actual_end_row = int(end_row_number) if end_row_number != -1 else total_rows

            # --- 修改：核心停止逻辑不再检查 read_mode ---
            if total_rows > 0 and current_exec_start_row > actual_end_row:
                msg = f"已到达或超过结束行 ({actual_end_row})，当前请求起始行 {current_exec_start_row}。工作流停止。"
                print(f"[QC.LoadExcelContent] INFO: {msg}")
                raise WorkflowStopRequested(msg) # 引发停止异常
            # ------------------------------------------

            # --- [如果没停止，继续执行] ---
            df = pd.read_excel(excel_file_path, header=None, index_col=None, engine='openpyxl', sheet_name=(int(sheet_name) if sheet_name.isdigit() else sheet_name))

            # [获取列，验证行号等...]
            try:
                col_idx = self._col_letter_to_index(column_letter)
                if col_idx < 0 or col_idx >= len(df.columns): raise ValueError(f"...")
                target_column = df.iloc[:, col_idx]
            except (ValueError, IndexError) as e: return {"result": ("Error", f"...")}

            if total_rows == 0: return {"ui": {"start_row_number": [1]}, "result": ("0", "警告：...")}
            if actual_end_row < 1: actual_end_row = 1

            original_start_row = current_exec_start_row; clamped = False
            if current_exec_start_row < 1: current_exec_start_row = 1; clamped = True
            # 下面的 > actual_end_row 理论上不会触发，因为前面有停止检查
            if current_exec_start_row > actual_end_row: current_exec_start_row = max(1, actual_end_row - read_count + 1); clamped = True
            if clamped: current_row_str = f"{current_exec_start_row} (修正自 {original_start_row})"
            else: current_row_str = str(current_exec_start_row)

            # [读取数据，处理排除项...]
            read_start_index = current_exec_start_row - 1
            read_end_index = min(read_start_index + read_count, total_rows)
            cell_contents = []
            if read_start_index < total_rows and read_start_index >= 0 and read_end_index > read_start_index:
                 cell_contents = target_column.iloc[read_start_index:read_end_index].tolist()
            string_contents = [str(content) if pd.notna(content) else "" for content in cell_contents]
            exclusions = [line.strip() for line in exclude_text.splitlines() if line.strip()]
            processed_contents = []
            if exclusions:
                # 使用列表推导式简化排除逻辑 (Python 3.8+)
                processed_contents = [''.join(c for c in content if not any(ex in c for ex in exclusions)) # 这种替换更复杂，还是用replace简单
                                      if exclusions else content for content in string_contents] # 不对， replace 简单
                temp_processed = []
                for content in string_contents:
                     temp_content = content
                     for ex in exclusions: temp_content = temp_content.replace(ex, "")
                     temp_processed.append(temp_content)
                processed_contents = temp_processed
            else:
                processed_contents = string_contents
            output_text = "\n".join(processed_contents)


            # --- 修改：计算下一个起始行号，固定使用 increment 逻辑 ---
            last_possible_start_row = max(1, actual_end_row - read_count + 1)
            if current_exec_start_row >= last_possible_start_row:
                # 如果当前行已经是最后一个能开始的位置或之后，则停止递增 (保持当前值)
                next_start_row = current_exec_start_row
            else:
                # 否则，正常增加 read_count
                next_start_row = current_exec_start_row + read_count
            #----------------------------------------------------------

            if next_start_row < 1: next_start_row = 1 # 安全检查

            # [返回结果和 UI 更新信息]
            return {
                "ui": {"start_row_number": [next_start_row]},
                "result": (current_row_str, output_text)
            }

        # --- 异常处理 (保持不变，捕获 WorkflowStopRequested 并重新抛出) ---
        except WorkflowStopRequested as e: raise e
        except FileNotFoundError: return {"result": ("Error", f"...")}
        except ImportError: return {"result": ("Error", f"...")}
        except pd.errors.EmptyDataError: return {"result": ("Error", f"...")}
        except ValueError as e: print(f"..."); traceback.print_exc(); return {"result": ("Error", f"...")}
        except KeyError as e: print(f"..."); traceback.print_exc(); return {"result": ("Error", f"...")}
        except Exception as e: print(f"..."); traceback.print_exc(); return {"result": ("Error", f"...")}


# --- ComfyUI 节点注册 (保持不变) ---
NODE_CLASS_MAPPINGS = { "QC.LoadExcelContent": LoadExcelContentAdvanced }
NODE_DISPLAY_NAME_MAPPINGS = { "QC.LoadExcelContent": "QC.LoadExcelContent" }

print("--- 加载自定义节点: QC.LoadExcelContent (ExcelUtils/QC.LoadExcelContent) v1.3 [Fixed Increment Mode] ---")