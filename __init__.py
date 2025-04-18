# __init__.py
# This file makes the directory a Python package.
# ComfyUI looks for NODE_CLASS_MAPPINGS and NODE_DISPLAY_NAME_MAPPINGS here.

# Import the node classes and mappings from your node file(s)
from .excel_loader_node import NODE_CLASS_MAPPINGS, NODE_DISPLAY_NAME_MAPPINGS

# Export them so ComfyUI can find them
__all__ = ['NODE_CLASS_MAPPINGS', 'NODE_DISPLAY_NAME_MAPPINGS']

print("--- Loading Custom Nodes: ComfyUI_ExcelLoader ---")