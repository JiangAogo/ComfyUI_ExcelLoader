import torch
import numpy as np
from PIL import Image
import io
import requests
import traceback

class LoadImageFromURL:
    """
    QC.LoadImageFromURL Node for ComfyUI
    
    从URL加载图片并转换为ComfyUI可识别的IMAGE格式。
    支持从Excel读取的URL，输出可直接连接到ComfyUI的其他图片处理节点。
    """
    
    def __init__(self):
        self.timeout = 30  # 请求超时时间（秒）
    
    @classmethod
    def INPUT_TYPES(cls):
        """
        定义节点的输入参数。
        """
        return {
            "required": {
                "url": ("STRING", {
                    "multiline": False, 
                    "default": "https://example.com/image.jpg",
                    "description": "图片的URL地址（支持http/https）"
                }),
            },
            "optional": {
                "timeout": ("INT", {
                    "default": 30,
                    "min": 5,
                    "max": 300,
                    "step": 5,
                    "description": "请求超时时间（秒）"
                }),
            }
        }
    
    RETURN_TYPES = ("IMAGE", "MASK")
    RETURN_NAMES = ("image", "mask")
    FUNCTION = "load_image"
    CATEGORY = "ExcelUtils"
    
    def load_image(self, url, timeout=30):
        """
        从URL加载图片并转换为ComfyUI的IMAGE格式
        
        Args:
            url: 图片URL地址
            timeout: 请求超时时间
            
        Returns:
            tuple: (image_tensor, mask_tensor)
                - image_tensor: 形状为 [1, H, W, 3] 的torch.Tensor，范围0-1
                - mask_tensor: 形状为 [1, H, W] 的torch.Tensor，范围0-1
        """
        try:
            # 验证URL
            if not url or not isinstance(url, str):
                raise ValueError("URL不能为空且必须是字符串")
            
            url = url.strip()
            if not url.startswith(('http://', 'https://')):
                raise ValueError("URL必须以 http:// 或 https:// 开头")
            
            # 从URL下载图片
            print(f"[QC.LoadImageFromURL] 正在从URL加载图片: {url}")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            response = requests.get(url, timeout=timeout, headers=headers, stream=True)
            response.raise_for_status()  # 检查HTTP错误
            
            # 检查内容类型
            content_type = response.headers.get('content-type', '')
            if not content_type.startswith('image/'):
                print(f"[QC.LoadImageFromURL] 警告: Content-Type是 '{content_type}'，可能不是图片")
            
            # 读取图片数据
            image_data = response.content
            image = Image.open(io.BytesIO(image_data))
            
            # 处理EXIF旋转
            try:
                from PIL import ImageOps
                image = ImageOps.exif_transpose(image)
            except Exception as e:
                print(f"[QC.LoadImageFromURL] EXIF处理警告: {e}")
            
            # 转换为RGB模式（ComfyUI标准格式）
            if image.mode == 'I':
                image = image.point(lambda i: i * (1 / 255))
            
            # 保存原始模式以处理透明度
            has_alpha = 'A' in image.getbands()
            
            # 转换为RGB
            if image.mode != 'RGB':
                # 如果有alpha通道，先保存
                if has_alpha:
                    alpha_channel = image.getchannel('A')
                image = image.convert('RGB')
            
            # 转换为numpy数组，范围0-1
            image_np = np.array(image).astype(np.float32) / 255.0
            
            # 转换为torch张量 [H, W, C]
            image_tensor = torch.from_numpy(image_np)
            
            # 添加batch维度 [1, H, W, C]
            image_tensor = image_tensor.unsqueeze(0)
            
            # 处理mask
            if has_alpha:
                # 如果有alpha通道，使用它作为mask
                mask_np = np.array(alpha_channel).astype(np.float32) / 255.0
                mask_tensor = 1.0 - torch.from_numpy(mask_np)  # ComfyUI的mask是反转的
            else:
                # 没有alpha通道，创建全0的mask（表示完全不透明）
                h, w = image_np.shape[:2]
                mask_tensor = torch.zeros((h, w), dtype=torch.float32)
            
            # 添加batch维度 [1, H, W]
            mask_tensor = mask_tensor.unsqueeze(0)
            
            print(f"[QC.LoadImageFromURL] 成功加载图片: {image_tensor.shape[2]}x{image_tensor.shape[1]}")
            
            return (image_tensor, mask_tensor)
            
        except requests.exceptions.Timeout:
            error_msg = f"请求超时: URL '{url}' 在 {timeout} 秒内未响应"
            print(f"[QC.LoadImageFromURL] 错误: {error_msg}")
            raise RuntimeError(error_msg)
            
        except requests.exceptions.RequestException as e:
            error_msg = f"网络请求失败: {e}"
            print(f"[QC.LoadImageFromURL] 错误: {error_msg}")
            raise RuntimeError(error_msg)
            
        except Image.UnidentifiedImageError:
            error_msg = f"无法识别的图片格式，URL可能不是有效的图片: {url}"
            print(f"[QC.LoadImageFromURL] 错误: {error_msg}")
            raise RuntimeError(error_msg)
            
        except Exception as e:
            print(f"[QC.LoadImageFromURL] 未知错误: {e}")
            traceback.print_exc()
            raise RuntimeError(f"加载图片失败: {e}")


# ComfyUI 节点注册
NODE_CLASS_MAPPINGS = {
    "QC.LoadImageFromURL": LoadImageFromURL
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "QC.LoadImageFromURL": "QC.LoadImageFromURL"
}

print("--- 加载自定义节点: QC.LoadImageFromURL (ExcelUtils/QC.LoadImageFromURL) v1.0 ---")