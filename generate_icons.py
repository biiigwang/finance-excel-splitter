#!/usr/bin/env python3
"""
生成各平台所需的图标文件，使用最近邻插值保持像素风格。
"""

from PIL import Image
import os
from pathlib import Path

# 原始图标路径
INPUT_ICON = "icon.png"

# 输出目录
OUTPUT_DIR = "icons"
Path(OUTPUT_DIR).mkdir(exist_ok=True)

# 需要生成的PNG尺寸（px）
PNG_SIZES = [16, 24, 32, 48, 64, 128, 256, 512, 1024]

def resize_pixel_art(image: Image.Image, size: int) -> Image.Image:
    """使用最近邻算法放大像素画，保持清晰边缘"""
    return image.resize((size, size), Image.Resampling.NEAREST)

def generate_png_icons(original: Image.Image):
    """生成所有尺寸的PNG图标"""
    for size in PNG_SIZES:
        resized = resize_pixel_art(original, size)
        output_path = os.path.join(OUTPUT_DIR, f"icon_{size}x{size}.png")
        resized.save(output_path, "PNG")
        print(f"✅ 生成: {output_path}")

def generate_ico(original: Image.Image):
    """生成Windows平台的ICO格式图标"""
    # 生成所有尺寸的图像用于ICO
    icon_images = []
    for size in [16, 24, 32, 48, 64, 128, 256]:
        icon_images.append(resize_pixel_art(original, size))

    output_path = os.path.join(OUTPUT_DIR, "app.ico")
    icon_images[0].save(
        output_path,
        format="ICO",
        sizes=[(s, s) for s in [16, 24, 32, 48, 64, 128, 256]],
        append_images=icon_images[1:]
    )
    print(f"✅ 生成Windows图标: {output_path}")

def generate_icns(original: Image.Image):
    """生成macOS平台的ICNS格式图标"""
    # macOS icns需要的尺寸
    icns_sizes = [16, 32, 64, 128, 256, 512, 1024]
    icon_images = []

    for size in icns_sizes:
        icon_images.append(resize_pixel_art(original, size))

    output_path = os.path.join(OUTPUT_DIR, "app.icns")
    icon_images[0].save(
        output_path,
        format="ICNS",
        sizes=[(s, s) for s in icns_sizes],
        append_images=icon_images[1:]
    )
    print(f"✅ 生成macOS图标: {output_path}")

def main():
    if not os.path.exists(INPUT_ICON):
        print(f"❌ 错误: 找不到原始图标文件 {INPUT_ICON}")
        return

    # 打开原始图标
    original = Image.open(INPUT_ICON)
    print(f"🔍 原始图标尺寸: {original.size}")

    # 生成各种格式
    generate_png_icons(original)
    generate_ico(original)
    generate_icns(original)

    print(f"\n🎉 所有图标已生成到 {OUTPUT_DIR}/ 目录")
    print("\n📋 图标清单:")
    for f in os.listdir(OUTPUT_DIR):
        print(f"   - {f}")

if __name__ == "__main__":
    main()