from PIL import Image

# 输入 PNG 文件
png_file = "test.png"
# 输出 ICO 文件
ico_file = "icon.ico"

# 生成多尺寸的 ICO，Windows 推荐尺寸
sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

img = Image.open(png_file)
img.save(ico_file, format='ICO', sizes=sizes)

print(f"已生成图标: {ico_file}")
