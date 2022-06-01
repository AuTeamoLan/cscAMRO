import os
import PIL.Image as Image


if __name__ == '__main__':
    # 如果当前位深是32的话，可以不用写转RGBA模式的这一句，但是写上也没啥问题
    # 从RGB（24位）模式转成RGBA（32位）模式
    pngs = os.listdir(os.path.join(os.getcwd(), '签名图片'))
    for png in pngs:
        img = Image.open(os.path.join(os.getcwd(), '签名图片', png)).convert('RGBA')
        W, L = img.size
        white_pixel = (255, 255, 255, 255)  # 白色
        for h in range(W):
            for i in range(L):
                if img.getpixel((h, i)) == white_pixel:
                    img.putpixel((h, i), (0, 0, 0, 0))   # 设置透明
        img.save(os.path.join(os.getcwd(), '新签名图片', png))
