import pyzbar.pyzbar as pyzbar
# import numpy
# from PIL import Image, ImageDraw, ImageFont
# import cv2

def gettiaoxingma(frame,*args,**kwargs):
    # frame = cv2.imread(frame)
    # gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    barcodes = pyzbar.decode(frame,*args,**kwargs)
    for barcode in barcodes:
        # 提取条形码的边界框的位置
        # 画出图像中条形码的边界框
        # (x, y, w, h) = barcode.rect
        # cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 255, 0), 2)
        # # 条形码数据为字节对象，所以如果我们想在输出图像上
        #  画出来，就需要先将它转换成字符串
        barcodeData = barcode.data.decode("utf-8")
        # 绘出图像上条形码的数据和条形码类型
        # barcodeType = barcode.type
        # 把cv2格式的图片转成PIL格式的图片然后在上标注二维码和条形码的内容
        # img_PIL = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
        # 参数（字体，默认大小）
        # font = ImageFont.truetype('STFANGSO.TTF', 25)
        # 字体颜色
        # fillColor = (0, 255, 0)
        # 文字输出位置
        # position = (x, y - 25)
        # 输出内容
        # strl = barcodeData
        # 需要先把输出的中文字符转换成Unicode编码形式(str.decode("utf-8))
        # 创建画笔
        # draw = ImageDraw.Draw(img_PIL)
        # draw.text(position, strl, font=font, fill=fillColor)
        # 使用PIL中的save方法保存图片到本地
        # img_PIL.save('结果图.jpg', 'jpeg')
        # 向终端打印条形码数据和条形码类型
        # print("扫描结果==》 类别： {0} 内容： {1}".format(barcodeType, barcodeData))
        # print(barcodeData)
        return barcodeData
# frame='D:\\PyQt5\\datikashibie\\piccut\\1\\0.jpg'
# gettiaoxingma(frame)