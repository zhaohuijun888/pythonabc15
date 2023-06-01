from pystrich.code128 import Code128Encoder
numb=input('''请输入考号后两位数字：''')
encoder = Code128Encoder(numb,options={"ttf_font":"C:/Windows/Fonts/SimHei.ttf","ttf_fontsize":22 ,"bottom_border":5,"height":150,"label_border":1})     							#生成
encoder.save(f"{numb}.png", bar_width=3)		#保存
