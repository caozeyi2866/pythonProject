#批量化提取图像中的文字


import easyocr,glob,os,re

reader = easyocr.Reader(['ch_sim','en'],gpu=False)
path_list=[]
path = "./"
path_list= glob.glob(os.path.join(path, "*.jpg"))
for i in path_list:
     result = reader.readtext(i)
     want=""
     for res in result:
          want=want+res[1]
     print(want)
     print("_____________________")