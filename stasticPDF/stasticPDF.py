import os
import PyPDF2
import time
# import pdfplumber

def get_all_file_by_type(path, type=(), get_all_dirs = True):  # 获得以type类型结尾的所有文件，返回一个list

    filelist = []

    for a, b, c in os.walk(path):
        for name in c:
            fname = os.path.join(a, name)
            if fname.endswith(type):
                filelist.append(fname)
        if not get_all_dirs:        # 仅在当前目录查找文件
            print("跳出循环")
            break
    print("总共有%d个文件"%filelist.__len__())
    return filelist

def compute_pdfpage(path, get_all_dirs = False):
    counts = 0
    type = ("PDF","pdf")
    file_list = get_all_file_by_type(path=path, type=type, get_all_dirs = get_all_dirs)
    res = []
    for pdf in file_list:
        name = pdf.replace(path,"").replace("\\","")
        dict = {}
        try:
            reader = PyPDF2.PdfFileReader(pdf)
            # 不解密可能会报错：PyPDF2.utils.PdfReadError: File has not been decrypted
            if reader.isEncrypted:
                reader.decrypt('')
            information = reader.getDocumentInfo()
            title =  {information.title}
            page_num = reader.getNumPages()
            page_1 = reader.getPage(0)
            if page_1.get('/Rotate', 0) in [90, 270]:
                # weight , hight = page_1['/MediaBox'][2]* 0.3528, page_1['/MediaBox'][3]* 0.3528
                w , h = float(page_1['/MediaBox'][2])* 0.3528, float(page_1['/MediaBox'][3])* 0.3528
                weight = round(float(page_1['/MediaBox'][2])* 0.3528,0)
                hight = round(float(page_1['/MediaBox'][3])* 0.3528,0)
            else:
                w , h = float(page_1['/MediaBox'][3])* 0.3528, float(page_1['/MediaBox'][2])* 0.3528
                weight = round(float(page_1['/MediaBox'][3])* 0.3528,0)
                hight = round(float(page_1['/MediaBox'][2])* 0.3528,0)
            if hight == 210.0 and weight ==297.0:
                dict = {"name":name,"type":"A4","pageNum:":page_num}
                res.append(dict)
            else:
                dict = {"name":name,"type":str(h)+"*"+str(w),"pageNum:":page_num}
                res.append(dict)
            counts += page_num
        except Exception as e:
            dict = {"name":name,"type":"统计出错！！！","pageNum:":page_num}
            res.append(dict)
            print("-"*70)
            print(pdf + "该文件出现异常，可能是权限问题")
            print(e)
            print("-"*70)
            pass
    return res

if __name__ == '__main__':
    path = r"E:\Chinadci\temp\stasticPDF\data"
    res = compute_pdfpage(path, get_all_dirs=True)
    print("统计结果：")
    with open("E:\\Chinadci\\temp\\stasticPDF\\data\\res.txt","w") as f:
        for i in res:
            f.write(str(i)+"\n")
            print(i)

