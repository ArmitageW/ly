# -*- coding: UTF-8 -*-

import requests
import subprocess
import xlwt
import xlrd
import os
import xlutils
from xlutils.copy import copy
import datetime
import sys
import time

insert_data={}

def wget(url, product):
    new_file_name = "./product-" + str(product) + "/" + str(product) + '.html'
    cmd_ret = subprocess.run(['wget', url, "-O", new_file_name])
    # print("cmd_ret is:" + str(cmd_ret.returncode))
    if cmd_ret.returncode != 0:
        print("ERROR: no such html  : " + url)
        return None

    fp = open(new_file_name, 'r')
    data = fp.read()
    fp.close()
    return data

def login_get_html(product_id, coockie):
    url = r"http://www.chusecang.com/?product-" + str(product_id) + r".html"
    headers = {"Cookie":coockie}
    ret_data = requests.get(url, headers=headers)
    return ret_data.text

def get_img(url, path):
    cmd = subprocess.run(['wget', url, "-O", path])
    if cmd.returncode != 0:
        print("ERROR: wget img faild, url:" + url)
        return None
    

def get_img_car(html_data, product):
    # print(html_data)
    index_intro = html_data.find(r"intro")
    
    img_index = html_data.find(r"img", index_intro)

    div_index = html_data.find(r"div", img_index)

    img_src_data = html_data[img_index:div_index]
    img_list = []
    index = 0
    while(img_src_data.find(r"http", index) != -1):
        http_index = img_src_data.find('http', index)
        jpg_index = img_src_data.find('jpg', http_index)
        img_url = img_src_data[http_index:jpg_index+3]
        img_list.append(img_url)
        index = jpg_index
    i = 1
    for img in img_list:
        url = img
        path = r"./product-" + str(product) + "/car_img/" + str(product) + "-" + str(i) + ".jpg"
        get_img(url, path)
        i += 1
    


def get_market_price(html_data):
    data_index = html_data.find(r"PRODUCT_HASH=")
    data_end =  html_data.find(r";", data_index)

    price_data = html_data[data_index:data_end]

 
def get_info(html_data):
    data_index = html_data.find(r"SPEC_HASH=")
    data_end =  html_data.find(r";", data_index)

    spec_data = html_data[data_index:data_end]

    
def get_detail_img(html_data):
    data_index = html_data.find(r"dtdetail")
    img_index = html_data.find(r"img", data_index)
    div_index = html_data.find(r"div", img_index)

    img_src_data = html_data[img_index:div_index]
    detail_img_list = []
    index = 0
    while(img_src_data.find(r"http", index) != -1):
        http_index = img_src_data.find('http', index)
        jpg_index = img_src_data.find('jpg', http_index)
        img_url = img_src_data[http_index:jpg_index+3]
        detail_img_list.append(img_url)
        index = jpg_index
    return detail_img_list


def get_pic_info(html_data):
    index = html_data.find(r"goods-detail-pic-thumbnail pics")

    data_index = html_data.find(r"<td img_id", index)
    data_end = html_data.find(r"<img border=", data_index)
    data = html_data[data_index:data_end]
    return data


def get_huohao(html_data):
    index = html_data.find(r"商品编号")
    data_index = html_data.find(r"span>", index)

    data_end = html_data.find(r"<", data_index)
    huo_hao = html_data[data_index+5: data_end]
    insert_data["label_1"] = huo_hao
    return huo_hao

def get_zhong_liang(html_data):
    index = html_data.find(r"商品重量")
    data_index = html_data.find(r"nowrap", index)

    data_end = html_data.find(r"<", data_index)
    zhong_liang = html_data[data_index+8: data_end]
    insert_data["label_3"] = zhong_liang
    return zhong_liang

def get_ping_pai(html_data):
   
    index = html_data.find(r"品牌：")
    data_index = html_data.find(r"span", index)

    data_end = html_data.find(r"<", data_index)
    ping_pai = html_data[data_index+5: data_end]
    insert_data["label_5"] = ping_pai
    return ping_pai

def get_gui_ge(html_data):
    index = html_data.find(r"请选择规格")

    data_index = html_data.find(r"<tr product", index)

    data_end = html_data.find(r"actbtn btn-fastbuy", data_index)
    gui_ge = html_data[data_index: data_end]

    pic_info = get_pic_info(html_data)
    pic_list = []
    while(pic_info.find(r'img_id=') != -1):
        tmp = {}
        imgid_index = pic_info.find(r'img_id=')
        imgid_end = pic_info.find(r"'", imgid_index+9)
        tmp["img_id"] = pic_info[imgid_index+8:imgid_end]

        b_src_index = pic_info.find(r'b_src="')
        b_src_end = pic_info.find(r'"', b_src_index+7)
        tmp["b_src"] = pic_info[b_src_index+7:b_src_end]

        c_src_index = pic_info.find(r'c_src="')
        c_src_end = pic_info.find(r'"', c_src_index+7)
        tmp["c_src"] = pic_info[c_src_index+7:c_src_end]

        pic_list.append(tmp)
        pic_info = pic_info[c_src_end:]

    ret_info = []
    while(gui_ge.find(r"product=") != -1):
        tmp = {}

        index = gui_ge.find(r'product="')
        product_end = gui_ge.find(r'">', index)
        tmp["product_id"] = gui_ge[index+9:product_end]
        
        bian_hao_index = gui_ge.find(r'<td>G')
        bian_hao_end = gui_ge.find(r'<', bian_hao_index+5)
        tmp["bian_hao"] = gui_ge[bian_hao_index+4:bian_hao_end]

        guige_index = gui_ge.find(r'left')
        guige_end = gui_ge.find(r'<', guige_index)
        tmp["guige"] = gui_ge[guige_index+6:guige_end]

        img_id_index = gui_ge.find(r'vids="')
        img_id_end = gui_ge.find(r'"', img_id_index+6)
        tmp["img_id"] = gui_ge[img_id_index+6:img_id_end]

        vip2_index = gui_ge.find(r"fontcolorOrange") 	
        vip2_end = gui_ge.find(r"<", vip2_index)
        vip2 = gui_ge[vip2_index+17:vip2_end]
        
        vip3_index = gui_ge.find(r"fontcolorOrange", vip2_end)
        vip3_end = gui_ge.find(r"<", vip3_index)
        vip3 = gui_ge[vip3_index+17:vip3_end]

        vip_index = gui_ge.find(r"fontcolorOrange", vip3_end)
        vip_end = gui_ge.find(r"<", vip_index)
        vip = gui_ge[vip_index+17:vip_end]

        tmp["vip"] = vip
        tmp["vip2"] = vip2
        tmp["vip3"] = vip3

        tmp["b_src"] = ""
        tmp["c_src"] = ""
        for i in pic_list:
            if i["img_id"] == tmp["img_id"]:
                tmp["b_src"] = i["b_src"]
                tmp["c_src"] = i["c_src"]
       
        mktprice_index = gui_ge.find(r"mktprice1") 
        mktprice_end = gui_ge.find(r"<", mktprice_index)
        mktprice = gui_ge[mktprice_index+11:mktprice_end]       
        tmp["mktprice"] = mktprice

        mprice_index = gui_ge.find(r"mprice='")
        mprice_end = gui_ge.find(r"'", mprice_index+8)
        tmp["mprice"] = gui_ge[mprice_index+8:mprice_end]
        ret_info.append(tmp)
        gui_ge=gui_ge[mprice_end:]

    return ret_info

    #insert_data["label_5"] = gui_ge

def get_mkprice(html_data):
    index = html_data.find(r"市场价")
    data_index = html_data.find(r"mktprice", index)

    data_end = html_data.find(r"<", data_index)
    mkprice = html_data[data_index+11: data_end]
    return mkprice

def get_ping_lei(html_data):
    index = html_data.find(r"您当前的位置")
    end = html_data.find(r"div", index)
    pinglei_index = html_data.find(r"首页", index)
    src_pinglei = html_data[pinglei_index:end]
    ping_lei = ""
    while(src_pinglei.find(r"title=") != -1):
        name_index = src_pinglei.find(r"title=")
        name_end = src_pinglei.find(r"<", name_index+9)
        name = src_pinglei[name_index+9:name_end]
        ping_lei += name
        ping_lei += r"/"
        src_pinglei = src_pinglei[name_end:]
    return ping_lei

def get_ping_ming(html_data):
    index = html_data.find(r'class="now"')
    end = html_data.find(r'<', index+12)
    ping_ming = html_data[index+12:end]
    return ping_ming
    
def get_mprice_rang(html_data):
    index = html_data.find(r"销售价")
    rang_index = html_data.find(r"￥", index)
    rang_end = html_data.find(r"<", rang_index+1)
    rang = html_data[rang_index+1:rang_end]
    return rang

def get_x_mprice_rang(html_data):
    rang_index = html_data.find(r"x-mprice")
    rang_end = html_data.find(r"<", rang_index)
    rang = html_data[rang_index+10:rang_end]
    return rang

def make_data(html_data, xml_name, product, down_img=False):
    xml_t = Controller_XML(xml_name)
    huo_hao = get_huohao(html_data)
    zhong_liang = get_zhong_liang(html_data)
    ping_pai = get_ping_pai(html_data)
    mktprice = get_mkprice(html_data)
    
    ping_info = get_ping_lei(html_data)
    ping_lei = ping_info[:-1]


    ping_ming = get_ping_ming(html_data)
    mprice_rang = get_mprice_rang(html_data)
    x_mprice_rang = get_x_mprice_rang(html_data)
    data_info = get_gui_ge(html_data)
    for i in data_info:
        tmp = {}
        """
        "箱规"         : 1
        "商品编号"      :2
        "重量（）g"     :3
        "品类"         :4
        "品牌"         :5
        "市场价"       :6
        "销售价"       :7
        "销售价（范围）"       :8
        "会员价"       :9
        "规格"         :10
        "规格名"       :11
        """
        tmp["product-id"] = product
        tmp["pin_ming"] = ping_ming                                 # TODO
        tmp["huo_hao"] = huo_hao
        tmp["bian_hao"] = i["bian_hao"]
        tmp["zhong_liang"] = zhong_liang
        tmp["ping_lei"] = ping_lei
        tmp["ping_pai"] = ping_pai
        tmp["mktprice"] = mktprice
        tmp["mprice_rang"] = mprice_rang
        tmp["vip_price"] = x_mprice_rang                                 # TODO
        tmp["gui_ge"] = i["bian_hao"]
        tmp["gui_ge_ming"] = i["guige"]
        tmp["mprice"] = i["mktprice"]
        tmp["guige_vip_price"] = i["vip"] + "\n" + i["vip2"] + "\n" +i["vip3"]                          # TODO
        xml_t.add_data(tmp)
        if down_img:
            url = i["b_src"]
            path = r"./product-" + str(product) + "/spec_img/" + str(product) + "-" + tmp["gui_ge"] + ".jpg" 
            get_img(url, path)
    

def get_xiang_qing_img(html_data, product):
    img_info=get_detail_img(html_data)
    i = 1
    for img in img_info:
        path = r"./product-" + str(product) + "/detail_img/" + str(product) + "-" + str(i) + ".jpg"
        get_img(img, path)
        i += 1

xml_format_list = ['default', ]


class xml_format_model(object):
    def __init__(self):
        self.sheet_name = 'test'
        self.lable_name = ["name", "age", "image_path"]


class xml_format_df(xml_format_model):
    def __init__(self):
        self.format_name = "default_test"
        self.sheet_name = 'test'
        self.lable_name = ["product-id", "品名", "箱规", "商品编号", "重量（）g", "品类", "品牌", "市场价", "销售价（范围）",
                           "品类会员价", "规格", "规格名", "销售价", "规格会员价"]

    def set_xml_style(self, name, height, bold=False, format_str=''):
        style = xlwt.XFStyle()  # 初始化样式

        font = xlwt.Font()  # 为样式创建字体
        font.name = name  # 'Times New Roman'
        font.bold = bold
        font.height = height

        borders = xlwt.Borders()  # 为样式创建边框
        borders.left = 6
        borders.right = 6
        borders.top = 6
        borders.bottom = 6

        style.font = font
        style.borders = borders
        style.num_format_str = format_str

        return style

    def init_sheet(self, sheet):
        # sheet.col(0).width = 200 * 50  # 设置第一列的列宽
        # sheet.col(1).width = 200 * 50
        # sheet.col(2).width = 400 * 50
        style = self.set_font()
        i = 0
        for name in self.lable_name:
            sheet.write(0, i, name, style)
            i += 1

    def set_font(self):
        font = xlwt.Font()
        # font.colour_index = 10
        font.bold = True

        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # pattern.pattern_fore_colour = 10

        borders = xlwt.Borders()  # 为样式创建边框
        borders.left = 6
        borders.right = 6
        borders.top = 6
        borders.bottom = 6

        al = xlwt.Alignment()
        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中

        style = xlwt.XFStyle()
        style.font = font
        style.pattern = pattern
        style.borders = borders
        style.alignment = al
        return style

    def set_data(self):
        """
        "箱规"         : 1
        "商品编号"      :2
        "重量（）g"     :3
        "品类"         :4
        "品牌"         :5
        "市场价"       :6
        "销售价"       :7
        "会员价"       :8
        "规格"         :9
        "规格名"       :10
        """
        src_data = {}
        src_data["label_1"] = "label_1"
        src_data["label_2"] = "label_2"
        src_data["label_3"] = "label_3"
        src_data["label_4"] = "label_4"
        src_data["label_5"] = "label_5"
        src_data["label_6"] = "label_6"
        src_data["label_7"] = "label_7"
        src_data["label_8"] = "label_8"
        src_data["label_9"] = "label_9"
        src_data["label_10"] = "label_10"

        return src_data


class Controller_XML(object):
    def __init__(self, xml_path, xml_format=xml_format_df()):
        self.workbook = xlwt.Workbook(encoding='utf-8')
        self.xml_path = xml_path
        self.xml_format = xml_format
        self.init_xml()

    def init_xml(self):
        if not os.path.exists(self.xml_path):
            sheet = self.workbook.add_sheet(self.xml_format.sheet_name)
            self.xml_format.init_sheet(sheet)
            self.workbook.save(self.xml_path)

    def init_format(self):
        # TODO
        pass

    def add_data(self, data=None):
        if data is None:
            data = self.xml_format.set_data()
        keys = list(data.keys())

        r_xls = xlrd.open_workbook(self.xml_path)
        row = r_xls.sheets()[0].nrows


        excel = copy(r_xls)
        sheet = excel.get_sheet(self.xml_format.sheet_name)

        for i in range(0, len(data)):
            sheet.write(row, i, data[keys[i]])
        excel.save(self.xml_path)


    def set_data(self):
        src_data = {}
        src_data[""]

def complite():
    xml_t = xml_c.Controller_XML("./test.xls")
    xml_t.add_data()


def main(args):
    product="all"
    coockie=""
    up_img=True
    up_xls=True
    img_type=""
    for i in args:
        if "product=" in i:
            product = i[8:]
        if "coockie=" in i:
            coockie = i[8:]
    if "--get-img" in args:
        img_type=""
        for i in args:
            if "--img-type=" in i:
                img_type=i[11:]
            if img_type not in ["detail", "car", "spec"]:
                img_type=""
                print("没有指定img type。或者给定的 img type 不在支持范围内。默认将更新所有图片")
    else:
        up_img=False
        
    if "--update-xls" not in args:
        up_xls=False
        

    if up_xls is True and up_img is False:
        if "-" in product:
            index = product.find("-")
            start = product[:index]
            end = product[index+1:]
            for i in range(int(start), int(end)+1):
                product_file_name = r"./product-" + str(i)
                subprocess.run(['mkdir', "-p", product_file_name + "/detail_img"])
                subprocess.run(['mkdir', "-p", product_file_name + "/car_img"])
                subprocess.run(['mkdir', "-p", product_file_name + "/spec_img"])

                url = "http://www.chusecang.com/?product-" + str(i) + ".html"
                html_data = wget(url, i)
                if html_data is None:
                    subprocess.run(['rm', "-rf", product_file_name])
                    continue
                make_data(html_data, product_file_name + "/" + str(i) + ".xls", i, False)
        else:
            product_file_name = r"./product-" + str(product)
            subprocess.run(['mkdir', "-p", product_file_name + "/detail_img"])
            subprocess.run(['mkdir', "-p", product_file_name + "/car_img"])
            subprocess.run(['mkdir', "-p", product_file_name + "/spec_img"])

            url = "http://www.chusecang.com/?product-" + str(product) + ".html"
            html_data = wget(url, product)
            if html_data is None:
                subprocess.run(['rm', "-rf", product_file_name])
                return -1;
            make_data(html_data, product_file_name + "/" + str(product) + ".xls", product, False)      

  
    if up_img is False and up_xls is False:
        if "-" in product:
            index = product.find("-")
            start = product[:index]
            end = product[index+1:]
            for i in range(int(start), int(end)+1):
                product_file_name = r"./product-" + str(i)
                subprocess.run(['mkdir', "-p", product_file_name + "/detail_img"])
                subprocess.run(['mkdir', "-p", product_file_name + "/car_img"])
                subprocess.run(['mkdir', "-p", product_file_name + "/spec_img"])
                
                url = "http://www.chusecang.com/?product-" + str(i) + ".html"
                # html_data = wget(url, i)
                html_data = login_get_html(i, coockie)
                if html_data is None:
                    subprocess.run(['rm', "-rf", product_file_name])
                    continue
                make_data(html_data, product_file_name + "/" + str(i) + ".xls", i, True)
                get_xiang_qing_img(html_data, i)
                get_img_car(html_data, i)
               
        else:
            product_file_name = r"./product-" + str(product)
            subprocess.run(['mkdir', "-p", product_file_name + "/detail_img"])
            subprocess.run(['mkdir', "-p", product_file_name + "/car_img"])
            subprocess.run(['mkdir', "-p", product_file_name + "/spec_img"])
            
            url = "http://www.chusecang.com/?product-" + str(product) + ".html"
            # html_data = wget(url, product)
            html_data = login_get_html(product, coockie)
            if html_data is None:
                subprocess.run(['rm', "-rf", product_file_name])
                return -1;
            make_data(html_data, product_file_name + "/" + str(product) + ".xls", product, True)

            get_xiang_qing_img(html_data, product)
            get_img_car(html_data, product)
        
if __name__ == "__main__":
    help_info ="help: \n \
        coockie=    若需要更新xls 获取会员价信息，必须有此参数。 \n \
        product=    输入product id， all 表示更新所有。范围表示方法：1-200 （不可以有空格，格式严格要求）\n \
        [--get-img]    可选参数，若有此参数，表示只更新图片。\n \
        [--img-type=]   在有设置--get-img的情况下此参数生效。type包含：detail (表示详情图)，car （表示轮播图）, spec （表示规格图）\n \
        [--update-xls] 有此设置表示只更新 xls。"
    
    #wget("http://www.chusecang.com/?product-752.html")
    if len(sys.argv) > 1:
        main(sys.argv[1:])
    else:
        print("Invalid parameter!")
        print(help_info)
        sys.exit(-1)


