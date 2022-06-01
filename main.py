# -*- coding: utf-8 -*-
import os
import random
import sys
import json
import logging
import arrow
import pandas as pd
import numpy as np
import xlwings as xw
from PIL import Image
from amro import AmroWeb


def update_datadf(amro):
    attn1 = amro.getattn1(4)
    attn2 = amro.getattn2(4)
    attn3 = amro.getattn3(4)
    emps = amro.getempsx(4)
    attn = pd.concat([attn1, attn2, attn3], join='outer', ignore_index=True)
    datadf = pd.merge(emps, attn, on='NAME')
    datadf['TEAM'] = ''
    teams = amro.getteams()
    for team in teams:
        datadf['TEAM'].loc[datadf['REF_PKID'] == team['PKID']] = team['TEXT']
    datadf['LG'] = datadf['LG'].str.replace('OTHER', '其他')
    datadf['LG'] = datadf['LG'].str.replace('FURLOUGH', '休假')
    datadf['LG'] = datadf['LG'].str.replace('TRAIN', '培训')
    datadf['LG'] = datadf['LG'].str.replace('SECOND', '借调')
    datadf['TEAM'].loc[datadf['MEMO1'].str.contains('排故组', na=False)] = '二队十一组'
    arrow.now().format('YYYY-MM-DD HH:mm:ss')
    # datadf.to_excel('a.xlsx')
    return datadf


def add_centerx(sht, target, filePath, scale):
    '''Excel智能居中插入图片
    优先级：match > width & height > column_width & row_height
    建议使用column_width或row_height，定义单元格最大宽或高
    :param sht: 工作表
    :param target: 目标单元格，字符串，如'A1'
    :param filePath: 图片绝对路径
    '''
    rng = sht.range(target)  # 目标单元格
    name = os.path.basename(filePath)  # 文件名
    pic_width, pic_height = Image.open(filePath).size  # 原图片宽高
    if rng.width/rng.height < pic_width/pic_height:
        height = rng.width/pic_width*pic_height*scale
        width = rng.width*scale
    else:
        height = rng.height*scale
        width = rng.height/pic_height*pic_width*scale
    left = rng.left + (rng.width - width) / 2  # 居中
    top = rng.top + (rng.height - height) / 2
    try:
        sht.pictures.add(filePath, left=left, top=top, width=width, height=height, scale=None, name=name)
    except Exception:  # 已有同名图片，采用默认命名
        pass


if __name__ == '__main__':
    logging.basicConfig(filename="error_log.txt",
                        filemode="a",
                        format="%(asctime)s %(name)s:%(levelname)s:%(message)s",
                        datefmt="%Y-%m-%d %H:%M:%S",
                        level=logging.INFO)
    # Define a Handler and set a format which output to console
    console = logging.StreamHandler()  # 定义console handler
    console.setLevel(logging.DEBUG)  # 定义该handler级别
    formatter = logging.Formatter('%(asctime)s  %(filename)s : %(levelname)s  %(message)s')  # 定义该handler格式
    console.setFormatter(formatter)
    # Create an instance
    logging.getLogger().addHandler(console)  # 实例化添加handler
    png_scale = png_scale_dong3 = None
    try:
        with open('config.json', 'r') as f:
            content = f.read()
            config = json.loads(content)
        png_scale = config.get('png_scale')
        png_scale_dong3 = config.get('png_scale_dong3')
    except Exception as e:
        logging.exception(e)
    if (png_scale is None) or (not isinstance(png_scale, (int, float))):
        png_scale = 1.6
        logging.warning('config.json文件中png_scale参数有误，程序按默认参数1.6执行')
    if (png_scale_dong3 is None) or (not isinstance(png_scale_dong3, (int, float))):
        png_scale_dong3 = 1.6
        logging.warning('config.json文件中png_scale_dong3参数有误，程序按默认参数1.6执行')
    try:
        amro = AmroWeb()
        amro.login()
        datadf = update_datadf(amro)
    except Exception as e:
        logging.exception('访问AMRO时出错，请检查网络连接！')
        sys.exit()
    try:
        files = os.listdir(os.path.join(os.getcwd(), '原始签到表'))
    except:
        logging.exception('未找到【原始签到表】文件夹!')
        sys.exit()
    try:
        pngs = os.listdir(os.path.join(os.getcwd(), '签名图片'))
    except:
        logging.exception('未找到【签名图片】文件夹!')
        sys.exit()
    newpath = os.path.join(os.getcwd(), arrow.now().format('YYYY-MM-DD'))
    while os.path.exists(newpath):
        newpath = newpath + "(1)"
    os.mkdir(newpath)
    app = xw.App(visible=False, add_book=False)
    err_name = []
    err_no_png = []
    for file in files:
        if file[-4:] == '.xls' or file[-5:] == '.xlsx':
            wb = app.books.open(os.path.join(os.getcwd(), '原始签到表', file))
            sht = wb.sheets['Sheet1']
            # sht.range('A2').value = str(sht.range('A2').value).replace('YYYY年MM月DD日', arrow.now().format('YYYY年MM月DD日'))
            sht.range('H2').value = arrow.now().format('YYYY年MM月DD日')
            for i in range(1, 60):
                name = sht.range('B{}'.format(i)).value
                if name is not None and name != '姓名':
                    a = datadf.loc[datadf['NAME'] == name].copy()
                    if len(a) == 1:  # len(a)==1 LG离岗备注+签名
                        if a.iloc[0]['LG'] is not np.nan:
                            lg_str = a.iloc[0]['LG']
                            sht.range('G{}'.format(i)).value = lg_str.replace('其他/', '')
                        if name+'.png' in pngs and a.iloc[0]['attn'] == 'Y':
                            add_centerx(sht, 'C{}'.format(i), os.path.join(os.getcwd(), '签名图片', name+'.png'), png_scale)
                            if a.iloc[0]['ATTNTIME'] is not np.nan:
                                sht.range('D{}'.format(i)).value = a.iloc[0]['ATTNTIME'][-8:]
                    if len(a) > 1:  # len(a)>1 LG离岗备注+不签名（存在错误打卡）
                        if a.iloc[0]['LG'] is not np.nan:
                            lg_str = a.iloc[0]['LG']
                            sht.range('G{}'.format(i)).value = lg_str.replace('其他/', '')
                    if len(a) == 0:
                        err_name.append(name)
                    if name+'.png' not in pngs:
                        err_no_png.append(name)
            wb.save(os.path.join(newpath, file))
            wb.close()

# 处理东三体温表
    try:
        dong3_files = os.listdir(os.path.join(os.getcwd(), '原始医学观察统计表'))
    except:
        logging.exception('未找到【原始医学观察统计表】文件夹!')
        sys.exit()
    for file in dong3_files:
        if file[-4:] == '.xls' or file[-5:] == '.xlsx':
            wb = app.books.open(os.path.join(os.getcwd(), '原始医学观察统计表', file))
            sht = wb.sheets['Sheet1']
            # sht.range('E2').value = arrow.now().format('YYYY年MM月DD日')
            for i in range(3, 300):
                name = sht.range('B{}'.format(i)).value
                if name is not None and name != '姓名':
                    a = datadf.loc[datadf['NAME'] == name].copy()
                    if len(a) > 0 and a.iloc[0]['LG'] is not np.nan:
                        lg_str = a.iloc[0]['LG']
                        sht.range('L{}'.format(i)).value = lg_str.replace('其他/', '')
                    else:
                        sht.range('I{}'.format(i)).value = random.randint(361, 366)/10
                        sht.range('J{}'.format(i)).value = '无'
                        if name + '.png' in pngs:
                            add_centerx(sht, 'K{}'.format(i), os.path.join(os.getcwd(), '签名图片', name + '.png'), png_scale_dong3)
            wb.save(os.path.join(newpath, file))
            wb.close()
    app.quit()
# 将错误写入日志
    if err_name or err_no_png:
        if err_name:
            err_name = ','.join(err_name)
            logging.info('在AMRO中未找到[' + err_name + ']，检查签到表中的名字与AMRO是否一致!')
        if err_no_png:
            err_no_png = ','.join(err_no_png)
            logging.info('在【签名图片】文件夹中未找到：[' + err_no_png + ']!')
