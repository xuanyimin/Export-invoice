# -*- coding: utf-8 -*-

import xlrd
import time
from lxml import etree
from config import Config
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

def to_xml(list,Kp):
    #处理XML头
    Version = etree.SubElement(Kp, 'Version')
    Version.text = '2.0'
    Fpxx = etree.SubElement(Kp, 'Fpxx')
    #处理xml发票张数
    Zsl = etree.SubElement(Fpxx, 'Zsl')#单据数量
    line = len(list)
    invoices =[]
    i = 0
    out_amount = bf = yf = zf = 0
    for data in range(0, line):
        in_xls_data = list[data]
        invoice = in_xls_data.get(u'海关报关单号')
        date = time.strftime('%Y%m%d%H%M%S',time.localtime(time.time()))
        if invoice in invoices:
            mixi(in_xls_data,Fp,Bz,out_amount,bf,yf,zf)
        else:
            out_amount = bf = yf = zf = 0
            i += 1
            invoices.append(invoice)
            # 发票头
            Fpsj = etree.SubElement(Fpxx, 'Fpsj')
            Fp = etree.SubElement(Fpsj, 'Fp')
            Djh = etree.SubElement(Fp, 'Djh')#单据号
            Djh.text = u'%s%s'%(date,i)
            Spbmbbh = etree.SubElement(Fp, 'Spbmbbh')#商品编码版本号
            Spbmbbh.text = u'16.0'
            Hsbz = etree.SubElement(Fp, 'Hsbz')#含税标志
            Hsbz.text = u'0'
            Gfmc = etree.SubElement(Fp, 'Gfmc')#购方名称
            Gfmc.text = u'1'
            Gfsh = etree.SubElement(Fp, 'Gfsh')#购方税号
            Gfsh.text = u''
            Gfdzdh = etree.SubElement(Fp, 'Gfdzdh')  # 购方地址电话
            Gfdzdh.text = u''
            Gfyhzh = etree.SubElement(Fp, 'Gfyhzh')  # 购方银行帐号
            Gfyhzh.text = u''
            Skr = etree.SubElement(Fp, 'Skr')#收款人
            Skr.text = u''
            Fhr = etree.SubElement(Fp, 'Fhr')#复核人
            Fhr.text = u'昊添财务'
            Bz = etree.SubElement(Fp, 'Bz')  # 复核人
            # 发票明细行
            mixi(in_xls_data,Fp,Bz,out_amount,bf,yf,zf)

    Zsl.text = str(i)

def exchange_rate(currency,date):
    currency = currency.split(' ')[1]
    month = date.replace('-','')[:6]
    print currency,month
    base_data = xlrd.open_workbook('base.xls')
    table3 = base_data.sheet_by_name(u'汇率')
    # 取得行数
    ncows3 = table3.nrows
    colnames3 = table3.row_values(0)
    for rownum in range(1, ncows3):
        row = table3.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames3)):
                app[colnames3[i]] = row[i]
            if str(int(app.get(u'月份'))) == month:
                return app.get(currency)

def mixi(in_xls_data,Fp,Bz,out_amount,bf,yf,zf):
    # 明细计算内容,
    out_amount += float(in_xls_data.get(u'成交金额'))  # 成交外币
    currency = in_xls_data.get(u'币种')
    mouth = in_xls_data.get(u'出口日期')
    rate = exchange_rate(currency,mouth) or 0
    if not rate:
        logger.exception (u'找不到%s所在的月份所对应%s汇率' % (mouth,currency))
    amount = float(in_xls_data.get(u'成交金额')) * float(rate)  # 成交人民币
    bf += float(in_xls_data.get(u'保费金额'))  # 保费
    yf += float(in_xls_data.get(u'运费金额'))  # 运费
    zf += float(in_xls_data.get(u'杂费金额'))  # 杂费
    Spxx = etree.SubElement(Fp, 'Spxx')
    Sph = etree.SubElement(Spxx, 'Sph')
    Kce = etree.SubElement(Sph, 'Kce')  # 扣除额
    Kce.text = u''
    Spbm = etree.SubElement(Sph, 'Spbm')  # 商品编码
    goods_code = base_date(in_xls_data.get(u'商品代码'), 1)
    if goods_code:
        Spbm.text = str(goods_code)
    else:
        logger.exception (u'找不到海关商品编码%s所对应商品税收编码' % in_xls_data.get(u'商品代码'))
    Dj = etree.SubElement(Sph, 'Dj')  # 单价
    Dj.text = str(amount / float(in_xls_data.get(u'数量')))
    Spmc = etree.SubElement(Sph, 'Spmc')  # 商品名称
    Spmc.text = in_xls_data.get(u'商品名称')
    Ggxh = etree.SubElement(Sph, 'Ggxh')  # 规格型号
    Ggxh.text = u''
    Slv = etree.SubElement(Sph, 'Slv')  # 税率
    Slv.text = u'0'
    Xh = etree.SubElement(Sph, 'Xh')  # 序号
    Xh.text = str(int(in_xls_data.get(u'序号')))
    Lslbz = etree.SubElement(Sph, 'Lslbz')  # 零标识，0出口退税，1免税
    Lslbz.text = u'1'
    Syyhzcbz = etree.SubElement(Sph, 'Syyhzcbz')  # 优惠政策标识：0不使用，1使用
    Syyhzcbz.text = u'1'
    Sl = etree.SubElement(Sph, 'Sl')  # 数量
    Sl.text = str(in_xls_data.get(u'数量'))
    Je = etree.SubElement(Sph, 'Je')  # 金额
    Je.text = str(amount)
    Yhzcsm = etree.SubElement(Sph, 'Yhzcsm')  # 优惠政策说明
    Yhzcsm.text = u''
    Qyspbm = etree.SubElement(Sph, 'Qyspbm')  # 企业商品编码
    Qyspbm.text = u''
    Jldw = etree.SubElement(Sph, 'Jldw')  # 计量单位
    Jldw.text = in_xls_data.get(u'计量单位')
    if in_xls_data.get(u'成交方式') == 'FOB':
        cjfs = 'FOB'
    if in_xls_data.get(u'成交方式') == 'CNF':
        cjfs = u'CNF；运费：%s' % yf
    # todo 更多成交方式u
    Bz.text = u'出口业务；出口销售总额:%s；币种:%s；成交方式:%s；合同号:%s；运单号:%s；目的港:%s；' % (
        out_amount, in_xls_data.get(u'币种'), cjfs, in_xls_data.get(u'进出口合同号') or '',in_xls_data.get(u'运输工具'),u'现在数据里找不到等已后再加')

def base_date(data,number):
    # 从薄名中取出基础数据
    base_data = xlrd.open_workbook('base.xls')
    table2 = base_data.sheet_by_name(u'编码')
    # 取得行数
    ncows2 = table2.nrows
    colnames2 = table2.row_values(0)
    for rownum in range(1, ncows2):
        row = table2.row_values(rownum)
        if row:
            app = []
            for i in range(len(colnames2)):
                app.append(str(row[i]))
            if app[0] == data:
                return app[number]

def outformxls():
    #处理EXCEL
    xls_data = xlrd.open_workbook('10137.xls')
    table = xls_data.sheets()[0]
    # 取得行数
    ncows = table.nrows
    colnames = table.row_values(0)
    list = []
    newcows = 0
    for rownum in range(1, ncows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            list.append(app)
            newcows += 1

    # 写XML
    Kp = etree.Element('Kp')
    to_xml(list,Kp)
    tree = etree.ElementTree(Kp)
    tree.write('tax_code.xml', pretty_print=True, xml_declaration=True, encoding='GBK')

if __name__ == "__main__":
    conf = Config()
    logger = conf.getLog()
    outformxls()