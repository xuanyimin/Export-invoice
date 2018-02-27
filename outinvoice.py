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
    Version.text = '3.0'
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
            mixi(in_xls_data,Spxx,Bz,out_amount,bf,yf,zf)
        else:
            out_amount = bf = yf = zf = 0
            i += 1
            invoices.append(invoice)
            # 发票头
            Fpsj = etree.SubElement(Fpxx, 'Fpsj')
            Fp = etree.SubElement(Fpsj, 'Fp')
            Djh = etree.SubElement(Fp, 'Djh')#单据号
            Djh.text = u'%s'%(in_xls_data.get(u'海关报关单号'))
            Spbmbbh = etree.SubElement(Fp, 'Spbmbbh')#商品编码版本号
            Spbmbbh.text = u'19.0'
            Hsbz = etree.SubElement(Fp, 'Hsbz')#含税标志
            Hsbz.text = u'0'
            Sgbz = etree.SubElement(Fp, 'Sgbz')  # 含税标志
            Hsbz.text = u'0'
            Gfmc = etree.SubElement(Fp, 'Gfmc')#购方名称
            Gfmc.text = in_xls_data.get(u'客户')
            Gfsh = etree.SubElement(Fp, 'Gfsh')#购方税号
            Gfsh.text = u''
            Gfdzdh = etree.SubElement(Fp, 'Gfdzdh')  # 购方地址电话
            Gfdzdh.text = u''
            Gfyhzh = etree.SubElement(Fp, 'Gfyhzh')  # 购方银行帐号
            Gfyhzh.text = u''
            Skr = etree.SubElement(Fp, 'Skr')#收款人
            Skr.text = u''
            Fhr = etree.SubElement(Fp, 'Fhr')#复核人
            Fhr.text = u''
            Bz = etree.SubElement(Fp, 'Bz')  # 复核人
            Spxx = etree.SubElement(Fp, 'Spxx')
            # 发票明细行
            mixi(in_xls_data,Spxx,Bz,out_amount,bf,yf,zf)

    Zsl.text = str(i)

def exchange_rate(currency,date):
    currency = currency.split(' ')[1]
    month = date.replace('-','')[:6]
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

def mixi(in_xls_data,Spxx,Bz,out_amount,bf,yf,zf):
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
    Ggxh.text = in_xls_data.get(u'规格型号') or ''
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
    bz = u'出口业务；出口总额:%s；'%out_amount
    if in_xls_data.get(u'币种'):
        bz = bz + u'币种:%s,' % in_xls_data.get(u'币种').split(' ')[1]
    if in_xls_data.get(u'成交方式'):
        bz = bz + u'成交方式:%s,' % in_xls_data.get(u'成交方式')
    if in_xls_data.get(u'保费金额') > 0:
        bz = bz + u'保费:%s,' % in_xls_data.get(u'保费金额')
    if in_xls_data.get(u'运费金额') > 0:
        bz = bz + u'运费:%s,' % in_xls_data.get(u'运费金额')
    if in_xls_data.get(u'进出口合同号'):
        bz = bz + u'合同号:%s,' % in_xls_data.get(u'进出口合同号')
    if in_xls_data.get(u'出口日期'):
        if len(bytes(bz.encode('GBK'))) + 11.0 > 130.0:
            pass
        else:
            mouth = in_xls_data.get(u'出口日期')
            currency = in_xls_data.get(u'币种')
            bz = bz + u'汇率:%s,' % exchange_rate(currency, mouth)
    if in_xls_data.get(u'加工贸易手册号'):
        if len(bytes(bz.encode('GBK'))) + len(bytes(in_xls_data.get(u'加工贸易手册号').encode('GBK'))) + 8 > 130:
            pass
        else:
            bz = bz + u'手册号:%s,' % in_xls_data.get(u'加工贸易手册号')
    if in_xls_data.get(u'装船口岸'):
        if len(bytes(bz.encode('GBK'))) + len(bytes(in_xls_data.get(u'装船口岸').encode('GBK'))) + 10 > 130:
            pass
        else:
            bz = bz + u'装船口岸:%s,' % in_xls_data.get(u'装船口岸')
    if in_xls_data.get(u'出口口岸'):
        if len(bytes(bz.encode('GBK'))) + len(bytes(in_xls_data.get(u'出口口岸').encode('GBK'))) + 10 > 130:
            pass
        else:
            bz = bz + u'出口口岸:%s,' % in_xls_data.get(u'出口口岸')
    if in_xls_data.get(u'目的地'):
        if len(bytes(bz.encode('GBK'))) + len(bytes(in_xls_data.get(u'目的地').encode('GBK'))) + 10 > 130:
            pass
        else:
            bz = bz + u'出口口岸:%s,' % in_xls_data.get(u'目的地')
    Bz.text = bz

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
            if app[0] == str(int(data)):
                return app[number]

def company_date(data):
    # 从薄名中取出基础数据
    base_data = xlrd.open_workbook('base.xls')
    table = base_data.sheet_by_name(u'公司信息')
    nrows = table.nrows
    colnames = table.row_values(0)
    for rownum in range(0, nrows):
        row = table.row_values(rownum)
        if row:
            app = []
            for i in range(len(colnames)):
                app.append(str(row[i]))
            if app[0] == data:
                return app[1]

def outformxls(select):
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
    if select == '1':
        Kp = etree.Element('Kp')
        to_xml(list,Kp)
        tree = etree.ElementTree(Kp)
        tree.write('out%s.xml'%time.strftime('%Y%m%d',time.localtime(time.time())), pretty_print=True, xml_declaration=True, encoding='GBK')
    else :
        business = etree.Element("business",  comment=u"发票开具", id="FPKJ")
        to_dzxml(list, business)
        tree = etree.ElementTree(business)
        tree.write('out%s.xml' % time.strftime('%Y%m%d', time.localtime(time.time())), pretty_print=True,
                   xml_declaration=True, encoding='GBK')

def to_dzxml(list,business):
    line = len(list)
    invoices =[]
    i = 0
    out_amount = amount = 0
    for data in range(0, line):
        in_xls_data = list[data]
        invoice = in_xls_data.get(u'海关报关单号')
        if invoice in invoices:
            (out_amount2, amount2) = dzmixi(in_xls_data, COMMON_FPKJ_XMXXS)
            out_amount += out_amount2
            amount += amount2
        else:
            i += 1
            invoices.append(invoice)
            # 发票头
            REQUEST_COMMON_FPKJ = etree.SubElement(business, 'REQUEST_COMMON_FPKJ')
            REQUEST_COMMON_FPKJ.set("class", "REQUEST_COMMON_FPKJ")
            COMMON_FPKJ_FPT = etree.SubElement(REQUEST_COMMON_FPKJ, 'COMMON_FPKJ_FPT')
            COMMON_FPKJ_FPT.set("class", "COMMON_FPKJ_FPT")
            FPQQLSH = etree.SubElement(COMMON_FPKJ_FPT, 'FPQQLSH')  # 开票请求流水号
            FPQQLSH.text = u'%s'%(in_xls_data.get(u'海关报关单号'))
            KPLX = etree.SubElement(COMMON_FPKJ_FPT, 'KPLX')#开票类型 0为蓝字，1为红字
            KPLX.text = u'0'
            XSF_NSRSBH = etree.SubElement(COMMON_FPKJ_FPT, 'XSF_NSRSBH')#销售方纳税人识别号
            XSF_NSRSBH.text = company_date(u'公司税号：')
            XSF_MC = etree.SubElement(COMMON_FPKJ_FPT, 'XSF_MC')  # 销售方名称
            XSF_MC.text = u'%s'% company_date(u'公司名称：')
            XSF_DZDH = etree.SubElement(COMMON_FPKJ_FPT, 'XSF_DZDH')#销售方地址、电话
            XSF_DZDH.text = u'%s'% company_date(u'公司地址、电话：')
            XSF_YHZH = etree.SubElement(COMMON_FPKJ_FPT, 'XSF_YHZH')#销售方银行帐号
            XSF_YHZH.text = u'%s'% company_date(u'开户行及帐号：')
            GMF_NSRSBH = etree.SubElement(COMMON_FPKJ_FPT, 'GMF_NSRSBH')  # 购买主纳税人识别号
            GMF_NSRSBH.text = u''
            GMF_MC = etree.SubElement(COMMON_FPKJ_FPT, 'GMF_MC')  # 购方名称
            GMF_MC.text = in_xls_data.get(u'客户')
            GMF_DZDH = etree.SubElement(COMMON_FPKJ_FPT, 'GMF_DZDH')#购方地址、电话
            GMF_DZDH.text = u''
            GMF_YHZH = etree.SubElement(COMMON_FPKJ_FPT, 'GMF_YHZH')#购方银行帐号
            GMF_YHZH.text = u''
            KPR = etree.SubElement(COMMON_FPKJ_FPT, 'KPR')  # 开票人
            KPR.text = u''
            SKR = etree.SubElement(COMMON_FPKJ_FPT, 'SKR')  # 收款人
            SKR.text = u''
            FHR = etree.SubElement(COMMON_FPKJ_FPT, 'FHR')  # 复核人
            FHR.text = u''
            YFP_DM = etree.SubElement(COMMON_FPKJ_FPT, 'YFP_DM')  # 原发票代码，红字必须
            YFP_DM.text = u''
            YFP_HM = etree.SubElement(COMMON_FPKJ_FPT, 'YFP_HM')  # 原发票号码，红字必须
            YFP_HM.text = u''
            BZ = etree.SubElement(COMMON_FPKJ_FPT, 'BZ')  # 备注
            BMB_BBH = etree.SubElement(COMMON_FPKJ_FPT, 'BMB_BBH')  # 版本号
            BMB_BBH.text = u'18.0'
            JSHJ = etree.SubElement(COMMON_FPKJ_FPT, 'JSHJ')  # 价税合计
            HJJE = etree.SubElement(COMMON_FPKJ_FPT, 'HJJE')  # 合计金额（不含税）
            HJSE = etree.SubElement(COMMON_FPKJ_FPT, 'HJSE')  # 合计税额
            HJSE.text = u'0.00'
            COMMON_FPKJ_XMXXS= etree.SubElement(REQUEST_COMMON_FPKJ, 'COMMON_FPKJ_XMXXS')
            COMMON_FPKJ_XMXXS.set("class", "COMMON_FPKJ_XMXX")
            COMMON_FPKJ_XMXXS.set("size", "1")
            # 发票明细行
            (out_amount,amount) = dzmixi(in_xls_data,COMMON_FPKJ_XMXXS)
        bz = u'出口业务;出口总额:%s;' % out_amount
        if in_xls_data.get(u'币种'):
            bz = bz + u'币种:%s,' % in_xls_data.get(u'币种').split(' ')[1]
        if in_xls_data.get(u'成交方式'):
            bz = bz + u'成交方式:%s,' % in_xls_data.get(u'成交方式')
        if in_xls_data.get(u'保费金额') > 0:
            bz = bz + u'保费:%s,' % in_xls_data.get(u'保费金额')
        if in_xls_data.get(u'运费金额') > 0:
            bz = bz + u'运费:%s,' % in_xls_data.get(u'运费金额')
        if in_xls_data.get(u'进出口合同号'):
            bz = bz + u'合同号:%s,' % in_xls_data.get(u'进出口合同号')
        else:
            logger.exception(u'找不到报关单%s所对应进出口合同号' % in_xls_data.get(u'海关报关单号'))
        if in_xls_data.get(u'加工贸易手册号'):
            if len(bytes(bz.encode('GBK'))) + len(bytes(in_xls_data.get(u'加工贸易手册号').encode('GBK'))) + 8 > 130:
                pass
            else:
                bz = bz + u'手册号:%s,' % in_xls_data.get(u'加工贸易手册号')
        if in_xls_data.get(u'目的地'):
            if len(bytes(bz.encode('GBK'))) + len(bytes(in_xls_data.get(u'目的地').encode('GBK'))) + 10> 130:
                pass
            else:
                bz = bz + u'出口口岸:%s,' % in_xls_data.get(u'目的地')
        if in_xls_data.get(u'出口日期'):
            if len(bytes(bz.encode('GBK'))) + 11.0 > 130.0 :
                pass
            else:
                mouth = in_xls_data.get(u'出口日期')
                currency = in_xls_data.get(u'币种')
                bz = bz + u'汇率:%s,' % exchange_rate(currency,mouth)
        if in_xls_data.get(u'装船口岸'):
            if len(bytes(bz.encode('GBK'))) + len(bytes(in_xls_data.get(u'装船口岸').encode('GBK'))) + 10 > 130:
                pass
            else:
                bz = bz + u'装船口岸:%s,' % in_xls_data.get(u'装船口岸')
        if in_xls_data.get(u'出口口岸'):
            if len(bytes(bz.encode('GBK'))) + len(bytes(in_xls_data.get(u'出口口岸').encode('GBK'))) + 10> 130:
                pass
            else:
                bz = bz + u'出口口岸:%s,' % in_xls_data.get(u'出口口岸')
        if len(bytes(bz.encode('GBK'))) > 130 :
            logger.exception(u'报关单%s的备注长度超过130个字节' % in_xls_data.get(u'海关报关单号'))
        BZ.text = bz
        HJJE.text = JSHJ.text = str(amount)

def dzmixi(in_xls_data,COMMON_FPKJ_XMXXS):
    # 明细计算内容,
    out_amount = float(in_xls_data.get(u'成交金额'))  # 成交外币
    currency = in_xls_data.get(u'币种')
    mouth = in_xls_data.get(u'出口日期')
    rate = exchange_rate(currency,mouth) or 0
    if not rate:
        logger.exception (u'找不到%s所在的月份所对应%s汇率' % (mouth,currency))
    amount = float(in_xls_data.get(u'成交金额')) * float(rate)  # 成交人民币
    COMMON_FPKJ_XMXX = etree.SubElement(COMMON_FPKJ_XMXXS, 'COMMON_FPKJ_XMXX')
    FPHXZ = etree.SubElement(COMMON_FPKJ_XMXX, 'FPHXZ') #发票行性质，0正常行，1折扣行，2被折扣行
    FPHXZ.text = u'0'
    XMMC = etree.SubElement(COMMON_FPKJ_XMXX, 'XMMC') #商品名称
    XMMC.text = in_xls_data.get(u'商品名称')
    GGXH = etree.SubElement(COMMON_FPKJ_XMXX, 'GGXH') # 规格型号
    GGXH.text = in_xls_data.get(u'规格型号') or ''
    DW = etree.SubElement(COMMON_FPKJ_XMXX, 'DW') #计量单位
    DW.text = in_xls_data.get(u'计量单位')
    SPBM = etree.SubElement(COMMON_FPKJ_XMXX, 'SPBM') #税收编码
    ZXBM = etree.SubElement(COMMON_FPKJ_XMXX, 'ZXBM') #企业编码
    ZXBM.text = u''
    YHZCBS = etree.SubElement(COMMON_FPKJ_XMXX, 'YHZCBS')# 优惠政策标识：0不使用，1使用
    YHZCBS.text = u'1'
    LSLBS = etree.SubElement(COMMON_FPKJ_XMXX, 'LSLBS')#零标识，0出口退税，1免税
    LSLBS.text = u'1'
    ZZSTSGL = etree.SubElement(COMMON_FPKJ_XMXX, 'ZZSTSGL')# 优惠政策说明？？
    ZZSTSGL.text = u''
    XMSL = etree.SubElement(COMMON_FPKJ_XMXX, 'XMSL')# 数量
    XMSL.text = str(in_xls_data.get(u'数量'))
    XMDJ = etree.SubElement(COMMON_FPKJ_XMXX, 'XMDJ') # 单价
    XMDJ.text = str(round(amount / float(in_xls_data.get(u'数量')),6))
    XMJE = etree.SubElement(COMMON_FPKJ_XMXX, 'XMJE') # 金额
    XMJE.text = str(round(amount,2))
    SE = etree.SubElement(COMMON_FPKJ_XMXX, 'SE') # 税额
    SE.text = u'0.00'
    SL = etree.SubElement(COMMON_FPKJ_XMXX, 'SL') # 税率
    SL.text = u'0'
    KCE = etree.SubElement(COMMON_FPKJ_XMXX, 'KCE') #扣除额
    KCE.text = u'0'

    goods_code = base_date(in_xls_data.get(u'商品代码'), 1)
    if goods_code:
        SPBM.text = str(goods_code)
    else:
        logger.exception (u'找不到海关商品编码%s所对应商品税收编码' % in_xls_data.get(u'商品代码'))
    return (out_amount,amount)

if __name__ == "__main__":
    conf = Config()
    logger = conf.getLog()
    print u"普通发票请选1，电子发票请选2"
    select = raw_input(u'select:')
    if select == '1' or select == '2':
        outformxls(select)
    else:
        print u'请重新运行并正确选择'
    print u'完成'
    time.sleep(10)