# -*- coding: utf-8 -*-
"""
financial analysis

《pandas入门》之read_excel()和to_excel()函数解析
https://blog.csdn.net/tongxinzhazha/article/details/78796952

Python中字典合并的四种方法
https://blog.csdn.net/jerry_1126/article/details/73017270

无息流动负债=应付账款+预收款项+应付职工薪酬+应交税费+其他应付款+其他流动负债
无息非流动负债=非流动负债合计-长期借款-应付债券
带息债务=负债合计-无息流动负债-无息非流动负债

财报下载
http://basic.10jqka.com.cn/api/stock/export.php?export=main&type=report&code=000498
http://basic.10jqka.com.cn/api/stock/export.php?export=debt&type=report&code=000498

export: main主要财务指标,debt资产负债表,benefit利润表,cash现金流量表,
        each每股能力,grow成长能力,pay偿债能力,operate运营能力
type:report报告期,year年度,simple单季度
code:股票代码

"""

import xlwings as xw
import pandas as pd
import datetime
import numpy as np
from urllib import request
import sys
import os
import re

def getdrive():
    '''
    获取驱动器 
    '''
    return sys.argv[0][:2]


def lgpdm(dm):
    '''
    长股票代码
    '''
    dm=re.findall('(\d{6})',dm)
    
    if len(dm)==0 :
        return None

    dm=dm[0] 

    return dm+('.SH' if dm[0]=='6' else '.SZ')


def mgpdm(dm):
    '''
    中股票代码
    '''
    dm=re.findall('(\d{6})',dm)
    
    if len(dm)==0 :
        return None
    dm=dm[0]
    return ('SH' if dm[0]=='6' else 'SZ')+dm

def sgpdm(dm):
    '''
    短股票代码
    '''
    dm=re.findall('(\d{6})',dm)
    
    if len(dm)==0 :
        return None

    return dm[0]


def flddic():
    flddic1={
            "科目\时间":"rq",
            "货币资金(元)":"hbzj",
            "交易性金融资产(元)":"jyxjrzc",
            "应收票据(元)":"yspj",
            "应收账款(元)":"yszk",
            "预付账款(元)":"yufzk",
            "应收利息(元)":"yslx",
            "其他应收款(元)":"qtysk",
            "存货(元)":"ch",
            "一年内到期的非流动资产(元)":"ynndqdfldzc",
            "其他流动资产(元)":"qtldzc",
            "流动资产合计(元)":"ldzchj",
            "可供出售金融资产(元)":"kgcsjrzc",
            "持有至到期投资(元)":"cyzdqtz",
            "长期股权投资(元)":"cqgqtz",
            "投资性房地产(元)":"tzxfdc",
            "固定资产(元)":"gdzc",
            "在建工程(元)":"zjgc",
            "工程物资(元)":"gcwz",
            "无形资产(元)":"wxzc",
            "商誉(元)":"sy",
            "长期待摊费用(元)":"cqdtfy",
            "递延所得税资产(元)":"dysdszc",
            "非流动资产合计(元)":"fldzchj",
            "资产总计(元)":"zczj",
            "短期借款(元)":"dqjk",
            "交易性金融负债(元)":"jyxjrfz",
            "应付票据(元)":"yfpj",
            "应付账款(元)":"yfzk",
            "预收账款(元)":"yuszk",
            "应付职工薪酬(元)":"yfzgxc",
            "应交税费(元)":"yjsf",
            "应付利息(元)":"yflx",
            "应付股利(元)":"yfgl",
            "其他应付款(元)":"qtyfk",
            "一年内到期的非流动负债(元)":"ynndqdfldfz",
            "其他流动负债(元)":"qtldfz",
            "流动负债合计(元)":"ldfzhj",
            "长期借款(元)":"cqjk",
            "应付债券(元)":"yfzq",
            "长期应付款(元)":"cqyfk",
            "专项应付款(元)":"zxyfk",
            "预计负债(元)":"yjfz",
            "递延所得税负债(元)":"dysdsfz",
            "其他非流动负债(元)":"qtfldfz",
            "非流动负债合计(元)":"fldfzhj",
            "负债合计(元)":"fzhj",
            "股本(股)":"gb",
            "股本(元)":"gb",
            "资本公积金(元)":"zbgjj",
            "减:库存股(元)":"kcg",
            "专项储备(元)":"zxcb",
            "盈余公积金(元)":"yygjj",
            "未分配利润(元)":"wfplr",
            "外币报表折算差额(元)":"wbbbzsce",
            "归属于母公司股东权益合计(元)":"gsymgsgdqyhj",
            "少数股东权益(元)":"ssgdqy",
            "股东权益合计(元)":"gdqyhj",
            "负债和股东权益总计(元)":"fzhgdqyzj",
            "现金及存放中央银行款项(元)":"xjjcfzyyxkx",
            "存放同业款项(元)":"cftykx",
            "发放贷款及垫款(元)":"ffdkjdk",
            "向中央银行借款(元)":"xzyyxjk",
            "同业及其他金融机构存放款项(元)":"tyjqtjrjgcfkx",
            "吸收存款(元)":"xsck"
            }
    
    flddic2={
            "科目\时间":"rq",
            "销售商品、提供劳务收到的现金(元)":"xssptglwsddxj",
            "收到的税费与返还(元)":"sddsfyfh",
            "支付的各项税费(元)":"zfdgxsf",
            "支付给职工以及为职工支付的现金(元)":"zfgzgyjwzgzfdxj",
            "经营现金流入(元)":"jyxjlr",
            "经营现金流出(元)":"jyxjlc",
            "经营现金流量净额(元)":"jyxjllje",
            "处置固定资产、无形资产的现金(元)":"czgdzcwxzcdxj",
            "购建固定资产和其他支付的现金(元)":"gjgdzchqtzfdxj",
            "投资支付的现金(元)":"tzzfdxj",
            "取得子公司现金净额(元)":"qdzgsxjje",
            "支付其他与投资的现金(元)":"zfqtytzdxj",
            "投资现金流入(元)":"tzxjlr",
            "投资现金流出(元)":"tzxjlc",
            "投资现金流量净额(元)":"tzxjllje",
            "吸收投资收到现金(元)":"xstzsdxj",
            "其中子公司吸收现金(元)":"qzzgsxsxj",
            "取得借款的现金(元)":"qdjkdxj",
            "收到其他与筹资的现金(元)":"sdqtyczdxj",
            "偿还债务支付现金(元)":"chzwzfxj",
            "分配股利、利润或偿付利息支付的现金(元)":"fpgllrhcflxzfdxj",
            "其中子公司支付股利(元)":"qzzgszfgl",
            "支付其他与筹资的现金(元)":"zfqtyczdxj",
            "筹资现金流入(元)":"czxjlr",
            "筹资现金流出(元)":"czxjlc",
            "筹资现金流量净额(元)":"czxjllje",
            "汇率变动对现金的影响(元)":"hlbddxjdyx",
            "现金及现金等价物净增加额(元)":"xjjxjdjwjzje",
            "客户存款和同业存放款项净增加额(元)":"khckhtycfkxjzje",	
            "客户贷款及垫款净增加额(元)":"khdkjdkjzje	"
            }
    
    flddic3={
            "科目\时间":"rq",
            "基本每股收益(元)":"eps",
            "净利润(元)":"jlr",
            "净利润同比增长率":"jlrtbzzl",
            "扣非净利润(元)":"kfjlr",
            "扣非净利润同比增长率":"kfjlr_g",
            "营业总收入(元)":"yyzsr",
            "营业总收入同比增长率":"yyzsr_g",
            "每股净资产(元)":"mgjzc",
            "净资产收益率":"roe",
            "净资产收益率-摊薄":"roe_tb",
            "资产负债比率":"zcfzbl",
            "每股资本公积金(元)":"mgzbgjj",
            "每股未分配利润(元)":"mgwfplr",
            "每股经营现金流(元)":"mgjyxjl",
            "销售毛利率":"xsmll",
            "存货周转率":"chzzl",
            "销售净利率":"xsjll",
            "每股收益同比增长率":"eps_g",
            "营业收入同比增长率":"yysr_g",
            "净资产收益率同比增长率":"jzcsyl_g",
            "流动比率":"ldbl",
            "速动比率":"sdbl",
            "保守速动比率":"bssdbl",
            "产权比率":"cqbl",
            "营业周期":"yyzq",
            "存货周转天数":"chzzts",
            "应收账款周转天数":"yszkzzts",
            }
    
    flddic4={
            "科目\时间":"rq",
            "净利润(元)":"jlr",
            "扣非净利润(元)":"kfjlr",
            "营业总收入(元)":"yyzsr",
            "营业收入(元)":"yysr",
            "营业总成本(元)":"yyzcb",
            "营业成本(元)":"yycb",
            "营业利润(元)":"yylr",
            "投资收益(元)":"tzsy",
            "其中：联营企业和合营企业的投资收益(元)":"lyqyhhyqydtzsy",
            "资产减值损失(元)":"zcjzss",
            "管理费用(元)":"glfy",
            "销售费用(元)":"xsfy",
            "财务费用(元)":"cwfy",
            "营业外收入(元)":"yywsr",
            "营业外支出(元)":"yywzc",
            "营业税金及附加(元)":"yysjjfj",
            "利润总额(元)":"lrze",
            "所得税(元)":"sds",
            "其他综合收益(元)":"qtzhsy",
            "综合收益总额(元)":"zhsyze",
            "归属于母公司股东的综合收益总额(元)":"gsymgsgddzhsyze",
            "归属于少数股东的综合收益总额(元)":"gsyssgddzhsyze",
            "利息收入(元)":"lxsr",
            "利息支出(元)":"lxzc",
            "利息净收入(元)":"lxjsr",
            "手续费及佣金净收入(元)":"sxfjyjjsr",
            "汇兑收益(元)":"hdsy",
            "公允价值变动收益(元)":"gyjzbdsy",
            "业务及管理费(元)":"ywjglf"
            }
    
    flddic5={
            "销售收现率":"xssxl",
            "净利现金率":"jlxjl",
            "固定资产周转率":"gdzczzl",
            "应收账款周转率":"yszkzzl",
            "总资产周转率":"zzczzl",
            "无息流动负债":"wxldfz",
            "无息非流动负债":"wxfldfz",
            "有息负债率":"yxfzl",
            "权益负债率":"qyfzl",
            "权益长期负债率":"qycqfzl",
            "有息负债(元)":"yxfz",
            "息税前利息保障倍数":"xsqlxbzbs",
            "经营现金流与流动负债比":"jyxjlyldfzb",
            "总资产收益率":"roa",
            "投资资产收益率":"roic",
            "投资资产(元)":"tzzc",
            "所得税率":"sdsl",
            "应收账款占总资产比率":"yszkzzzcbl",
            "存货占总资产比率":"chzzzcbl",
            "商誉占总资产比率":"syzzzcbl",
            "应收票据占总资产比率":"yspjzzzcbl",
            "预付账款占总资产比率":"yfzkzzzcbl",
            "固定资产占总资产比率":"gdzczzzcbl",
            "流动资产占总资产比率":"ldzczzzcbl",
            "非流动资产占总资产比率":"fldzczzzcbl",
            "流动负债占总资产比率":"ldfzzzzcbl",
            "非流动负债占资产比率":"fldfzzzcbl",
            "股东权益占总资产比率":"gdqyzzzcbl"
            }

    flddic6={
            "平均固定资产":"pjgdzc",
            "平均总资产":"pjzczj",
            "平均应收账款":"pjyszk"}
            
    
    flddic={}
    flddic.update(flddic1)
    flddic.update(flddic2)
    flddic.update(flddic3)
    flddic.update(flddic4)
    flddic.update(flddic5)
    flddic.update(flddic6)

    return flddic    


def read_xls(xlsfn):
        
    wb = xw.Book(xlsfn)
    
    data = wb.sheets[0].range('A2').options(pd.DataFrame,expand='table').value
    
    '''下面的语句很重要，MultiIndex转换成Index'''
    data.columns=[e[0] for e in data.columns]
    
    df=data.T
    fldn=flddic()
    
    df.columns=[fldn[e] for e in df.columns]
    df.index.name='rq'

    xw.apps[0].quit()

    return df    

def delefld(df1,df2):
    '''
    df1.columns - df2.columns
    '''
    df1cols=df1.columns
    df2cols=df2.columns
    col=[]
    for e in df1cols:
        if e not in df2cols:
            col.append(e)
    return col

def dfeval(df,form):
    '''
    公式计算
    '''
    return df.eval(form,inplace=True)

def pjval(df,fld):
    '''
    期末、期初平均值计算
    '''
    val1='%s_1' % fld
    form='pj%s = (%s + %s_1)/2' % (fld,fld,fld)
    df[val1]=df[fld].shift(-1)
    df.eval(form,inplace=True)
    
def zzl(df,fld):
    '''
    增长率计算
    '''
    val1='%s_1' % fld
    form='%s_g = (%s / %s_1 - 1) * 100' % (fld,fld,fld)
    df[val1]=df[fld].shift(-1)
    df.eval(form,inplace=True)
    
    
if __name__ == '__main__':
    
    formulas=[
            ["销售收现率","xssxl = jyxjllje / yysr * 100"],
            ["净利现金率","jlxjl = jyxjllje / jlr * 100"],
            ["固定资产周转率","gdzczzl = yyzsr / pjgdzc "],
            ["应收账款周转率","yszkzzl = yysr / pjyszk "],
            ["无息流动负债","wxldfz = yszk + yuszk + yfzgxc + yjsf + qtyfk + qtldfz "],
            ["无息非流动负债","wxfldfz = fldfzhj - cqjk - yfzq "],
            ["有息负债","yxfz = fzhj - wxldfz - wxfldfz "]
            ]
    '''        
            ,
            "总资产周转率":"zzczzl",
            "有息负债率":"yxfzl",
            "权益负债率":"qyfzl",
            "权益长期负债率":"qycqfzl",
            "有息负债(元)":"yxfz",
            "息税前利息保障倍数":"xsqlxbzbs",
            "经营现金流与流动负债比":"jyxjlyldfzb",
            "总资产收益率":"roa",
            "投资资产收益率":"roic",
            "投资资产(元)":"tzzc",
            "流动比率":"ldbl",
            "速动比率":"sdbl",
            "所得税率":"sdsl",
            "应收账款占总资产比率":"yszkzzzcbl",
            "存货占总资产比率":"chzzzcbl",
            "商誉占总资产比率":"syzzzcbl",
            "应收票据占总资产比率":"yspjzzzcbl",
            "预付账款占总资产比率":"yfzkzzzcbl",
            "固定资产占总资产比率":"gdzczzzcbl",
            "流动资产占总资产比率":"ldzczzzcbl",
            "非流动资产占总资产比率":"fldzczzzcbl",
            "流动负债占总资产比率":"ldfzzzzcbl",
            "非流动负债占资产比率":"fldfzzzcbl",
            "股东权益占总资产比率":"gdqyzzzcbl"]

        '''
    
    xlsfn=r'D:\公司研究\新华制药\000756_debt_year.xls'
    debt=read_xls(xlsfn)

    xlsfn=r'D:\公司研究\新华制药\000756_main_year.xls'
    main=read_xls(xlsfn)
    
    xlsfn=r'D:\公司研究\新华制药\000756_benefit_year.xls'
    benefit=read_xls(xlsfn)

    xlsfn=r'D:\公司研究\新华制药\000756_cash_year.xls'
    cash=read_xls(xlsfn)

    df=main.join(debt)
    
    df=df.join(benefit[delefld(benefit,df)])
    
    df=df.join(cash)
    
    df['yfzq']=np.nan
    #将空用一极小值代替，便于后面的计算，后面再改回
    fillna=0.0000001    
    df=df.replace(False,fillna) 
    
    
    for fld in ['gdzc','zczj','yszk'] :
        pjval(df,fld)

    for fld in ['jlr','yysr']:
        zzl(df,fld)
        
    for gs in formulas:
        dfeval(df,gs[1])

#    df=df.replace(fillna,np.nan) 
