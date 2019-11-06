
# coding: utf-8

# import datetime
# from datetime import timedelta
# import pandas as pd
# import numpy as np

# In[ ]:

def ROESelect(universe,date):
    """
    给定股票列表和日期，返回ROE价值投资法选择的股票
    Args:
        universe (list or str): 股票列表（有后缀）
        date (str or datetime): 常见日期格式支持
    Returns:
        list: ROE价值投资法选择的股票
        
    Examples:
        >> universe = DynamicUniverse('A')
        >> buy_list = LongGeSelect(universe,'20170101')
    """
    trade_date = date if isinstance(date,datetime.datetime) else datetime.datetime.strptime(date,'%Y%m%d')
    trade_date = trade_date.strftime('%Y%m%d')
    date_previous = str(int(trade_date)-10000)
    data_NIncomeAttrP = DataAPI.FdmtISQPITGet(secID=universe,endDate=trade_date,field=u"secID,endDate,NIncomeAttrP",pandas="1") # 获取交易日之前每个季度归母净利润数据
    data_NIncomeAttrP = pd.concat(data_NIncomeAttrP[data_NIncomeAttrP['secID']==stock][:3] for stock in data_NIncomeAttrP['secID'].unique()
                             if len(data_NIncomeAttrP[data_NIncomeAttrP['secID']==stock])>=3) # 保留每只股票近3个季度归母净利润

    data_NIncomeAttrP_previous = DataAPI.FdmtISQPITGet(secID=universe,endDate=date_previous,field=u"secID,endDate,NIncomeAttrP",pandas="1") # 同样获取前一个交易年度的数据
    data_NIncomeAttrP_previous = pd.concat(data_NIncomeAttrP_previous[data_NIncomeAttrP_previous['secID']==stock][:3] for stock in data_NIncomeAttrP_previous['secID'].unique()
                                          if len(data_NIncomeAttrP_previous[data_NIncomeAttrP_previous['secID']==stock])>=3) 
 
    if len(data_NIncomeAttrP_previous.secID.unique()) < len(data_NIncomeAttrP.secID.unique()): # 将两个年度的公司保持一致
        data_NIncomeAttrP = data_NIncomeAttrP[data_NIncomeAttrP['secID'].isin(list(data_NIncomeAttrP_previous.secID.unique()))]
    else:
        data_NIncomeAttrP_previous = data_NIncomeAttrP_previous[data_NIncomeAttrP_previous['secID'].isin(list(data_NIncomeAttrP.secID.unique()))]
    data_NIncomeAttrP['NIncomeAttrP_previous'] = list(data_NIncomeAttrP_previous.NIncomeAttrP) # 计算近三个报告期归母净利润的同比增速
    data_NIncomeAttrP['growth'] = data_NIncomeAttrP.NIncomeAttrP/data_NIncomeAttrP.NIncomeAttrP_previous-1 
    NIncome_list = list(data_NIncomeAttrP.groupby('secID').min()[data_NIncomeAttrP.groupby('secID').min()['growth']>0].index) # 获取近三个年度归母净利润增速都大于0%的公司列表
    
    ROE = DataAPI.FdmtIndiRtnGet(secID=universe,endDate=trade_date,reportType=u"A",field=u"secID,endDate,ROE",pandas="1") # 获取交易日之前每个年报ROE数据
    try:
        ROE = pd.concat(ROE[ROE['secID']==stock][:5] for stock in ROE['secID'].unique() if len(ROE[ROE['secID']==stock])>=5) # 保留每个公司近5个年度ROE
    except ValueError:
        return []
    ROE = ROE.groupby('secID').mean() # 计算5年ROE均值
    ROE_PB = pd.merge(ROE,DataAPI.MktStockFactorsOneDayGet(secID=universe,tradeDate=trade_date,field=u"secID,endDate,PB",pandas="1"),left_index=True,right_on='secID') # 获取每个公司PB现值
    ROE_PB['perc'] = ROE_PB['ROE']/ROE_PB['PB'] # 计算ROE均值与PB现值比
    ROE_PB = ROE_PB[ROE_PB['perc']>4] # 保留比值大于4的公司
    ROE_list = list(ROE_PB.secID) # 将列表保存
    
    Div = DataAPI.EquDivGet(secID=universe,endDate=trade_date,field=u"secID,endDate,perCashDiv",pandas="1") # 获取交易日前每次分红数据
    Div['year'] = Div.endDate.str[:4].astype('int') 
    Div = Div.dropna()
    year_list = range(int(trade_date[:4])-5,int(trade_date[:4])) # 筛选出近五年有分红的
    Div = Div[Div.year.isin(year_list)]
    Div = Div.groupby('secID').sum()[['perCashDiv']] # 算出五年总分红
    EPS = DataAPI.FdmtIndiQGet(secID=universe,endDate=trade_date,reportType='Q4',field=u"secID,endDate,EPS",pandas="1").dropna() # 获取交易日前每个年度EPS数据
    EPS = pd.concat(EPS[EPS['secID']==stock][:5] for stock in EPS['secID'].unique() if len(EPS[EPS['secID']==stock])>=5).groupby('secID').sum() # 保留每只股票近5个年度EPS数据并计算总和
    Div_EPS = pd.merge(Div,EPS,left_index=True,right_index=True)
    Div_EPS['perc'] = Div_EPS.perCashDiv/Div_EPS.EPS # 计算过去5年分红率
    Div_list = list(Div_EPS[Div_EPS['perc']>0.1].index) # 筛选出分红率大于10%的公司
    
    buy_list = list(set(NIncome_list).intersection(ROE_list).intersection(Div_list))
    
    return buy_list


# In[ ]:

start = '2012-03-12' # 回测的起止时间是2012年3月12日至2018年1月17日
end = '2018-1-17'
benchmark = '000002.ZICN' # 参考标准为全部A股
universe = set_universe('A') # 股票池为全部A股
capital_base = 10000000 # 初始本金为100万元
freq = 'd'                                 # 策略类型，'d'表示日间策略使用日线回测，'m'表示日内策略使用分钟线回测
refresh_rate = 90                           # 调仓频率，表示执行handle_data的时间间隔，若freq = 'd'时间间隔的单位为交易日，若freq = 'm'时间间隔为分钟

def initialize(account):                   # 初始化虚拟账户状态
    pass

def handle_data(account):                  # 每个交易日的买入卖出指令

    buy_list = ROESelect(account.universe,account.previous_date)

    for stk in account.security_position:
        if stk not in buy_list:
            order_to(stk, 0)

    if len(buy_list) > 0:
        price = account.reference_price
        for stk in buy_list:
            if not np.isnan(price[stk]) and not price[stk] == 0:
                if stk not in account.security_position:
                    order(stk,account.cash/len(buy_list)/price[stk])

