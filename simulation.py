__author__ = 'Mercury'
#coding=UTF-8

import os
import re
import sys

import xlrd
import datetime


def read_file(file_name, file_path):
    '''用于读取excel中的数据文件，最终依次返回价差序列、标准差和价格以及时间信息，
    即依次返回diff_msg, sd_msg, price_a, price_h, time_msg'''
    # 打开数据文件
    os.chdir(file_path)
    input_file = xlrd.open_workbook(file_name)
    # 依次读入价差序列表、标准差表、价格表
    series_diff = input_file.sheet_by_index(0)
    series_sd = input_file.sheet_by_index(1)
    series_price = input_file.sheet_by_index(2)

    # 从价差序列表中读取时间信息和价差序列以及A股和H股名称
    beta_msg = series_diff.col_values(0)[0]
    name_msg = series_diff.col_values(1)[0]
    time_msg = series_diff.col_values(0)[1:]
    diff_mgs = series_diff.col_values(1)[1:]

    # 从标准差表中读取标准差信息
    sd_msg = series_sd.col_values(1)[1:]

    # 从价格表中读取A股和H股的价格信息
    price_a = series_price.col_values(1)[1:]
    price_h = series_price.col_values(2)[1:]
    return diff_mgs, sd_msg, price_a, price_h, time_msg, name_msg, beta_msg


def trans_time(time_msg):
    '''用于将读取的数据文件中的日期数据转换为更易于理解的日期信息,返回日期的字符串数组'''
    start_date = datetime.date(1899, 12, 31).toordinal() - 1
    dict_date = []
    for date_data in time_msg:
        if isinstance(date_data, float):
            date_data = int(date_data)
        d = datetime.date.fromordinal(start_date + date_data)
        dict_date.append(d.strftime('%Y-%m-%d'))
    return dict_date


def get_xls_file(path):     # 从文件路径下读取所有的xls文件，并将文件名存入字典
    dir_file = os.listdir(path)
    pattern = re.compile(r'.*\.xlsx*$')
    dict_xls = []
    for f in dir_file:
        dict_xls.append(pattern.findall(f))
    result = [dict_xls[k][0] for k in range(len(dict_xls))
              if len(dict_xls[k]) != 0]
    return result


# 尝试运行，以及模拟交易的逻辑判断，收益分析部分
if __name__ == '__main__':
    # 从文件中获取数据
    reload(sys)
    sys.setdefaultencoding('utf-8')
    dir_path = r'C:\Users\Mercury\Desktop\project'
    list_file = get_xls_file(dir_path)

    # 设置阀值
    constrain = 1     # 终止的上下限
    trigger = 0.9      # 触发建仓的阀值

    # 定义一些需要使用的全局变量
    beta = 0
    cost_this = 0
    cost_total = 0
    trade_cost = 0.002
    in_trade = False        # 判断是否已经建仓
    sabh = False        # 判断建仓的方式sell a buy h
    shba = False        # 判断建仓的方式sell h buy a
    total_revenue = 0       # 所有交易完成之后计算总的利润
    total_cost = 0      # 所有交易完成后计算所有的成本

     # 交易信息以文件的方式输出
    log_file = 'trade_log.txt'
    trade_log = open(log_file, "w")

    # 交易循环
    def trade_iter(diff_msg, sd_msg, stocks_price, str_time):
        # 具有global属性的数值初始化
        len_of_date = len(str_time)
        revenue = 0
        buy_a = 0
        sell_a = 0
        buy_h = 0
        sell_h = 0
        number = 100        # 买的基础数量

        # 买卖的数量比函数，如果为a股则返回1，如果为h股则返回1.9
        def beta(item):
            if item == 'a':
                return 1.0
            else:
                global beta
                return beta

        # 建仓函数
        def trade_begin(buy_item, sell_item, date):
            buy = stocks_price[buy_item][date] * number * beta(buy_item)
            sell = stocks_price[sell_item][date] * number * beta(sell_item)

            global cost_this
            global cost_total
            global in_trade
            global shba
            global sabh
            cost_this += (buy + (buy + sell) * trade_cost)
            cost_total += cost_this
            in_trade = True
            if buy_item == 'a':
                shba = True
            else:
                sabh = True
            return buy, sell

        # 平仓函数
        def trade_end(buy_item, sell_item, date):
            buy = stocks_price[buy_item][date] * number * beta(buy_item)
            sell = stocks_price[sell_item][date] * number * beta(sell_item)

            global cost_this
            global cost_total
            global in_trade
            global shba
            global sabh
            cost_this += (buy + (buy + sell) * trade_cost)
            cost_total += (buy + (buy + sell) * trade_cost)
            in_trade = False

            if buy_item == 'a':
                sabh = False
            else:
                shba = False
            return buy, sell

        for trade_date in range(len_of_date):
            diff_divide_sd = diff_msg[trade_date] / sd_msg[trade_date]        # 计算价差序列是标准差的多少倍

            if not in_trade:        # 建仓
                if trigger < diff_divide_sd < constrain:     # 判断是否在买A卖H的建仓区间内
                    buy_a, sell_h = trade_begin('a', 'h', trade_date)
                    trade_log.write('建仓时间：%s，做多A做空H\n' % str_time[trade_date])
                elif -constrain < diff_divide_sd < -trigger:       # 判断是否在卖A买H的建仓区间内
                    buy_h, sell_a = trade_begin('h', 'a', trade_date)
                    trade_log.write('建仓时间：%s，做多H做空A\n' % str_time[trade_date])
            else:       # 平仓
                if trade_date == len_of_date - 1:     # 如果已经到了模拟的最后一天不管如何都收盘
                    if shba:        # 已经买A卖H，则反向操作
                        buy_h, sell_a = trade_end('h', 'a', trade_date)
                        global cost_this
                        revenue_this = (sell_a - buy_a) + (sell_h - buy_h)\
                            - (sell_a + buy_a + sell_h + buy_h) * trade_cost
                        revenue += revenue_this
                        rate_of_revenue = revenue_this / cost_this * 100
                        cost_this = 0
                        trade_log.write('平仓时间：%s，本次交易收益 %f，本次交易收益率 %f%%\n'
                                        % (str_time[trade_date], revenue_this, rate_of_revenue))
                    elif sabh:      # 已经卖A买H，则反向操作
                        buy_a, sell_h = trade_end('a', 'h', trade_date)
                        global cost_this
                        revenue_this = (sell_a - buy_a) + (sell_h - buy_h)\
                            - (sell_a + buy_a + sell_h + buy_h) * trade_cost
                        revenue += revenue_this
                        rate_of_revenue = revenue_this / cost_this * 100
                        cost_this = 0
                        trade_log.write('平仓时间：%s，本次交易收益 %f，本次交易收益率 %f%%\n'
                                        % (str_time[trade_date], revenue_this, rate_of_revenue))
                elif trigger < diff_divide_sd < constrain or -constrain < diff_divide_sd < -trigger:
                    continue

                # 判断是否需要在今天或者明天平仓
                elif diff_msg[trade_date] * diff_msg[trade_date + 1] < 0:
                    square_diff = pow(diff_msg[trade_date], 2) - pow(diff_msg[trade_date + 1], 2)
                    if square_diff >= 0:
                        continue
                    else:
                        if shba:        # 已经买A卖H，则反向操作
                            buy_h, sell_a = trade_end('h', 'a', trade_date)
                            global cost_this
                            revenue_this = (sell_a - buy_a) + (sell_h - buy_h)\
                                - (sell_a + buy_a + sell_h + buy_h) * trade_cost
                            revenue += revenue_this
                            rate_of_revenue = revenue_this / cost_this * 100
                            cost_this = 0
                            trade_log.write('平仓时间：%s，本次交易收益 %f，本次交易收益率 %f%%\n'
                                            % (str_time[trade_date], revenue_this, rate_of_revenue))
                        elif sabh:      # 已经卖A买H，则反向操作
                            buy_a, sell_h = trade_end('a', 'h', trade_date)
                            global cost_this
                            revenue_this = (sell_a - buy_a) + (sell_h - buy_h)
                            revenue += revenue_this
                            rate_of_revenue = revenue_this / cost_this * 100
                            cost_this = 0
                            trade_log.write('平仓时间：%s，本次交易收益 %f，本次交易收益率 %f%%\n'
                                            % (str_time[trade_date], revenue_this, rate_of_revenue))
                elif diff_msg[trade_date] * diff_msg[trade_date - 1] < 0:
                    square_diff = pow(diff_msg[trade_date], 2) - pow(diff_msg[trade_date - 1], 2)
                    if square_diff >= 0:
                        if shba:        # 已经买A卖H，则反向操作
                            buy_h, sell_a = trade_end('h', 'a', trade_date)
                            global cost_this
                            revenue_this = (sell_a - buy_a) + (sell_h - buy_h)\
                                - (sell_a + buy_a + sell_h + buy_h) * trade_cost
                            revenue += revenue_this
                            rate_of_revenue = revenue_this / cost_this * 100
                            cost_this = 0
                            trade_log.write('平仓时间：%s，本次交易收益 %f，本次交易收益率 %f%%\n'
                                            % (str_time[trade_date], revenue_this, rate_of_revenue))
                        elif sabh:      # 已经卖A买H，则反向操作
                            buy_a, sell_h = trade_end('a', 'h', trade_date)
                            global cost_this
                            revenue_this = (sell_a - buy_a) + (sell_h - buy_h)
                            revenue += revenue_this
                            rate_of_revenue = revenue_this / cost_this * 100
                            cost_this = 0
                            trade_log.write('平仓时间：%s，本次交易收益 %f，本次交易收益率 %f%%\n'
                                            % (str_time[trade_date], revenue_this, rate_of_revenue))
                    else:
                        continue
                elif diff_divide_sd > constrain or diff_divide_sd < -constrain:
                    if shba:        # 已经买A卖H，则反向操作
                        buy_h, sell_a = trade_end('h', 'a', trade_date)
                        global cost_this
                        revenue_this = (sell_a - buy_a) + (sell_h - buy_h)\
                            - (sell_a + buy_a + sell_h + buy_h) * trade_cost
                        revenue += revenue_this
                        rate_of_revenue = revenue_this / cost_this * 100
                        cost_this = 0
                        trade_log.write('平仓时间：%s，本次交易收益 %f，本次交易收益率 %f%%\n'
                                        % (str_time[trade_date], revenue_this, rate_of_revenue))
                    elif sabh:      # 已经卖A买H，则反向操作
                        buy_a, sell_h = trade_end('a', 'h', trade_date)
                        revenue_this = (sell_a - buy_a) + (sell_h - buy_h)
                        revenue += revenue_this
                        rate_of_revenue = revenue_this / cost_this * 100
                        cost_this = 0
                        trade_log.write('平仓时间：%s，本次交易收益 %f，本次交易收益率 %f%%\n'
                                        % (str_time[trade_date], revenue_this, rate_of_revenue))
        return revenue, cost_total
    for xls_file in list_file:
        diff_msgs, sd_msgs, price_as, price_hs, time_msgs, name_this, beta_this = read_file(xls_file, dir_path)
        global beta
        beta = beta_this
        str_times = trans_time(time_msgs)
        stocks = {'a': price_as, 'h': price_hs}
        trade_log.write(name_this)
        trade_log.write('\n')

        # 作图流程
#        pl.figure(figsize=(60, 10), dpi=80)       # 设置图的长宽
#        x_axle = np.linspace(0, len(str_times), len(str_times))     # 设置x轴
#
 #       up_constrain = [constrain * i for i in sd_msgs]     # 设置平仓上界列表
  #      up_trigger = [trigger * i for i in sd_msgs]     # 设置开仓上界列表
   #     bottom_trigger = [-trigger * i for i in sd_msgs]        # 设置开仓下界列表
    #    bottom_constrain = [-constrain * i for i in sd_msgs]        # 设置平仓下界列表
#
 #       diff_curve = np.array(diff_msgs)        # 将价差序列列表转换为np列表
  #      up_constrain_curve = np.array(up_constrain)     # 设置平仓上界列表
   #     up_trigger_curve = np.array(up_trigger)     # 设置建仓上界列表
    #    zero_curve = np.array([0] * len(str_times))         # 设置横轴列表
     #   bottom_trigger_curve = np.array(bottom_trigger)        # 设置建仓下界列表
      #  bottom_constrain_curve = np.array(bottom_constrain)        # 设置平仓下界列表
#
 #       pl.xlim(x_axle.min() * 1.1, x_axle.max() * 1.1)     # 设定x轴的画图范围
  #      # 设置x轴的标记
   #     pl.xticks([j for j in range(0, len(str_times), 30)], [str_times[j] for j in range(0, len(str_times), 30)])
#
 #       pl.ylim(min(bottom_constrain_curve.min(), diff_curve.min()) * 1.5,        # 设置y轴的画图范围
  #              max(up_constrain_curve.max(), diff_curve.max()) * 1.2)
   #     pl.yticks([-0.2, 0, 0.2],        # 设置y轴的标记
    #              [r'$-0.2$', r'$0$', r'$0.2$'])
#
 #       # 画出价差序列曲线
  #      pl.plot(x_axle, diff_curve, color='blue', linewidth=1.5,
   #             linestyle='-', label='diff_curve')
    #    # 画出平仓上界曲线
     #   pl.plot(x_axle, up_constrain_curve, color='red', linewidth=1,
      #          linestyle='-.', label='constrain')
        # 画出开仓上界曲线
       # pl.plot(x_axle, up_trigger_curve, color='green', linewidth=1,
        #        linestyle='--', label='trigger')
        ## 画出横轴
#        pl.plot(x_axle, zero_curve, color='black', linewidth=1, linestyle='-')
 #       # 画出开仓下界曲线
  #      pl.plot(x_axle, bottom_trigger_curve, color='green', linewidth=1, linestyle='--')
        # 画出平仓下界曲线
   #     pl.plot(x_axle, bottom_constrain_curve, color='red', linewidth=1, linestyle='--')
    #    # 设置曲线说明
     #   pl.legend(loc='upper right')

        # 显示图片
      #  pl.show()

        ultimate_revenue, ultimate_cost = trade_iter(diff_msgs, sd_msgs, stocks, str_times)
        total_revenue += ultimate_revenue
        total_cost += ultimate_cost

        if ultimate_cost == 0:
            trade_log.write('最终收益及总收益率均为0\n')
        else:
            trade_log.write('最终收益：%f，总收益率为：%f%%\n'
                            % (ultimate_revenue, ultimate_revenue / ultimate_cost * 100))
        trade_log.write('\n')
        # 重新初始化一些全局变量
        cost_total = 0
        cost_this = 0
        sabh = False
        shba = False
        in_trade = False

    trade_log.write('所有交易最终的总收益：%f, 总收益率为：%f%%\n'
                    % (total_revenue, total_revenue / total_cost * 100))
    trade_log.close()