__author__ = 'Mercury'
#coding=UTF-8

import os
import re
import sys
import numpy as np
import pylab as pl
import xlrd
import datetime


def read_file(file_name, file_path):
    '''用于读取excel中的数据文件，最终依次返回价差序列、标准差和价格、时间信息、交易股票名称以及beta系数信息，
    即依次返回diff_msg, sd_msg, price_a, price_h, time_msg, name_msg, beta_msg'''
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
    list_date = []
    for date_data in time_msg:
        if isinstance(date_data, float):
            date_data = int(date_data)
        d = datetime.date.fromordinal(start_date + date_data)
        list_date.append(d.strftime('%Y-%m-%d'))
    return list_date


def get_xls_file(path):
    '''从文件路径下读取所有的xls文件，并将文件名存入列表'''
    dir_file = os.listdir(path)
    pattern = re.compile(r'.*\.xlsx*$')
    list_xls = []
    for f in dir_file:
        list_xls.append(pattern.findall(f))
    result = [list_xls[k][0] for k in range(len(list_xls))
              if len(list_xls[k]) != 0]
    return result


def diff_div_sd(diff_list, sd_list):
    '''计算价差序列除以标准差序列得到的结果，以列表的形式返回'''
    diff_div_sd_list = [diff_list[i]/sd_list[i] for i in range(len(diff_list))]
    return diff_div_sd_list


def trade_begin(price_dict, buy_item, sell_item, beta_dict,
                trade_number, trade_cost_total, trade_log_list,
                trade_cost_rate, trade_date):
    '''开始交易的函数，计算买卖股票的数量及其相应的交易成本，
    将交易开始日期添加到交易日志列表中，并设置交易状态，返回
    cost_this, cost_total, trade_log_list, trade_state, in_trade'''
    buy = price_dict[buy_item][trade_date] * trade_number * beta_dict[buy_item]
    sell = price_dict[sell_item][trade_date] * trade_number * beta_dict[sell_item]
    cost_this = (buy + sell) * trade_cost_rate + buy        # 计算本次的交易成本
    new_trade_cost_total = trade_cost_total + cost_this     # 计算总的交易成本

    trade_state = {'shba': False, 'sabh': False}        # 初始化交易状态细节的字典
    if buy_item == 'a':             # 设置交易状态细节
        trade_state['shba'] = True
    else:
        trade_state['sabh'] = True
    in_trade = True         # 设置交易状态

    trade_log_list.append([trade_date])        # 写入交易日志列表
    trade_log_list[-1].append(buy_item)

    return cost_this, new_trade_cost_total, trade_log_list, trade_state, in_trade


def trade_end(price_dict, trade_state,
              beta_dict, trade_number, trade_cost_total,
              trade_cost_this, revenue_total, trade_log_list,
              trade_date, trade_cost_rate):
    '''交易结束的函数，计算买卖股票的数量及其 交易成本，以及交易利润，
    将交易结束日期、交易成本、交易收益添加到交易日志列表中，并设置交易状态，
    返回 cost_total, revenue_total, trade_log_list, trade_state, in_trade'''
    if trade_state['sabh']:
        buy_item = 'a'
        sell_item = 'h'
    else:
        buy_item = 'h'
        sell_item = 'a'

    buy = price_dict[buy_item][trade_date] * trade_number * beta_dict[buy_item]
    sell = price_dict[sell_item][trade_date] * trade_number * beta_dict[sell_item]

    cost_this = (buy + sell) * trade_cost_rate + buy + trade_cost_this          # 计算整个完整交易的成本
    new_trade_cost_total = trade_cost_total + (buy + sell) * trade_cost_rate + buy         # 计算至今为止的交易成本

    trade_state['sabh'] = False     # 还原未交易的状态
    trade_state['shba'] = False
    in_trade = False

    last_trade_date = trade_log_list[-1][0]
    revenue_this = (price_dict[buy_item][last_trade_date] - price_dict[buy_item][trade_date]) *\
        trade_number * beta_dict[buy_item] +\
        (price_dict[sell_item][trade_date] - price_dict[sell_item][last_trade_date]) *\
        trade_number * beta_dict[sell_item]                              # 计算本次完整交易的收益
    new_revenue_total = revenue_total + revenue_this        # 计算总的交易收益
    rate_of_revenue = revenue_this / cost_this * 100

    trade_log_list[-1].append(trade_date)           # 将交易信息添加到交易日志列表中
    trade_log_list[-1].append(revenue_this)
    trade_log_list[-1].append(rate_of_revenue)

    return new_trade_cost_total, new_revenue_total, trade_log_list, trade_state, in_trade


def print_log_by_trade(output_path, trade_log_list, time_str, stock_price_dict,
                       total_cost, total_revenue, name_msg, print_model='a+', file_name='trade_log'):
    '''打印交易日志，包含所有的组合的交易日志'''
    cwd = os.getcwd()
    if not(cwd == output_path):
        os.chdir(output_path)
    file_name_txt = file_name + '.txt'
    output_file = open(file_name_txt, print_model)

    rate_of_rev = total_revenue / total_cost * 100      # 计算收益率

    output_file.write(name_msg)
    for ith_trade in trade_log_list:
        output_file.write('\n')
        try:
            if ith_trade[1] == 'a':
                output_file.write('建仓时间：%s，买A卖H，买入价：%f，卖出价：%f\n'
                                  % (time_str[ith_trade[0]], stock_price_dict['a'][ith_trade[0]],
                                     stock_price_dict['h'][ith_trade[0]]))
                output_file.write('平仓时间：%s，卖A买H，卖出价：%f，买入价：%f，本次收益为：%f，收益率为：%f%%\n'
                                  % (time_str[ith_trade[2]], stock_price_dict['a'][ith_trade[2]],
                                     stock_price_dict['h'][ith_trade[2]], ith_trade[3], ith_trade[4]))
            else:
                output_file.write('建仓时间：%s，买H卖A，买入价：%f，卖出价：%f\n'
                                  % (time_str[ith_trade[0]], stock_price_dict['h'][ith_trade[0]],
                                     stock_price_dict['a'][ith_trade[0]]))
                output_file.write('平仓时间：%s，卖H买A，卖出价：%f，买入价：%f，本次收益为：%f，收益率为：%f%%\n'
                                  % (time_str[ith_trade[2]], stock_price_dict['h'][ith_trade[2]],
                                     stock_price_dict['a'][ith_trade[2]], ith_trade[3], ith_trade[4]))
        except IOError as err:
            print 'Error, unable to write log' + str(err)

    output_file.write('本次交易总收益为：%f,总的收益率为 %f%%\n'
                      % (total_revenue, rate_of_rev))

    output_file.write('\n|' + '-'*10 + 'NEXT-TRADE' + '-'*10 + '|\n\n')
    output_file.close()


def print_log_total(output_path, total_cost, total_revenue, file_name='trade_log'):
    '''用于打印整个交易总和最终所得到的收益和收益率'''
    cwd = os.getcwd()
    if not(cwd == output_path):
        os.chdir(output_path)

    rate_of_rev = total_revenue / total_cost * 100

    file_name_txt = file_name + '.txt'
    output_file = open(file_name_txt, 'a+')
    output_file.write('交易总计收益为：%f，交易总计收益率为：%f%%\n' % (total_revenue, rate_of_rev))
    output_file.close()


def plot_trade_procedure(diff_msg, sd_msg, constrain,
                         trigger, output_path, name_msg,
                         date_list):
    '''用于画出价差序列曲线，触发以及终止交易的区间曲线'''
    pl.figure(figsize=(30, 10), dpi=80)     # 设置图的长宽
    x_axle = np.linspace(0, len(diff_msg), len(diff_msg))   # 设置x轴

    up_constrain = np.array([i * constrain for i in sd_msg])        # 设置终止上界列表
    up_trigger = np.array([i * trigger for i in sd_msg])        # 设置触发上界列表
    bottom_trigger = np.array([i * (-trigger) for i in sd_msg])     # 设置触发下界列表
    bottom_constrain = np.array([i * (-constrain) for i in sd_msg])     # 设置终止下界列表
    diff_curve = np.array(diff_msg)     # 设置价差序列列表

    pl.xlim(x_axle.min(), x_axle.max() * 1.2)       # 设置x轴的画图范围
    # 设置x轴的标记
    pl.xticks([j for j in range(0, len(date_list), 22)], [date_list[j] for j in range(0, len(date_list), 22)])

    pl.ylim(min(bottom_constrain.min(), diff_curve.min()) * 1.5,
            max(up_constrain.max(), diff_curve.max()) * 1.5)        # 设置y轴的画图范围
    pl.yticks([-0.2, 0, 0.2], [r'$-0.2$', r'$0$', r'$0.2$'])        # 设置y轴的标记

    # 画出价差序列
    pl.plot(x_axle, diff_curve, color='blue', linewidth=1.5,
            linestyle='-', label='diff_curve')
    # 画出平仓上界曲线
    pl.plot(x_axle, up_constrain, color='red', linewidth=1,
            linestyle='--', label='constrain')
    # 画出触发上界曲线
    pl.plot(x_axle, up_trigger, color='green', linewidth=1,
            linestyle='--', label='trigger')
    # 画出平仓下界曲线
    pl.plot(x_axle, bottom_constrain, color='red', linewidth=1,
            linestyle='--')
    # 画出触发下界曲线
    pl.plot(x_axle, bottom_trigger, color='green', linewidth=1,
            linestyle='--')

    # 设置曲线说明的位置
    pl.legend(loc='upper right')

    # 调节轴线的位置
    ax = pl.gca()
    ax.spines['right'].set_color('none')        # 去除右边轴线
    ax.spines['top'].set_color('none')      # 去除上方轴线
    ax.xaxis.set_ticks_position('bottom')
    ax.spines['bottom'].set_position(('data', 0))
    ax.yaxis.set_ticks_position('left')
    ax.spines['left'].set_position(('data', 0))

    # 保存图片
    file_save = output_path + '\\' + name_msg + '.jpg'
    pl.savefig(file_save)


def trade_iteration(input_path, output_path):
    reload(sys)
    sys.setdefaultencoding('utf-8')
    xls_name_list = get_xls_file(input_path)

    # 初始化全局变量
    constrain = 3.0
    trigger = 1.5
    trade_num = 100
    rate_of_cost = 0.0004
    cost = .0       # 所有交易的总成本
    rev = .0        # 所有交易的总成本

    for xls_name in xls_name_list:      # 文件处理循环
        # 获取文件中的数据
        diff_msg, sd_msg, price_a, price_h, time_msg, name_msg, beta_msg = read_file(xls_name, input_path)

        # 计算价差序列除以标准差序列
        diff_div_sd_list = diff_div_sd(diff_msg, sd_msg)

        # 将日期型数据转换为字符，并加入总时间列表，供打印日志使用
        time_str = trans_time(time_msg)

        # 设置beta系数字典
        beta_dict = {'a': 1, 'h': beta_msg}

        stocks_price_dict = {'a': price_a, 'h': price_h}

        # 初始化域变量
        in_trade = False        # 用于判断是否处于交易中
        trade_state = {'sabh': False, 'shba': False}        # 用于判断所处的交易状态
        cost_total = 0
        revenue_total = 0
        trade_log_list = []     # 用于记录交易记录
        len_trade_day = len(time_str)       # 总的交易天数

        # 日期循环
        for ith_day in range(len_trade_day):
            if not in_trade:
                if (ith_day != 0) and (diff_div_sd_list[ith_day - 1] > constrain or diff_div_sd_list[-1] < -constrain):
                    continue
                elif trigger < diff_div_sd_list[ith_day] < constrain:
                    cost_this, cost_total, trade_log_list, trade_state, in_trade = trade_begin(stocks_price_dict,
                                                                                               'h', 'a', beta_dict,
                                                                                               trade_num, cost_total,
                                                                                               trade_log_list,
                                                                                               rate_of_cost, ith_day)
                elif -constrain < diff_div_sd_list[ith_day] < -trigger:
                    cost_this, cost_total, trade_log_list, trade_state, in_trade = trade_begin(stocks_price_dict,
                                                                                               'a', 'h', beta_dict,
                                                                                               trade_num, cost_total,
                                                                                               trade_log_list,
                                                                                               rate_of_cost, ith_day)
            else:
                if ith_day == len_trade_day - 1:  # 如果此时处于交易的最后一天，则无论如何都平仓
                    cost_total, revenue_total, trade_log_list, trade_state, in_trade = trade_end(stocks_price_dict,
                                                                                                 trade_state,
                                                                                                 beta_dict,
                                                                                                 trade_num,
                                                                                                 cost_total,
                                                                                                 cost_this,
                                                                                                 revenue_total,
                                                                                                 trade_log_list,
                                                                                                 ith_day,
                                                                                                 rate_of_cost)
                elif diff_msg[ith_day] * diff_msg[ith_day - 1] <= 0:  # 判断价差曲线是否突破横轴，及其突破的方式
                    judge_index = pow(diff_msg[ith_day], 2) - pow(diff_msg[ith_day - 1], 2)
                    if judge_index >= 0:
                        cost_total, revenue_total, trade_log_list, trade_state, in_trade = trade_end(stocks_price_dict,
                                                                                                     trade_state,
                                                                                                     beta_dict,
                                                                                                     trade_num,
                                                                                                     cost_total,
                                                                                                     cost_this,
                                                                                                     revenue_total,
                                                                                                     trade_log_list,
                                                                                                     ith_day,
                                                                                                     rate_of_cost)
                elif diff_msg[ith_day] * diff_msg[ith_day + 1] <= 0:  # 判断价差曲线是否突破横轴，及其突破的方式
                    judge_index = pow(diff_msg[ith_day], 2) - pow(diff_msg[ith_day + 1], 2)
                    if judge_index <= 0:
                        cost_total, revenue_total, trade_log_list, trade_state, in_trade = trade_end(stocks_price_dict,
                                                                                                     trade_state,
                                                                                                     beta_dict,
                                                                                                     trade_num,
                                                                                                     cost_total,
                                                                                                     cost_this,
                                                                                                     revenue_total,
                                                                                                     trade_log_list,
                                                                                                     ith_day,
                                                                                                     rate_of_cost)
                elif diff_div_sd_list[ith_day] < -constrain or diff_div_sd_list[ith_day] > constrain:
                    # 判断价差序列除标准差序列的结果是否突破了波动的上下界
                    cost_total, revenue_total, trade_log_list, trade_state, in_trade = trade_end(stocks_price_dict,
                                                                                                 trade_state,
                                                                                                 beta_dict,
                                                                                                 trade_num,
                                                                                                 cost_total,
                                                                                                 cost_this,
                                                                                                 revenue_total,
                                                                                                 trade_log_list,
                                                                                                 ith_day,
                                                                                                 rate_of_cost)
                else:
                    continue

        plot_trade_procedure(diff_msg, sd_msg, constrain, trigger, output_path, name_msg, time_str)     # 作图
        if xls_name_list.index(xls_name) == 0:
            print_log_by_trade(output_path, trade_log_list, time_str, stocks_price_dict,
                               cost_total, revenue_total, name_msg, 'w+')
        else:
            print_log_by_trade(output_path, trade_log_list, time_str, stocks_price_dict,
                               cost_total, revenue_total, name_msg)
        cost += cost_total
        rev += revenue_total
    print_log_total(output_path, cost, rev)

if __name__ == '__main__':
    path_input = r'C:\Users\Mercury\Desktop\stock\project'
    path_output = r'C:\Users\Mercury\Desktop\stock\project'
    trade_iteration(path_input, path_output)