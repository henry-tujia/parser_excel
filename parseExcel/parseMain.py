from addict import Dict
import pandas
from utils import Config
import os
import numpy


def thousand_sep(num):
    # num = ('%.2f'%num)
    if isinstance(num, str):
        num = float(num)
    return str(0) if num == 0 else format(num, '0,.2f')


def parse_main(root_path, combine_name, budget_name,year):
    root_path = root_path
    config = Config(os.path.join(root_path, combine_name).replace('\\', '/'),
                    os.path.join(root_path, budget_name).replace('\\', '/'),year)

    res_df = config.get_res_df()
    res_sheet = config.get_res_sheet()

    paras = Dict()

    paras["利润实现情况"] = str(year)+"年，蓝色创投实现利润总额{0}万元，实现全年预算利润总额{1}万元的{2}；实现净利润{3}万元，实现全年预算净利润{4}万元的{5}。".format(
        thousand_sep(res_df.loc['利润总额', 'value']), thousand_sep(
            res_df.loc['利润总额', 'budget_value']),
        res_df.loc['利润总额', 'percent'],
        thousand_sep(res_df.loc['净利润', 'value']), thousand_sep(
            res_df.loc['净利润', 'budget_value']),
        res_df.loc['净利润', 'percent'])
    paras["营业收入"] = str(year)+"年，蓝色创投实现营业收入{0}万元，实现全年预算营业收入{1}万元的{2}。".format(thousand_sep(res_df.loc['营业收入', 'value']),
                                                                        thousand_sep(
                                                                            res_df.loc['营业收入', 'budget_value']),
                                                                        res_df.loc['营业收入', 'percent'])
    paras["管理费用"] = str(year)+"年,管理费用累计发生{0}万元，完成全年预算{1}万元的{2}。".format(thousand_sep(res_df.loc['管理费用', 'value']),
                                                                  thousand_sep(
                                                                      res_df.loc['管理费用', 'budget_value']),
                                                                  res_df.loc['管理费用', 'percent'])
    paras["财务费用"] = str(year)+"年，财务费用累计发生{0}万元，完成全年预算{1}万元的{2}，主要原因为南海恒蓝少数股东享有的损益。".format(
        thousand_sep(res_df.loc['财务费用', 'value']), thousand_sep(
            res_df.loc['财务费用', 'budget_value']),
        res_df.loc['财务费用', 'percent'])
    paras["投资收益"] = str(year)+"年，投资收益累计实现{0}万元，完成全年预算{1}万元的{2}，主要原因为项目收益尚未实现。".format(
        thousand_sep(res_df.loc['投资收益', 'value']),
        thousand_sep(res_df.loc['投资收益', 'budget_value']),
        res_df.loc['投资收益', 'percent'])
    paras["公允价值变动损益"] = str(year)+"年，公允价值变动损益累计实现{0}万元，完成全年预算{1}万元的{2}，主要原因为实行新金融工具会计准则核算，预算投资项目出现的公允价值变动。".format(
        thousand_sep(res_df.loc['公允价值变动收益', 'value']), thousand_sep(
            res_df.loc['公允价值变动收益', 'budget_value']),
        res_df.loc['公允价值变动收益', 'percent'])

    return paras, res_sheet
