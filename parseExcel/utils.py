from addict import Dict
import pandas
import numpy
from decimal import Decimal


class Config:
    def __init__(self, combine_file, budget_file,year):
        self.combine_file = combine_file
        self.budget_file = budget_file
        self.keywords = ['管理费用', '营业收入', '净利润', '利润总额', '财务费用', '投资收益', '公允价值变动收益']
        self.sheet_name = ["利润表（新）", "利润预算表"]
        self.sheet_name_sample = "管理费用明细表"
        self.combine_df = None
        self.budget_df = None
        self.combine_sheet = None
        self.budget_sheet = None
        self.res_sheet = None
        self.year = str(year)
        self._openfile()
        self._processdf()
        self._opensheet()
        self._processsheet()

    def _openfile(self):
        combine_file = pandas.read_excel(self.combine_file, self.sheet_name[0])
        combine_file = combine_file.drop(index=0)
        combine_file.columns = combine_file.iloc[0]
        combine_file = combine_file.drop(index=1)
        combine_file = combine_file.reset_index()
        combine_file = combine_file.drop(axis=1, columns=["index"])
        combine_file_0 = combine_file.iloc[:, 0:6]
        combine_file_1 = combine_file.iloc[:, 6:]
        self.combine_df = pandas.concat([combine_file_0, combine_file_1])

        budget_file = pandas.read_excel(self.budget_file, self.sheet_name[1])
        budget_file = budget_file.drop(index=0).drop(index=1)
        budget_file.columns = budget_file.iloc[0]
        budget_file = budget_file.drop(index=2).drop(index=3)
        budget_file = budget_file.reset_index()
        self.budget_df = budget_file.drop(axis=1, columns=["index"])

    def _processdf(self):
        res_df = pandas.DataFrame(index=self.keywords, columns=["value", "budget_value"])

        for keyword in self.keywords:
            value_index = self.combine_df['项   目'].astype(str).str.contains(keyword)
            value = self.combine_df[value_index]['本年累计数']  # 本年累计数

            budget_value_index = self.budget_df['项         目'].astype(str).str.contains(keyword)
            budget_value = self.budget_df[budget_value_index]['本年预算数']

            res_df['value'][keyword] = str(Decimal(value.values[0]/ 10000).quantize(Decimal("0.01"),rounding="ROUND_HALF_UP"))
            res_df['budget_value'][keyword] = budget_value.values[0]


        res_df["percent"] = res_df.apply(
            lambda x: str(Decimal(100*float(x['value']) / x['budget_value']).quantize(Decimal("0.01"),
                                                                                 rounding="ROUND_HALF_UP")) + "%",
            axis=1)
        self.res_df = res_df

    def _opensheet(self):
        combine_sheet = pandas.read_excel(self.combine_file, self.sheet_name_sample)
        combine_sheet.columns = combine_sheet.iloc[1]
        combine_sheet = combine_sheet.drop(index=0).drop(index=1).reset_index()
        self.combine_sheet = combine_sheet.loc[:, ["项目", "本年累计数"]]
        budget_sheet = pandas.read_excel(self.budget_file, self.sheet_name_sample)
        budget_sheet.columns = budget_sheet.iloc[1]
        budget_sheet = budget_sheet.drop(index=0).drop(index=1).reset_index()
        self.budget_sheet = budget_sheet.loc[:, ["项目", self.year+"年预算"]]

        self.res_sheet = pandas.merge(self.combine_sheet, self.budget_sheet, on="项目", how='outer')

    def _processsheet(self):
        self.res_sheet = self.res_sheet.fillna(0)
        # cz.dropna(axis=0, how='any', inplace=True)

        self.res_sheet = self.res_sheet[(self.res_sheet['本年累计数'] != 0) | (self.res_sheet[self.year+"年预算"] != 0)]
        self.res_sheet['本年累计数'] = self.res_sheet['本年累计数'].map(lambda x: str(Decimal(x/10000+0.00000001).quantize(Decimal("0.01"),
                                                                            rounding="ROUND_HALF_UP")))
        self.res_sheet[self.year+"年预算"] = self.res_sheet[self.year+"年预算"].map(lambda x: str(Decimal(x+0.00000001).quantize(Decimal("0.01"),
                                                                            rounding="ROUND_HALF_UP")))
        self.res_sheet["percent"] = self.res_sheet.apply(
            lambda x: str(Decimal(100*float(x['本年累计数']) / float(x[self.year+"年预算"])+0.00000001).quantize(Decimal("0.01"),
                                                                            rounding="ROUND_HALF_UP")) + "%" if float(x[self.year+"年预算"]) != 0 else 'N/A', axis=1)
        columns = ['项目', self.year+"年预算", '本年累计数', 'percent']
        self.res_sheet = self.res_sheet[columns]

    def get_res_sheet(self):
        return self.res_sheet

    def get_res_df(self):
        return self.res_df
