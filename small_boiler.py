# -*- coding: utf-8 -*-
"""
Created on Wed Dec 30 12:47:37 2020

@author: stinc
"""
from quarter_calculation import CalculatorQuarter
import numpy as np
import math as mt
import pandas as pd
from pprint import pprint
import openpyxl


class CalculatorBoiler:

    def __init__(self, file_name, file_out_name):
        self.file_in = file_name
        self.file_out = file_out_name

    def initializing_file(self):
        self.file_source_data = self.file_in
        # Load spreadsheet
        self.xl_source_data = pd.ExcelFile(self.file_source_data)

    def initializing_source_data(self):
        # Блок ввода
        self.df_source_data = self.xl_source_data.parse('Котельная')
        self.array_source_data = np.array(pd.DataFrame(self.df_source_data))
        self.df_pipes_data = self.xl_source_data.parse('information_pipes')
        self.array_pipes_data = np.array(pd.DataFrame(self.df_pipes_data))
        calc = CalculatorQuarter(self.file_in, self.file_out)
        calc.run()
        self.vis = calc.vis
        self.y_g = calc.y_g
        self.delta_y = calc.delta_y
        self.number_quarter = calc.calc.calc.calc.cal.source_data['number of quarters']
        self.V_small_boiler_houses = calc.calc.calc.calc.cal.V_small_boiler_houses
        self.beginnings_sections = []
        self.finish_sections = []
        self.parcel_name = []
        self.lenth_d = []
        self.number_boilers = []
        for index, beginning in enumerate(self.array_source_data[:, 0]):
            if np.isnan(beginning):
                break
            self.beginnings_sections.append(beginning)
            self.finish_sections.append(self.array_source_data[:, 1][index])
            self.parcel_name.append(self.array_source_data[:, 2][index])
            self.lenth_d.append(self.array_source_data[:, 3][index])
            self.number_boilers.append(self.array_source_data[:, 4][index])
        self.dict_heights = {}
        for index, numbers_knot in enumerate(list(self.array_source_data[:, 5])):
            self.dict_heights[numbers_knot] = list(self.array_source_data[:, 6])[index]
        self.sites_dP = []
        for index, node in enumerate(self.array_source_data[:, 7]):
            if type(node) == float:
                break
            self.sites_dP.append(node)
        self.dP = self.array_source_data[:, 8][0]
        self.k = self.array_source_data[:, 9][0]
        self.Dn_delta = list(self.array_pipes_data[:, 0])
        self.Dn = list(self.array_pipes_data[:, 1])
        self.delta = list(self.array_pipes_data[:, 2])
        self.Dvn = list(self.array_pipes_data[:, 3])

    def calculate(self):
        self.dict_lenth_d = {}
        self.dict_lenth_r = {}
        for index, lenth in enumerate(self.lenth_d):
            self.dict_lenth_d[self.parcel_name[index]] = lenth
            self.dict_lenth_r[self.parcel_name[index]] = self.k * lenth
        self.V_all_boilers_house = self.V_small_boiler_houses / self.number_quarter
        num_max_boiler = 0
        for num in self.number_boilers:
            if num > num_max_boiler:
                num_max_boiler = num
        self.V_one_boiler = self.V_all_boilers_house / num_max_boiler
        self.V = []
        for n in self.number_boilers:
            V = n * self.V_one_boiler
            self.V.append(V)
        feed_index = 0
        for index, bg in enumerate(self.beginnings_sections):
            if bg == 1:
                feed_index = index
            elif bg == '1':
                feed_index = index
        self.Dvn.reverse()
        self.Dn_delta.reverse()
        self.dict_V = {}
        self.dict_Dvn = {}
        self.dict_Dn_delta = {}
        self.dict_Re = {}
        self.dict_Re2 = {}
        self.dict_lambda = {}
        self.dict_dh = {}
        self.dict_hydrostatic_head = {}
        self.dict_dP = {}
        self.dict_dP_h = {}
        for index, name in enumerate(self.parcel_name):
            self.dict_lenth_r[name] = self.dict_lenth_r[name]
            self.dict_V[name] = self.V[index]
            self.dict_Dvn[name] = self.Dvn[0]
            self.dict_Dn_delta[name] = self.Dn_delta[0]
            re, re_2, lambda_, dh, hydrostatic_head, dP, dP_h = self.calculate_dP_h(self.Dvn[0], self.V[index],
                                                                                    self.dict_lenth_r[name],
                                                                                    self.dict_heights, index)
            self.dict_Re[name] = re
            self.dict_Re2[name] = re_2
            self.dict_lambda[name] = lambda_
            self.dict_dh[name] = dh
            self.dict_hydrostatic_head[name] = hydrostatic_head
            self.dict_dP[name] = dP
            self.dict_dP_h[name] = dP_h
        logic = False
        for i_d, dv in enumerate(self.Dvn):
            for index, name in enumerate(self.parcel_name):
                self.dict_lenth_r[name] = self.dict_lenth_r[name]
                self.dict_V[name] = self.V[index]
                self.dict_Dvn[name] = dv
                self.dict_Dn_delta[name] = self.Dn_delta[i_d]
                re, re_2, lambda_, dh, hydrostatic_head, dP, dP_h = self.calculate_dP_h(dv, self.V[index],
                                                                                        self.dict_lenth_r[name],
                                                                                        self.dict_heights, index)
                self.dict_Re[name] = re
                self.dict_Re2[name] = re_2
                self.dict_lambda[name] = lambda_
                self.dict_dh[name] = dh
                self.dict_hydrostatic_head[name] = hydrostatic_head
                self.dict_dP[name] = dP
                self.dict_dP_h[name] = dP_h
                self.sum_dP = self.check_sum_dP()
                if self.sum_dP is not None:
                    logic = True
                    break
            if logic is True:
                break

    def check_sum_dP(self):
        sum_dP = 0
        for key, value in self.dict_dP_h.items():
            for name in self.sites_dP:
                if name == key:
                    sum_dP += value
                    if sum_dP > self.dP:
                        return None
        return sum_dP

    def calculate_dP_h(self, dv, V, lenth, dict_heights, index):
        if V == 0:
            return [0] * 4
        re = 0.0354 * V / (dv * self.vis)
        re_2 = re * 0.01 / dv
        if re == 0 and re_2 == 0:
            lambda_ = 0
        elif re <= 2000:
            lambda_ = 64 / re
        elif re <= 4000:
            lambda_ = 0.0025 * re ** 0.333
        elif re < 100000 and re_2 < 23:
            lambda_ = 0.3164 / re ** 0.25
        elif re > 100000 and re_2 < 23:
            lambda_ = 1 / (1.821 * mt.log10(re) - 1.64) ** 2
        elif re_2 > 23 and re > 4000:
            lambda_ = 0.11 * (0.01 / dv + 68 / re) ** 0.25
        else:
            raise ValueError('Не найден соответствующий режим транспортировки газа: ValueError')
        delta_h = dict_heights[self.beginnings_sections[index]] - \
                  dict_heights[self.finish_sections[index]]
        hydrostatic_head = self.delta_y * delta_h
        dP = 626.1 * lambda_ * V ** 2 * self.y_g * lenth * 1.1 / (dv ** 5)
        dP_h = hydrostatic_head + dP
        return re, re_2, lambda_, delta_h, hydrostatic_head, dP, dP_h

    def create_arrive(self):
        names = ['Имя', 'Lд',
                 'Lр', 'V', 'Dн×δ', 'Dвн', 'Re', 'Re*(n/d)', 'λ', 'delta_h',
                 '+/-Нг', 'dP', 'dP_h']
        letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
        ar = {}
        for index, name in enumerate(names):
            if name == 'Имя':
                ar[name] = [letters[index], self.parcel_name]
            elif name == 'V':
                ar[name] = [letters[index], self.V]
            else:
                ar[name] = [letters[index], []]
        for index, name in enumerate(self.parcel_name):
            ar['Lд'][1].append(self.dict_lenth_r[name])
            ar['Lр'][1].append(self.dict_lenth_r[name])
            ar['Dн×δ'][1].append(self.dict_Dn_delta[name])
            ar['Dвн'][1].append(self.dict_Dvn[name])
            ar['Re'][1].append(self.dict_Re[name])
            ar['Re*(n/d)'][1].append(self.dict_Re2[name])
            ar['λ'][1].append(self.dict_lambda[name])
            ar['delta_h'][1].append(self.dict_dh[name])
            ar['+/-Нг'][1].append(self.dict_hydrostatic_head[name])
            ar['dP'][1].append(self.dict_dP[name])
            ar['dP_h'][1].append(self.dict_dP_h[name])
        return ar

    def output_excel(self):
        wb = openpyxl.load_workbook(filename=self.file_out)
        sheet = wb['Котельная']
        ar = self.create_arrive()
        for key, value in ar.items():
            sheet[f'{value[0]}1'] = key
            iter = 2
            for num in value[1]:
                sheet[f'{value[0]}{iter}'] = num
                iter += 1
        sheet[f'N1'] = 'Сумма перепадов'
        sheet['N2'] = self.sum_dP
        wb.save(self.file_out)

    def run(self):
        self.initializing_file()
        self.initializing_source_data()
        self.calculate()
        self.output_excel()


calc = CalculatorBoiler('source_data.xlsx', 'result.xlsx')
calc.run()
