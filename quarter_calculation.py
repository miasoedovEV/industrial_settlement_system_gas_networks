# -*- coding: utf-8 -*-
"""
Created on Tue Dec 29 09:44:49 2020

@author: stinc
"""
from in_house_check import CalculatorHousePipes
import numpy as np
import math as mt
import pandas as pd
from pprint import pprint
import openpyxl


class CalculatorQuarter:

    def __init__(self, file_name, file_out_name):
        self.file_in = file_name
        self.file_out = file_out_name

    def initializing_file(self):
        self.file_source_data = self.file_in
        # Load spreadsheet
        self.xl_source_data = pd.ExcelFile(self.file_source_data)

    def initializing_source_data(self):
        # Блок ввода
        self.df_source_data = self.xl_source_data.parse('Квартал')
        self.array_source_data = np.array(pd.DataFrame(self.df_source_data))
        self.df_pipes_data = self.xl_source_data.parse('information_pipes')
        self.array_pipes_data = np.array(pd.DataFrame(self.df_pipes_data))
        self.calc = CalculatorHousePipes(self.file_in, self.file_out)
        self.vis, self.V_house, self.i_d_hous, self.y_g, self.delta_y = self.calc.run()
        self.k_n_crash_1 = list(self.array_source_data[:, 0][1:])
        self.parcel_name_crash_1 = list(self.array_source_data[:, 1][1:])
        self.lenth_d_crash_1 = list(self.array_source_data[:, 2][1:])
        self.k_n_crash_2 = list(self.array_source_data[:, 4][1:])
        self.parcel_name_crash_2 = list(self.array_source_data[:, 5][1:])
        self.lenth_d_crash_2 = list(self.array_source_data[:, 6][1:])
        self.k_n = list(self.array_source_data[:, 8][1:])
        self.parcel_name = list(self.array_source_data[:, 9][1:])
        self.lenth_d = list(self.array_source_data[:, 10][1:])
        self.k = self.array_source_data[:, 11][0]
        self.dP = self.array_source_data[:, 12][0]
        self.df_pipes_data = self.xl_source_data.parse('information_pipes')
        self.array_pipes_data = np.array(pd.DataFrame(self.df_pipes_data))
        self.Dn_delta = list(self.array_pipes_data[:, 0])
        self.Dn = list(self.array_pipes_data[:, 1])
        self.delta = list(self.array_pipes_data[:, 2])
        self.Dvn = list(self.array_pipes_data[:, 3])
        self.sites_dP_crash_1 = []
        self.sites_dP_crash_2 = []
        for index, node in enumerate(self.array_source_data[:, 3][1:]):
            if type(node) == float:
                break
            self.sites_dP_crash_1.append(node)
            self.sites_dP_crash_2.append(self.array_source_data[:, 7][1:][index])

    def calculate(self):
        self.dict_lenth_d = {}
        for index, value in enumerate(self.lenth_d):
            self.dict_lenth_d[self.parcel_name[index]] = value
        self.dict_lenth_d_crash_1 = {}
        for index, value in enumerate(self.lenth_d_crash_1):
            self.dict_lenth_d_crash_1[self.parcel_name_crash_1[index]] = value
        self.dict_lenth_d_crash_2 = {}
        for index, value in enumerate(self.lenth_d_crash_2):
            self.dict_lenth_d_crash_2[self.parcel_name_crash_2[index]] = value
        self.Dvn.reverse()
        self.Dn_delta.reverse()
        self.dict_lenth_r_crash_1, self.dict_V_crash_1, self.dict_Dvn_crash_1, self.dict_Dn_delta_crash_1, \
        self.dict_Re_crash_1, self.dict_Re2_crash_1, \
        self.dict_lambda_crash_1, self.dict_dP_crash_1, self.sum_dP_1 = self.calculate_crash(self.k_n_crash_1,
                                                                                             self.parcel_name_crash_1,
                                                                                             self.lenth_d_crash_1,
                                                                                             self.sites_dP_crash_1)
        self.dict_lenth_r_crash_2, self.dict_V_crash_2, self.dict_Dvn_crash_2, self.dict_Dn_delta_crash_2, \
        self.dict_Re_crash_2, self.dict_Re2_crash_2, \
        self.dict_lambda_crash_2, self.dict_dP_crash_2, self.sum_dP_2 = self.calculate_crash(self.k_n_crash_2,
                                                                                             self.parcel_name_crash_2,
                                                                                             self.lenth_d_crash_2,
                                                                                             self.sites_dP_crash_2)
        self.choose_diametre_mode()
        self.lenth_r = [ld * self.k for ld in self.lenth_d]
        self.V = []
        for index, k in enumerate(self.k_n):
            if "'" in self.parcel_name[index]:
                self.V.append(self.V_house)
            elif np.isnan(k):
                self.V.append(0)
            else:
                self.V.append(self.V_house * k)
        self.dict_lenth_r = {}
        self.dict_V = {}
        self.dict_Re = {}
        self.dict_Re2 = {}
        self.dict_lambda = {}
        self.dict_dP = {}
        for index, name in enumerate(self.parcel_name):
            name_2 = self.check_dictionary_exists(name, self.dict_Dvn)
            self.dict_lenth_r[name] = self.lenth_r[index]
            self.dict_V[name] = self.V[index]
            re, re_2, lambda_, dP = self.calculate_dP_h(self.dict_Dvn[name_2], self.V[index], self.lenth_r[index])
            self.dict_Re[name] = re
            self.dict_Re2[name] = re_2
            self.dict_lambda[name] = lambda_
            self.dict_dP[name] = dP
        self.dict_Dn_delta_crash_1, self.dict_Dvn_crash_1, self.dict_Re_crash_1, self.dict_Re2_crash_1, \
        self.dict_lambda_crash_1, self.dict_dP_crash_1, \
        self.sum_dP_1 = self.last_calculate_crash(self.dict_Dvn, self.dict_V_crash_1, self.dict_lenth_r_crash_1,
                                                  self.sites_dP_crash_1, self.dict_Dn_delta)
        self.dict_Dn_delta_crash_2, self.dict_Dvn_crash_2, self.dict_Re_crash_2, self.dict_Re2_crash_2, \
        self.dict_lambda_crash_2, self.dict_dP_crash_2, \
        self.sum_dP_2 = self.last_calculate_crash(self.dict_Dvn, self.dict_V_crash_2, self.dict_lenth_r_crash_2,
                                                  self.sites_dP_crash_2, self.dict_Dn_delta)

    def last_calculate_crash(self, dict_Dvn, dict_V, dict_lenth_r, sites_dP_crash, dict_Dn_delta):
        dict_new_Dn_delta = {}
        dict_new_Dvn = {}
        dict_Re = {}
        dict_Re2 = {}
        dict_lambda = {}
        dict_dP = {}
        for key, v in dict_V.items():
            key_2 = self.check_dictionary_exists(key, dict_Dvn)
            dict_new_Dn_delta[key] = dict_Dn_delta[key_2]
            dict_new_Dvn[key] = dict_Dvn[key_2]
            re, re_2, lambda_, dP = self.calculate_dP_h(dict_Dvn[key_2], v, dict_lenth_r[key])
            dict_Re[key] = re
            dict_Re2[key] = re_2
            dict_lambda[key] = lambda_
            dict_dP[key] = dP
        sum_dP = self.check_sum_dP(dict_dP, sites_dP_crash)
        return dict_new_Dn_delta, dict_new_Dvn, dict_Re, dict_Re2, dict_lambda, dict_dP, sum_dP

    def choose_diametre_mode(self):
        self.dict_Dvn = {}
        self.dict_Dn_delta = {}
        for key, dvn in self.dict_Dvn_crash_1.items():
            key = self.check_dictionary_exists(key, self.dict_Dvn_crash_2)
            if dvn > self.dict_Dvn_crash_2[key]:
                self.dict_Dvn[key] = dvn
            else:
                self.dict_Dvn[key] = self.dict_Dvn_crash_2[key]
        for key, dv in self.dict_Dvn.items():
            for index, dvn in enumerate(self.Dvn):
                if dv == dvn:
                    self.dict_Dn_delta[key] = self.Dn_delta[index]

    def check_dictionary_exists(self, key, dict_Dvn):
        if key in dict_Dvn:
            return key
        else:
            list_node = key.split('-')
            key = list_node[1] + '-' + list_node[0]
            return key

    def calculate_crash(self, k_n, parcel_name, lenth_d, sites_dP_crash):
        lenth_r = [ld * self.k for ld in lenth_d]
        V = []
        for index, k in enumerate(k_n):
            if "'" in parcel_name[index]:
                V.append(self.V_house)
            else:
                V.append(self.V_house * k)
        dict_lenth_r, dict_V, dict_Dvn, dict_Dn_delta, \
        dict_Re, dict_Re2, dict_lambda, dict_dP = self.calculate_smallest_dv(parcel_name, lenth_r, self.Dvn[0], V, 0)
        for d_i, dv in enumerate(self.Dvn):
            for index, name in enumerate(parcel_name):
                if "'" in name:
                    continue
                dict_lenth_r[name] = lenth_r[index]
                dict_V[name] = V[index]
                dict_Dvn[name] = dv
                dict_Dn_delta[name] = self.Dn_delta[d_i]
                re, re_2, lambda_, dP = self.calculate_dP_h(dv, V[index], lenth_r[index])
                dict_Re[name] = re
                dict_Re2[name] = re_2
                dict_lambda[name] = lambda_
                dict_dP[name] = dP
                sum_dP = self.check_sum_dP(dict_dP, sites_dP_crash)
                if sum_dP is not None:
                    return dict_lenth_r, dict_V, dict_Dvn, dict_Dn_delta, \
                           dict_Re, dict_Re2, dict_lambda, dict_dP, sum_dP

    def check_sum_dP(self, dict_dP, sites_dP_crash):
        sum_dP = 0
        for name in sites_dP_crash:
            sum_dP += dict_dP[name]
        if sum_dP > self.dP:
            return None
        else:
            return sum_dP

    def calculate_smallest_dv(self, parcel_name, lenth_r, dv, V, d_i):
        dict_lenth_r = {}
        dict_V = {}
        dict_Dvn = {}
        dict_Dn_delta = {}
        dict_Re = {}
        dict_Re2 = {}
        dict_lambda = {}
        dict_dP = {}
        for index, name in enumerate(parcel_name):
            dict_lenth_r[name] = lenth_r[index]
            dict_V[name] = V[index]
            if "'" in name:
                dict_Dvn[name] = self.Dvn[self.i_d_hous]
                dict_Dn_delta[name] = self.Dn_delta[self.i_d_hous]
                re, re_2, lambda_, dP = self.calculate_dP_h(self.Dvn[self.i_d_hous], V[index], lenth_r[index])
            else:
                dict_Dvn[name] = dv
                dict_Dn_delta[name] = self.Dn_delta[d_i]
                re, re_2, lambda_, dP = self.calculate_dP_h(dv, V[index], lenth_r[index])
            dict_Re[name] = re
            dict_Re2[name] = re_2
            dict_lambda[name] = lambda_
            dict_dP[name] = dP
        return dict_lenth_r, dict_V, dict_Dvn, dict_Dn_delta, dict_Re, dict_Re2, dict_lambda, dict_dP

    def calculate_dP_h(self, dv, V, lenth):
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
        dP = 626.1 * lambda_ * V ** 2 * self.y_g * lenth * 1.1 / (dv ** 5)
        return re, re_2, lambda_, dP

    def find_crash_node(self, k_n, parcel_name):
        for index, k in enumerate(k_n):
            if k == 0:
                return parcel_name[index]

    def output_excel(self):
        wb = openpyxl.load_workbook(filename=self.file_out)
        sheet = wb['Квартал']
        names = ['№ участков', 'Lдейств,м', 'Lпривед,км', 'V м3/ч',
                 'Dн×δ', 'Dвн', 'Re', 'Re*(n/d)', 'λ', 'dP', 'sum_dP']
        letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
        sheet[
            'D1'] = f'Внутриквартальный при нерабочем {self.find_crash_node(self.k_n_crash_1, self.parcel_name_crash_1)}'
        for index, name in enumerate(names):
            sheet[f'{letters[index]}2'] = name
        i = 3
        for index, name in enumerate(self.parcel_name_crash_1):
            for key, value in self.dict_lenth_d_crash_1.items():
                if name == key:
                    sheet[f'A{i}'] = key
                    sheet[f'B{i}'] = value
                    sheet[f'C{i}'] = self.dict_lenth_r_crash_1[key]
                    sheet[f'D{i}'] = self.dict_V_crash_1[key]
                    sheet[f'E{i}'] = self.dict_Dn_delta_crash_1[key]
                    sheet[f'F{i}'] = self.dict_Dvn_crash_1[key]
                    sheet[f'G{i}'] = self.dict_Re_crash_1[key]
                    sheet[f'H{i}'] = self.dict_Re2_crash_1[key]
                    sheet[f'I{i}'] = self.dict_lambda_crash_1[key]
                    sheet[f'J{i}'] = self.dict_dP_crash_1[key]
                    i += 1
        sheet[f'K3'] = self.sum_dP_1
        letters = ['M', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']
        sheet[
            'R1'] = f'Внутриквартальный при нерабочем {self.find_crash_node(self.k_n_crash_2, self.parcel_name_crash_2)}'
        for index, letter in enumerate(letters):
            sheet[f'{letter}2'] = names[index]
        i = 3
        for index, name in enumerate(self.parcel_name_crash_2):
            for key, value in self.dict_lenth_d_crash_2.items():
                if name == key:
                    sheet[f'M{i}'] = key
                    sheet[f'O{i}'] = value
                    sheet[f'P{i}'] = self.dict_lenth_r_crash_2[key]
                    sheet[f'Q{i}'] = self.dict_V_crash_2[key]
                    sheet[f'R{i}'] = self.dict_Dn_delta_crash_2[key]
                    sheet[f'S{i}'] = self.dict_Dvn_crash_2[key]
                    sheet[f'T{i}'] = self.dict_Re_crash_2[key]
                    sheet[f'U{i}'] = self.dict_Re2_crash_2[key]
                    sheet[f'V{i}'] = self.dict_lambda_crash_2[key]
                    sheet[f'W{i}'] = self.dict_dP_crash_2[key]
                    i += 1
        sheet[f'X3'] = self.sum_dP_2
        del names[-1]
        letters = ['Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI']
        sheet['AE1'] = "Стационарный режим"
        for index, letter in enumerate(letters):
            sheet[f'{letter}2'] = names[index]
        i = 3
        for index, name in enumerate(self.parcel_name):
            for key, value in self.dict_lenth_d.items():
                if name == key:
                    key_2 = self.check_dictionary_exists(key, self.dict_Dn_delta)
                    sheet[f'Z{i}'] = key
                    sheet[f'AA{i}'] = value
                    sheet[f'AB{i}'] = self.dict_lenth_r[key]
                    sheet[f'AC{i}'] = self.dict_V[key]
                    sheet[f'AD{i}'] = self.dict_Dn_delta[key_2]
                    sheet[f'AE{i}'] = self.dict_Dvn[key_2]
                    sheet[f'AF{i}'] = self.dict_Re[key]
                    sheet[f'AG{i}'] = self.dict_Re2[key]
                    sheet[f'AH{i}'] = self.dict_lambda[key]
                    sheet[f'AI{i}'] = self.dict_dP[key]
                    i += 1
        wb.save(self.file_out)

    def run(self):
        self.initializing_file()
        self.initializing_source_data()
        self.calculate()
        self.output_excel()


# calc = CalculatorQuarter('source_data.xlsx', 'result.xlsx')
# calc.run()
