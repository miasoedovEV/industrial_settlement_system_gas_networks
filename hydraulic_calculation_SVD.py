# -*- coding: utf-8 -*-
"""
Created on Fri Nov 27 17:02:49 2020

@author: stinc
"""
from hydraulic_calculation_SND import HydraulicsСalculatorSND
import numpy as np
import math as mt
import pandas as pd
from pprint import pprint
import openpyxl


def find_nodes(data, direction):
    for node_information in enumerate(data):
        if node_information[1] == direction or node_information[1] == 'Ист':
            yield node_information[0]


class HydraulicsСalculatorSVD:

    def __init__(self, file_name, file_out_name):
        self.file_in = file_name
        self.file_out = file_out_name

    def initializing_file(self):
        self.file_source_data = self.file_in
        # Load spreadsheet
        self.xl_source_data = pd.ExcelFile(self.file_source_data)

    def initializing_source_data(self):
        # Блок ввода
        self.df_source_data_SND = self.xl_source_data.parse('СНД')
        self.array_source_data_SND = np.array(pd.DataFrame(self.df_source_data_SND))
        self.df_source_data = self.xl_source_data.parse('СВД')
        self.array_source_data = np.array(pd.DataFrame(self.df_source_data))  # Масив с исходными данными
        self.df_pipes_data = self.xl_source_data.parse('information_pipes')
        self.array_pipes_data = np.array(pd.DataFrame(self.df_pipes_data))
        self.Vpp = {}
        for index, node in enumerate(self.array_source_data[:, 0]):
            if np.isnan(node):
                break
            self.Vpp[int(node)] = self.array_source_data[:, 1][index]
        self.classes_site = list(self.array_source_data[:, 2])
        self.beginnings_sections = list(self.array_source_data[:, 3])
        self.finish_sections = list(self.array_source_data[:, 4])
        self.parcel_name = list(self.array_source_data[:, 5])
        self.lenth_d = self.array_source_data[:, 6]
        self.list_site_category = list(self.array_source_data[:, 7])
        self.pressure_after_grs = self.array_source_data[:, 8][0]
        self.density_gas = self.array_source_data[:, 9][0]
        self.calc = HydraulicsСalculatorSND(self.file_in, self.file_out)
        self.dict_V_prg, self.V, self.vis = self.calc.run()
        self.V_bakeries_hour = self.V[1]
        self.V_laundromats_hour = self.V[2]
        self.V_big_boiler_houses = self.V[3]
        self.Dn_delta = list(self.array_pipes_data[:, 0])
        self.Dn = list(self.array_pipes_data[:, 1])
        self.delta = list(self.array_pipes_data[:, 2])
        self.Dvn = list(self.array_pipes_data[:, 3])
        self.allowable_pressure_PP1 = self.array_source_data[:, 11][0]
        self.allowable_pressure_PP2 = self.array_source_data[:, 12][0]
        self.k = self.array_source_data[:, 10][0]

    def calculation_SVD(self):
        self.reduced_length = [ld * self.k for i, ld in enumerate(self.lenth_d)]
        self.dict_length = {}
        self.dict_reduced_length = {}
        for index, length in enumerate(self.lenth_d):
            self.dict_length[self.parcel_name[index]] = length
            self.dict_reduced_length[self.parcel_name[index]] = self.reduced_length[index]

        # Расчёт аварий
        self.Dvn.reverse()
        self.Dn_delta.reverse()
        dv_good = self.Dvn[-3::]
        for d in dv_good:
            print(self.Dvn)
            checker = False
            self.dict_length_po, self.dict_reduced_lengt_po, self.dict_V_po, self.dict_Re_crash_po, self.dict_Re2_crash_po, \
            self.dict_Dn_delta_po, self.dict_finish_dvn_po, self.dict_lambda_crash_po, self.dict_Pn_node_crash_po, \
            self.dict_Pk_node_crash_po, self.dict_dP_crash_po, self.name_out_node_po = self.calculation_crash('ПО', d)
            self.dict_length_pr, self.dict_reduced_lengt_pr, self.dict_V_pr, self.dict_Re_crash_pr, self.dict_Re2_crash_pr, \
            self.dict_Dn_delta_pr, self.dict_finish_dvn_pr, self.dict_lambda_crash_pr, self.dict_Pn_node_crash_pr, \
            self.dict_Pk_node_crash_pr, self.dict_dP_crash_pr, self.name_out_node_pr = self.calculation_crash('ПР', d)
            direction_N = self.calculate_stationary_mode()
            if self.N > 10:
                name_reduction_node, name_new_node_1, name_new_node_2 = self.make_good_N(direction_N)
                self.remark_main_list(name_reduction_node, direction_N, name_new_node_1, name_new_node_2)
                self.Dvn.reverse()
                self.Dn_delta.reverse()
                self.dict_length_po, self.dict_reduced_lengt_po, self.dict_V_po, self.dict_Re_crash_po, self.dict_Re2_crash_po, \
                self.dict_Dn_delta_po, self.dict_finish_dvn_po, self.dict_lambda_crash_po, self.dict_Pn_node_crash_po, \
                self.dict_Pk_node_crash_po, self.dict_dP_crash_po, self.name_out_node_po = self.recalculation_crash(
                    'ПО')
                if self.dict_length_po is None:
                    self.back_last_values()
                    continue
                self.dict_length_pr, self.dict_reduced_lengt_pr, self.dict_V_pr, self.dict_Re_crash_pr, self.dict_Re2_crash_pr, \
                self.dict_Dn_delta_pr, self.dict_finish_dvn_pr, self.dict_lambda_crash_pr, self.dict_Pn_node_crash_pr, \
                self.dict_Pk_node_crash_pr, self.dict_dP_crash_pr, self.name_out_node_pr = self.recalculation_crash(
                    'ПР')
                if self.dict_length_pr is None:
                    self.back_last_values()
                    continue
                checker = self.check_dv()
                if checker:
                    self.back_last_values()
                    continue
                self.calculate_stationary_mode()
                break
            else:
                break
        self.output_excel()

    def back_last_values(self):
        self.list_site_category = list(self.array_source_data[:, 7])
        self.classes_site = list(self.array_source_data[:, 2])
        self.parcel_name = list(self.array_source_data[:, 5])
        self.beginnings_sections = list(self.array_source_data[:, 3])
        self.finish_sections = list(self.array_source_data[:, 4])
        self.reduced_length = [ld * self.k for i, ld in enumerate(self.lenth_d)]
        self.dict_length = {}
        self.dict_reduced_length = {}
        for index, length in enumerate(self.lenth_d):
            self.dict_length[self.parcel_name[index]] = length
            self.dict_reduced_length[self.parcel_name[index]] = self.reduced_length[index]
        self.Dvn.reverse()
        self.Dn_delta.reverse()

    def check_dv(self):
        for index, fs in enumerate(self.finish_sections):
            for i, bs in enumerate(self.beginnings_sections):
                if fs == bs and self.dict_finish_dvn[self.parcel_name[index]] < self.dict_finish_dvn[
                    self.parcel_name[i]]:
                    return True

    def calculate_stationary_mode(self):
        # Определение расходов при стационарном режиме
        self.dict_V = {}
        self.name_site_highest_consumption = []
        consumption = 0
        for index, class_ in enumerate(self.classes_site):
            if class_ != 'Туп':
                continue
            if 'Пек' in self.finish_sections[index]:
                self.dict_V[self.parcel_name[index]] = self.V_bakeries_hour
            elif 'ПП' in self.finish_sections[index]:
                self.dict_V[self.parcel_name[index]] = self.Vpp[int(self.finish_sections[index][-1])]
            elif 'Прач' in self.finish_sections[index]:
                self.dict_V[self.parcel_name[index]] = self.V_laundromats_hour
            elif 'ТЭЦ' in self.finish_sections[index]:
                self.dict_V[self.parcel_name[index]] = self.V_big_boiler_houses
            elif 'ПРГ' in self.finish_sections[index]:
                self.dict_V[self.parcel_name[index]] = self.dict_V_prg[(self.finish_sections[index][-1])]
            if self.dict_V[self.parcel_name[index]] > consumption:
                self.name_site_highest_consumption = [index, self.parcel_name[index]]
                consumption = self.dict_V[self.parcel_name[index]]
        self.site_category = set(self.list_site_category)
        for all_category in self.site_category:
            if all_category == 0:
                continue
            for index, category in enumerate(self.list_site_category):
                if category == all_category:
                    required_expense = 0
                    sources = 0
                    for i, bs in enumerate(self.beginnings_sections):
                        if str(self.finish_sections[index]) == str(bs):
                            required_expense += self.dict_V[self.parcel_name[i]]
                    for i, fs in enumerate(self.finish_sections):
                        if str(self.finish_sections[index]) == str(fs):
                            sources += 1
                    transit_one = required_expense / sources
                    self.dict_V[self.parcel_name[index]] = transit_one

        # Определение параметров при стационарном режиме
        self.dict_Re = {}
        self.dict_Re2 = {}
        self.dict_finish_dvn = {}
        self.dict_Dn_delta = {}
        self.dict_lambda = {}
        self.dict_Pn_node = {}
        self.dict_Pk_node = {}
        self.dict_dP = {}
        for index, name in enumerate(self.parcel_name):
            for key, value in self.dict_finish_dvn_po.items():
                if name == key:
                    self.dict_finish_dvn[key] = value
                    self.dict_Dn_delta[key] = self.dict_Dn_delta_po[key]
        for index, name in enumerate(self.parcel_name):
            for key, value in self.dict_finish_dvn_pr.items():
                if name == key:
                    self.dict_finish_dvn[key] = value
                    self.dict_Dn_delta[key] = self.dict_Dn_delta_pr[key]
        site_category = list(set(self.list_site_category))
        site_category.reverse()
        for category in enumerate(site_category):
            for index, cat in enumerate(self.list_site_category):
                if category[1] != cat:
                    continue
                for key, dv in self.dict_finish_dvn.items():
                    if self.parcel_name[index] != key:
                        continue
                    if category[0] == 0:
                        self.dict_Pn_node[key] = self.pressure_after_grs
                    else:
                        required_Pk = 0
                        sources = 0
                        for i_f, fs in enumerate(self.finish_sections):
                            if str(self.beginnings_sections[index]) == str(fs):
                                required_Pk += self.dict_Pk_node[self.parcel_name[i_f]]
                                sources += 1
                        Pk = required_Pk / sources
                        self.dict_Pn_node[key] = Pk
                    re = 0.0354 * self.dict_V[key] / (dv * self.vis)
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
                    Pk = ((self.dict_Pn_node[key] + 0.1013) ** 2 - 1.2687 * 10 ** (-4) * lambda_ * self.dict_V[
                        key] ** 2 / dv ** 5 * self.density_gas * self.dict_reduced_length[key])
                    Pk = Pk ** 0.5
                    Pk -= 0.1013
                    dP = self.dict_Pn_node[key] - Pk
                    self.dict_Re[key] = re
                    self.dict_Re2[key] = re_2
                    self.dict_lambda[key] = lambda_
                    self.dict_Pk_node[key] = Pk
                    self.dict_dP[key] = dP
        sum_dP_po = 0
        sum_dP_pr = 0
        for index, site in enumerate(self.classes_site):
            for key, dp in self.dict_dP.items():
                if site == 'ПО' and key == self.parcel_name[index]:
                    sum_dP_po += dp
                elif site == 'ПР' and key == self.parcel_name[index]:
                    sum_dP_pr += dp
        if sum_dP_po > sum_dP_pr:
            self.N = (sum_dP_po - sum_dP_pr) * 100 / sum_dP_po
            direction_N = 'ПР'
        elif sum_dP_pr > sum_dP_po:
            self.N = (sum_dP_pr - sum_dP_po) * 100 / sum_dP_pr
            direction_N = 'ПО'
        else:
            self.N = 0
        return direction_N

    def remark_main_list(self, name_reduction_node, direction_N, name_new_node_1, name_new_node_2):
        cat = 10 ** 25
        quantity_PO = 0
        quantity_PR = 0
        for class_ in self.classes_site:
            if class_ == 'ПО':
                quantity_PO += 1
            elif class_ == 'ПР':
                quantity_PR += 1
        for index, category in enumerate(self.list_site_category):
            if quantity_PR == quantity_PO and self.classes_site[index] == 'Ист':
                self.list_site_category[index] = category + 1
            if self.classes_site[index] == direction_N and category < cat:
                cat = category
            if self.classes_site[index] == direction_N:
                self.list_site_category[index] = category + 1
        del self.list_site_category[name_reduction_node[0]]
        del self.classes_site[name_reduction_node[0]]
        del self.parcel_name[name_reduction_node[0]]
        del self.finish_sections[name_reduction_node[0]]
        del self.beginnings_sections[name_reduction_node[0]]
        del self.reduced_length[name_reduction_node[0]]
        self.list_site_category.append(cat + 1)
        self.list_site_category.append(cat)
        self.classes_site.append(direction_N)
        self.classes_site.append(direction_N)
        self.parcel_name.append(name_new_node_1)
        self.parcel_name.append(name_new_node_2)
        b1, f1 = self.parcel_name[-2].split('-')
        b2, f2 = self.parcel_name[-1].split('-')
        self.beginnings_sections.append(b1)
        self.finish_sections.append(f1)
        self.beginnings_sections.append(b2)
        self.finish_sections.append(f2)
        self.reduced_length.append(self.dict_reduced_length[name_new_node_1])
        self.reduced_length.append(self.dict_reduced_length[name_new_node_2])

    def recalculation_crash(self, dir):
        list_category, parcel_name, beginnings_sections, finish_sections, name_out_node = self.create_classes_site(dir)
        dict_V = {}
        for category, index in list_category:
            if self.classes_site[index] != 'Туп':
                continue
            if 'Пек' in self.finish_sections[index]:
                dict_V[parcel_name[index][1]] = self.V_bakeries_hour
            elif 'ПП' in self.finish_sections[index]:
                dict_V[parcel_name[index][1]] = self.Vpp[int(self.finish_sections[index][-1])]
            elif 'Прач' in self.finish_sections[index]:
                dict_V[parcel_name[index][1]] = self.V_laundromats_hour
            elif 'ТЭЦ' in self.finish_sections[index]:
                dict_V[parcel_name[index][1]] = self.V_big_boiler_houses
            elif 'ПРГ' in self.finish_sections[index]:
                dict_V[parcel_name[index][1]] = self.dict_V_prg[(self.finish_sections[index][-1])]
        site_category = []
        for category, index in list_category:
            if type(category) != int:
                continue
            site_category.append(category)
        site_category.sort()
        site_category = set(site_category)
        for all_category in site_category:
            if all_category == 0:
                continue
            for category, index in list_category:
                if category == all_category:
                    required_expense = 0
                    sources = 0
                    for i, bs in enumerate(beginnings_sections):
                        if str(finish_sections[index]) == str(bs):
                            required_expense += dict_V[parcel_name[i][1]]
                    for fs in finish_sections:
                        if str(finish_sections[index]) == str(fs):
                            sources += 1
                    transit_one = required_expense / sources
                    dict_V[parcel_name[index][1]] = transit_one
        site_category = list(site_category)
        site_category.sort(reverse=True)

        # Расчёт давления для аварий
        dict_length = {}
        dict_reduced_length = {}
        dict_Re = {}
        dict_Re2 = {}
        dict_Dn_delta = {}
        dict_finish_dvn = {}
        dict_lambda = {}
        dict_Pn_node = {}
        dict_Pk_node = {}
        dict_dP = {}
        for category in enumerate(site_category):
            for cat, index in list_category:
                if category[1] != cat:
                    continue
                if category[0] == 0:
                    dict_Pn_node[parcel_name[index][1]] = self.pressure_after_grs
                else:
                    index_Pn = self.find_Pn(index, beginnings_sections, finish_sections)
                    dict_Pn_node[parcel_name[index][1]] = dict_Pk_node[parcel_name[index_Pn][1]]
                re = 0.0354 * dict_V[parcel_name[index][1]] / (self.dict_finish_dvn[parcel_name[index][0]] * self.vis)
                re_2 = re * 0.01 / self.dict_finish_dvn[parcel_name[index][0]]
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
                    lambda_ = 0.11 * (0.01 / self.dict_finish_dvn[parcel_name[index][0]] + 68 / re) ** 0.25
                else:
                    raise ValueError('Не найден соответствующий режим транспортировки газа: ValueError')
                Pk = ((dict_Pn_node[parcel_name[index][1]] + 0.1013) ** 2 - 1.2687 * 10 ** (-4) * lambda_ * dict_V[
                    parcel_name[index][1]] ** 2 / self.dict_finish_dvn[parcel_name[index][0]] ** 5 * self.density_gas *
                      self.dict_reduced_length[parcel_name[index][0]])
                Pk = Pk ** 0.5
                Pk -= 0.1013
                if 'ПП1' in parcel_name[index][0] and Pk < self.allowable_pressure_PP1:
                    dv, dn, re, re_2, lambda_, Pk = self.recalculate_dv_crash(dict_Pn_node[parcel_name[index][1]],
                                                                              dict_V[parcel_name[index][1]],
                                                                              self.dict_reduced_length[
                                                                                  parcel_name[index][0]],
                                                                              parcel_name[index][0],
                                                                              self.classes_site[index])
                    if dv is None:
                        return [None] * 12
                elif 'ПП2' in parcel_name[index][0] and Pk < self.allowable_pressure_PP2:
                    dv, dn, re, re_2, lambda_, Pk = self.recalculate_dv_crash(dict_Pn_node[parcel_name[index][1]],
                                                                              dict_V[
                                                                                  parcel_name[index][1]],
                                                                              self.dict_reduced_length[
                                                                                  parcel_name[index][0]],
                                                                              parcel_name[index][0],
                                                                              self.classes_site[index])
                    if dv is None:
                        return [None] * 12
                elif Pk < 0.31 and self.classes_site[index] == 'Туп' and 'ПП1' not in parcel_name[index][0] and \
                        'ПП2' not in parcel_name[index][0]:
                    dv, dn, re, re_2, lambda_, Pk = self.recalculate_dv_crash(dict_Pn_node[parcel_name[index][1]],
                                                                              dict_V[parcel_name[index][1]],
                                                                              self.dict_reduced_length[
                                                                                  parcel_name[index][0]],
                                                                              parcel_name[index][0],
                                                                              self.classes_site[index])
                    if dv is None:
                        return [None] * 12
                else:
                    dv, dn = self.dict_finish_dvn[parcel_name[index][0]], self.dict_Dn_delta[parcel_name[index][0]],
                dP = dict_Pn_node[parcel_name[index][1]] - Pk
                dict_length[parcel_name[index][1]] = self.dict_length[parcel_name[index][0]]
                dict_reduced_length[parcel_name[index][1]] = self.dict_reduced_length[parcel_name[index][0]]
                dict_Dn_delta[parcel_name[index][1]] = dn
                dict_finish_dvn[parcel_name[index][1]] = dv
                dict_Re[parcel_name[index][1]] = re
                dict_Re2[parcel_name[index][1]] = re_2
                dict_lambda[parcel_name[index][1]] = lambda_
                dict_Pk_node[parcel_name[index][1]] = Pk
                dict_dP[parcel_name[index][1]] = dP
        return dict_length, dict_reduced_length, dict_V, dict_Re, dict_Re2, dict_Dn_delta, dict_finish_dvn, dict_lambda, \
               dict_Pn_node, dict_Pk_node, dict_dP, name_out_node

    def make_good_N(self, direction_N):
        name_reduction_node = []
        for index, site in enumerate(self.classes_site):
            if site == direction_N and self.finish_sections[index] == self.beginnings_sections[
                self.name_site_highest_consumption[0]]:
                name_reduction_node = [index, self.parcel_name[index]]
        name_new_node_1 = f"{self.beginnings_sections[name_reduction_node[0]]}-{self.beginnings_sections[name_reduction_node[0]]}'"
        name_new_node_2 = f"{self.beginnings_sections[name_reduction_node[0]]}'-{self.finish_sections[name_reduction_node[0]]}"
        list_caf = [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]
        check = False
        self.dict_dP.pop(name_reduction_node[1])
        for caf in list_caf:
            self.dict_length[name_new_node_1] = self.dict_length[name_reduction_node[1]] * caf
            self.dict_length[name_new_node_2] = self.dict_length[name_reduction_node[1]] * (1 - caf)
            self.dict_reduced_length[name_new_node_1] = self.dict_reduced_length[name_reduction_node[1]] * caf
            self.dict_reduced_length[name_new_node_2] = self.dict_reduced_length[name_reduction_node[1]] * (1 - caf)
            self.dict_V[name_new_node_1] = self.dict_V[name_reduction_node[1]]
            self.dict_V[name_new_node_2] = self.dict_V[name_reduction_node[1]]
            index_Pn = self.find_Pn(name_reduction_node[0], self.beginnings_sections, self.finish_sections)
            self.dict_Pn_node[name_new_node_1] = self.dict_Pk_node[self.parcel_name[index_Pn]]
            re = 0.0354 * self.dict_V[name_new_node_1] / (self.dict_finish_dvn[name_reduction_node[1]] * self.vis)
            re_2 = re * 0.01 / self.dict_finish_dvn[name_reduction_node[1]]
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
                lambda_ = 0.11 * (0.01 / self.dict_finish_dvn[name_reduction_node[1]] + 68 / re) ** 0.25
            else:
                raise ValueError('Не найден соответствующий режим транспортировки газа: ValueError')
            Pk = ((self.dict_Pn_node[name_new_node_1] + 0.1013) ** 2 - 1.2687 * 10 ** (-4) * lambda_ * self.dict_V[
                name_new_node_1] ** 2 / self.dict_finish_dvn[name_reduction_node[1]] ** 5 * self.density_gas *
                  self.dict_reduced_length[name_new_node_1])
            Pk = Pk ** 0.5
            Pk -= 0.1013
            dP = self.dict_Pn_node[name_new_node_1] - Pk
            self.dict_Re[name_new_node_1] = re
            self.dict_Re2[name_new_node_1] = re_2
            self.dict_Dn_delta[name_new_node_1] = self.dict_Dn_delta[name_reduction_node[1]]
            self.dict_finish_dvn[name_new_node_1] = self.dict_finish_dvn[name_reduction_node[1]]
            self.dict_lambda[name_new_node_1] = lambda_
            self.dict_Pk_node[name_new_node_1] = Pk
            self.dict_dP[name_new_node_1] = dP
            for i, dv in enumerate(self.Dvn):
                self.dict_Pn_node[name_new_node_2] = self.dict_Pk_node[name_new_node_1]
                re = 0.0354 * self.dict_V[name_new_node_2] / (dv * self.vis)
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
                Pk = ((self.dict_Pn_node[name_new_node_2] + 0.1013) ** 2 - 1.2687 * 10 ** (-4) * lambda_ * self.dict_V[
                    name_new_node_2] ** 2 / dv ** 5 * self.density_gas * self.dict_reduced_length[name_new_node_2])
                if Pk < 0:
                    continue
                Pk = Pk ** 0.5
                Pk -= 0.1013
                dP = self.dict_Pn_node[name_new_node_2] - Pk
                self.dict_Re[name_new_node_2] = re
                self.dict_Re2[name_new_node_2] = re_2
                self.dict_Dn_delta[name_new_node_2] = self.Dn_delta[i]
                self.dict_finish_dvn[name_new_node_2] = dv
                self.dict_lambda[name_new_node_2] = lambda_
                self.dict_Pk_node[name_new_node_2] = Pk
                self.dict_dP[name_new_node_2] = dP
                sum_dP_po = 0
                sum_dP_pr = 0
                for index, site in enumerate(self.classes_site):
                    for key, dp in self.dict_dP.items():
                        if site == 'ПО' and key == self.parcel_name[index]:
                            sum_dP_po += dp
                        elif site == 'ПР' and key == self.parcel_name[index]:
                            sum_dP_pr += dp
                if direction_N == 'ПО':
                    sum_dP_po += self.dict_dP[name_new_node_1] + self.dict_dP[name_new_node_2]
                elif direction_N == 'ПР':
                    sum_dP_pr += self.dict_dP[name_new_node_1] + self.dict_dP[name_new_node_2]
                if sum_dP_po > sum_dP_pr:
                    self.N = (sum_dP_po - sum_dP_pr) * 100 / sum_dP_po
                elif sum_dP_pr > sum_dP_po:
                    self.N = (sum_dP_pr - sum_dP_po) * 100 / sum_dP_pr
                else:
                    self.N = 0
                if self.N < 10:
                    check = True
                    break
            if check:
                list_dict = [self.dict_length, self.dict_reduced_length,
                             self.dict_V, self.dict_Re,
                             self.dict_Re2, self.dict_Dn_delta, self.dict_finish_dvn,
                             self.dict_lambda, self.dict_Pn_node, self.dict_Pk_node]
                for dict_ in list_dict:
                    dict_.pop(name_reduction_node[1])
                break
        return name_reduction_node, name_new_node_1, name_new_node_2

    def calculation_crash(self, dir, d):
        list_category, parcel_name, beginnings_sections, finish_sections, name_out_node = self.create_classes_site(dir)
        dict_V = {}
        for category, index in list_category:
            if self.classes_site[index] != 'Туп':
                continue
            if 'Пек' in self.finish_sections[index]:
                dict_V[parcel_name[index][1]] = self.V_bakeries_hour
            elif 'ПП' in self.finish_sections[index]:
                dict_V[parcel_name[index][1]] = self.Vpp[int(self.finish_sections[index][-1])]
            elif 'Прач' in self.finish_sections[index]:
                dict_V[parcel_name[index][1]] = self.V_laundromats_hour
            elif 'ТЭЦ' in self.finish_sections[index]:
                dict_V[parcel_name[index][1]] = self.V_big_boiler_houses
            elif 'ПРГ' in self.finish_sections[index]:
                dict_V[parcel_name[index][1]] = self.dict_V_prg[(self.finish_sections[index][-1])]
        site_category = []
        for category, index in list_category:
            if type(category) != int:
                continue
            site_category.append(category)
        site_category.sort()
        site_category = set(site_category)
        for all_category in site_category:
            if all_category == 0:
                continue
            for category, index in list_category:
                if category == all_category:
                    required_expense = 0
                    sources = 0
                    for i, bs in enumerate(beginnings_sections):
                        if str(finish_sections[index]) == str(bs):
                            required_expense += dict_V[parcel_name[i][1]]
                    for fs in finish_sections:
                        if str(finish_sections[index]) == str(fs):
                            sources += 1
                    transit_one = required_expense / sources
                    dict_V[parcel_name[index][1]] = transit_one
        site_category = list(site_category)
        site_category.sort(reverse=True)

        # Расчёт давления для аварий
        dict_length = {}
        dict_reduced_length = {}
        dict_Re = {}
        dict_Re2 = {}
        dict_Dn_delta = {}
        dict_finish_dvn = {}
        dict_lambda = {}
        dict_Pn_node = {}
        dict_Pk_node = {}
        dict_dP = {}
        for category in enumerate(site_category):
            for cat, index in list_category:
                if category[1] != cat:
                    continue
                for i, dv in enumerate(self.Dvn):
                    if category[0] == 0:
                        dict_Pn_node[parcel_name[index][1]] = self.pressure_after_grs
                    else:
                        index_Pn = self.find_Pn(index, beginnings_sections, finish_sections)
                        dict_Pn_node[parcel_name[index][1]] = dict_Pk_node[parcel_name[index_Pn][1]]
                    re = 0.0354 * dict_V[parcel_name[index][1]] / (dv * self.vis)
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
                    Pk = ((dict_Pn_node[parcel_name[index][1]] + 0.1013) ** 2 - 1.2687 * 10 ** (-4) * lambda_ * dict_V[
                        parcel_name[index][1]] ** 2 / dv ** 5 * self.density_gas *
                          self.reduced_length[index])
                    if Pk < 0:
                        continue
                    Pk = Pk ** 0.5
                    Pk -= 0.1013
                    dP = dict_Pn_node[parcel_name[index][1]] - Pk
                    dict_length[parcel_name[index][1]] = self.dict_length[parcel_name[index][0]]
                    dict_reduced_length[parcel_name[index][1]] = self.dict_reduced_length[parcel_name[index][0]]
                    dict_Dn_delta[parcel_name[index][1]] = self.Dn_delta[i]
                    dict_finish_dvn[parcel_name[index][1]] = dv
                    dict_Re[parcel_name[index][1]] = re
                    dict_Re2[parcel_name[index][1]] = re_2
                    dict_lambda[parcel_name[index][1]] = lambda_
                    dict_Pk_node[parcel_name[index][1]] = Pk
                    dict_dP[parcel_name[index][1]] = dP
                    if self.classes_site[index] == 'Ист' and dv == d:
                        break
                    if self.classes_site[index] != 'Туп' and self.classes_site[index] != 'Ист' and dv == d:
                        break
                    if 'ПП1' in parcel_name[index][0] and Pk >= self.allowable_pressure_PP1:
                        break
                    if 'ПП2' in parcel_name[index][0] and Pk >= self.allowable_pressure_PP2:
                        break
                    if Pk >= 0.45 and self.classes_site[index] == 'Туп' and 'ПП1' not in parcel_name[index][0] and \
                            'ПП2' not in parcel_name[index][0]:
                        break
        return dict_length, dict_reduced_length, dict_V, dict_Re, dict_Re2, dict_Dn_delta, dict_finish_dvn, dict_lambda, \
               dict_Pn_node, dict_Pk_node, dict_dP, name_out_node

    def recalculate_dv_crash(self, Pn, V, rl, name, class_):
        for index, dv in enumerate(self.Dvn):
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
            Pk = ((Pn + 0.1013) ** 2 - 1.2687 * 10 ** (-4) * lambda_ * V ** 2 / dv ** 5 * self.density_gas * rl)
            if Pk < 0:
                continue
            Pk = Pk ** 0.5
            Pk -= 0.1013
            dn = self.Dn_delta[index]
            if 'ПП1' in name and Pk >= self.allowable_pressure_PP1:
                return dv, dn, re, re_2, lambda_, Pk
            if 'ПП2' in name and Pk >= self.allowable_pressure_PP2:
                return dv, dn, re, re_2, lambda_, Pk
            if Pk >= 0.41 and class_ == 'Туп' and 'ПП1' not in name and \
                    'ПП2' not in name:
                return dv, dn, re, re_2, lambda_, Pk
        else:
            return [None] * 6

    def find_Pn(self, index, beginnings_sections, finish_sections):
        for i, fn in enumerate(finish_sections):
            if fn == beginnings_sections[index]:
                return i

    def create_classes_site(self, dir):
        list_ist = []
        list_po = []
        list_pr = []
        list_tup = []
        for index, direction in enumerate(self.classes_site):
            if direction == 'Ист':
                list_ist.append([self.list_site_category[index], index])
            elif direction == 'ПО':
                list_po.append([self.list_site_category[index], index])
            elif direction == 'ПР':
                list_pr.append([self.list_site_category[index], index])
            else:
                list_tup.append([self.list_site_category[index], index])
        if dir == 'ПО':
            list_pr.sort(reverse=True)
            list_po.sort()
            list_all = list_pr + list_po + list_ist
        else:
            list_pr.sort()
            list_po.sort(reverse=True)
            list_all = list_po + list_pr + list_ist
        list_category = []
        for index, category in enumerate(list_all):
            if index == 0:
                category[0] = '-'
                name_out_node = self.parcel_name[category[1]]
                list_category.append(category)
                continue
            category[0] = index
            list_category.append(category)
        list_category += list_tup
        parcel_name = []
        for index, name in enumerate(self.parcel_name):
            if dir == 'ПО' and self.classes_site[index] == 'ПР':
                new_name = name[::-1]
            elif dir == 'ПР' and self.classes_site[index] == 'ПО':
                new_name = name[::-1]
            else:
                new_name = name
            parcel_name.append([name, new_name])
        beginnings_sections = []
        finish_sections = []
        for index, name in enumerate(parcel_name):
            if index == list_category[0][1]:
                beginnings_sections.append('a')
                finish_sections.append('a')
                continue
            if dir == 'ПО' and self.classes_site[index] == 'ПР':
                s = name[1].split('-')
                beginnings_sections.append(s[0])
                finish_sections.append(s[1])
            elif dir == 'ПР' and self.classes_site[index] == 'ПО':
                s = name[1].split('-')
                beginnings_sections.append(s[0])
                finish_sections.append(s[1])
            else:
                s = name[0].split('-')
                beginnings_sections.append(s[0])
                finish_sections.append(s[1])
        return list_category, parcel_name, beginnings_sections, finish_sections, name_out_node

    def output_excel(self):
        wb = openpyxl.load_workbook(filename=self.file_out)
        sheet = wb['СВД']
        names = ['№ участков', 'Lдейств,м', 'Lпривед,км', 'Qрасч м3/ч',
                 'Dн×δ', 'Dвн', 'Re', 'Re*(n/d)', 'λ',
                 'Pн', 'Рк', 'DP', 'H']
        letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
        sheet['D1'] = 'Стационарный режим'
        for index, name in enumerate(names):
            sheet[f'{letters[index]}2'] = name
        i = 3
        class_set = list(set(self.classes_site))
        class_set.sort()
        for class_original in class_set:
            for index, class_ in enumerate(self.classes_site):
                for key, value in self.dict_length.items():
                    if self.parcel_name[index] == key and class_ == class_original:
                        sheet[f'A{i}'] = key
                        sheet[f'B{i}'] = value
                        sheet[f'C{i}'] = self.dict_reduced_length[key]
                        sheet[f'D{i}'] = self.dict_V[key]
                        sheet[f'E{i}'] = self.dict_Dn_delta[key]
                        sheet[f'F{i}'] = self.dict_finish_dvn[key]
                        sheet[f'G{i}'] = self.dict_Re[key]
                        sheet[f'H{i}'] = self.dict_Re2[key]
                        sheet[f'I{i}'] = self.dict_lambda[key]
                        sheet[f'J{i}'] = self.dict_Pn_node[key]
                        sheet[f'K{i}'] = self.dict_Pk_node[key]
                        sheet[f'L{i}'] = self.dict_dP[key]
                        i += 1
        sheet[f'M3'] = self.N
        letters = ['O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
        sheet['S1'] = f'Авария на {self.name_out_node_po}'
        for index, letter in enumerate(letters):
            sheet[f'{letter}2'] = names[index]
        i = 3
        for key, value in self.dict_length_po.items():
            sheet[f'O{i}'] = key
            sheet[f'P{i}'] = value
            sheet[f'Q{i}'] = self.dict_reduced_lengt_po[key]
            sheet[f'R{i}'] = self.dict_V_po[key]
            sheet[f'S{i}'] = self.dict_Dn_delta_po[key]
            sheet[f'T{i}'] = self.dict_finish_dvn_po[key]
            sheet[f'U{i}'] = self.dict_Re_crash_po[key]
            sheet[f'V{i}'] = self.dict_Re2_crash_po[key]
            sheet[f'W{i}'] = self.dict_lambda_crash_po[key]
            sheet[f'X{i}'] = self.dict_Pn_node_crash_po[key]
            sheet[f'Y{i}'] = self.dict_Pk_node_crash_po[key]
            sheet[f'Z{i}'] = self.dict_dP_crash_po[key]
            i += 1
        letters = ['AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM']
        sheet['AF1'] = f'Авария на {self.name_out_node_pr}'
        for index, letter in enumerate(letters):
            sheet[f'{letter}2'] = names[index]
        i = 3
        for key, value in self.dict_length_pr.items():
            sheet[f'AB{i}'] = key
            sheet[f'AC{i}'] = value
            sheet[f'AD{i}'] = self.dict_reduced_lengt_pr[key]
            sheet[f'AE{i}'] = self.dict_V_pr[key]
            sheet[f'AF{i}'] = self.dict_Dn_delta_pr[key]
            sheet[f'AG{i}'] = self.dict_finish_dvn_pr[key]
            sheet[f'AH{i}'] = self.dict_Re_crash_pr[key]
            sheet[f'AI{i}'] = self.dict_Re2_crash_pr[key]
            sheet[f'AJ{i}'] = self.dict_lambda_crash_pr[key]
            sheet[f'AK{i}'] = self.dict_Pn_node_crash_pr[key]
            sheet[f'AL{i}'] = self.dict_Pk_node_crash_pr[key]
            sheet[f'AM{i}'] = self.dict_dP_crash_pr[key]
            i += 1
        wb.save(self.file_out)

    def run(self):
        self.initializing_file()
        self.initializing_source_data()
        self.calculation_SVD()
        self.output_excel()
        return self.vis


# calc_svd = HydraulicsСalculatorSVD('source_data.xlsx', 'result.xlsx')
# calc_svd.run()
