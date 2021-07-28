# -*- coding: utf-8 -*-
"""
Created on Wed Nov  4 13:52:53 2020

@author: stinc
"""
from gas_network_low_pressure import GasExpenseCalculator
import numpy as np
import math as mt
import pandas as pd
from pprint import pprint
import openpyxl


def find_nodes(data, direction):
    for node_information in data:
        if node_information[0] == direction:
            yield node_information[1]


class HydraulicsСalculatorSND:
    def __init__(self, file_name, file_out_name=None):
        self.file_in = file_name
        self.file_out = file_out_name

    def initializing_file(self):
        self.file_source_data = self.file_in
        # Load spreadsheet
        self.xl_source_data = pd.ExcelFile(self.file_source_data)

    def initializing_source_data(self):
        # Блок ввода
        self.df_source_data = self.xl_source_data.parse('СНД')
        self.array_source_data = np.array(pd.DataFrame(self.df_source_data))  # Масив с исходными данными
        self.df_gas_data = self.xl_source_data.parse('gas_composition')
        self.array_gas_data = np.array(pd.DataFrame(self.df_gas_data))  # Масив с исходными данными газа
        self.df_pipes_data = self.xl_source_data.parse('information_pipes')
        self.array_pipes_data = np.array(pd.DataFrame(self.df_pipes_data))
        self.parcel_name = []
        for value in self.array_source_data[:, 5]:
            if type(value) == str:
                self.parcel_name.append(value)
            elif np.isnan(value):
                continue
            else:
                self.parcel_name.append(value)
        print(self.parcel_name)
        self.lenth_d = []
        for value in self.array_source_data[:, 6]:
            if np.isnan(value):
                continue
            else:
                self.lenth_d.append(value)
        print(self.lenth_d)
        self.k = []
        for value in self.array_source_data[:, 7]:
            if np.isnan(value):
                continue
            else:
                self.k.append(value)
        print(self.k)
        self.directions = []
        for value in self.array_source_data[:, 1]:
            if value == 'по' or value == 'пр' or value == 'т':
                self.directions.append(value)
            elif np.isnan(value):
                continue
            else:
                self.directions.append(value)
        print(self.directions)
        self.site_category = []
        for value in self.array_source_data[:, 2]:
            if np.isnan(value):
                continue
            else:
                self.site_category.append(value)
        self.beginnings_sections = []
        for value in self.array_source_data[:, 3]:
            if np.isnan(value):
                continue
            else:
                self.beginnings_sections.append(value)
        self.finish_sections = []
        for value in self.array_source_data[:, 4]:
            if np.isnan(value):
                continue
            else:
                self.finish_sections.append(value)
        self.number_prgs = int(self.array_source_data[:, 10][0])
        self.prg_node_numbers = self.array_source_data[:, 11][:self.number_prgs]
        self.coefficient_increase_length = self.array_source_data[:, 12][0]
        self.pressure_drop = self.array_source_data[:, 13][0]
        self.a_all_gas = self.array_gas_data[:, 1]
        self.density_st = self.array_gas_data[:, 2]
        self.density_n = self.array_gas_data[:, 3]
        self.Dn_delta = self.array_pipes_data[:, 0]
        self.Dn = self.array_pipes_data[:, 1]
        self.delta = self.array_pipes_data[:, 2]
        self.Dvn = self.array_pipes_data[:, 3]
        self.ring_numbers = []
        for value in self.array_source_data[:, 0]:
            if value == 'т':
                self.ring_numbers.append(value)
            elif np.isnan(value):
                continue
            else:
                self.ring_numbers.append(value)
        self.set_ring_numbers = set(self.ring_numbers)
        self.dict_heights_dictionary = dict()
        for index, node in enumerate(self.array_source_data[:, 8]):
            if np.isnan(node):
                break
            self.dict_heights_dictionary[int(node)] = self.array_source_data[:, 9][index]
        self.y_g = self.array_source_data[:, 14][0]
        self.density_air = self.array_source_data[:, 15][0]
        self.cal = GasExpenseCalculator(self.file_in, self.file_out)
        self.V = self.cal.run()

    def calculation_SND(self):
        self.given_length = [ld * self.k[i] for i, ld in enumerate(self.lenth_d)]
        print(f'Приведдёные длины участков {self.given_length}')
        list_for_sort = []
        self.sum_given_length = 0
        for i, gl in enumerate(self.given_length):
            if self.parcel_name[i] in list_for_sort:
                continue
            list_for_sort.append(self.parcel_name[i])
            self.sum_given_length += gl
        print(len(list_for_sort), len(self.lenth_d))
        print(f'Сумма приведенных длин участков = {self.sum_given_length}')
        self.specific_travel_expense = self.V[0] / self.sum_given_length
        print(f'Удельный путевой расход = {self.specific_travel_expense}')
        self.travelV = [self.specific_travel_expense * gl for gl in self.given_length]
        print(f'Путевой расход = {self.travelV}')
        self.dict_travelV = {}
        list_for_sort = []
        for index, tv in enumerate(self.travelV):
            if self.parcel_name[index] in list_for_sort:
                continue
            list_for_sort.append(self.parcel_name[index])
            self.dict_travelV[self.parcel_name[index]] = tv

        # Определение транзитныйх расходов
        self.set_site_category = set(self.site_category)
        self.dict_transit_expense = {}
        for all_category in self.set_site_category:
            for index, category in enumerate(self.site_category):
                if category == all_category and all_category == 0:
                    self.dict_transit_expense[index] = 0
                elif category == all_category:
                    required_expense = 0
                    sources = 0
                    list_for_sort_required_expense = []
                    list_for_sort_sources = []
                    for i, bs in enumerate(self.beginnings_sections):
                        if self.parcel_name[i] in list_for_sort_required_expense:
                            continue
                        if self.finish_sections[index] == bs:
                            required_expense += self.travelV[i] + self.dict_transit_expense[i]
                        list_for_sort_required_expense.append(self.parcel_name[i])
                    for i, fs in enumerate(self.finish_sections):
                        if self.parcel_name[i] in list_for_sort_sources:
                            continue
                        if self.finish_sections[index] == fs:
                            sources += 1
                        list_for_sort_sources.append(self.parcel_name[i])
                    transit_one = required_expense / sources
                    self.dict_transit_expense[index] = transit_one
        self.transit_expense = []
        for key, value in self.dict_transit_expense.items():
            self.transit_expense.append([key, value])
        self.transit_expense.sort()
        print(f'Транзитные расходы = {self.transit_expense}')
        self.dict_transit_expense = {}
        list_for_sort = []
        for te in self.transit_expense:
            if self.parcel_name[te[0]] in list_for_sort:
                continue
            list_for_sort.append(self.parcel_name[te[0]])
            self.dict_transit_expense[self.parcel_name[te[0]]] = te[1]

        # Проверка на правильность расчёта транзитных расходов
        list_for_sort = []
        self.sum_travelV = 0
        for i, tv in enumerate(self.travelV):
            if self.parcel_name[i] in list_for_sort:
                continue
            list_for_sort.append(self.parcel_name[i])
            self.sum_travelV += tv
        self.required_expense = 0
        for prg_node_number in self.prg_node_numbers:
            list_for_sort = []
            for index, tv in enumerate(self.travelV):
                if self.parcel_name[index] in list_for_sort:
                    continue
                list_for_sort.append(self.parcel_name[index])
                if self.beginnings_sections[index] == prg_node_number:
                    self.required_expense += tv + self.transit_expense[index][1]
        if round(self.sum_travelV, 8) != round(self.required_expense, 8):
            raise ValueError('Не сходятся транзиты!!! Проверьте правильность внесённых данных!!!: ValueError')

        # Расчётный расход
        self.estimated_expenses = []
        for index, tv in enumerate(self.travelV):
            estimated_expense = self.transit_expense[index][1] + 0.5 * tv
            self.estimated_expenses.append(estimated_expense)
        print(f'Расчётные расходы = {self.estimated_expenses}')
        list_for_sort = []
        self.dict_estimated_expenses = {}
        for index, ee in enumerate(self.estimated_expenses):
            if self.parcel_name[index] in list_for_sort:
                continue
            list_for_sort.append(self.parcel_name[index])
            self.dict_estimated_expenses[self.parcel_name[index]] = ee

        # Расчётные длины участков
        self.calculated_lengths_sections = []
        for ld in self.lenth_d:
            calculated_length_sections = ld * self.coefficient_increase_length
            self.calculated_lengths_sections.append(calculated_length_sections)
        print(f'Расчётные длины участков = {self.calculated_lengths_sections}')
        # Ориентировочные удельные потери по длине участка газопровода
        self.approximate_specific_losses = []
        for lp in self.calculated_lengths_sections:
            r_ud = self.pressure_drop / lp
            self.approximate_specific_losses.append(r_ud)
        print(f'Ориентировочные удельные потери по длине участка газопровода = {self.approximate_specific_losses}')

        # Определение диаметров труб
        # Расчёт вязкости
        self.a_density_st = []
        for index, den in enumerate(self.density_st):
            a_den_st = den * self.a_all_gas[index]
            self.a_density_st.append(a_den_st)
        print(f'а*ρcт = {self.a_density_st}')

        self.a_density_n = []
        for index, den in enumerate(self.density_n):
            print(den, self.a_all_gas[index])
            a_den_n = den * self.a_all_gas[index]
            self.a_density_n.append(a_den_n)
        print(f'а*ρн = {self.a_density_n}')

        self.mixture_density_st = sum(self.a_density_st)
        self.mixture_density_n = sum(self.a_density_n)
        print(f'Плотность смеси ст = {self.mixture_density_st}')
        print(f'Плотность смеси н = {self.mixture_density_n}')

        self.p_pk = 0.1737 * (26.831 - self.mixture_density_st)
        print(f'Pпк  = {self.p_pk}')

        self.t_pk = 155.24 * (0.564 + self.mixture_density_st)
        print(f'Tпк  = {self.t_pk}')

        self.p_pr = 0.101325 / self.p_pk
        print(f'Pпр  = {self.p_pr}')

        self.t_pr = 273 / self.t_pk
        print(f'Tпр  = {self.t_pr}')

        self.my_gas = 5.1 * 10**(-6) * (1 + self.mixture_density_n * (1.1 - 1.25 * self.mixture_density_n)) * \
                      (0.037 + self.t_pr * (1 - 0.104 * self.t_pr)) * (1 + (self.p_pr**2 / (30 * (self.t_pr - 1))))
        print(f'μ  = {self.my_gas}')

        self.vis = self.my_gas / self.mixture_density_n
        print(f'ν = {self.vis}')
        print(self.dict_heights_dictionary)
        list_data_node_from_rings = []
        for ring_number in self.set_ring_numbers:
            list_node = []
            for index, ring in enumerate(self.ring_numbers):
                if ring_number != ring:
                    continue
                list_node.append([self.directions[index], index])
            list_data_node_from_rings.append(list_node)
        self.dict_dP_h = dict()
        self.dict_Re = {}
        self.dict_Re2 = {}
        self.dict_lambda = {}
        self.dict_finish_dvn = {}
        self.dict_dh = {}
        self.dict_Hg = {}
        self.dict_dP = {}
        self.dict_lambda = {}
        self.dict_N = {}
        self.dict_Dn_delta = {}
        list_for_sort = []
        self.Dn_delta = list(self.Dn_delta)
        self.Dn_delta.reverse()
        self.Dvn = list(self.Dvn)
        self.Dvn.sort()
        for number, data in enumerate(list_data_node_from_rings):
            dict_dP_h_node_po = {}
            dict_dP_h_node_pr = {}
            dict_dP_h_node_t = {}
            for i, dv in enumerate(self.Dvn):
                list_dP_h_node_all = []
                for node in data:
                    if self.parcel_name[node[1]] in self.dict_dP_h:
                        list_dP_h_node_all.append(self.dict_dP_h[self.parcel_name[node[1]]])
                        if node[0] == 'по':
                            dict_dP_h_node_po[self.parcel_name[node[1]]] = self.dict_dP_h[self.parcel_name[node[1]]]
                        else:
                            dict_dP_h_node_pr[self.parcel_name[node[1]]] = self.dict_dP_h[self.parcel_name[node[1]]]
                        continue
                    if self.parcel_name[node[1]] in list_for_sort:
                        if node[0] == 'по':
                            list_dP_h_node_all.append(dict_dP_h_node_po[self.parcel_name[node[1]]])
                        else:
                            list_dP_h_node_all.append(dict_dP_h_node_pr[self.parcel_name[node[1]]])
                        continue
                    dP_h = self.calculate_dP_h(dv, node, i)
                    index = i
                    while 0 > dP_h:
                        index -= 1
                        dP_h = self.calculate_dP_h(self.Dvn[index], node, index)
                    while dP_h > self.pressure_drop:
                        index += 1
                        dP_h = self.calculate_dP_h(self.Dvn[index], node, index)
                    list_dP_h_node_all.append(dP_h)
                    if node[0] == 'по':
                        dict_dP_h_node_po[self.parcel_name[node[1]]] = dP_h
                    elif node[0] == 'пр':
                        dict_dP_h_node_pr[self.parcel_name[node[1]]] = dP_h
                    elif node[0] == 'т':
                        dict_dP_h_node_t[self.parcel_name[node[1]]] = dP_h
                    else:
                        raise ValueError(
                            'Вы неверно обозначили направление, введите либо "по" либо "пр" либо "т": ValueError')
                if sum(dict_dP_h_node_po.values()) > sum(dict_dP_h_node_pr.values()):
                    N = (sum(dict_dP_h_node_po.values()) - sum(dict_dP_h_node_pr.values())) * 100 / sum(
                        dict_dP_h_node_po.values())
                    dict_dP_h_node_po = {}
                    list_for_sort = []
                    for node in find_nodes(data, 'пр'):
                        list_for_sort.append(self.parcel_name[node])
                elif sum(dict_dP_h_node_pr.values()) > sum(dict_dP_h_node_po.values()):
                    N = (sum(dict_dP_h_node_pr.values()) - sum(dict_dP_h_node_po.values())) * 100 / sum(
                        dict_dP_h_node_pr.values())
                    dict_dP_h_node_pr = {}
                    list_for_sort = []
                    for node in find_nodes(data, 'по'):
                        list_for_sort.append(self.parcel_name[node])
                else:
                    N = 0
                if dict_dP_h_node_t:
                    for key, dP_h_t in dict_dP_h_node_t.items():
                        self.dict_dP_h[key] = dP_h_t
                    break
                elif N < 10:
                    list_for_sort = []
                    for num, node in enumerate(data):
                        self.dict_dP_h[f'{self.parcel_name[node[1]]}'] = list_dP_h_node_all[num]
                    self.dict_N[number + 1] = N
                    break
        print(f'Наружний диаметр для участков {self.dict_Dn_delta}')
        print(f'Внутренние диметры трубопроводов для участков {self.dict_finish_dvn}')
        print(f'Число Рейнольдса для участков {self.dict_Re}')
        print(f'Лямбда для участков {self.dict_lambda}')
        print(f'Разница высот между узлами на участках {self.dict_dh}')
        print(f'Гидростатическое давление на участках {self.dict_Hg}')
        print(f'Давление на участках {self.dict_dP}')
        print(f'Полный перепад давления на участках {self.dict_dP_h}')
        print(f'Невязка на узлах {self.dict_N}')
        self.dict_V_prg = {}
        i = 1
        for prg_node_number in self.prg_node_numbers:
            list_for_sort = []
            required_expense = 0
            for index, tv in enumerate(self.travelV):
                if self.parcel_name[index] in list_for_sort:
                    continue
                list_for_sort.append(self.parcel_name[index])
                if self.beginnings_sections[index] == prg_node_number:
                    required_expense += tv + self.transit_expense[index][1]
            self.dict_V_prg[f'{i}'] = required_expense
            i += 1
        self.output_excel()
        return self.dict_V_prg, self.V, self.vis

    def calculate_dP_h(self, dv, node, i):
        self.re = 0.0354 * self.estimated_expenses[node[1]] / (dv * self.vis)
        self.re_2 = self.re * 0.01 / dv
        if self.re == 0 and self.re_2 == 0:
            lambda_ = 0
        elif self.re <= 2000:
            lambda_ = 64 / self.re
        elif self.re <= 4000:
            lambda_ = 0.0025 * self.re ** 0.333
        elif self.re < 100000 and self.re_2 < 23:
            lambda_ = 0.3164 / self.re ** 0.25
        elif self.re > 100000 and self.re_2 < 23:
            lambda_ = 1 / (1.821 * mt.log10(self.re) - 1.64) ** 2
        elif self.re_2 > 23 and self.re > 4000:
            lambda_ = 0.11 * (0.01 / dv + 68 / self.re) ** 0.25
        else:
            raise ValueError('Не найден соответствующий режим транспортировки газа: ValueError')
        delta_h = self.dict_heights_dictionary[self.beginnings_sections[node[1]]] - \
                  self.dict_heights_dictionary[self.finish_sections[node[1]]]
        hydrostatic_head = (self.density_air - self.y_g) * 9.81 * delta_h
        dP = 626.1 * lambda_ * self.estimated_expenses[node[1]] ** 2 * self.y_g * self.lenth_d[
            node[1]] * 1.1 / (dv ** 5)
        dP_h = hydrostatic_head + dP
        self.dict_Dn_delta[self.parcel_name[node[1]]] = self.Dn_delta[i]
        self.dict_finish_dvn[self.parcel_name[node[1]]] = dv
        self.dict_Re[self.parcel_name[node[1]]] = self.re
        self.dict_Re2[self.parcel_name[node[1]]] = self.re_2
        self.dict_lambda[self.parcel_name[node[1]]] = lambda_
        self.dict_dh[self.parcel_name[node[1]]] = delta_h
        self.dict_Hg[self.parcel_name[node[1]]] = hydrostatic_head
        self.dict_dP[self.parcel_name[node[1]]] = dP
        return dP_h

    def output_excel(self):
        wb = openpyxl.load_workbook(filename=self.file_out)
        sheet = wb['СНД']
        names = ['№ участков', 'V_уд. Пут', 'V путев', 'Транзитн расход',
                 'Расчётный расход', 'Dн×δ', 'Dвн', 'Re', 'Re*(n/d)', 'λ',
                 '+/-Нг', 'dP', 'dP_h', 'номер кольца', 'Н', '№ Прг', 'Расход Прг']
        letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
        list_dict = [self.dict_travelV, self.dict_travelV, self.dict_transit_expense,
                     self.dict_estimated_expenses, self.dict_Dn_delta, self.dict_finish_dvn,
                     self.dict_Re, self.dict_Re2, self.dict_lambda, self.dict_Hg,
                     self.dict_dP, self.dict_dP_h]
        for index, name in enumerate(names):
            sheet[f'{letters[index]}1'] = name
        sheet['B2'] = self.specific_travel_expense
        for index, letter in enumerate(letters):
            if index == 0 or index == 1:
                continue
            elif letter == 'N' or letter == 'O':
                continue
            elif letter == 'P' or letter == 'Q':
                continue
            i = 2
            for key, value in list_dict[index - 1].items():
                sheet[f'A{i}'] = key
                sheet[f'{letter}{i}'] = value
                i += 1
        i = 2
        for key, value in self.dict_N.items():
            sheet[f'N{i}'] = key
            sheet[f'O{i}'] = value
            i += 1
        i = 2
        for key, value in self.dict_V_prg.items():
            sheet[f'P{i}'] = key
            sheet[f'Q{i}'] = value
            i += 1
        wb.save(self.file_out)

    def run(self):
        self.initializing_file()
        self.initializing_source_data()
        return self.calculation_SND()


# TODO надо сделать оптимизацию кода
calc = HydraulicsСalculatorSND('source_data.xlsx', 'result.xlsx')
calc.run()
