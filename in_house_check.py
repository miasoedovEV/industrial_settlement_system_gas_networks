# -*- coding: utf-8 -*-
"""
Created on Thu Dec 24 19:24:29 2020

@author: stinc
"""
from hydraulic_calculation_SVD import HydraulicsСalculatorSVD
import numpy as np
import math as mt
import pandas as pd
from pprint import pprint
import openpyxl


class CalculatorHousePipes:

    def __init__(self, file_name, file_out_name):
        self.file_in = file_name
        self.file_out = file_out_name

    def initializing_file(self):
        self.file_source_data = self.file_in
        # Load spreadsheet
        self.xl_source_data = pd.ExcelFile(self.file_source_data)

    def initializing_source_data(self):
        # Блок ввода
        self.df_source_data = self.xl_source_data.parse('Дом')
        self.array_source_data = np.array(pd.DataFrame(self.df_source_data))
        self.df_pipes_data = self.xl_source_data.parse('information_pipes')
        self.array_pipes_data = np.array(pd.DataFrame(self.df_pipes_data))
        self.calc = HydraulicsСalculatorSVD(self.file_in, self.file_out)
        self.vis = self.calc.run()
        self.types_pipes = []
        self.beginnings_sections = []
        self.finish_sections = []
        self.parcel_name = []
        self.lenth_d = []
        self.a = []
        self.number_apartments = []
        for index, beginning in enumerate(self.array_source_data[:, 1]):
            if np.isnan(beginning):
                break
            self.types_pipes.append(self.array_source_data[:, 0][index])
            self.beginnings_sections.append(beginning)
            self.finish_sections.append(self.array_source_data[:, 2][index])
            self.parcel_name.append(self.array_source_data[:, 3][index])
            self.lenth_d.append(self.array_source_data[:, 4][index])
            self.a.append(self.array_source_data[:, 5][index])
            self.number_apartments.append(self.array_source_data[:, 6][index])
        self.reference_point = list(self.array_source_data[:, 7])[1::]
        self.dict_heights_floor = {}
        self.dict_heights_land = {}
        for index, numbers_knot in enumerate(list(self.array_source_data[:, 8])[1::]):
            if 'пол' in self.reference_point[index]:
                self.dict_heights_floor[numbers_knot] = list(self.array_source_data[:, 9])[1::][index]
            else:
                self.dict_heights_land[numbers_knot] = list(self.array_source_data[:, 9])[1::][index]
        self.delta_P = list(self.array_source_data[:, 12])[2]
        self.V_stove = list(self.array_source_data[:, 16])[0]
        self.V_water_heater = list(self.array_source_data[:, 16])[1]
        self.delta_y = list(self.array_source_data[:, 18])[2]
        self.y_g = list(self.array_source_data[:, 18])[1]
        self.quantity_aparnaments = []
        self.k = []
        for index, quantity in enumerate(self.array_source_data[:, 19]):
            if np.isnan(quantity):
                break
            self.quantity_aparnaments.append(quantity)
            self.k.append(self.array_source_data[:, 20][index])
        self.Dn_delta = list(self.array_pipes_data[:, 0])
        self.Dn = list(self.array_pipes_data[:, 1])
        self.delta = list(self.array_pipes_data[:, 2])
        self.Dvn = list(self.array_pipes_data[:, 3])

    def find_k(self, na):
        for i, qa in enumerate(self.quantity_aparnaments):
            if na == qa:
                k = self.k[i]
                return k
            elif na < qa:
                ns = [self.quantity_aparnaments[i - 1], qa]
                ks = [self.k[i - 1], self.k[i]]
                k = np.interp(na, ns, ks)
                return k

    def calculate(self):
        self.lenth_r = [lenth * self.a[index] for index, lenth in enumerate(self.lenth_d)]
        self.V = []
        self.K = []
        self.V_house = 0
        sum_V = self.V_water_heater + self.V_stove
        for index, typ in enumerate(self.types_pipes):
            if 'гвн' in typ:
                self.V.append(self.V_water_heater)
                k = self.find_k(self.number_apartments[index])
                self.K.append(k)
                continue
            elif 'пл' in typ:
                self.V.append(self.V_stove)
                k = self.find_k(self.number_apartments[index])
                self.K.append(k)
                continue
            k = self.find_k(self.number_apartments[index])
            V = self.number_apartments[index] * k * sum_V
            self.K.append(k)
            self.V.append(V)
            if 'пода' in typ:
                self.V_house = V
        feed_index = 0
        for index, typ in enumerate(self.types_pipes):
            if 'под' in typ:
                feed_index = index
        self.dict_branches = self.get_branches(feed_index=feed_index)
        for key in self.dict_branches.keys():
            self.dict_branches[key].insert(0, [feed_index, self.finish_sections[feed_index]])
        self.dict_Dn_delta = {}
        self.dict_finish_dvn = {}
        self.dict_Re = {}
        self.dict_Re2 = {}
        self.dict_lambda = {}
        self.dict_dh = {}
        self.dict_Hg = {}
        self.dict_dP = {}
        self.dict_dP_h = {}
        self.dves_room = self.Dvn[-8::]
        self.dves_room.reverse()
        self.dnes_room = self.Dn_delta[-8::]
        self.dnes_room.reverse()
        self.Dvn.reverse()
        self.Dn_delta.reverse()
        self.calculate_smaller_diametre()
        self.dict_sum_dp_h = {}
        logic_value = False
        list_good_branches = []
        for i, dv_room in enumerate(self.dves_room):
            for i_dv, dv in enumerate(self.Dvn):
                for item, branch in self.dict_branches.items():
                    if item in list_good_branches:
                        continue
                    list_sum_dP_h, list_good_branches = self.choose_diametre(dv_room, branch, i, i_dv, dv,
                                                                             list_good_branches)
                    if list_sum_dP_h is not None:
                        logic_value = True
                        break
                if logic_value is True:
                    break
            if logic_value is True:
                break
        self.i_d_house = 0
        for index, typ in enumerate(self.types_pipes):
            if 'подач' in typ:
                for i, dn_delta in enumerate(self.Dn_delta):
                    if dn_delta == self.dict_Dn_delta[self.parcel_name[index]]:
                        self.i_d_house = i

    def calculate_smaller_diametre(self):
        for i_dv, dv in enumerate(self.Dvn):
            for branch in self.dict_branches.values():
                for index, fs in branch:
                    if 'внут' in self.types_pipes[index]:
                        self.calculate_dP_h(self.dves_room[0], index, 0, self.dict_heights_floor)
                    elif 'стояк' in self.types_pipes[index]:
                        self.calculate_dP_h(self.dves_room[1], index, 1, self.dict_heights_land)
                    elif 'развод' in self.types_pipes[index]:
                        self.calculate_dP_h(dv, index, i_dv, self.dict_heights_land)
                    else:
                        self.calculate_dP_h(dv, index, i_dv, self.dict_heights_land)

    def choose_diametre(self, dv_room, branch, i, i_dv, dv, list_good_branches):
        for index, fs in branch:
            if 'внут' in self.types_pipes[index]:
                self.calculate_dP_h(dv_room, index, i, self.dict_heights_floor)
            elif 'стояк' in self.types_pipes[index]:
                self.calculate_dP_h(self.dves_room[i + 1], index, i + 1, self.dict_heights_land)
            elif 'развод' in self.types_pipes[index]:
                self.calculate_dP_h(dv, index, i_dv, self.dict_heights_land)
            else:
                self.calculate_dP_h(dv, index, i_dv, self.dict_heights_land)
            list_sum_dP_h, list_good_branches = self.calculate_sum_dP_h(list_good_branches)
            if list_sum_dP_h is not None:
                return list_sum_dP_h, list_good_branches
        else:
            return None, list_good_branches

    def calculate_sum_dP_h(self, list_good_branches):
        list_sum_dP_h = []
        sum_dP_h = 0
        for item, branch in self.dict_branches.items():
            if item in list_good_branches:
                continue
            for index, fs in branch:
                for key, dP_h in self.dict_dP_h.items():
                    if self.parcel_name[index] == key:
                        sum_dP_h += dP_h
            if sum_dP_h < self.delta_P:
                list_good_branches.append(item)
                self.dict_sum_dp_h[item] = sum_dP_h
                list_sum_dP_h.append(sum_dP_h)
            sum_dP_h = 0
        if len(list_good_branches) == len(self.dict_branches.keys()):
            return list_sum_dP_h, list_good_branches
        else:
            return None, list_good_branches

    def get_branches(self, feed_index):
        finish_sections = self.finish_sections[feed_index]
        dict_branches_value = {}
        dict_branches = {}
        list_repits = []
        for ind, typ in enumerate(self.types_pipes):
            for index, bg in enumerate(self.beginnings_sections):
                number_branches, number_finish_section, list_repits, last_nodes = self.check_node(dict_branches, bg,
                                                                                                  finish_sections, ind,
                                                                                                  list_repits,
                                                                                                  index,
                                                                                                  dict_branches_value)

                if number_branches is None:
                    continue
                if number_branches not in dict_branches_value:
                    dict_branches_value[number_branches] = number_finish_section
                if dict_branches_value[number_branches] == bg:
                    if last_nodes is not None:
                        dict_branches[number_branches] = last_nodes
                        dict_branches[number_branches].append([index, self.finish_sections[index]])
                    elif number_branches not in dict_branches:
                        dict_branches[number_branches] = []
                        dict_branches[number_branches].append([index, self.finish_sections[index]])
                    else:
                        dict_branches[number_branches].append([index, self.finish_sections[index]])
            for key, value in dict_branches.items():
                dict_branches_value[key] = value[-1][-1]
            list_repits = []
        return dict_branches

    def check_node(self, dict_branches, bg, finish_sections, ind, list_repits, index, dict_branches_value):
        returned_key = 0
        itr = 0
        itr_index = -1
        for rep in list_repits:
            if rep == bg:
                itr += 1
                itr_index -= 1
        for key, list_node in dict_branches.items():
            returned_key = key
            if bg == list_node[itr_index][-1] and dict_branches_value[key] == bg:
                if bg not in list_repits or 'внут' in self.types_pipes[index]:
                    list_repits.append(bg)
                    return key, list_node[-1][-1], list_repits, None
                else:
                    list_repits.append(bg)
                    last_nodes = self.find_branches(dict_branches, index)
                    return key + itr, list_node[itr_index][-1], list_repits, last_nodes
        else:
            if ind == 0 and bg == finish_sections:
                return returned_key + 1, finish_sections, list_repits, None
            else:
                return None, None, list_repits, None

    def find_branches(self, dict_branches, index):
        for key, branch in dict_branches.items():
            for i, fs in branch:
                if self.beginnings_sections[i] == self.beginnings_sections[index]:
                    return branch[:-1]
        else:
            return None

    def calculate_dP_h(self, dv, i, i_d, dict_heights):
        self.re = 0.0354 * self.V[i] / (dv * self.vis)
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
        delta_h = dict_heights[self.beginnings_sections[i]] - \
                  dict_heights[self.finish_sections[i]]
        hydrostatic_head = self.delta_y * delta_h
        dP = 626.1 * lambda_ * self.V[i] ** 2 * self.y_g * self.lenth_r[i] * 1.1 / (dv ** 5)
        dP_h = hydrostatic_head + dP
        self.dict_Dn_delta[self.parcel_name[i]] = self.Dn_delta[i_d]
        self.dict_finish_dvn[self.parcel_name[i]] = dv
        self.dict_Re[self.parcel_name[i]] = self.re
        self.dict_Re2[self.parcel_name[i]] = self.re_2
        self.dict_lambda[self.parcel_name[i]] = lambda_
        self.dict_dh[self.parcel_name[i]] = delta_h
        self.dict_Hg[self.parcel_name[i]] = hydrostatic_head
        self.dict_dP[self.parcel_name[i]] = dP
        self.dict_dP_h[self.parcel_name[i]] = dP_h

    def create_arrive(self):
        names = ['Тип', 'Имя', 'Lд', 'а, доли',
                 'Lр', 'n', 'K, доли', 'V', 'Dн×δ', 'Dвн', 'Re', 'Re*(n/d)', 'λ', 'delta_h',
                 '+/-Нг', 'dP', 'dP_h']
        letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
        ar = {}
        for index, name in enumerate(names):
            if name == 'Тип':
                ar[name] = [letters[index], self.types_pipes]
            elif name == 'Имя':
                ar[name] = [letters[index], self.parcel_name]
            elif name == 'Lд':
                ar[name] = [letters[index], self.lenth_d]
            elif name == 'а, доли':
                ar[name] = [letters[index], self.a]
            elif name == 'Lр':
                ar[name] = [letters[index], self.lenth_r]
            elif name == 'n':
                ar[name] = [letters[index], self.number_apartments]
            elif name == 'K, доли':
                ar[name] = [letters[index], self.K]
            elif name == 'V':
                ar[name] = [letters[index], self.V]
            else:
                ar[name] = [letters[index], []]
        for index, name in enumerate(self.parcel_name):
            ar['Dн×δ'][1].append(self.dict_Dn_delta[name])
            ar['Dвн'][1].append(self.dict_finish_dvn[name])
            ar['Re'][1].append(self.dict_Re[name])
            ar['Re*(n/d)'][1].append(self.dict_Re2[name])
            ar['λ'][1].append(self.dict_lambda[name])
            ar['delta_h'][1].append(self.dict_dh[name])
            ar['+/-Нг'][1].append(self.dict_Hg[name])
            ar['dP'][1].append(self.dict_dP[name])
            ar['dP_h'][1].append(self.dict_dP_h[name])
        return ar

    def output_excel(self):
        wb = openpyxl.load_workbook(filename=self.file_out)
        sheet = wb['Дом']
        ar = self.create_arrive()
        for key, value in ar.items():
            sheet[f'{value[0]}1'] = key
            iter = 2
            for num in value[1]:
                sheet[f'{value[0]}{iter}'] = num
                iter += 1
        letter = ['R', 'S']
        names = ['Номер ветви', 'Сумма перепадов давлений']
        for index, name in enumerate(names):
            sheet[f'{letter[index]}1'] = name
        iter = 2
        for key, value in self.dict_sum_dp_h.items():
            sheet[f'R{iter}'] = key
            sheet[f'S{iter}'] = value
            iter += 1
        wb.save(self.file_out)

    def run(self):
        self.initializing_file()
        self.initializing_source_data()
        self.calculate()
        self.output_excel()
        return self.vis, self.V_house, self.i_d_house, self.y_g, self.delta_y

# calc = CalculatorHousePipes('source_data.xlsx', 'result.xlsx')
# calc.run()
