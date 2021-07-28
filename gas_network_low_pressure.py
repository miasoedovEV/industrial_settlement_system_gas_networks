# -*- coding: utf-8 -*-
"""
Created on Sat Sep 19 19:28:22 2020

@author: stinc
"""
import numpy as np
import math as mt
import pandas as pd
import openpyxl


class GasExpenseCalculator:
    def __init__(self, file_name, file_out_name):
        self.file_in = file_name
        self.file_out = file_out_name

    def initializing_file(self):
        self.file_source_data = self.file_in
        # Load spreadsheet
        self.xl_source_data = pd.ExcelFile(self.file_source_data)

    def initializing_source_data(self):
        # Блок ввода
        self.df_source_data = self.xl_source_data.parse('master data')
        self.array_source_data = np.array(pd.DataFrame(self.df_source_data))  # Масив с исходными данными
        self.df_maximum_coefficients = self.xl_source_data.parse('maximum_coefficients')
        self.array_maximum_coefficients = np.array(
            pd.DataFrame(self.df_maximum_coefficients))  # масив с коэффициентами часового максимума
        self.df_temperature_correction_factor = self.xl_source_data.parse('temperature correction factor')
        self.array_temperature_correction_factor = np.array(pd.DataFrame(
            self.df_temperature_correction_factor))  # Масив с таблицей для поправочного коэффициента температуры
        self.source_data = {}
        for index, value in enumerate(self.array_source_data[:, 4]):
            self.source_data[f'{self.array_source_data[:, 3][index]}'] = value
        number_of_quarters = int(self.source_data['number of quarters'])
        self.S = self.array_source_data[:, 1]
        self.S = list(self.S[0:number_of_quarters])  # площади районов
        self.df_source_data = self.xl_source_data.parse('k_for_enterprises')
        self.array_k_for_enterprises = np.array(
            pd.DataFrame(self.df_source_data))  # Масив с Коэффициенты часового максимума расходов газа для предприятий
        self.k_for_enterprises = {}
        for index, value in enumerate(self.array_k_for_enterprises[:, 1]):
            self.k_for_enterprises[f'{self.array_k_for_enterprises[:, 2][index]}'] = value
        self.list_name_result = []
        self.list_result = []

    def counting_cost(self):
        # Блок расчётов
        # 1.1 Определение численности населения
        self.N = 0
        for s in self.S:
            n = s * self.source_data['population density']
            self.N += n
        self.list_name_result.append(f'Общее число житилей  во всех кварталах = ')
        self.list_result.append(self.N)
        print(f'Общее число житилей  во всех кварталах = {self.N}')
        # 1.2 Определение годового расхода газа на бытовые и коммунально-бытовые нуж-ды населения
        # 1.2.1 Определение годового расхода газа на бытовые нужды населения
        self.V_food_year = (self.N * self.source_data['a1'] * self.source_data['q1']) / self.source_data[
            'Qn']  # годового расхода газа для приготовления пищи
        self.list_name_result.append(f'Годовой расход газа для приготовления пищи =')
        self.list_result.append(self.V_food_year)
        print(f'Годовой расход газа для приготовления пищи = {self.V_food_year}')

        self.V_hot_water_year = (self.N * self.source_data['a2'] * self.source_data['q2']) / self.source_data[
            'Qn']  # годового расхода газа для приготовления горячей воды
        self.list_name_result.append(f'Годовой расход газа для приготовления горячей воды =')
        self.list_result.append(self.V_hot_water_year)
        print(f'Годовой расход газа для приготовления горячей воды {self.V_hot_water_year}')

        # 1.2.2 Определение годового расхода газа на коммунально-бытовое потребление
        # а) Больницы
        self.V_hospitals_year = (12 * self.N * self.source_data['a_hospitals'] / 1000) * (
                (self.source_data['q_food_hospitals'] + self.source_data['q_water_hospitals']) /
                self.source_data['Qn'])
        self.list_name_result.append(f'Годовой расход газа на коммунально-бытовое потребление: Больницы =')
        self.list_result.append(self.V_hospitals_year)
        print(f'Годовой расход газа на коммунально-бытовое потребление: Больницы {self.V_hospitals_year}')

        # б) Столовые и рестораны
        self.V_restaurants_and_cafes = (self.source_data['a_canteens'] * self.N * 360 * self.source_data[
            'q_canteens']) / self.source_data['Qn']
        self.list_name_result.append(f'Годовой расход газа на коммунально-бытовое потребление: Рестораны и столовые:')
        self.list_result.append(self.V_restaurants_and_cafes)
        print(
            f'Годовой расход газа на коммунально-бытовое потребление: Рестораны и столовые {self.V_restaurants_and_cafes}')

        # В) Прачеченые
        self.V_laundromats = (self.source_data['a_laundries'] * self.N * 100 * self.source_data['q_laundries']) / (
                self.source_data['Qn'] * 1000)
        self.list_name_result.append(f'Годовой расход газа на коммунально-бытовое потребление: Прачечные ')
        self.list_result.append(self.V_laundromats)
        print(f'Годовой расход газа на коммунально-бытовое потребление: Прачечные {self.V_laundromats}')

        # Г) Пекарни
        self.V_bakeries = (self.source_data['G'] * self.N * 360 * self.source_data['q_bakeries']) / (
                self.source_data['Qn'] * 1000)
        self.list_name_result.append(f'Годовой расход газа на коммунально-бытовое потребление: Пекарни ')
        self.list_result.append(self.V_bakeries)
        print(f'Годовой расход газа на коммунально-бытовое потребление: Пекарни {self.V_bakeries}')

        # Д) Бани
        self.V_bathhouses = (self.source_data['a_baths'] * self.N * 52 * self.source_data['q_baths']) / \
                            self.source_data['Qn']
        self.list_name_result.append(f'Годовогой расход газа на коммунально-бытовое потребление: Бани ')
        self.list_result.append(self.V_bathhouses)
        print(f'Годовогой расход газа на коммунально-бытовое потребление: Бани {self.V_bathhouses}')
        # Е) Мелкая городская промышленность
        self.V_urban_industry = (
                                        self.V_hospitals_year + self.V_restaurants_and_cafes + self.V_bakeries + self.V_bathhouses + self.V_laundromats) * \
                                self.source_data['gas_utilization_rate_year']
        self.list_name_result.append(
            f'Годовой расход газа на коммунально-бытовое потребление: Мелкая городская промышленность')
        self.list_result.append(self.V_urban_industry)
        print(
            f'Годовой расход газа на коммунально-бытовое потребление: Мелкая городская промышленность {self.V_urban_industry}')
        # 1.3.3 Распределение расхода газа на отопление
        # Определение поправочного коэффициента температуры
        self.quantity_people = self.array_maximum_coefficients[:, 0]
        self.Ks = self.array_maximum_coefficients[:, 1]
        for i, q in enumerate(self.quantity_people):
            if q == mt.ceil(self.N) / 1000:
                self.k = self.Ks[i]
                break
            elif q > mt.ceil(self.N) / 1000:
                ns = [self.quantity_people[i - 1], q]
                ks = [self.Ks[i - 1], self.Ks[i]]
                self.k = np.interp(mt.ceil(self.N) / 1000, ns, ks)
                break
        self.list_name_result.append(f'Коэффициент часового максимума')
        self.list_result.append(self.k)
        print(f'Коэффициент часового максимума {self.k}')

        # Определение часового расхода газа для приготовления пищи
        self.V_food_hour = self.V_food_year * self.k
        self.list_name_result.append(f'Часовой расход газа для приготовления пищи на каждый квартал')
        self.list_result.append(self.V_food_hour)
        print(f'Часовой расход газа для приготовления пищи на каждый квартал {self.V_food_hour}')

        # Определение часового расхода газа для приготовления воды
        self.V_hot_water_hour = self.V_hot_water_year * self.k
        self.list_name_result.append(f'Часовой расход газа для приготовления воды на каждый квартал')
        self.list_result.append(self.V_hot_water_hour)
        print(f'Часовой расход газа для приготовления воды на каждый квартал {self.V_hot_water_hour}')

        # 1.3.2 Определение часового расхода газа на комунально-бытовое потребление

        # а)  Больница
        self.V_hospitals_hour = self.V_hospitals_year * self.k
        self.list_name_result.append(f'Часовой расход газа на коммунально-бытовое потребление: Больницы')
        self.list_result.append(self.V_hospitals_hour)
        print(f'Часовой расход газа на коммунально-бытовое потребление: Больницы {self.V_hospitals_hour}')

        # б)  Столовые и рестораны
        self.V_restaurants_and_cafes_hour = self.V_restaurants_and_cafes * self.k_for_enterprises['canteens']
        self.list_name_result.append(f'Часовой расход газа на коммунально-бытовое потребление: Столовые и рестораны')
        self.list_result.append(self.V_restaurants_and_cafes_hour)
        print(
            f'Часовой расход газа на коммунально-бытовое потребление: Столовые и рестораны {self.V_restaurants_and_cafes_hour}')

        # в) Прачечные
        self.V_laundromats_hour = self.V_laundromats * self.k_for_enterprises['laundries']
        self.list_name_result.append(f'Часовой расход газа на коммунально-бытовое потребление: Прачечные ')
        self.list_result.append(self.V_laundromats_hour)
        print(f'Часовой расход газа на коммунально-бытовое потребление: Прачечные {self.V_laundromats_hour}')

        # г) Бани
        self.V_bathhouses_hour = self.V_bathhouses * self.k_for_enterprises['baths']
        self.list_name_result.append(f'Часовой расход газа на коммунально-бытовое потребление: Бани ')
        self.list_result.append(self.V_bathhouses_hour)
        print(f'Часовой расход газа на коммунально-бытовое потребление: Бани {self.V_bathhouses_hour}')

        # д) Пекарни
        self.V_bakeries_hour = self.V_bakeries * self.k_for_enterprises['bakeries']
        self.list_name_result.append(f'Часовой расход газа на коммунально-бытовое потребление: Пекарни ')
        self.list_result.append(self.V_bakeries_hour)
        print(f'Часовой расход газа на коммунально-бытовое потребление: Пекарни {self.V_bakeries_hour}')

        #  Мелкая городская промышленность
        self.V_urban_industry_hour = (
                                             self.V_hospitals_hour + self.V_restaurants_and_cafes_hour + self.V_bakeries_hour + self.V_bathhouses_hour + self.V_laundromats_hour) * \
                                     self.source_data['gas_utilization_rate_year']
        self.list_name_result.append(
            f'Часовой расход газа на коммунально-бытовое потребление: Мелкая городская промышленность ')
        self.list_result.append(self.V_urban_industry_hour)
        print(
            f'Часовой расход газа на коммунально-бытовое потребление: Мелкая городская промышленность {self.V_urban_industry_hour}')

        # 1.3.3 Распределение расхода газа на отопление
        # Определение поправочного коэффициента температуры
        Tn = self.array_temperature_correction_factor[:, 0]
        A_kafs = self.array_temperature_correction_factor[:, 1]
        for i, t in enumerate(Tn):
            if t == self.source_data['tn']:
                self.A = A_kafs[i]
                break
            elif t < self.source_data['tn']:
                ts = [t, Tn[i - 1]]
                As = [A_kafs[i], A_kafs[i - 1]]
                self.A = np.interp(self.source_data['tn'], ts, As)
                break
        self.list_name_result.append(f'поправочный коэффициент температуры')
        self.list_result.append(self.A)
        print(f'поправочный коэффициент температуры {self.A}')

        self.V_heating_hour = (self.A * self.N * self.source_data['q0'] * (
                self.source_data['tv'] - self.source_data['tn']) * self.source_data['Vn']) / (
                                      self.source_data['kpd'] * self.source_data['Qn'])
        self.list_name_result.append(f'Часовой расход газа на на отопление')
        self.list_result.append(self.V_heating_hour)
        print(f'Часовой расход газа на на отопление {self.V_heating_hour}')

        # 1.3.4 Определение расхода газа на вентиляцию
        self.V_ventilation_hour = (self.N * self.source_data['qv'] * (
                self.source_data['tv'] - self.source_data['t_middle']) * self.source_data['Vn']) / (
                                          self.source_data['Qn'] * self.source_data['kpd'])
        self.list_name_result.append(f'Часовой расход газа на  вентиляцию')
        self.list_result.append(self.V_ventilation_hour)
        print(f'Часовой расход газа на  вентиляцию {self.V_ventilation_hour}')

        # 1.3.5 Определение расхода газа на горячее водоснабжение
        self.V_hot_water_supply_hour = (self.source_data['Kc'] * self.N * (
                self.source_data['a'] + self.source_data['b']) * (60 - self.source_data['tx']) *
                                        self.source_data['Cv']) / (
                                               24 * self.source_data['Qn'] * self.source_data['kpd'])
        self.list_name_result.append(f'Часовой расход газа на Горячее водоснабжение')
        self.list_result.append(self.V_hot_water_supply_hour)
        print(f'Часовой расход газа на Горячее водоснабжение {self.V_hot_water_supply_hour}')

        # 1.3.6. Определение расхода газа на крупные и мелкие котельные
        # a) Расход газа на мелкие котельные
        self.V_small_boiler_houses = self.source_data['a_use_small_boiler'] * (
                self.V_heating_hour + self.V_ventilation_hour + self.V_hot_water_supply_hour)
        self.list_name_result.append(f'Часовой расход газа на  мелкие котельные')
        self.list_result.append(self.V_small_boiler_houses)
        print(f'Часовой расход газа на  мелкие котельные {self.V_small_boiler_houses}')

        # б) Расход газа на крупные котельные
        self.V_big_boiler_houses = self.source_data['a_use_big_boiler'] * (
                self.V_heating_hour + self.V_ventilation_hour + self.V_hot_water_supply_hour)
        self.list_name_result.append(f'Часовой расход газа на крупные котельные')
        self.list_result.append(self.V_big_boiler_houses)
        print(f'Часовой расход газа на крупные котельные {self.V_big_boiler_houses}')

        # 1.3.7 Расчетные расходы на сеть низкого давления
        self.V_bit = self.V_food_hour + self.V_hot_water_hour
        self.list_name_result.append(f'Часовой расход газа на быт')
        self.list_result.append(self.V_bit)
        print(f'Часовой расход газа на быт {self.V_bit}')

        self.Vsnd = self.V_bit + self.V_hospitals_hour + self.V_restaurants_and_cafes_hour + self.V_bathhouses_hour + self.V_small_boiler_houses + self.V_urban_industry_hour
        self.list_name_result.append(f'Часовой расход газа на сеть низкого давления')
        self.list_result.append(self.Vsnd)
        print(f'Часовой расход газа на сеть низкого давления {self.Vsnd}')
        self.output_excel()
        return self.Vsnd, self.V_bakeries_hour, self.V_laundromats_hour, self.V_big_boiler_houses

    def output_excel(self):
        wb = openpyxl.load_workbook(filename=self.file_out)
        sheet = wb['Расходы']
        for index, name in enumerate(self.list_name_result):
            sheet[f'A{index + 1}'] = name
            sheet[f'B{index + 1}'] = self.list_result[index]
        wb.save(self.file_out)

    def run(self):
        self.initializing_file()
        self.initializing_source_data()
        return self.counting_cost()


cal = GasExpenseCalculator('source_data.xlsx', 'result.xlsx')
V = cal.run()

