import pandas as pd
from datetime import datetime, date
import dateutil.parser
import re
import ast
import sys

def earlier_date(reg_list):
    reglist_num = list(enumerate(reg_list, 0))
    dates_num = []
    for i in range(len(reglist_num)):
        if type(reglist_num[i][1]) == datetime:
            dates_num.append(reglist_num[i])
    if dates_num == []:
        return 'Собств', None
    earlier_date = min(dates_num, key=lambda i: i[1])
    reg_types = ['Собств', 'ИП', 'ЕЗП', 'ЕЗП_ИП', 'Пред_собств', 'Пред_ИП']
    earlier = reg_types[earlier_date[0]]
    return earlier, earlier_date[1]

def define_date(timeStampt):
    if type(timeStampt) == str:
        return None
    else:
        return timeStampt
#     try:
#         return dateutil.parser.parse(date)
#     except:
#         return None

def prev_cicle(all_prev, all_reg, result):
    last = all_prev[-1]
    if len(all_prev) < 2:
        result.append(last)
        # arcpy.AddMessage("circle "+str(result))
        return result[::-1]
    z = 0
    for i in all_prev[::-1][1:]:
        if last in all_reg['Предыдущие номера'].loc[all_reg['Кадастровый номер'] == i].iloc[0].split(', '):
            result.append(last)
            return prev_cicle(all_prev=all_prev[:-1-z], all_reg=all_reg, result=result)
        z += 1

def prev_zu(prev_list, all_reg, ip_reg, all_prev, dates, ip_dates, no_xml, no_xml_ip, prev=None, prev_date=None, ip_date=None):
    if prev_list == ['-']:
        if len(all_prev) > 0:
            return all_prev[-1], prev, prev_date, ip_date, all_prev, no_xml, no_xml_ip
        else:
            return None, None, None, None, [], None, None
    for prev_cad in prev_list:
        if prev_cad in all_reg['Кадастровый номер'].to_list():
            prev_date = define_date(all_reg['Дата_собственности'].loc[all_reg['Кадастровый номер'] == prev_cad].iloc[0])
            prev = all_reg['Собственность'].loc[all_reg['Кадастровый номер'] == prev_cad].iloc[0]
            if prev_date != None:
                if prev_date < datetime(2016, 1, 1, 0, 0):
                    all_prev.append(prev_cad)
                    return prev_cad, prev, prev_date, ip_date, all_prev, no_xml, no_xml_ip
        else:
            no_xml.update([prev_cad])
            prev_date = None
        if prev_cad in ip_reg['Кадастровый номер'].to_list():
            ip_date = define_date(
                ip_reg['Дата_собственности'].loc[ip_reg['Кадастровый номер'] == prev_cad].iloc[0])
            if ip_date != None:
                if ip_date < datetime(2016, 1, 1, 0, 0):
                    all_prev.append(prev_cad)
                    return prev_cad, prev, prev_date, ip_date, all_prev, no_xml, no_xml_ip
        else:
            no_xml_ip.update([prev_cad])
            ip_date = None
        if prev_cad in all_reg['Кадастровый номер'].to_list():
            if prev_date == None:
                dates.append(datetime(2040, 1, 1, 0, 0))
            else:
                dates.append(prev_date)
            if ip_date == None:
                ip_dates.append(datetime(2040, 1, 1, 0, 0))
            else:
                ip_dates.append(ip_date)
            all_prev.append(prev_cad)
            prev_list = all_reg['Предыдущие номера'].loc[all_reg['Кадастровый номер'] == prev_cad].iloc[0].split(', ')
            cad, prev, prev_date, ip_date, all_prev, no_xml, no_xml_ip = prev_zu(prev_list=prev_list, all_reg=all_reg,
                                                                                 ip_reg=ip_reg, prev=prev, prev_date=prev_date,
                                                                                 ip_date=ip_date, all_prev=all_prev, dates=dates,
                                                                                 ip_dates=ip_dates, no_xml=no_xml, no_xml_ip=no_xml_ip)
            if prev_date != None:
                if prev_date < datetime(2016, 1, 1, 0, 0):
                    all_prev = prev_cicle(all_prev, all_reg, result=[])
                    return cad, prev, prev_date, ip_date, all_prev, no_xml, no_xml_ip
            if ip_date != None:
                if ip_date < datetime(2016, 1, 1, 0, 0):
                    all_prev = prev_cicle(all_prev, all_reg, result=[])
                    return cad, prev, prev_date, ip_date, all_prev, no_xml, no_xml_ip
        else:
            continue
    # arcpy.AddMessage(str(prev_list) + " " + str(all_prev))
    if len(all_prev) > 0:
        dates_num = list(enumerate(dates, 0))
        ip_dates_num = list(enumerate(ip_dates, 0))
        if len(dates) != 0:
            earlier_regdate = min(dates_num, key=lambda i: i[1])
        else:
            earlier_regdate = [0, datetime(2040, 1, 1, 0, 0)]
            dates = [datetime(2040, 1, 1, 0, 0)]
        if len(ip_dates) != 0:
            earlier_ip_regdate = min(ip_dates_num, key=lambda i: i[1])
        else:
            earlier_ip_regdate = [0, datetime(2040, 1, 1, 0, 0)]
            ip_dates = [datetime(2040, 1, 1, 0, 0)]
        if earlier_regdate[1] <= earlier_ip_regdate[1]:
            prev_date = dates[earlier_regdate[0]]
            ip_date = ip_dates[earlier_regdate[0]]
            prev_cad = all_prev[earlier_regdate[0]]
            all_prev = all_prev[:earlier_regdate[0]+1]
        else:
            prev_date = dates[earlier_ip_regdate[0]]
            ip_date = ip_dates[earlier_ip_regdate[0]]
            prev_cad = all_prev[earlier_ip_regdate[0]]
            all_prev = all_prev[:earlier_ip_regdate[0]+1]
        if prev_date == datetime(2040, 1, 1, 0, 0):
            prev_date = None
        if ip_date == datetime(2040, 1, 1, 0, 0):
            ip_date = None
        all_prev = prev_cicle(all_prev, all_reg, result=[])
        return prev_cad, prev, prev_date, ip_date, all_prev, no_xml, no_xml_ip
    else:
        if prev_date == datetime(2040, 1, 1, 0, 0):
            prev_date = None
        if ip_date == datetime(2040, 1, 1, 0, 0):
            ip_date = None
        all_prev.append(prev_cad)
        return prev_cad, prev, prev_date, ip_date, all_prev, no_xml, no_xml_ip

def main(input_xlsx, output_xlsx):
    xls = pd.ExcelFile(input_xlsx)  # Название исходного файла
    result_name = output_xlsx  # Название файла с результатом
    du_cads = pd.read_excel(xls, 'Двойной учет')
    all_reg = pd.read_excel(xls, 'Собственность')
    ip_reg = pd.read_excel(xls, 'История права')
    ezp_base = pd.read_excel(xls, 'Двойной учет')
    result = []
    du_cads_list = filter(lambda x: re.search(r':', x), set(du_cads['Кадастровый номер'].to_list()))
    xml_exists = True
    for cad in du_cads_list:
        # arcpy.AddMessage("УЧАСТОК " + cad)
        xml_exists = True
        reg = None
        reg_date = None
        ip_date = None
        prev_cad = None
        prev = None
        prev_date = None
        prev_ip = None
        all_prev = None
        no_xml = None
        no_xml_ip = None
        ezp = None
        ezp_reg = None
        ezp_date = None
        ezp_ip_date = None
        if cad in all_reg['Кадастровый номер'].to_list():
            reg = all_reg['Собственность'].loc[all_reg['Кадастровый номер'] == cad].iloc[0]
            reg_date = define_date(
                all_reg['Дата_собственности'].loc[all_reg['Кадастровый номер'] == cad].iloc[0])
            if reg_date != None:
                if reg_date < datetime(2016, 1, 1, 0, 0):
                    earlier, earlier_regdate = earlier_date(
                        [reg_date, ip_date, ezp_date, ezp_ip_date, prev_date, prev_ip])
                    result.append([cad, reg, reg_date, ip_date, ezp, ezp_reg, ezp_date, ezp_ip_date,
                                  all_prev, prev_cad, prev, prev_date, prev_ip, earlier, earlier_regdate])
                    continue
        else:
            reg_date = 'Нет выписки'
            xml_exists = False
        if cad in ip_reg['Кадастровый номер'].to_list():
            ip_date = define_date(
                ip_reg['Дата_собственности'].loc[ip_reg['Кадастровый номер'] == cad].iloc[0])
            if ip_date != None:
                if ip_date < datetime(2016, 1, 1, 0, 0):
                    earlier, earlier_regdate = earlier_date(
                        [reg_date, ip_date, ezp_date, ezp_ip_date, prev_date, prev_ip])
                    result.append([cad, reg, reg_date, ip_date, ezp, ezp_reg, ezp_date, ezp_ip_date,
                                  all_prev, prev_cad, prev, prev_date, prev_ip, earlier, earlier_regdate])
                    continue
        else:

            ip_date = 'Нет выписки'
        if cad in ezp_base['Кадастровый номер'].to_list():
            ezp = ezp_base['ЕЗП'].loc[ezp_base['Кадастровый номер']
                                      == cad].iloc[0]
            if ezp != '-':
                if ezp in all_reg['Кадастровый номер'].to_list():
                    ezp_reg = all_reg['Собственность'].loc[all_reg['Кадастровый номер'] == ezp].iloc[0]
                    ezp_date = define_date(
                        all_reg['Дата_собственности'].loc[all_reg['Кадастровый номер'] == ezp].iloc[0])
                    if ezp_date != None:
                        if ezp_date < datetime(2016, 1, 1, 0, 0):
                            earlier, earlier_regdate = earlier_date(
                                [reg_date, ip_date, ezp_date, ezp_ip_date, prev_date, prev_ip])
                            result.append([cad, reg, reg_date, ip_date, ezp, ezp_reg, ezp_date, ezp_ip_date,
                                          all_prev, prev_cad, prev, prev_date, prev_ip, earlier, earlier_regdate])
                            continue
                else:
                    ezp_date = 'Нет выписки'
                if ezp in ip_reg['Кадастровый номер'].to_list():
                    ezp_ip_date = define_date(
                        ip_reg['Дата_собственности'].loc[ip_reg['Кадастровый номер'] == ezp].iloc[0])
                    if ezp_ip_date != None:
                        if ezp_ip_date < datetime(2016, 1, 1, 0, 0):
                            earlier, earlier_regdate = earlier_date(
                                [reg_date, ip_date, ezp_date, ezp_ip_date, prev_date, prev_ip])
                            result.append([cad, reg, reg_date, ip_date, ezp, ezp_reg, ezp_date, ezp_ip_date,
                                          all_prev, prev_cad, prev, prev_date, prev_ip, earlier, earlier_regdate])
                            continue
                else:
                    ezp_ip_date = 'Нет выписки'
        if xml_exists:
            prev_list = all_reg['Предыдущие номера'].loc[all_reg['Кадастровый номер'] == cad].iloc[0].split(', ')
            prev_cad, prev, prev_date, prev_ip, all_prev, no_xml, no_xml_ip = prev_zu(
                prev_list, all_reg, ip_reg, all_prev=[], dates=[], ip_dates=[], no_xml=set(), no_xml_ip=set())
        if cad == prev_cad:
            prev_cad, prev, prev_date, prev_ip, all_prev, no_xml, no_xml_ip = None, None, None, None, None, None, None
        if no_xml == set():
            no_xml = None
        if no_xml_ip == set():
            no_xml_ip = None
        if all_prev == []:
            all_prev = None
        earlier, earlier_regdate = earlier_date(
            [reg_date, ip_date, ezp_date, ezp_ip_date, prev_date, prev_ip])
        if type(all_prev) == list:
            all_prev = ', '.join(all_prev)
        if type(no_xml) == set:
            no_xml = ', '.join(no_xml)
        if type(no_xml_ip) == set:
            no_xml_ip = ', '.join(no_xml_ip)
        result.append([cad, reg, reg_date, ip_date, ezp, ezp_reg, ezp_date, ezp_ip_date, all_prev,
                      prev_cad, prev, prev_date, prev_ip, earlier, earlier_regdate, no_xml, no_xml_ip])
    writer = pd.ExcelWriter(f'{result_name}.xlsx')
    df = pd.DataFrame(result, columns=['Кадастровый номер', 'Собств', 'Дата', 'Ип', 'ЕЗП',
                                       'ЕЗП_Собств', 'ЕЗП_Дата', 'ЕЗП_ИП', 'Прев_цикл',
                                       'Прев_кад', 'Прев_собств', 'Прев_дата', 'Прев_ип',
                                       'Ранняя_Собств', 'Ранняя_Дата', 'Прев_без_выписок',
                                       'Прев_без_выписок_ИП'])
    df.to_excel(writer, sheet_name='Собств', index=False)
    xls.close()
    writer.close()
    print('ГОТОВО!')
    show = pd.DataFrame(result, columns=['cad',
                                        'reg',
                                        'reg_date',
                                        'ip_date',
                                        'ezp',
                                        'ezp_reg',
                                        'ezp_date',
                                        'ezp_ip_date',
                                        'all_prev',
                                        'prev_cad',
                                        'prev',
                                        'prev_date',
                                        'prev_ip',
                                        'earlier',
                                        'earlier_regdate',
                                        'no_xml',
                                        'no_xml_ip'])
    # show

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
