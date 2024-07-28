import xlsxwriter
import xlrd


class Employee:
    def __init__(self, name_of_employee):
        self.name = name_of_employee
        self.id = name_of_employee
        self.total_money = 0
        self.total_works = 0
        self.general_works = {}
        self.additional_works = {}

    def set_general_works(self, dictionary_to_set):
        self.general_works = dictionary_to_set.copy()

    def set_additional_works(self, dictionary_to_set):
        self.additional_works = dictionary_to_set.copy()


# {'code':[amount, total_money_per_code]
def update_dict(dictionary_to_update, manip_code, amount_of_manip, price_of_manip):
    dictionary_to_work = dictionary_to_update.copy()
    if dictionary_to_work.get(manip_code) is None:
        dictionary_to_work.update(
            {manip_code: [amount_of_manip, price_of_manip]})
    else:
        amount_of_manip_old = dictionary_to_update.get(manip_code)[0]
        price_of_manip_old = dictionary_to_update.get(manip_code)[1]
        dictionary_to_work.update(
            {manip_code: [amount_of_manip_old + amount_of_manip, price_of_manip_old + price_of_manip]})

    return dictionary_to_work


def is_float_digit(n: str) -> bool:
    try:
        float(n)
        return True
    except ValueError:
        return False


def first_scan(name_of_file):
    workbook = xlrd.open_workbook(name_of_file)
    sheet = workbook.sheet_by_name('Sheet1')

    check_id = 0
    employee_name_id = 12
    check_all_employees = []
    list_of_all_employees = []

    for rx in range(sheet.nrows):
        split_line = []
        for cell in sheet.row(rx):
            cell_value = cell.value
            split_line.append(cell_value)

        if is_float_digit(split_line[check_id]):
            if split_line[employee_name_id] not in check_all_employees and split_line[employee_name_id] != "":
                new_employee = Employee(split_line[employee_name_id])
                list_of_all_employees.append(new_employee)
                check_all_employees.append(split_line[employee_name_id])

    return list_of_all_employees


def scan_file_for_people(name_of_file, list_of_employees, dates_mass, choice, need_to_write):
    workbook = xlrd.open_workbook(name_of_file)
    sheet = workbook.sheet_by_name('Sheet1')

    date_id = 0
    if choice == "дата сдачи":
        date_id = 3
    elif choice == "дата отправки":
        date_id = 2
    elif choice == "дата техника":
        date_id = 23
    else:
        date_id = 3
    check_id = 0
    manip_id = 6
    manip_amount_id = 9
    manip_price_id = 11
    employee_name_id = 12

    list_of_all_employees = []
    for employee in list_of_employees:
        person = Employee(employee)
        list_of_all_employees.append(person)
    dict_of_general_manip = {}
    dict_of_additional_manip = {}

    dict_of_general_manip_zeros = {}
    dict_of_additional_manip_zeros = {}

    for rx in range(sheet.nrows):
        split_line = []
        for cell in sheet.row(rx):
            cell_value = cell.value
            split_line.append(cell_value)
        if is_float_digit(split_line[check_id]) and split_line[manip_price_id] > 0 \
                and len(split_line[employee_name_id]) > 0:
            date_1 = split_line[date_id].split("-")
            flag = False
            for date_curr in dates_mass:
                date_2 = str(date_curr).split("-")
                if date_1[0] == date_2[2] and date_1[1] == date_2[1] and date_1[2] == date_2[0]:
                    flag = True

            if flag:
                manip_to_work = split_line[manip_id]
                # manip_to_work = [2161500, 2161501, 2161503, 1161500, 1161505]

                if manip_to_work % 100 == 0.0:
                    if dict_of_general_manip.get(manip_to_work) is None:
                        dict_of_general_manip.update({manip_to_work:
                                                          [split_line[manip_amount_id], split_line[manip_price_id]]})
                    else:
                        dict_of_general_manip = update_dict(dict_of_general_manip, manip_to_work,
                                                            split_line[manip_amount_id], split_line[manip_price_id])
                else:
                    if dict_of_additional_manip.get(manip_to_work) is None:
                        dict_of_additional_manip.update({manip_to_work:
                                                             [split_line[manip_amount_id], split_line[manip_price_id]]})
                    else:
                        dict_of_additional_manip = update_dict(dict_of_additional_manip, manip_to_work,
                                                               split_line[manip_amount_id], split_line[manip_price_id])

    general_sorted = sorted(dict_of_general_manip.items(), key=lambda item: item[1][1])
    general_sorted.reverse()
    general_sorted = dict(general_sorted)
    additional_sorted = sorted(dict_of_additional_manip.items(), key=lambda item: item[1][1])
    additional_sorted.reverse()
    additional_sorted = dict(additional_sorted)

    for manip_to_work in general_sorted:
        dict_of_general_manip_zeros.update({manip_to_work: [0, 0]})
    for manip_to_work in additional_sorted:
        dict_of_additional_manip_zeros.update({manip_to_work: [0, 0]})

    if need_to_write:
        for i in range(len(list_of_all_employees)):
            list_of_all_employees[i].set_additional_works(dict_of_additional_manip_zeros)
            list_of_all_employees[i].set_general_works(dict_of_general_manip_zeros)

    return list_of_all_employees, general_sorted, additional_sorted


def write_file(name_of_file, list_of_all_employees, dates_mass, choice, modified, text_name):
    list_of_employees, general_stored, additional_stored = scan_file_for_people(name_of_file, list_of_all_employees,
                                                                                dates_mass, choice, modified)
    workbook = xlrd.open_workbook(name_of_file)
    sheet = workbook.sheet_by_name('Sheet1')

    date_id = 0
    if choice == "дата сдачи":
        date_id = 3
    elif choice == "дата отправки":
        date_id = 2
    elif choice == "дата техника":
        date_id = 23
    else:
        date_id = 3

    # curr_date = datetime.now().month
    check_id = 0
    manip_id = 6
    manip_name_id = 7
    manip_amount_id = 9
    price_id = 11
    employee_name_id = 12
    # list_of_employees = ["Вяткин Э.В.", "Великотрав А.В.", "Мирзоян Б.Б.", "Белоконь А.В.", "Шевцова А.И.",
    #                     "Артемьева А.А.", "Александрова О.А.", "Акимова И.Г", "Акимова И.Г.", "Мацевич Д.Г."]

    if modified:
        all_together_employee = Employee("Итог")
        all_together_employee.general_works = general_stored
        all_together_employee.additional_works = additional_stored
        list_of_employees.insert(0, all_together_employee)
    price_list = {}

    for rx in range(sheet.nrows):
        split_line = []
        for cell in sheet.row(rx):
            cell_value = cell.value
            split_line.append(cell_value)

        if is_float_digit(split_line[check_id]):
            date_1 = split_line[date_id].split("-")
            if split_line[price_id] != 0:
                flag = False
                for date_curr in dates_mass:
                    date_2 = str(date_curr).split("-")
                    if date_1[0] == date_2[2] and date_1[1] == date_2[1] and date_1[2] == date_2[0]:
                        flag = True

                if flag:
                    for employee in list_of_employees:
                        if split_line[employee_name_id] == employee.id:
                            service = price_list.get(split_line[manip_id])
                            if service is None:
                                price_list.update(
                                    {split_line[manip_id]: [split_line[manip_name_id], split_line[price_id] /
                                                            split_line[manip_amount_id]]})

                            if split_line[manip_id] % 10 == 0:
                                employee.general_works = update_dict(employee.general_works, split_line[manip_id],
                                                                     split_line[manip_amount_id],
                                                                     split_line[price_id])
                            else:
                                employee.additional_works = update_dict(employee.additional_works, split_line[manip_id],
                                                                        split_line[manip_amount_id],
                                                                            split_line[price_id])
                            break

    for employee in list_of_employees:
        general_works_for_employee = employee.general_works
        for id_of_task in general_works_for_employee:
            employee.total_works += general_works_for_employee.get(id_of_task)[0]
            employee.total_money += general_works_for_employee.get(id_of_task)[1]

        additional_works_for_employee = employee.additional_works
        for id_of_task in additional_works_for_employee:
            employee.total_works += additional_works_for_employee.get(id_of_task)[0]
            employee.total_money += additional_works_for_employee.get(id_of_task)[1]

    workbook = xlsxwriter.Workbook(text_name)
    worksheet = workbook.add_worksheet()
    row_to_write = 0
    column_to_write = 0

    for employee in list_of_employees:
        worksheet.write(row_to_write, column_to_write + 1, F"{employee.id}")
        worksheet.write(row_to_write + 1, column_to_write, "Код")
        worksheet.write(row_to_write + 1, column_to_write + 1, "Название")
        worksheet.write(row_to_write + 1, column_to_write + 2, "Кол-во")
        worksheet.write(row_to_write + 1, column_to_write + 3, "цена")
        if modified:
            worksheet.write(row_to_write + 2, column_to_write + 1, "Основные Позиции")

        column_to_write += 5

    column_to_write = 0

    for employee in list_of_employees:
        if modified:
            row_to_write = 3
        else:
            row_to_write = 2

        tasks = employee.general_works
        total_price = employee.total_money
        total_works = employee.total_works
        for id_to_write in tasks:
            #print(id_to_write, price_list.get(id_to_write))

            name_of_task = price_list.get(id_to_write)[0]
            number_of_tasks = tasks.get(id_to_write)[0]
            total_money_per_task = tasks.get(id_to_write)[1]

            worksheet.write_number(row_to_write, column_to_write, int(id_to_write))
            worksheet.write(row_to_write, column_to_write + 1, F'{name_of_task}')
            worksheet.write_number(row_to_write, column_to_write + 2, int(number_of_tasks))
            worksheet.write_number(row_to_write, column_to_write + 3, int(total_money_per_task))

            row_to_write += 1

        if modified:
            worksheet.write(row_to_write + 1, column_to_write + 1, "Второстепенные Позиции")
            row_to_write += 2

        tasks = employee.additional_works
        for id_to_write in tasks:
            name_of_task = price_list.get(id_to_write)[0]
            number_of_tasks = tasks.get(id_to_write)[0]
            total_money_per_task = tasks.get(id_to_write)[1]

            worksheet.write_number(row_to_write, column_to_write, int(id_to_write))
            worksheet.write(row_to_write, column_to_write + 1, F'{name_of_task}')
            worksheet.write_number(row_to_write, column_to_write + 2, int(number_of_tasks))
            worksheet.write_number(row_to_write, column_to_write + 3, int(total_money_per_task))

            row_to_write += 1

        worksheet.write(row_to_write, column_to_write, "Итог")
        worksheet.write_number(row_to_write, column_to_write + 3, int(total_price))
        worksheet.write(row_to_write + 1, column_to_write, "Кол-во работ")
        worksheet.write_number(row_to_write + 1, column_to_write + 3, int(total_works))

        column_to_write += 5

    workbook.close()
