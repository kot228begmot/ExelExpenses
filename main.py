import openpyxl
class ExelExpenses:
    def __init__(self):
        self. _book = openpyxl.open("C:/Users/apec9/Dropbox/ExelExpenses/test.xlsx", read_only=True)
        self.list_number = self._book.worksheets[0]
        self._storage_of_category = {'еда': {},
                                     'Инвестиции': {},
                                     'Алкоголь': {},
                                     'Транспорт': {},
                                     'Продукты': {},
                                     'Рестораны': {},
                                     'Одежда': {},
                                     'Здоровье, красота, гигиена': {},
                                     'Кино, театры, музеи, развлечения': {},
                                     'Связь': {},
                                     'Квартира': {},
                                     'Творчество, книги, обучение, спорт': {},
                                     'Прочее': {}}
    def expenses_processing(self, month, year):
        error_positions = []
        self.find_border_of_the_month(month, year)



    def find_border_of_the_month(self, month, year):
        _position_of_start_line = 2
        _position_of_end_line = self.list_number.max_row
        _position_of_general_data_column = 2
        flag_find_year = False
        start_year_line = 0
        end_month_line = 0
        if self.list_number[_position_of_start_line][_position_of_general_data_column].value.year == year:
            start_year_line = _position_of_start_line
        else:
            while flag_find_year == False:
                if self.list_number[_position_of_start_line][_position_of_general_data_column].value.year == year - 1 and self.list_number[_position_of_start_line+ 1][_position_of_general_data_column].value.year == year:
                    start_year_line = _position_of_start_line
                    flag_find_year == True
                    break
                middle = (_position_of_start_line + _position_of_end_line) // 2
                if self.list_number[middle][_position_of_general_data_column].value.year < year:
                    _position_of_start_line = middle
                else:
                    _position_of_end_line = middle









            """
            if self.list_number[_position_of_start_list_line + 1][2].value.month == month and self.list_number[_position_of_start_list_line + 1][2].value.day == 1:
                return _position_of_start_list_line + 1
            middle = (_position_of_start_list_line + end_of_all_list_line) // 2
            if exel_list[middle][2].value.month < month:
                _position_of_start_list_line = middle
            else:
                end_of_all_list_line = middle
        return _position_of_start_list_line
            """






def category_counter(month: int, year: int, write_exel: bool = False):
    """
    :param month: месяц, для которого надо произвести расчёты
    :param year: год, для которого надо произвести расчёты
    :return: словарь категорий, подкатегорий и сумм
    """
    book = openpyxl.open("C:/Users/apec9/Dropbox/ExelExpenses/test.xlsx", read_only=True)
    daily_expenses = book.worksheets[0]  # позиционирование на листе
    # all_category = {'Продукты':{},'Транспорт':{}}
    all_category = {'транспорт': {},
                    'еда': {},
                    'здоровье, красота, гигиена': {},
                    'инвестиции': {},
                    'квартира': {},
                    'продукты': {},
                    'прочее': {},
                    'рестораны': {},
                    'связь': {},
                    'алкоголь': {},
                    'кино, театры, музеи': {},
                    'одежда': {},
                    'творчество, книги, обучение': {}}
    border_of_month = find_border_of_month(month, year, daily_expenses)
    for row in range(border_of_month[0], border_of_month[1] + 1):
        category = daily_expenses[row][4].value.lower()
        #_category = daily_expenses[row][5].value.lower()
        _category = daily_expenses[row][5].value
        if _category is None and all_category[category] == {}:
            all_category[category].setdefault('noncategory', daily_expenses[row][6].value)
        elif _category is None and all_category[category] != {}:
            all_category[category]['noncategory'] = daily_expenses[row][6].value + all_category[category]['noncategory']
        elif _category.lower() in all_category[category].keys():
            all_category[category][_category.lower()] = all_category[category][_category.lower()] + daily_expenses[row][6].value
        else:
            all_category[category].setdefault(_category.lower(), daily_expenses[row][6].value)
    book.close()
    if write_exel:
        write_category(month, year, all_category)
    return all_category


def find_border_of_month(month: int, year: int, exel_list):
    """
    :param month: месяц, для которого надо найти границы
    :param year: год месяца, для которого надо найти границы
    :return: возвращает list с двумя числами: начало месяца в строке, конец месяца в строке
    """
    start_list = left_bound(month, year, exel_list)
    end_list =  right_bound(month, year, exel_list, start_list)
    return [start_list,end_list]


def left_bound(month: int, year: int, exel_list):
    left = 2
    right = exel_list.max_row
    while exel_list[right][2].value.month - exel_list[left][2].value.month != 0:
        if exel_list[left + 1][2].value.month == month and exel_list[left + 1][2].value.day == 1:
            return left + 1
        middle = (left + right )//2
        if exel_list[middle][2].value.month < month:
            left = middle
        else:
            right = middle
    return left

def right_bound(month: int, year: int, exel_list, start_list: int):
    left = start_list #необходимо, чтобы всегда присутствовал первое число месяца
    right = exel_list.max_row - 1
    while exel_list[right][2].value.month - exel_list[left][2].value.month != 0:
        if exel_list[left + 1][2].value.month == month + 1 and exel_list[left + 1][2].value.day == 1:
            return left +1
        middle = (left + right )//2
        if exel_list[middle][2].value.month <= month:
            left = middle
        else:
            right = middle
    return right


def write_category(month: int, year: int, all_category):
    """
    :param month: месяц подсчётов
    :param year: год подсчётов
    :param all_category: набор категорий
    :return:
    """
    #book = openpyxl.open("test.xlsx", read_only=False)
    book = openpyxl.open("C:/Users/apec9/Dropbox/ExelExpenses/test.xlsx", read_only=False)
    list2 = book.worksheets[1]
    begin_category_list_col = 10
    for column in range(begin_category_list_col + 2, list2.max_column):
        if list2[2][column].value.month == month and list2[2][column].value.year == year:
            number_of_column = column - begin_category_list_col
    #row = 2
    row = 3
    while row <= list2.max_row + 1 or len(all_category) != 0:
        for category in all_category.keys():
            leven_length = levenstein(list2[row][begin_category_list_col].value, category)
            if leven_length == 0 or leven_length == 1:
                list2[row][begin_category_list_col + number_of_column].value = sum(all_category[category].values())
                position = 0
                while len(all_category[category]) != 0:
                    leven_length_ = levenstein(list2[row + 1][begin_category_list_col + 1].value,
                                               list(all_category[category].keys())[position])
                    if leven_length_ == 0 or leven_length_ ==1 and \
                            position <= len(all_category[category]):
                        list2[row + 1][begin_category_list_col + number_of_column].value = all_category[category][
                            list(all_category[category].keys())[position]]
                        all_category[category].pop(list(all_category[category].keys())[position])
                        row += 1
                        position = 0
                    elif leven_length_ == None and \
                            list(all_category[category].keys())[position] == 'noncategory':
                        list2[row + 1][begin_category_list_col + number_of_column].value = all_category[category][
                            list(all_category[category].keys())[position]]
                        all_category[category].pop(list(all_category[category].keys())[position])
                        row += 1
                        position = 0
                    elif position + 1 == len(all_category[category]):
                        list2[row + 1][begin_category_list_col + number_of_column].value = 0
                        position = 0
                        row += 1
                    else:
                        position += 1
                all_category.pop(category)
                break
        row += 1

    book.save("test.xlsx")
    book.close()
    return print('Расходы записаны')


def levenstein(A, B):
    if A == None:
        return None
    F = [[(i + j) if i * j == 0 else 0 for j in range(len(B) + 1)]
         for i in range(len(A) + 1)]
    for i in range(1, len(A) + 1):
        for j in range(1, len(B) + 1):
            if A[i - 1] == B[j - 1]:
                F[i][j] = F[i - 1][j - 1]
            else:
                F[i][j] = 1 + min(F[i - 1][j], F[i][j - 1], F[i - 1][j - 1])
    return F[len(A)][len(B)]




#print(levenstein('Жкх', 'жкх'))
xui = ExelExpenses()
a = xui.expenses_processing(3, 2022)



#category_counter(11, 2022, True)