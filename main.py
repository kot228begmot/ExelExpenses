import openpyxl


def category_counter(month: int, year: int, write_exel: bool = False):
    """
    :param month: месяц, для которого надо произвести расчёты
    :param year: год, для которого надо произвести расчёты
    :return: словарь категорий, подкатегорий и сумм
    """
    book = openpyxl.open("test.xlsx", read_only=True)
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
    data_stack = []
    for row in range(2, exel_list.max_row + 1):
        print(row)
        if exel_list[row][2].value == None or exel_list[row][2].value == None:
            return data_stack
        elif exel_list[row][2].value.month == month and exel_list[row][2].value.year == year:
            if len(data_stack) == 0 or len(data_stack) == 1:
                data_stack.append(row)
            else:
                data_stack.pop()
                data_stack.append(row)
    return data_stack


def write_category(month: int, year: int, all_category):
    """
    :param month: месяц подсчётов
    :param year: год подсчётов
    :param all_category: набор категорий
    :return:
    """
    book = openpyxl.open("test.xlsx", read_only=False)
    list2 = book.worksheets[1]
    begin_category_list_col = 10
    for column in range(begin_category_list_col + 2, list2.max_column):
        if list2[2][column].value.month == month and list2[2][column].value.year == year:
            number_of_column = column - begin_category_list_col
    #row = 2
    row = 3
    while row <= list2.max_row + 1 or len(all_category) != 0:
        print(row)
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


category_counter(6, 2022, True)

