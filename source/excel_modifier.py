from openpyxl import load_workbook


def options():
    workbook = load_workbook('database.xlsx')
    worksheet = workbook.active

    name_cache = {}
    removed_names = {}
    i = 2

    for row in range(2, worksheet.max_row + 1):
        # Put all <name>: <row #> in name_cache
        name_cache[worksheet.cell(row=row, column=1).value] = i
        i += 1

    whitelist = []
    blacklist = []
    default_start = start_date = worksheet.cell(row=1, column=2).value
    default_end = end_date = worksheet.cell(row=1, column=worksheet.max_column - 1).value

    while True:
        ans = input('Total Messages for a person? ("total <name>")\n'
                    'Whitelist("add <keyword>")\n'
                    'Blacklist("remove <keyword>"\n'
                    'Within a given time period ("time start <yyyy/mm/dd>" or "time end <yyyy/mm/dd>")? \n'
                    'Get statistics of a certain day ("day <yyyy/mm/dd>">\n'
                    'Done?\n').lower()

        info = ans.split()

        if info == []:
            # If nothing was entered
            pass

        elif 'done' in info[0] or 'exit' in info[0]:
            break

        elif 'add' in info[0]:
            # Adds a keyword to the whitelist that has to appear in a person's name
            whitelist.append(info[1])

        elif 'remove' in info[0]:
            # Adds a keyword to the blacklist that has to appear in a person's name
            blacklist.append(info[1])

        elif 'time' == info[0]:
            # Limit the time in the graph
            try:
                if info[2][:4].isnumeric() and info[2][4] == info[2][7] == '/' and info[2][5:7].isnumeric() and info[2][8:].isnumeric():
                    if 'start' == info[1]:
                        if info[2] < end_date:
                            start_date = info[2]
                        else:
                            print('Start date exceeds the end date!')
                    elif 'end' == info[1]:
                        if info[2] > start_date:
                            end_date = info[2]
                        else:
                            print('End date exceeds the start date!')
            except: print('Invalid Date!')

        elif 'total' in info[0]:
            # Get total reiceved msgs from a person
            try:
                attempted_people = []

                for name in name_cache:
                    # Tests all names in database
                    if info[1] in name.lower():
                        # If name matches the keyword
                        attempted_people.append(name)

                if len(attempted_people) == 0:
                    # If no results match
                    print('Person does not exist!')
                elif len(attempted_people) == 1:
                    # If only one result appears
                    print('You have a total of: ' + str(worksheet.cell(row=name_cache[attempted_people[0]], column=worksheet.max_column).value) + ' messages from ' + attempted_people[0] + '\n')
                else:
                    # If multiple results appear
                    print('Did you mean:')

                    for i in range(len(attempted_people)):
                        # Post all the results of people that showed up and asks user to pick one
                        print(str(i + 1) + '. ' + attempted_people[i] + '\n')

                    ans = input('Enter the number of the person you want to search, otherwise enter any other key\n')
                    if ans.isnumeric() and int(ans) - 1 <= len(attempted_people):
                        print('You have a total of: ' + str(worksheet.cell(row=name_cache[attempted_people[int(ans) - 1]], column=worksheet.max_column).value) + ' messages from ' + attempted_people[int(ans) - 1] + '\n')
            except:
                # If invalid information was entered
                print('Person does not exist!\n')


    edited = False

    if start_date != default_start:
        remove_from_start(worksheet, start_date)
        edited = True
    if end_date != default_end:
        remove_from_end(worksheet, end_date)
        edited = True
    if whitelist != []:
        check_names(worksheet, name_cache, removed_names, list=whitelist, key='whitelist')
        edited = True
    if blacklist != []:
        check_names(worksheet, name_cache, removed_names, list=whitelist, key='blacklist')
        edited = True

    if edited == True:
        #remove_empty(worksheet, name_cache, removed_names)

        ans = input('What would you like to name the file?\n')
        workbook.save(ans.split('.')[0] + '.xlsx')


def remove_empty(worksheet):
    pass


def check_names(worksheet, name_cache, removed_names, list=None, key=None):
    if key is None or list is None:
        raise Exception('Invalid key!')

    for name in name_cache.copy():

        if key == 'whitelist':
            has = whitelist_keys(list, name.lower())
        else:
            has = blacklist_keys(list, name.lower())

        if not has:
            delete_row(worksheet, name_cache[name])
            removed_names[name] = name_cache[name]
            name_cache.pop(name)


def whitelist_keys(whitelist, name):
    """
    Remove row if name does not have a keyword in the whitelist

    :param worksheet:
    :param whitelist:
    :return:
    """
    for key in whitelist:
        if key.lower() in name:
            return True
    return False


def blacklist_keys(blacklist, name):
    """
    Remove row if name has a keyword in the whitelist

    :param worksheet:
    :param whitelist:
    :return:
    """

    for key in blacklist:
        if not key.lower() in name:
            return True
    return False


def remove_from_start(worksheet, start_date):
    """
    Remove dates before <start_date> in worksheet

    :param worksheet:
    :param start_date:
    :return:
    """
    row = 1
    col = 2
    cell = ''

    while cell != 'Total':
        cell = worksheet.cell(row=row, column=col).value

        if cell < start_date:
            delete_column(worksheet, col)
            col += 1
        else:
            break


def remove_from_end(worksheet, end_date):
    """
    Remove dates after <end_date> in worksheet

    :param worksheet:
    :param end_date:
    :return:
    """
    row = 1
    col = worksheet.max_column - 1
    cell = 'Total'

    while cell != '':
        cell = worksheet.cell(row=row, column=col).value

        if cell > end_date:
            delete_column(worksheet, col)
            col -= 1
        else:
            break


def delete_row(worksheet, row):
    """
    Deletes row <row> in the worksheet

    :param worksheet:
    :param row:
    :return:
    """

    for col in range(1, worksheet.max_row):
        worksheet.cell(row=row, column=col).value = None


def delete_column(worksheet, col):
    """
    Deletes column <col> in the worksheet

    :param worksheet:
    :param col:
    :return:
    """
    for row in range(1, worksheet.max_row):
        worksheet.cell(row=row, column=col).value = None
