import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter


def split_years(full_file, year_file, start, end, initials, year):

    year_book = openpyxl.load_workbook(year_file)
    full_book = openpyxl.load_workbook(full_file)
    year_sheet = year_book.get_sheet_by_name(name='TB from Prior Mgmt')
    full_sheet = full_book.get_sheet_by_name(name='TB from Prior Mgmt')
    full_range = range(column_index_from_string(start), column_index_from_string(end) + 3)

    for index in range(1, 105):
        for j in range(4, 279):
            full = full_range[index - 1]

            year_sheet[get_column_letter(index) + str(j - 2)].value = full_sheet[get_column_letter(full) + str(j)].value

    year_book.save(
        'C:\Users\T530\Desktop\Banta\\' + initials + '\\' + initials + year +
        ' Banta Management Chart of Account Mapping .xlsx')


def move_trial_balances(initials, length):
    target = openpyxl.load_workbook('C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping.xlsx')
    target_sheet = target.get_sheet_by_name(name='TB from Prior Mgmt')

    source = openpyxl.load_workbook('C:\Users\T530\Desktop\Banta\\' + initials + '\\' + initials + 'TB201312.xlsx')
    source_sheet = source.get_sheet_by_name(name='Sheet1')

    for i in range(14, length, 1):
        j = i-9
        account_num_column = get_column_letter(1)
        account_name_column = get_column_letter(2)
        debit_column = get_column_letter(4)
        credit_column = get_column_letter(5)
        ending_column = get_column_letter(7)

        target_sheet[account_num_column + str(j)].value = source_sheet['A' + str(i)].value
        target_sheet[account_name_column + str(j)].value = source_sheet['E' + str(i)].value
        target_sheet[debit_column + str(j)].value = source_sheet['H' + str(i)].value
        target_sheet[credit_column + str(j)].value = source_sheet['I' + str(i)].value
        if target_sheet[debit_column + str(j)].value or target_sheet[debit_column + str(j)].value == 0:
            target_sheet[ending_column + str(j)].value = target_sheet[debit_column + str(j)].value
        elif target_sheet[credit_column + str(j)].value or target_sheet[credit_column + str(j)].value == 0:
            target_sheet[ending_column + str(j)].value = target_sheet[credit_column + str(j)].value * -1

    for index in range(0, column_index_from_string('OR') / 8, 1):

        account_num_column = get_column_letter(index * 8 + 9)
        account_name_column = get_column_letter(index * 8 + 10)
        opening_column = get_column_letter(index * 8 + 11)
        debit_column = get_column_letter(index * 8 + 12)
        credit_column = get_column_letter(index * 8 + 13)
        net_change_column = get_column_letter(index * 8 + 14)
        ending_column = get_column_letter(index * 8 + 15)

        year = 2014 + (index/12)
        month = (index % 12) + 1
        print str(month) + str(year)

        if month < 10:
            source = openpyxl.load_workbook('C:\Users\T530\Desktop\Banta\\'+initials+'\\'+initials+'TB' + str(year) +
                                            '0' + str(month) + '.xlsx')

        else:
            source = openpyxl.load_workbook('C:\Users\T530\Desktop\Banta\\' + initials + '\\'+initials+'TB' + str(year)
                                            + str(month) + '.xlsx')

        source_sheet = source.get_sheet_by_name(name='Sheet1')

        for i in range(14, length, 1):
            j = i - 9
            target_sheet[account_num_column + str(j)].value = source_sheet['A' + str(i)].value
            target_sheet[account_name_column + str(j)].value = source_sheet['E' + str(i)].value

            if int(target_sheet[account_num_column + str(j)].value) == 1996:
                retained_row = j

            target_sheet[opening_column + str(j)].value = target_sheet[get_column_letter(index * 8 + 7) + str(j)].value

            if index % 12 == 0 and int(target_sheet[account_num_column + str(j)].value) > 1999:
                target_sheet[opening_column + str(retained_row)].value += target_sheet[get_column_letter(index * 8 + 7)
                                                                                       + str(j)].value
                target_sheet[opening_column + str(j)].value = 0

            target_sheet[debit_column + str(j)].value = source_sheet['H' + str(i)].value

            target_sheet[credit_column + str(j)].value = source_sheet['I' + str(i)].value

            if index < 1:
                if target_sheet['D' + str(i)].value:
                    target_sheet[ending_column + str(j)].value = target_sheet['D' + str(i)].value
                elif target_sheet['E' + str(i)].value:
                    target_sheet[ending_column + str(j)].value = target_sheet['E' + str(i)].value * -1

            if target_sheet[debit_column + str(j)].value or target_sheet[debit_column + str(j)].value == 0:
                target_sheet[ending_column + str(j)].value = target_sheet[debit_column + str(j)].value
            elif target_sheet[credit_column + str(j)].value or target_sheet[credit_column + str(j)].value == 0:
                target_sheet[ending_column + str(j)].value = target_sheet[credit_column + str(j)].value * -1
            else:
                break

            target_sheet[net_change_column + str(j)].value = target_sheet[ending_column + str(j)].value - target_sheet[
                opening_column + str(j)].value

    target.save('C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping - '+initials+'.xlsx')

    split_years('C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping - '+initials+'.xlsx',
                'C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping 2014.xlsx',
                'A', 'CZ', initials, '2014')
    split_years('C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping - '+initials+'.xlsx',
                'C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping 2015.xlsx',
                'CR', 'GS', initials, '2015')
    split_years('C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping - '+initials+'.xlsx',
                'C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping 2016.xlsx',
                'GK', 'KK', initials, '2016')
    split_years('C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping - '+initials+'.xlsx',
                'C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping 2017.xlsx',
                'KC', 'OC', initials, '2017')
    split_years('C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping - '+initials+'.xlsx',
                'C:\Users\T530\Desktop\Banta\Banta Management Chart of Account Mapping 2018.xlsx',
                'NU', 'RS', initials, '2018')


Inputs = {"HI": 162, "MB": 307, "SI": 274}

for initial in Inputs.keys():
    print initial
    move_trial_balances(initial, Inputs[initial])
