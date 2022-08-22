def if_function_1(ws, iterator_col, iterator_row, list_id, list_times, list_money, money, times,
                  excel_max_row, money_last, break_num):
    # print(iterator_row)
    # print(iterator_col)
    result = ws[iterator_col + str(iterator_row)].value
    if excel_max_row == iterator_row and '编号' in ws[iterator_col + str(1)].value:
        money += int(ws[money_last + str(iterator_row)].value)
        times += 1
        list_id.append(result)

        list_times.append(times)
        times = 0
        list_money.append(money)
        money = 0
    else:
        result_extra = ws[iterator_col + str(iterator_row + 1)].value
        if '编号' in ws[iterator_col + str(1)].value and result_extra != result:
            money += int(ws[money_last + str(iterator_row)].value)
            times += 1
            list_id.append(result)
            list_times.append(times)
            times = 0
            list_money.append(money)
            money = 0
            break_num = 1
        elif '缴费金额' in ws[iterator_col + str(1)].value:
            money += int(result)
        elif '缴费日期' in ws[iterator_col + str(1)].value:
            times += 1

    return list_id, list_times, list_money, money, times, break_num
