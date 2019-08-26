import pandas as pd
import datetime
import calendar

nhom_fileName = 'nhom.xlsx'
tkb_fileName = 'TKB.xlsx'
speaker_fileName = 'nguoi_noi.xlsx'

data_nhom = pd.read_excel(nhom_fileName)
data_TKB = pd.read_excel(tkb_fileName)
lich_nguoi_noi = pd.read_excel(speaker_fileName)
mon = list(data_nhom.head(0))[1:]
lich = lich_nguoi_noi.copy()
now = datetime.datetime.now()
thu = list(calendar.day_name)


def search_class(mon, nhom):
    for i, ma_mon in enumerate(data_TKB['Subject']):
        ma_nhom = data_TKB['Group ID'][i]
        if ma_mon == mon or ma_nhom == nhom:
            return data_TKB.iloc[i]
    return None


def to_day_string(day):
    return thu[day - 2]


def xep_lich_bt(lop):
    if lop == len(data_nhom):
        return True
    ten_lop, nhom = data_nhom.iloc[lop][0], data_nhom.iloc[lop][1:]
    for i in range(len(nhom)):
        # Lịch học của lớp đó
        lich_hoc = search_class(mon[i], nhom[i])

        for j in range(len(lich_nguoi_noi['Day'])):
            # Tìm người nói có lịch phù hợp với kíp đó
            thu_lop = to_day_string(lich_hoc['Day'])
            thu_nguoi_noi = datetime.datetime.strptime(lich_nguoi_noi['Day'][j], '%d/%m').replace(now.year).strftime(
                '%A')
            if thu_lop == thu_nguoi_noi and \
                    lich_nguoi_noi.iat[j, lich_hoc['Time']] != 0:
                # Giảm số người nói đi 1
                lich_nguoi_noi.iat[j, lich_hoc['Time']] -= 1
                # Lưu ngày cũ
                lich_old = lich.iat[j, lich_hoc['Time']]
                lich.iat[j, lich_hoc['Time']] += ten_lop + ' '
                # Quay lui lớp tiếp theo
                if xep_lich_bt(lop + 1):
                    return True

                lich.iat[j, lich_hoc['Time']] = lich_old
                lich_nguoi_noi.iat[j, lich_hoc['Time']] += 1
    return False


def xep_lich():
    # Tạo một DataFrame lich là một lich_nguoi_noi với data rống
    for i in lich.iloc[:, 1:]:
        for j in range(len(lich.iloc[:, 1:][i])):
            lich[i][j] = ''
    return xep_lich_bt(0)


print("Xếp lịch thành công" if xep_lich() else "Xếp lịch không thành công")

lich.to_excel('output.xlsx', index=False)