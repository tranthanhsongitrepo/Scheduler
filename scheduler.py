import pandas as pd

nhom_fileName = 'nhom.xlsx'
tkb_fileName = 'TKB.xlsx'
speaker_fileName = 'Speaker.xlsx'

data_nhom = pd.read_excel(nhom_fileName)
data_TKB = pd.read_excel(tkb_fileName)
speaker_schedule = pd.read_excel(speaker_fileName)
mon = list(data_nhom.head(0))[1:]
lich = speaker_schedule.copy()


def search_class(mon, nhom):
    for i, ma_mon in enumerate(data_TKB['Subject']):
        ma_nhom = data_TKB['Group ID'][i]
        if ma_mon == mon or ma_nhom == nhom:
            return i
    return -1


def xep_lich_bt(lop):
    if lop == len(data_nhom):
        return True
    ten_lop, nhom = data_nhom.iloc[lop][0], data_nhom.iloc[lop][1:]
    for i in range(len(nhom)):
        tt = search_class(mon[i], nhom[i])
        # When is the class
        data = data_TKB.iloc[tt]

        # No of speakers
        for j in range(len(speaker_schedule['Day'])):
            # Get the number of speakers scheduled at that class's day and time
            if speaker_schedule['Day'][j] == data['Day'] and \
                    (speaker_schedule.iloc[j])[data['Time']] != 0:
                # Skip if there are no speakers
                # Else speaker found
                # Reduce available speakers
                speaker_schedule.iat[j, data['Time']] -= 1
                # Save the date
                lich_old = lich.iat[j, data['Time']]
                lich.iat[j, data['Time']] += ten_lop + ' '
                # Call recursion with the next class
                if xep_lich_bt(lop + 1):
                    return True
                # Revert changes
                lich.iat[j, data['Time']] = lich_old
                speaker_schedule.iat[j, data['Time']] += 1
    return False
    # for day, speakers in (speaker_schedule[lop['Day']]):
    #     if speakers != 0:
    #         return xep_lich(lop, )
    # return -1


def xep_lich():
    for i in lich.iloc[:, 1:]:
        for j in range(len(lich.iloc[:, 1:][i])):
            lich[i][j] = ''
    return xep_lich_bt(0)


print("Xếp lịch thành công" if xep_lich() else "Xếp lịch không thành công")
lich.to_excel('output.xlsx', index=False)

