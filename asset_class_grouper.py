import openpyxl
from openpyxl.utils import column_index_from_string

def main():
    workbook = openpyxl.load_workbook('CUX_新资产台账_021121.xlsx')
    sheet = workbook['CUX_新资产台账_021121']
    keywords = ['邮政专用汽车', '升降设备', '输送设备', '搬运装卸设备', '金融设备', '营业投递设备']

    # BH行对应“所属专业”，获取后去重，然后去掉None，每个专业一张表
    names = set([cell.value for cell in list(sheet.columns)[column_index_from_string('BH')-1][7:]])    
    names.remove(None)

    workbooks = dict()
    TITLE_LINE = 6
    for name in names:
        workbooks[name] = openpyxl.Workbook()
        workbooks[name].active.append([cell.value for cell in list(sheet.rows)[TITLE_LINE]])

    for row in list(sheet.rows)[7:]:
        # C列为资产类别名称
        asset_tag = row[column_index_from_string('C')-1].value
        
        # BH列为所属专业
        belong_to = row[column_index_from_string('BH')-1].value

        # F列为原值
        original = row[column_index_from_string('F')-1].value

        if asset_tag is None or belong_to is None or original is None or original <= 0:
            continue

        asset_tag = asset_tag.strip()
        belong_to = belong_to.strip()

        found = False
        for keyword in keywords:
            if keyword in asset_tag:
                found = True
                break
        
        if found:
            active_sheet = workbooks[belong_to].active
            active_sheet.append([cell.value for cell in row])

    for n, wb in workbooks.items():
        wb.save('{}.xlsx'.format(n))

if __name__ == '__main__':
    main()
        


