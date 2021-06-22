from  openpyxl import  Workbook

import download

if __name__ == '__main__':
    url = input("url: ")
    title, tabs, opendoc_params = download.initial_fetch(url)
    wb = Workbook()
    for tab in tabs:
        tab_id = tab["id"]
        name = tab["name"]
        sheet_content, max_col = download.read_sheet(tab_id, opendoc_params)
        row = []
        ws = wb.create_sheet(name)
        for k, v in sheet_content.items():
            if (int(k) % max_col == 0 and k != '0'):
                ws.append(row)
                row=[]
            if '2' in v:
                row.append(v['2'][1])
            else:
                row.append("")
    empty_ws = wb["Sheet"]
    wb.remove(empty_ws)
    wb.save('%s.xlsx' % title)