from easy_excel_tool import Excel

if __name__ == '__main__':
    excel = Excel('./test.xlsx')
    # excel.create_excel(inplace=True)
    # excel.add_sheet(sheet_name=[
    #     'donation_information',
    # ])
    data = [
        {'a': 1, 'b': 'DG文漢三 \x08🌹', 'c': '88'},
        {'a': 7, 'b': 9, 'c': 66},
    ]
    excel.write_excel(
        data,
        mode='w+',
        sheet_name='Sheet1',
        fill_column={'fill_column': 'fill'},
        header=True,
        inplace=False,
    )
