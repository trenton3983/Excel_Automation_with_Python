import win32com.client as win32
import sys
import os
from pythoncom import com_error

win32c = win32.constants


def pivot_table(wb, ws1, pt_ws, num_test_cond, ws_name, pt_name, pt_rows, pt_filters, pt_fields):
    """
    wb = workbook1 reference
    ws1 = worksheet1
    pt_ws = pivot table worksheet number
    num_test_cond = number of unique TestCondition1 values (use for determine row locations)
    ws_name = worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_filters, pt_fields: values selected for filling the pivot tables
    """

    pt_loc = len(pt_filters) + 2  # the table begins 2 rows below the filters

    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C1', TableName=pt_name)

    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    """Sets the rows and filters of the pivot table"""

    for field_list, field_c in ((pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_c
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    """Sets the Values of the pivot table"""

    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1], field[2])

    pt_ws.PivotTables(pt_name).ShowValuesRow = False
    pt_ws.PivotTables(pt_name).ColumnGrand = False

    """Hides details under each row section - generic form"""

    row_first_chart_value = pt_loc + 1
    row_last_chart_value = row_first_chart_value + num_test_cond  # this is used to hide detail in the PT

    '''
    for x in range(row_first_chart_value, row_last_chart_value, 1):
        pt_ws.Cells(x, 1).Select()
        cell_value = (str(pt_ws.Cells(x, 1).Value)).rstrip('0').rstrip('.')
        pt_ws.PivotTables(pt_name).PivotFields('TestCondition1').PivotItems(f'{cell_value}').ShowDetail = False
    '''

    return row_first_chart_value, row_last_chart_value


def run_excel():
    filename = os.path.join(f_path, f_name)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    try:
        wb = excel.Workbooks.Open(filename)
    except com_error as e:
        if e.excepinfo[5] == -2146827284:
            print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
        else:
            raise e
        sys.exit(1)

    ws1 = wb.Sheets('Dual_Edge_Power')

    """Determine number of unique TestCondition1 values"""
    last_r = ws1.UsedRange.Rows.Count
    column_values = set()
    for x in range(2, last_r + 1, 1):
        column_values.add(ws1.Range(f'N{x}').Value)
    unique_test_conditions = len(column_values)

    worksheet_name2 = 'Average_of_Max'
    wb.Sheets.Add().Name = worksheet_name2
    ws2 = wb.Sheets(worksheet_name2)
    row_first_chart_value2, _ =\
        pivot_table(wb, ws1, ws2, unique_test_conditions,
                    ws_name=worksheet_name2,
                    pt_name='PivotTable1',
                    pt_rows=['TestCondition1', 'TestCondition2'],
                    pt_filters=['Aux(V)', 'Main(V)', 'Temp(C)'],
                    pt_fields=[['Max_I_TC1(A)', 'Avgerage of Max_I_TC1(A)', win32c.xlAverage],
                               ['Max_I_TC2(A)', 'Avgerage of Max_I_TC2(A)', win32c.xlAverage],
                               ['Total_Pwr(W)', 'Avgerage of Total_Pwr(W)', win32c.xlAverage]])

    worksheet_name3 = 'Max_of_Max'
    wb.Sheets.Add().Name = worksheet_name3
    ws3 = wb.Sheets(worksheet_name3)
    row_first_chart_value3, _ =\
        pivot_table(wb, ws1, ws3, unique_test_conditions,
                    ws_name=worksheet_name3,
                    pt_name='PivotTable2',
                    pt_rows=['TestCondition1', 'TestCondition2'],
                    pt_filters=[],
                    pt_fields=[['Max_I_TC1(A)', 'Max of Max_I_TC1(A)', win32c.xlMax],
                               ['Max_I_TC2(A)', 'Max of Max_I_TC2(A)', win32c.xlMax],
                               ['Total_Pwr(W)', 'Max of Total_Pwr(W)', win32c.xlMax]])

    pt_row_positions = [row_first_chart_value2, row_first_chart_value3]
    ev_report_table(excel, wb, ws1, ws3, pt_row_positions)


def ev_report_table(excel, wb, ws1, ws3, row_list):
    """
    This method creats the report table
    :param excel: Excel object
    :param wb: Excel workbook object
    :param ws1: Excel worksheet objects
    :param ws3: Excel worksheet objects
    :param row_list: List of fist/last data cells for each pivot table
    :return:
    """

    used = ws3.UsedRange  # create the UsedRange object
    ws3_first_data_row = row_list[1]
    ws3_last_data_row = used.Row + used.Rows.Count - 1  # used.Row -> returns first row used, Rows.Count -> rows used

    """Create Report Table worksheet"""
    worksheet_ev_table = 'EV_Report_Table'
    wb.Sheets.Add().Name = worksheet_ev_table
    ws4 = wb.Sheets(worksheet_ev_table)

    """Table Headers"""
    table_headers = ['Test Point', 'Typical I (A)', 'Maximum I (A)', 'Max Total Power (W)', 'Spec', 'Status']

    for x, col_header in enumerate(table_headers):
        cell = ws4.Cells(2, x + 2)
        cell.Value = col_header
        cell.Select()
        excel.Selection.Font.Bold = True

    """Create a list of Test Point(s) (table row headers) from ws3 (Max of Max)"""
    test_points = []
    for x in range(ws3_first_data_row, ws3_last_data_row + 1):
        test_points.append(ws3.Cells(x, 1).Value)

    """Write Table row header formulas"""
    for x in range(ws3_first_data_row, ws3_last_data_row + 1):
        ws4.Cells(x, 2).Value = f'=Max_of_Max!A{x}'

    """Set column filters and copy specification (tolerance) value"""
    def get_tol_address(i, test_condition):
        """
        Set column filters and copy specification (tolerance) value
        :param i: test condition index of the list being fed into this method
        :param test_condition: this is the parameter being used to set the Excel filter
        :return: returns a Excel cell address for the tolerance data
        """

        ws1.Activate()  # set the worksheet to be filtered as active

        if i % 2 == 0:
            col_select = [14, 'Q1']
        else:
            col_select = [19, 'V1']

        ws1.UsedRange.AutoFilter(col_select[0])  # remove column filter (set to all)
        ws1.UsedRange.AutoFilter(col_select[0], test_condition)  # set specific Column filter
        ws1.Range(col_select[1]).End(win32c.xlDown).Select()  # select last cell in the I_Tol column (Q or V)
        cell_address = excel.Selection.Address  # get the selected cell address
        ws1.Range(col_select[1]).End(win32c.xlUp).Select()  # reset to rol 1
        ws1.UsedRange.AutoFilter(col_select[0])  # remove column filter (set to all)

        return cell_address

    """Write Maximum I (A), Typical I (A), and Max Total Power (W) cell equations"""
    for x, row_header in enumerate(test_points):
        tol_address = get_tol_address(x, row_header)
        if x % 2 == 0:
            ws4.Cells(x + 3, 4).Value =\
                f'=GETPIVOTDATA("Max of Max_I_TC1(A)",Max_of_Max!R2C1,"TestCondition1","{row_header}")'
            ws4.Cells(x + 3, 3).Value =\
                f'=GETPIVOTDATA("Avgerage of Max_I_TC1(A)",Average_of_Max!R5C1,"TestCondition1","{row_header}")'
            ws4.Cells(x + 3, 5).Value =\
                f'=GETPIVOTDATA("Max of Total_Pwr(W)",Max_of_Max!R2C1,"TestCondition1","{row_header}")'
        else:
            ws4.Cells(x + 3, 4).Value =\
                f'=GETPIVOTDATA("Max of Max_I_TC2(A)",Max_of_Max!R2C1,"TestCondition1","{test_points[x - 1]}")'
            ws4.Cells(x + 3, 3).Value =\
                f'=GETPIVOTDATA("Avgerage of Max_I_TC2(A)",Average_of_Max!R5C1,"TestCondition1","{test_points[x - 1]}")'

        ws4.Cells(x + 3, 6).Value = f'=Dual_Edge_Power!{tol_address}'
        ws4.Cells(x + 3, 7).Value = f'=if(D{x + 3}>F{x + 3},"Fail","Pass")'

    ws4.Activate()  # set the worksheet with the EV report table as active

    """Format the Table"""
    """Merge and Center Max Total Power (W) column"""
    for x in range(ws3_first_data_row, ws3_last_data_row + 1, 2):
        ws4.Range(f'E{x}:E{x + 1}').Merge()
        ws4.Range(f'E{x}:E{x + 1}').HorizontalAlignment = win32c.xlCenter
        ws4.Range(f'E{x}:E{x + 1}').VerticalAlignment = win32c.xlCenter

    last_table_row = len(test_points) + 2  # because data begins on row 2

    """Format Numbers"""
    ws4.Range(f'C3:F{last_table_row}').Select()
    excel.Selection.NumberFormat = "0.000"

    """Autofit Width"""
    ws4.Columns('B:G').EntireColumn.AutoFit()

    """Conditional Formatting"""

    'Dictionary values: Condition, Font Color, Fill Color'
    conditional_formatting = {'Pass': [-16752384, 13561798],
                              'Fail': [-16383844, 13551615]}

    for k, v in conditional_formatting.items():
        ws4.Range(f'G3:G{last_table_row}').Select()
        excel.Selection.FormatConditions.Add(Type=win32c.xlTextString, TextOperator=win32c.xlContains, String=k)
        excel.Selection.FormatConditions(excel.Selection.FormatConditions.Count).SetFirstPriority()
        excel.Selection.FormatConditions(1).Font.Color = v[0]
        excel.Selection.FormatConditions(1).Interior.PatternColorIndex = win32c.xlAutomatic
        excel.Selection.FormatConditions(1).Interior.Color = v[1]

    """Set the Table Boarders"""
    border_types = {win32c.xlEdgeLeft: win32c.xlMedium, win32c.xlEdgeTop: win32c.xlMedium,
                    win32c.xlEdgeBottom: win32c.xlMedium,
                    win32c.xlEdgeRight: win32c.xlMedium, win32c.xlInsideVertical: win32c.xlThin,
                    win32c.xlInsideHorizontal: win32c.xlThin}

    '''Boarders for entire table'''
    for k, v in border_types.items():
        ws4.Range(f'B2:G{last_table_row}').Select()
        excel.Selection.Borders(k).Weight = v

    '''Boarders for top row'''
    for k, v in border_types.items():
        ws4.Range(f'B2:G2').Select()
        excel.Selection.Borders(k).Weight = win32c.xlMedium


if __name__ == "__main__":

    """Test Data"""

    f_name = 'pivot_tables.xlsx'
    f_path = r'C:\PythonProjects\Excel_Automation_with_Python\Pivot_Tables_and_Table'

    run_excel()
