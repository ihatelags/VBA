Dim xl
Dim xlBook
  
set xl = createobject("Excel.Application")

xl.Application.Visible = True
xl.DisplayAlerts = False
  
Set xlBook = xl.Workbooks.Open("K:\S&CM ÄËß ÎÒÄÅËÀ ÏÐÎÄÀÆ\09 Ìåð÷åíäàéçèíã\ÇÀßÂÊÈ íà ïîäêëþ÷åíèå\Äàííûå\CRM data.xlsx", 0, False)

current_month = Month(Now)

Dim filter_array()
x = current_month - 2
n = 11
ReDim filter_array(n)
For i = 0 To n
    If x - i > 0 Then
        filter_array(i) = "[Date].[by Week].[W Month].&[" & Year(Now) & right("00" & x - i, 2) & "]"
    Else
        filter_array(i) = "[Date].[by Week].[W Month].&[" & Year(Now) - 1 & right("00" & 12 + x - i, 2) & "]"
    End If
Next

'ActiveWorkbook.Connections("SWE_OLAP_Beiersdorf_STD SalesWorks").Delete

'ActiveWorkbook.Connections.Add "SWE_OLAP_Beiersdorf_STD SalesWorks", _
'    "Description", _
'    Array( _
'    "OLEDB;Provider=MSOLAP.8;Password=skVw&*73XO;Persist Security Info=True;Locale Identifier=1049;User ID=DATACENTER\akonkov;Initial Catalog=SWE_OLAP_Beiersdorf_" _
'    , _
'    "STD;Data Source=https://beiersdorf.datacenter.ssbs.com.ua/olap/msmdpump.dll;Location=https://beiersdorf.datacenter.ssbs.com.ua/o" _
'    , _
'    "lap/msmdpump.dll;MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;Update Isolation Level=2" _
'    ), Array("SalesWorks"), xlCmdCube

xl.Sheets(1).PivotTables("PivotTable1").PivotFields("[Date].[by Week].[W Year]").VisibleItemsList = Array("")
xl.Sheets(1).PivotTables("PivotTable1").PivotFields("[Date].[by Week].[W Month]").CubeField.IncludeNewItemsInFilter = False
xl.Sheets(1).PivotTables("PivotTable1").PivotFields("[Date].[by Week].[W Month]").VisibleItemsList = Array("[Date].[by Week].[W Month].&[" & Year(Now) & right("00" & x, 2) & "]")
xl.Sheets(1).PivotTables("PivotTable1").PivotFields("[Date].[by Week].[W Month]").VisibleItemsList = filter_array
xl.Sheets(1).PivotTables("PivotTable1").PivotFields("[Date].[by Week].[W Year]").PivotItems("[Date].[by Week].[W Year].&[2019]").DrilledDown = True
xl.Sheets(1).PivotTables("PivotTable1").PivotFields("[Date].[by Week].[W Year]").PivotItems("[Date].[by Week].[W Year].&[2020]").DrilledDown = True
xl.Sheets(1).PivotTables("PivotTable1").PivotFields("[Date].[by Week].[Week]").VisibleItemsList = Array("")
xl.Sheets(1).PivotTables("PivotTable1").PivotFields("[Date].[by Week].[Day]").VisibleItemsList = Array("")
xl.Sheets(1).PivotTables("PivotTable1").PivotFields("[Date].[by Week].[Day]").VisibleItemsList = Array("")
xl.Sheets(2).Delete
xl.Sheets.Add, xl.Sheets(1)
xl.Sheets(1).PivotTables("PivotTable1").TableRange1.Copy
xl.Sheets(2).Range("A1").PasteSpecial -4163, -4142

xlBook.save
xl.ActiveWindow.close True
xl.Quit

Set xlBook = Nothing
Set xl = Nothing

