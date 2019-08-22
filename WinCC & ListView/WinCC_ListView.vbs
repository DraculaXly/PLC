'Main Button function
'Call several function models
Sub OnClick(ByVal Item)
    Item.Enabled = False                    
    'Define database connection
    Dim conn
    Set conn = CreateObject("ADODB.Connection")

    'Define database sql
    Dim sSQL
    sSQL = "Tag:R,'SystemArchive\P1206A/Motor.Start#Value','2019-08-01 00:00:00.000','2019-08-22 00:00:00.000'"
    'sSQL:Modify by yourself

    'Set connection
    Dim oRs
    Set oRs = WinCCDataSourceAccess(conn, sSQL)

    'Define ListView Object And clear the content
    Dim ListViewT
    Set ListViewT = HMIRuntime.Screens("ZZZ Zaoyan").ScreenItems("Control1")
    ListViewT.ListItems.clear

    'Add column header
    AddColumnHeader oRs, ListViewT

    'Fill data
    FillListView oRs, ListViewT

    'Release connection source
    WinCCDataSourceClose oRs, conn
    Item.Enabled = True
End Sub


'Sub function 1:WinCCDataSourceAccess
'Access WinCC data and return data
Function WinCCDataSourceAccess(connObj, pSQL)
    On Error Resume Next

    'Read the local WinCC datasource name
    Dim DataSourceNameRT, DataConnectionName
    Set DataSourceNameRT = HMIRuntime.Tags("@DatasourceNameRT")
    DataSourceNameRT.Read 

    'Define database connection string
    Dim sPro, sDsn, sSer
    sPro = "Provider=WinCCOLEDBProvider.1;"
    sDsn = "Catalog=" & DatasourceNameRT.Value & ";"
    sSer = "Data Source=.\WinCC"
    DataConnectionName = sPro + sDsn + sSer

    'Define sql
    Dim sSQL
    sSQL = pSQL

    'Establish connection
    Dim oRs, oCom, conn
    Set conn = CreateObject("ADODB.Connection")
    Set conn = connObj
    conn.ConnectionString = DataConnectionName
    conn.CursorLocation = 3
    conn.Open

    'Create query command
    Set oRs = CreateObject("ADODB.Recordset")
    Set oCom = CreateObject("ADODB.Command")
    oCom.CommandType = 1
    Set oCom.ActiveConnection = conn
    oCom.CommandText = sSQL

    'Excute the query
    Set oRs = oCom.Execute

    'Return the result
    Set WinCCDataSourceAccess = oRs
    If err.number <> 0 Then 
    MsgBox "error Code" & Err.Number & "Source:" & Err.Source & "error description" & Err.Description
    err.clear
    End If 

    On Error Goto 0
End Function


'Sub function 2:AddColumnHeader
'ListView add column headers by using WinCC datasource column headers
Function AddColumnHeader(pRecordset, pListView)
    On Error Resume Next

    'Create record
    Dim oRs,columnCount
    Set oRs = CreateObject("ADODB.Recordset")
    Set oRs = pRecordset

    'Get columns
    columnCount = oRs.Fields.Count

    'Define ListView object
    Dim ListViewT
    Set ListViewT = pListView
    ListViewT.View = 3
    'Notice!
    ListViewT.ColumnHeaders.Clear

    'Use database column names to create ListView headers
    Dim AddColumnIndex
    For AddColumnIndex = 0 To columnCount - 1
    ListViewT.ColumnHeaders.Add ,,CStr(oRs.Fields(AddColumnIndex).Name)
    Next

    'Error message
    If err.number <> 0 Then
    MsgBox "AddColumnHeader函数报错，Source:"  & Err.Source & vbCr & "Error description:" & Err.Description
    err.clear
    End If

    On Error Goto 0
End Function 


'Sub function 3:FillListView
'Fill ListView items by using the return query data
Function FillListView(pRecordset,pListView)
    On Error Resume Next

    'Get data
    Dim recordsCount,oRs
    Set oRs = CreateObject("ADODB.Recordset")
    Set oRs = pRecordset
    recordsCount = oRs.RecordCount

    'Locate to first record
    If (recordsCount > 0) Then
    oRs.MoveFirst

    'Define max lines
    Dim maxLine,n
    maxLine = 10
    n = 0

    'Fill data
    Do While (Not oRs.EOF And n < maxLine)
    n = n + 1

    Dim oItem,ListViewT
    Set ListViewT = pListView
    Set oItem = ListViewT.ListItems.Add()
    oItem.text = oRs.Fields(0).Value
    oItem.SubItems(1) = oRs.Fields(1).Value
    oItem.SubItems(2) = FormatNumber(oRs.Fields(2).Value, 4)
    oItem.SubItems(3) = Hex(oRs.Fields(3).Value)
    oItem.SubItems(4) = Hex(oRs.Fields(4).Value)
    oRs.MoveNext
    Loop
    End If

    'Error message
    If err.number <> 0 Then
    MsgBox  "填充FillListView函数发生错误，Source:" & Err.Source & vbCr & "Error description:" & Err.Description
    err.clear
    End If

    On Error Goto 0
End Function


'Sub function 4:WinCCDataSourceClose
'Close connection
Function WinCCDataSourceClose(pRecordset,connObj)
    'Get connection
    Dim oRs, conn
    Set oRs = pRecordset
    Set conn = connObj

    'Close connection and release source
    oRs.Close
    Set oRs = Nothing
    conn.Close
    Set conn = Nothing
End Function