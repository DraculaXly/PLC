'DTPicker1 change vbs
'Get the picked datetime
Sub Change(ByVal Item)     
    Dim A,B
    Set A = HMIRuntime.Screens("ZZZ Zaoyan").ScreenItems("DTPicker1")
    Set B = HMIRuntime.Tags("strBeginTime")
    B.Read 
    B.Value = A.Value
    B.Write 
End Sub


'DTPicker2 change vbs
'Get the picked datetime
Sub Change(ByVal Item)     
    Dim A,B
    Set A = HMIRuntime.Screens("ZZZ Zaoyan").ScreenItems("DTPicker2")
    Set B = HMIRuntime.Tags("strEndTime")
    B.Read 
    B.Value = A.Value
    B.Write 
End Sub
   

'Excute the query function and write satisfied data to Excel sheet
Sub OnClick(ByVal Item)                            
    On Error Resume Next

    Dim sPro,sDsn,sSer,sCon,conn,sSql,oRs,oCom
    Dim tagDSNName
    Dim m,i
    Dim LocalBeginTime,LocalEndTime,UTCBeginTime, UTCEndTime,sVal
    Dim objExcelApp,objExcelBook,objExcelSheet,sheetname
    Item.Enabled = False
 
    sheetname = "Sheet1"

    'Use template
    Set objExcelApp = CreateObject("Excel.Application")
        objExcelApp.Visible = False
        objExcelApp.Workbooks.Open "C:\Test.xlsx"
        objExcelApp.Worksheets(sheetname).Activate
    
    'Get the begin and end time
    'Get the interval time
    Set tagDSNName = HMIRuntime.Tags("@DatasourceNameRT")
        tagDSNName.Read 
    Set LocalBeginTime = HMIRuntime.Tags("strBeginTime")
        LocalBeginTime.Read 
    Set LocalEndTime = HMIRuntime.Tags("strEndTime")
        LocalEndTime.Read
        UTCBeginTime = DateAdd("h", -8, LocalBeginTime.Value)
        UTCEndTime= DateAdd("h", -8, LocalEndTime.Value)
        UTCBeginTime = Year(UTCBeginTime) & "-" & Month(UTCBeginTime) & "-" & Day(UTCBeginTime) & " " & Hour(UTCBeginTime) & ":" & Minute(UTCBeginTime) & ":" & Second(UTCBeginTime)
        UTCEndTime = Year(UTCEndTime) & "-" & Month(UTCEndTime) & "-" & Day(UTCEndTime) & " " & Hour(UTCEndTime) & ":" & Minute(UTCEndTime) & ":" & Second(UTCEndTime)
        HMIRuntime.Trace "UTC Begin Time: " & UTCBeginTime & vbCrLf
        HMIRuntime.Trace "UTC end Time: " & UTCEndTime & vbCrLf
    Set sVal = HMIRuntime.Tags("sVal")
        sVal.Read
        
    'Establish connection and do the query
    sPro = "Provider=WinCCOLEDBProvider.1;"
    sDsn = "Catalog=" & tagDSNName.Value & ";"
    sSer = "Data Source=.\WinCC"
    sCon = sPro + sDsn + sSer
    Set conn = CreateObject("ADODB.Connection")
        conn.ConnectionString = sCon
        conn.CursorLocation = 3
        conn.Open
    
    'User-defined SQL
    'sSql = "Tag:R,'PVArchive\NewTag','" & UTCBeginTime & "','" & UTCEndTime & "'"
    'sSql = "Tag:R,'PVArchive\NewTag','0000-00-00 00:10:00.000','0000-00-00 00:00:00.000'"
    'sSql = "Tag:R,'PVArchive\NewTag';'PVArchive\NewTag_1','" & UTCBeginTime & "','" & UTCEndTime & "',"
    'sSql = "Tag:R,'PVArchive\NewTag','" & UTCBeginTime & "','" & UTCEndTime & "','order by Timestamp DESC','TimeStep=" & sVal.Value & ",1"
    sSql = "Tag:R,'SystemArchive\P1206A/Motor.Start#Value','" & UTCBeginTime & "','" & UTCEndTime & "',"
    sSql = sSql + "'order by Timestamp ASC','TimeStep=" & sVal.Value & ",1'"
    
    Set oRs = CreateObject("ADODB.Recordset")
    Set oCom = CreateObject("ADODB.Command")
        oCom.CommandType = 1
    Set oCom.ActiveConnection = conn
        oCom.CommandText = sSql
    
    'Get the date and put into the worksheet
    Set oRs = oCom.Execute
        m = oRs.RecordCount
    If (m > 0) Then
            objExcelApp.Worksheets(sheetname).cells(2,1).value = oRs.Fields(0).Name
            objExcelApp.Worksheets(sheetname).cells(2,2).value = oRs.Fields(1).Name
            objExcelApp.Worksheets(sheetname).cells(2,3).value = oRs.Fields(2).Name
            objExcelApp.Worksheets(sheetname).cells(2,4).value = oRs.Fields(3).Name
            objExcelApp.Worksheets(sheetname).cells(2,5).value = oRs.Fields(4).Name
            oRs.MoveFirst  
            i = 3  
        Do While Not oRs.EOF                              
            objExcelApp.Worksheets(sheetname).cells(i,1).value = oRs.Fields(0).Value
            'objExcelApp.Worksheets(sheetname).cells(i,2).value = GetLocalDate(oRs.Fields(1).Value) 
            objExcelApp.Worksheets(sheetname).cells(i,2).value = oRs.Fields(1).Value
            objExcelApp.Worksheets(sheetname).cells(i,3).value = oRs.Fields(2).Value
            objExcelApp.Worksheets(sheetname).cells(i,4).value = oRs.Fields(3).Value
            objExcelApp.Worksheets(sheetname).cells(i,5).value = oRs.Fields(4).Value
            oRs.MoveNext
            i = i + 1
        Loop
        oRs.Close
    Else
        MsgBox "No data found!"
        item.Enabled = True
        Set oRs = Nothing
        conn.Close
        Set conn = Nothing
        objExcelApp.Workbooks.Close
        objExcelApp.Quit
        Set objExcelApp = Nothing
        Exit Sub
    End If

    'Release the source
    Set oRs = Nothing
        conn.Close
    Set conn = Nothing
    
    'Write date to Excel
    Dim patch,filename
    filename = CStr(Year(Now)) & CStr(Month(Now)) & CStr(Day(Now)) & CStr(Hour(Now)) + CStr(Minute(Now)) & CStr(Second(Now))
    patch = "C:\" & filename & ".xlsx"	
    objExcelApp.ActiveWorkbook.SaveAs patch
    objExcelApp.Workbooks.Close
    objExcelApp.Quit
    Set objExcelApp = Nothing
    MsgBox "Well done!"
    Item.Enabled = True
End Sub