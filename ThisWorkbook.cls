Private Sub Workbook_Open()
    alertTime = Now + TimeValue("00:00:01")
    Application.OnTime alertTime, "EventMacro"
End Sub
