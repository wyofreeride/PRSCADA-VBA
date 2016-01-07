Public Function iH_ePodData(sDate As String, sTag As String, ByRef sResults() As String, iRecords As Integer) As Integer
Dim Conn As New ADODB.Connection
Dim RecSet As New ADODB.Recordset
Dim strConn As String
Dim StartTime As Double
Dim EndTime As Double
Dim sSqlQry As String

On Error GoTo ErrorHandler

Debug.Print

strConn = "Provider=ihOLEDB.iHistorian.1;Data Source=iHistorian;User Id=;Password="

' open connection
StartTime = Timer
Conn.ConnectionString = strConn
Conn.Open
EndTime = Timer
Debug.Print "Connect took " & EndTime - StartTime & " S"

StartTime = Timer

'Create the Qry
'sDate = "1/01/2016 00:00:00"
sSql = "SELECT timestamp, value, quality FROM ihRawdata  " & _
"WHERE ( tagname LIKE '" & sTag & "' " & _
"AND samplingmode=calculated " & _
"AND CalculationMode=Average " & _
"AND intervalmilliseconds=15m " & _
"AND timestamp>='" & sDate & "' " & _
"AND timestamp<='" & DateAdd("d", 1, sDate) & "' ) "

'Debug.Print sSql
RecSet.Open sSql, Conn
EndTime = Timer
Debug.Print "Qry took " & EndTime - StartTime & " S"


' print the output
iRecords = 0
StartTime = Timer
ReDim sResults(97, 2)
Do While Not RecSet.EOF
   sResults(iRecords, 0) = RecSet.Fields(0)
   sResults(iRecords, 1) = RecSet.Fields(1)
   sResults(iRecords, 2) = RecSet.Fields(2)
   RecSet.MoveNext
   iRecords = iRecords + 1
Loop
EndTime = Timer
Debug.Print "Record Read took " & EndTime - StartTime & " S"


' print a count of rows
Debug.Print iRecords & " Rows returned"
' close the recordset
RecSet.Close
Set RecSet = Nothing

' close the connection
Conn.Close
Exit Function

ErrorHandler:
Debug.Print Err.Description
Resume Next
End Function
