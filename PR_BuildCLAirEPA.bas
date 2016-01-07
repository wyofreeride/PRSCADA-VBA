Public Function PR_BuildClAirEPA(sReportDate As String, iRig As Integer, iEngine As Integer, Optional bBackground As Boolean = False) As Integer
'Create an EPA and HSE log file from the ePod log files
'  Create the blank file
'  Generate the required Header and date time fields
'  Open the ePod log file
'  run through the file and extract data as close as possible to the 15 minute interval
'  instert the raw data from the ePod log into the HSE file
'  After all data is inserted, calculate the additional fields (control status etc)

'Verify the Monico Log files
'Returns and Integer Code for each file indicating the quality of the file
'0 = File Created with no Issues
'x = Files had error (Binary code)

Dim ExcelApp As Excel.Application
Dim Report As Excel.Workbook
Dim ePodReport As Excel.Workbook

Dim sFilePath As String
Dim sFileName As String

Dim sePodFilePath As String
Dim sePodFileName As String

Dim sLogFileName As String

Dim sRigName(4) As String
Dim sEngineSerial(4, 4) As String

Dim iLog As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim kStart As Integer

Dim sTemp1 As String
Dim sTemp2 As String
Dim sTemp3 As String
Dim sATemp() As String

Dim sePodPumpOutput As String
Dim sePodBoostPressure As String
Dim sePodExhTemp As String

Dim sTemp As String
Dim sDateCheck As String
Dim sStartDate As String

' Define the constants used for file system objects
Const ForReading = 1
Const ForAppending = 8

Set fs = CreateObject("Scripting.FileSystemObject")

On Error Resume Next

sRigName(1) = "Unit-116"
sRigName(2) = "Unit-124"
sRigName(3) = "Unit-125"

sEngineSerial(1, 1) = "LLA00788"
sEngineSerial(1, 2) = "LLA00987"
sEngineSerial(1, 3) = "LLA00786"
sEngineSerial(2, 1) = "LLA02667"
sEngineSerial(2, 2) = "LLA02675"
sEngineSerial(2, 3) = "LLA02672"
sEngineSerial(3, 1) = "LLA02891"
sEngineSerial(3, 2) = "LLA02896"
sEngineSerial(3, 3) = "LLA02902"

'Create the source and destination paths
sFilePath = "\\PRSCADA\d_SA\EmissionsData\MonicoToProcessTEST2\"
sePodFilePath = "\\PRSCADA\D_SA\EmissionsData\ePodVerifiedTEST2\"

'Create the log files
sLogFileName = "\\PRSCADA\D_SA\EmissionsData\MonicoLogsTEST2\PR_BuildCLAirEPA_Log_" & Format(Now(), "mmddyy") & ".txt"

iLog = FreeFile
Open sLogFileName For Append As #iLog

On Error GoTo ErrorHandler

If Not bBackground Then
   Load frmStatus
   frmStatus.Show
   AddLog "Starting Clean Air EPA File Creation at " & Now()
   AddLog ""
End If
Print #iLog, "Starting Clean Air EPA File Creation at " & Now()
Debug.Print "Starting Clean Air EPA File Creation at " & Now()

'Start Excel
Set ExcelApp = CreateObject("Excel.Application")

'Generate the Report File Name
sFileName = "Pinedale-" & sRigName(iRig) & "-" & iEngine & "-" & Format(sReportDate, "yyyymmdd") & "0000" & "-CLAirEPA.csv"

'Check for pre-existing file, move it to the previous folder if there is one
If FileExists(sFilePath & sFileName) Then
   'Make sure previous folder has been created
   If Not FolderExists(sFilePath & "Previous\") Then
      MkDir sFilePath & "Previous\"
   End If
   'Check if previous file already exists, delete it if so
   If FileExists(sFilePath & "Previous\" & sFileName) Then
      Debug.Print "Previous file deleted"
      Kill sFilePath & "Previous\" & sFileName
   End If
   'Rename the file (moving it to the previous directoy
   Debug.Print "Moved existing file to previous"
   Print #iLog, "Moved existing file to Previous\ folder"
   Name sFilePath & sFileName As sFilePath & "Previous\" & sFileName
End If
   



'Open and Verify the file
If Not bBackground Then AddLog "Building " & sFileName
Print #iLog, "Building " & sFileName
Debug.Print "Building " & sFileName

'Open the Existing Report file
If Not bBackground Then AddLog "Building " & sFileName
Print #iLog, "Building " & sFileName
Set Report = ExcelApp.Workbooks.Add
With Report.Sheets(1)
   
   '**********************************************************************
   'Insert Header Row
   '**********************************************************************
   'A1 Date
   'B1 Time
   'C1 DateTime
   'D1 SerialNumber
   'E1 UnitNumber
   'F1 CA_BoostPressure
   'G1 CA_PumpOutput
   'H1 CA_CatalystInletTemp
   'I1 EngineRunStatus
   'J1 EngineControlledStatus
   'K1 Low Boost Urea Cutoff Setpoint
   .Range("A1").value = "Date"
   .Range("B1").value = "Time"
   .Range("C1").value = "DateTime"
   .Range("D1").value = "SerialNumber"
   .Range("E1").value = "Unit Number"
   .Range("F1").value = "CA_BoostPressure"
   .Range("G1").value = "CA_PumpOutput"
   .Range("H1").value = "CA_CatalystInletTemp"
   .Range("I1").value = "EngineRunStatus"
   .Range("J1").value = "EngineControlledStatus"
   .Range("K1").value = "Low Boost Cutoff"
    
   '**********************************************************************
   'Insert Timestamps
   'Insert Serial Numbers
   'Insert Unit Number
   '**********************************************************************
   sStartDate = sReportDate & " 12:00 AM"
   sTemp1 = "QEP " & Replace(sRigName(iRig), "-", " ") & " Engine " & iEngine
   j = 0
   For i = 2 To 97
      'calc what the next timestamp should be
      sDateCheck = Format(DateAdd("n", 15 * j, sStartDate), "m/d/yyyy h:mm")
      .Cells(i, 1).value = Format(sDateCheck, "m/d/yyyy")
      .Cells(i, 2).value = Format(sDateCheck, "h:mm:SS")
      .Cells(i, 3).value = sDateCheck
      .Cells(i, 4).value = sEngineSerial(iRig, iEngine)
      'Insert Unit Number
      .Cells(i, 5).value = sTemp1
      'Insert Low Boost Cutoff
      .Cells(i, 11).value = 2.5
      j = j + 1
   Next i
   If Not bBackground Then AddLog "  Insert Timestamps, S/N's and Unit Numbers"
   
   '**********************************************************************
   'Save the file
   '**********************************************************************
   ExcelApp.DisplayAlerts = False
   Report.SaveAs sFilePath & sFileName, xlCSV
   ExcelApp.DisplayAlerts = True
   If Not bBackground Then AddLog "File Saved to preserve corrections"

   '**********************************************************************
   'Fill in missing data from the ePod data log if possible
   '**********************************************************************
   'get the epod file name
   sATemp = Split(sFileName, "-")
   sePodFileName = "Pinedale-" & sRigName(iRig) & "-" & iEngine & "-" & sATemp(4) & "-ePod.csv"
   
   'Check for ePod File to exist
   If Not bBackground Then AddLog "  Checking for ePod file"

   If FileExists(sePodFilePath & sePodFileName) Then
      Set ePodReport = ExcelApp.Workbooks.Open(sePodFilePath & sePodFileName)
      kStart = 4
      For j = 2 To 97
         sDateCheck = Format(Report.Sheets(1).Cells(j, 2).value, "h:mm:ss")
         'Data is missing, gather everything from the ePodFile for this timestamp
         'First, find the row with a close timestamp
         If Not bBackground Then AddLog "   Searching for Missing Data in ePod File - " & sDateCheck
         For k = kStart To 18000
            
            If ePodReport.Sheets(1).Cells(k, 2).value <> "" Then
               'Insert a log update every 100 lines searched
               If (k Mod 100) = 0 Then If Not bBackground Then AddLog "   Checking line " & k
               
               If (Report.Sheets(1).Cells(j, 2).value - ePodReport.Sheets(1).Cells(kStart, 2).value) < (-15 / 1440) Then
                 'ExcelApp.Visible = True
                 'ExcelApp.Visible = False
                 If Not bBackground Then AddLog "  ePod file doesn't start until after the missing record"
                 Exit For
               End If
               
               'See if we are within 14 minutes of the requested time
               If (ePodReport.Sheets(1).Cells(k, 2).value - Report.Sheets(1).Cells(j, 2).value >= 0) And (ePodReport.Sheets(1).Cells(k, 2).value - Report.Sheets(1).Cells(j, 2).value < (15 / 1440)) Then
                  'save this row to start again on for next missing record
                  kStart = k
                  
                  'Clear the variables
                  sePodPumpOutput = ""
                  sePodBoostPressure = ""
                  sePodExhTemp = ""
                  sTemp = ""

                  'Read values from ePod report row
                  sePodPumpOutput = ePodReport.Sheets(1).Cells(k, 8).value
                  sePodBoostPressure = ePodReport.Sheets(1).Cells(k, 9).value
                  sePodExhTemp = ePodReport.Sheets(1).Cells(k, 12).value
                  
                  'get the timestamp and format it
                  'Insert ePod file timestamp for debugging
                  'sTemp = Format(ePodReport.Sheets(1).Cells(k, 2).value, "h:mm:ss")
                  'Report.Sheets(1).Cells(j, 14).value = sTemp
                  'If Not bBackground Then addlog "  Found Missing Data in ePod File at line " & k
            
                  'Insert ePod file device state for debugging
                  sTemp = ePodReport.Sheets(1).Cells(k, 3).value
                  Report.Sheets(1).Cells(j, 13).value = sTemp
                  
                  'Insert data into EPA Report
                  Report.Sheets(1).Cells(j, 6).value = sePodBoostPressure
                  Report.Sheets(1).Cells(j, 7).value = sePodPumpOutput
                  Report.Sheets(1).Cells(j, 8).value = sePodExhTemp

                  Exit For
               ElseIf (ePodReport.Sheets(1).Cells(k, 2).value - Report.Sheets(1).Cells(j, 2).value) > (15 / 1440) Then
                  'Gap in Data?
                  If Not bBackground Then AddLog "  No data for time frame - " & sDateCheck
                  Debug.Print "No data for time frame - " & sDateCheck
                  Exit For
               End If
            Else
               'Blank line reached, see if next two are blank, then exit for loop
               If ePodReport.Sheets(1).Cells(k + 1, 2).value = "" And ePodReport.Sheets(1).Cells(k + 2, 2).value = "" Then Exit For
            End If
         Next k
      Next j
      
      'close the epod file
      ePodReport.Close False
   Else
      If Not bBackground Then AddLog "  Unable to Find ePod file to pull from"
      Print #iLog, "  Unable to Find ePod file to pull from"
      .Cells(2, 12).value = "Unable to locate ePod file as data source"
      PR_BuildClAirEPA = 1
   End If
   
   '**********************************************************************
   'Insert Run Status based on Boost
   '**********************************************************************
   For i = 2 To 97
      'Check that a value is available for boost pressure, skip if not
      If Trim(.Cells(i, 6).value) <> "" Then
         If Trim(.Cells(i, 6).value) > 0.6 Then
            .Cells(i, 9).value = "RUNNING"
         Else
            .Cells(i, 9).value = "STOPPED"
         End If
      End If
   Next i
   
   '**********************************************************************
   'Insert Control Status Calculation
   '**********************************************************************
   For i = 2 To 97
      'Check for a record
      If Trim(.Cells(i, 1).value) <> "" And _
         Trim(.Cells(i, 6).value) <> "" And _
         Trim(.Cells(i, 7).value) <> "" And _
         Trim(.Cells(i, 8).value) <> "" And _
         Trim(.Cells(i, 9).value) <> "" Then
         'Figure out what the control status should be
         If Trim(.Cells(i, 9).value) = "RUNNING" Then
            If Val(Trim(.Cells(i, 8).value)) >= 270 Then
               If Val(Trim(.Cells(i, 6).value)) > Val(Trim(.Cells(i, 11).value)) Then
                  If Val(Trim(.Cells(i, 7).value)) > 0 Then
                     sTemp1 = "CONTROLLED"
                  Else
                     sTemp1 = "ALARM"
                  End If
               Else
                  sTemp1 = "CONTROLLED"
               End If
            Else
               sTemp1 = "CONTROLLED"
            End If
         Else
            sTemp1 = "CONTROLLED"
         End If
         
         'Insert the control status
         .Cells(i, 10).value = sTemp1
      End If
   Next i
  
  '**********************************************************************
   'Save the file
   '**********************************************************************
   ExcelApp.DisplayAlerts = False
   Report.SaveAs sFilePath & sFileName, xlCSV
   ExcelApp.DisplayAlerts = True
   If Not bBackground Then AddLog "File Saved to preserve corrections"
  
End With

Report.Close False
ExcelApp.Quit

If Not bBackground Then AddLog ""
If Not bBackground Then AddLog "Completed File creation at " & Now()
If Not bBackground Then AddLog "********************************************************************************" & vbCrLf

Print #iLog, "Completed File creation at " & Now()
Debug.Print "Completed File creation at " & Now()
Close #iLog

On Error Resume Next

Exit Function

ErrorHandler:

If Not bBackground Then AddLog "Log Creation Code Error Encountered: " & Err.Number & " - " & Err.Description & " - " & Err.Source
Debug.Print "##Log Creation Code Error Encountered: " & Err.Number & " - " & Err.Description & " - " & Err.Source
If Not bBackground Then AddLog ""
Resume Next

End Function

