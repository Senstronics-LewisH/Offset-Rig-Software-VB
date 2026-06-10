Attribute VB_Name = "CSVFile"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function FindExcelFile() As Boolean

Dim FilePath As String


FilePath = "\\USVR8\Results\Production\Offset Check Results\"

        
    FileName = FilePath & WorksOrder & ".xls"
        
If Len(Dir(FileName)) > 0 Then
    FindExcelFile = True
Else
    FindExcelFile = False
End If
  
End Function
Public Function FindPODInExcelFile() As Boolean

Dim FilePath As String

On Error GoTo errhandler

FilePath = "\\USVR8\Results\Production\Offset Check Results\"
        
    FileName = FilePath & WorksOrder & ".xls"

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(FileName)

objWorkbook.Worksheets(1).Activate
 If objWorkbook.Worksheets(1).Cells(10, 2).Value = "1" Then
        FindPODInExcelFile = True
    Else
        FindPODInExcelFile = False
    End If

    objExcel.Quit
    Set objExcel = Nothing
    Exit Function

errhandler:
    MsgBox "ERROR ACCESSING FILE"
    objExcel.Quit
    Set objExcel = Nothing
    Exit Function

End Function
Public Function FindResults()

Dim FilePath As String

On Error GoTo errhandler

FilePath = "\\USVR8\Results\Production\Offset Check Results\"

       
    FileName = FilePath & WorksOrder & ".xls"

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(FileName)

objWorkbook.Worksheets(1).Activate
 
 MainForm.PassedBox = objWorkbook.Worksheets(1).Cells(1, 2).Value
 MainForm.FailedBox = objWorkbook.Worksheets(1).Cells(2, 2).Value
     
    If MainForm.FailedBox = 0 Then
        MainForm.PercentBox = "0"
    Else
            TotalGood = MainForm.PassedBox
            TotalBad = MainForm.FailedBox
            TotalParts = TotalGood + TotalBad
            PercentFail = Format$((TotalBad / (TotalBad + TotalGood) * 100), "0")
            MainForm.PercentBox = PercentFail
    End If

    If MainForm.PercentBox > 5 Then
        MainForm.PercentBox.BackColor = &HFF&
    Else
        MainForm.PercentBox.BackColor = &HFFFFFF
    End If

    objExcel.Quit
    Set objExcel = Nothing
    Exit Function

errhandler:
    MsgBox "ERROR ACCESSING FILE"
    objExcel.Quit
    Set objExcel = Nothing
    Exit Function

End Function
Public Sub CreateExcel()

Dim NewFile As New Excel.Application

FilePath = "\\USVR8\Results\Production\Offset Check Results\"
    
    Set xl = New Excel.Application
    Set xlwbook = xl.Workbooks.Add
    Set xlsheet = xlwbook.Sheets.item(1)
    
    xlsheet.Cells(1, 1) = "PASSED"
    xlsheet.Cells(2, 1) = "FAILED"
    xlsheet.Cells(1, 2) = "0"
    xlsheet.Cells(2, 2) = "0"
    xlsheet.Cells(3, 1) = Now
    xlsheet.Cells(4, 1) = "OFFSET"
    xlsheet.Cells(4, 2) = MainForm.OffsetTargetDisplay
    xlsheet.Cells(5, 1) = "FULL SCALE"
    xlsheet.Cells(5, 2) = FullScale
    xlsheet.Cells(6, 1) = "OUTPUT TYPE"
    xlsheet.Cells(6, 2) = OutputType
    xlsheet.Cells(7, 1) = "CALIBRATED PRESSURE"
    xlsheet.Cells(7, 2) = BarcodeFSPressure
    xlsheet.Cells(7, 3) = Units
    xlsheet.Cells(8, 1) = "OFFSET LOW LIMIT"
    xlsheet.Cells(8, 2) = MainForm.LowerLimitDisplay
    xlsheet.Cells(9, 1) = "OFFSET HIGH LIMIT"
    xlsheet.Cells(9, 2) = MainForm.UpperLimitDisplay
    xlsheet.Cells(10, 1) = "POD ON?"
    xlsheet.Cells(10, 2) = MainForm.PODCheck
    
    xlsheet.Cells(12, 1) = "SENSOR ID"
    xlsheet.Cells(12, 2) = "STATUS"
    xlsheet.Cells(12, 3) = "OFFSET 1"
    xlsheet.Cells(12, 4) = "OFFSET 1 ERROR"
    xlsheet.Cells(12, 5) = "CALIBRATED OFFSET"
    xlsheet.Cells(12, 6) = "OFFSET 1 - CAL OFFSET"
    xlsheet.Cells(12, 7) = "OFFSET 2"
    xlsheet.Cells(12, 8) = "OFFSET 2 ERROR"
    xlsheet.Cells(12, 9) = "OFF2 - OFF1 ERROR"
    xlsheet.Cells(12, 10) = "VOUT 2 OUTPUT"
    xlsheet.Cells(12, 11) = "STC GND"
    xlsheet.Cells(12, 12) = "STC VS"
    xlsheet.Cells(12, 13) = "STC V1"
    xlsheet.Cells(12, 14) = "STC V2"
    xlsheet.Cells(12, 15) = "CURRENT"
    xlsheet.Cells(12, 16) = "O-RING RESULT"
    xlsheet.Cells(12, 17) = "RESTRICTOR RESULT"
    xlsheet.Cells(12, 18) = "RESTRICTOR WELDED"
    xlsheet.Cells(12, 19) = "CORRECT UNION"
        
        
    For i = 13 To 2013  'increased from 13 to 813 by DW
        xlsheet.Cells(i, 1) = i - 12
    Next
    
    xlsheet.Columns("A:P").AutoFit
    xlsheet.Columns("A:P").HorizontalAlignment = xlCenter
    
    FileName = FilePath & WorksOrder & ".xls"
    
    xl.ActiveWorkbook.SaveAs FileName:=FileName, FileFormat:=56
    
    xl.Quit
    Set xl = Nothing

End Sub
Public Function Update25DayHoldResult() As Boolean

Dim FilePath As String
Dim i As Integer

On Error GoTo errhandler

FilePath = "\\USVR8\Results\Production\Offset Check Results\"

    FileName = FilePath & WorksOrder & ".xls"
    
    Set xl = New Excel.Application
    Set xlwbook = xl.Workbooks.Open(FileName)
    Set xlsheet = xlwbook.Sheets.item(1)

    Transducer = MainForm.SensorID
    Row = Transducer + 12

    InitialResult = xlsheet.Cells(Row, 2).Value

    If InitialResult = "" Then
        MsgBox "PART HAS NOT BEEN THROUGH OFFSET 1"
        Update25DayHoldResult = False
    Else
        Update25DayHoldResult = True
        Update25DayHoldResult = True
        xlsheet.Cells(Row, 7) = MainForm.VOUT1OutputDisplay
        xlsheet.Cells(Row, 8) = MainForm.VOUT1OUTPUTERRORDISPLAY
        OffDif = xlsheet.Cells(Row, 8) - xlsheet.Cells(Row, 4)
        Vout1OriginalOutputDisplay = xlsheet.Cells(Row, 4)
        xlsheet.Cells(Row, 9) = OffDif
        
        MainForm.VoutDiffDisplay = Format$(OffDif, "0.000")
        
        MainForm.Vout1OriginalOutputDisplay = xlsheet.Cells(Row, 3)
    
        If OffDif > 0.15 Or OffDif < -0.1 Then
            xlsheet.Cells(Row, 2) = "FailedSecondOffset1"
            SensorStatus(MainForm.SensorID) = FailedSecondOffset1
            MainForm.OffDifFail.Visible = True
        End If
         
        xl.DisplayAlerts = False
        xl.ActiveWorkbook.SaveAs FileName:=FileName, FileFormat:=56
    End If
    
    xl.Quit
    Set xl = Nothing
    Exit Function
    
errhandler:
    MsgBox " ERROR SAVING DATA TO FILE"
    xl.Quit
    Set xl = Nothing
    Exit Function

End Function
Public Sub UpdateExcelWithIdResults()

Dim FilePath As String
Dim i As Integer

On Error GoTo errhandler

FilePath = "\\USVR8\Results\Production\Offset Check Results\"

       
AddToHistoryLogCDrive "Open Excel"
       
    FileName = FilePath & WorksOrder & ".xls"
    
    Set xl = New Excel.Application
    Set xlwbook = xl.Workbooks.Open(FileName)
    Set xlsheet = xlwbook.Sheets.item(1)

AddToHistoryLogCDrive "Input Data"

    Transducer = MainForm.SensorID
    Row = Transducer + 12

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'LH 11/06/25 - Changes for Milwaukee
    ' If the sensor id is <= 3 and the part number is in the list of part numbers with QR codes, ask the operator to scan the QR code.
    If MainForm.PartNumber = "%PTT14043" Or MainForm.PartNumber = "%PTT14044" Then
        If MainForm.PartCount <= 3 Then
            MainForm.PartCount = MainForm.PartCount + 1
            Dim QrValue
            QrValue = ""
            While QrValue = ""
                'Open the scan form
                QrValue = InputBox("PLEASE SCAN THE QR CODE ON THE LABEL", "QR Scan", "", 7000, 5000)
                
                'Check that the operator scanned something
                If QrValue = "" Then
                    'They didn't scan it, warn them that it wasnt scanned.
                    Response = MsgBox("You didn't scan anything! Try again or STOP CALL WAIT", 16, "Failed to Scan")
                    QrValue = ""
                End If
                If Len(QrValue) < 70 Then
                    'They scanned or entered something too short, warn them.
                    Response = MsgBox("Scan too short! Try again or STOP CALL WAIT", 16, "Scan too short")
                    QrValue = ""
                End If
            Wend
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''

    InitialResult = xlsheet.Cells(Row, 2).Value
         
    xlsheet.Cells(Row, 2) = SensorStatusEnumToString(SensorStatus(MainForm.SensorID))
            
    If SensorStatus(MainForm.SensorID) = PASSED Then
        If InitialResult = "" Then
            MainForm.PassedBox = MainForm.PassedBox + 1
        Else
            If InitialResult = "PASSED" Then
            Else
                MainForm.FailedBox = MainForm.FailedBox - 1
                MainForm.PassedBox = MainForm.PassedBox + 1
            End If
        End If
    Else
        If InitialResult = "" Then
            MainForm.FailedBox = MainForm.FailedBox + 1
        Else
            If InitialResult = "PASSED" Then
                MainForm.FailedBox = MainForm.FailedBox + 1
                MainForm.PassedBox = MainForm.PassedBox - 1
            End If
        End If
    End If
    
    If MainForm.FailedBox = 0 Then
        MainForm.PercentBox = "0"
    Else
        TotalGood = MainForm.PassedBox
        TotalBad = MainForm.FailedBox
        TotalParts = TotalGood + TotalBad
        PercentFail = Format$((TotalBad / (TotalBad + TotalGood) * 100), "0")
        MainForm.PercentBox = PercentFail
    End If
    
    If MainForm.PercentBox > 5 Then
        MainForm.PercentBox.BackColor = &HFF&
    Else
        MainForm.PercentBox.BackColor = &HFFFFFF
    End If

    If Post25DayTest = False And Retests = False Then
    
        If xlsheet.Cells(Row, 3) = "" Then
            xlsheet.Cells(Row, 3) = MainForm.VOUT1OutputDisplay
        Else
            OldReading = xlsheet.Cells(Row, 3)
            For i = 20 To 100
                Column = i
            
                If xlsheet.Cells(Row, Column) = "" Then
                    xlsheet.Cells(Row, 3) = MainForm.VOUT1OutputDisplay
                    xlsheet.Cells(Row, Column) = OldReading
                    Exit For
                End If
            Next
        End If
        
        xlsheet.Cells(1, 2) = MainForm.PassedBox
        xlsheet.Cells(2, 2) = MainForm.FailedBox
        
        xlsheet.Cells(Row, 4) = MainForm.VOUT1OUTPUTERRORDISPLAY
        xlsheet.Cells(Row, 5) = MainForm.OffsetFromCalDisplay
        xlsheet.Cells(Row, 6) = MainForm.OFFSETDIFFDISPLAY
        xlsheet.Cells(Row, 10) = MainForm.VOUT2OutputDisplay
        xlsheet.Cells(Row, 11) = MainForm.STCGNDDisplay
        xlsheet.Cells(Row, 12) = MainForm.STCVSDisplay
        xlsheet.Cells(Row, 13) = MainForm.STCVOUT1Display
        xlsheet.Cells(Row, 14) = MainForm.STCVOUT2Display
        xlsheet.Cells(Row, 15) = MainForm.CurrentDisplay
        xlsheet.Cells(Row, 16) = ORingResult
        xlsheet.Cells(Row, 17) = RestrictorResult
        xlsheet.Cells(Row, 18) = RestrictorWelded
        xlsheet.Cells(Row, 19) = CorrectUnion
        
        xlsheet.Cells(10, 2) = MainForm.PODCheck
   
    End If
    
AddToHistoryLogCDrive "Close Excel"
    
    xl.DisplayAlerts = False
    xl.ActiveWorkbook.SaveAs FileName:=FileName, FileFormat:=56
    xlwbook.Close
    xl.Quit
 

  If Not xlsheet Is Nothing Then Set xlsheet = Nothing
  If Not xlwbook Is Nothing Then Set xlwbook = Nothing
  If Not xl Is Nothing Then Set xl = Nothing

    
AddToHistoryLogCDrive "Excel Closed"
AddToHistoryLogCDrive "Waiting to exit sub"
    Exit Sub
    
errhandler:
    MsgBox " ERROR SAVING DATA TO FILE"
    xl.Quit
    Set xl = Nothing
    Exit Sub
    

End Sub
Public Function OpenExcelFile()

    Dim FilePath As String

    FilePath = "\\USVR8\Results\Production\Offset Check Results\"
       
    FileName = FilePath & WorksOrder & ".xls"

ShellExecute 0, vbNullString, FileName, vbNullString, vbNullString, vbNormalFocus

End Function

Public Function SensorStatusEnumToString(ByVal SensorStatus As SensorStatusEnum) As String

 Select Case SensorStatus
    Case PASSED
        SensorStatusEnumToString = "PASSED"
    Case FailedOffset1
        SensorStatusEnumToString = "FailedOffset1"
    Case FailedOffToCal
        SensorStatusEnumToString = "FailedOffToCal"
    Case FailedSecondOffset1
        SensorStatusEnumToString = "FailedSecondOffset1"
    Case FailedSTC
        SensorStatusEnumToString = "FailedSTC"
    Case FailedCurrent
        SensorStatusEnumToString = "FailedCurrent"
    Case FailedSwitch
        SensorStatusEnumToString = "FailedSwitch"
    Case FailedTemp
        SensorStatusEnumToString = "FailedTemp"
    Case FailedOring
        SensorStatusEnumToString = "FailedOring"
    Case FailedRestPresent
        SensorStatusEnumToString = "FailedRestPresent"
    Case FailedRestweld
        SensorStatusEnumToString = "FailedRestWeld"
    Case FailedWrongUnion
        SensorStatusEnumToString = "FailedWrongUnion"
    Case DidNotPassCalibration
        SensorStatusEnumToString = "DidNotPassCalibration"

    End Select
End Function
Public Sub PrintLabel()
    Dim FileLine As String
    Dim FileHandle As Integer
    Dim WorkOrder As String
    
    On Error GoTo errhandler
   
    WorkOrder = (Mid$(MainForm.WorksOrderBarcode, 5, 15))

    IDNumber = MainForm.SensorID
    
    FileName = "M:\system\load\Vborders.txt"
    
    FileHandle = FreeFile
    Open FileName For Output As #FileHandle
    
    FileLine = WorkOrder & vbTab & IDNumber
    Print #FileHandle, FileLine
    
    Close #FileHandle

    Shell ("C:\liveorders\PrintLabel.bat")

    Exit Sub
    
errhandler:

    MsgBox "Error Saving Print File"

End Sub
