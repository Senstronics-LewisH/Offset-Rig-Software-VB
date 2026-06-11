Attribute VB_Name = "CSVFile"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function FindExcelFile() As Boolean
    Dim FilePath As String
    Dim FileName As String

    FilePath = ResultsPath
    FileName = FilePath & WorksOrder & ".xls"
        
    If Len(Dir(FileName)) > 0 Then
        FindExcelFile = True
    Else
        FindExcelFile = False
    End If
End Function

Public Function FindPODInExcelFile() As Boolean
    Dim FilePath As String
    Dim FileName As String
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    
    On Error GoTo errhandler
    
    FilePath = ResultsPath
    FileName = FilePath & WorksOrder & ".xls"
    
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(FileName)
    Set xlSheet = xlBook.Worksheets(1)
    
    If xlSheet.Cells(10, 2).Value = "1" Then
        FindPODInExcelFile = True
    Else
        FindPODInExcelFile = False
    End If
    
    Set xlSheet = Nothing
    xlBook.Close SaveChanges:=False
    Set xlBook = Nothing
    xlApp.Quit
    Set xlApp = Nothing
    Exit Function

errhandler:
    MsgBox "ERROR ACCESSING FILE"
    Set xlSheet = Nothing
    If Not xlBook Is Nothing Then
        On Error Resume Next
        xlBook.Close SaveChanges:=False
        On Error GoTo 0
        Set xlBook = Nothing
    End If
    If Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.Quit
        On Error GoTo 0
        Set xlApp = Nothing
    End If
    FindPODInExcelFile = False
End Function

Public Function FindResults()
    Dim FilePath As String
    Dim FileName As String
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim TotalGood As Long
    Dim TotalBad As Long
    Dim TotalParts As Long
    Dim PercentFail As String
    
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    
    On Error GoTo errhandler
    
    FilePath = ResultsPath
    FileName = FilePath & WorksOrder & ".xls"
    
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(FileName)
    Set xlSheet = xlBook.Worksheets(1)
    
    MainForm.PassedBox = xlSheet.Cells(1, 2).Value
    MainForm.FailedBox = xlSheet.Cells(2, 2).Value
    
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
    
    Set xlSheet = Nothing
    xlBook.Close SaveChanges:=False
    Set xlBook = Nothing
    xlApp.Quit
    Set xlApp = Nothing
    Exit Function
    
errhandler:
    MsgBox "ERROR ACCESSING FILE"
    Set xlSheet = Nothing
    If Not xlBook Is Nothing Then
        On Error Resume Next
        xlBook.Close SaveChanges:=False
        On Error GoTo 0
        Set xlBook = Nothing
    End If
    If Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.Quit
        On Error GoTo 0
        Set xlApp = Nothing
    End If
End Function

Public Sub CreateExcel()
    Dim FilePath As String
    Dim FileName As String
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim i As Long
    
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    
    On Error GoTo errhandler
    
    FilePath = ResultsPath
    FileName = FilePath & WorksOrder & ".xls"
    
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets.Item(1)
    
    xlSheet.Cells(1, 1) = "PASSED"
    xlSheet.Cells(2, 1) = "FAILED"
    xlSheet.Cells(1, 2) = "0"
    xlSheet.Cells(2, 2) = "0"
    xlSheet.Cells(3, 1) = Now
    xlSheet.Cells(4, 1) = "OFFSET"
    xlSheet.Cells(4, 2) = MainForm.OffsetTargetDisplay
    xlSheet.Cells(5, 1) = "FULL SCALE"
    xlSheet.Cells(5, 2) = FullScale
    xlSheet.Cells(6, 1) = "OUTPUT TYPE"
    xlSheet.Cells(6, 2) = OutputType
    xlSheet.Cells(7, 1) = "CALIBRATED PRESSURE"
    xlSheet.Cells(7, 2) = BarcodeFSPressure
    xlSheet.Cells(7, 3) = Units
    xlSheet.Cells(8, 1) = "OFFSET LOW LIMIT"
    xlSheet.Cells(8, 2) = MainForm.LowerLimitDisplay
    xlSheet.Cells(9, 1) = "OFFSET HIGH LIMIT"
    xlSheet.Cells(9, 2) = MainForm.UpperLimitDisplay
    xlSheet.Cells(10, 1) = "POD ON?"
    xlSheet.Cells(10, 2) = MainForm.PODCheck
    
    xlSheet.Cells(12, 1) = "SENSOR ID"
    xlSheet.Cells(12, 2) = "STATUS"
    xlSheet.Cells(12, 3) = "OFFSET 1"
    xlSheet.Cells(12, 4) = "OFFSET 1 ERROR"
    xlSheet.Cells(12, 5) = "CALIBRATED OFFSET"
    xlSheet.Cells(12, 6) = "OFFSET 1 - CAL OFFSET"
    xlSheet.Cells(12, 7) = "OFFSET 2"
    xlSheet.Cells(12, 8) = "OFFSET 2 ERROR"
    xlSheet.Cells(12, 9) = "OFF2 - OFF1 ERROR"
    xlSheet.Cells(12, 10) = "VOUT 2 OUTPUT"
    xlSheet.Cells(12, 11) = "STC GND"
    xlSheet.Cells(12, 12) = "STC VS"
    xlSheet.Cells(12, 13) = "STC V1"
    xlSheet.Cells(12, 14) = "STC V2"
    xlSheet.Cells(12, 15) = "CURRENT"
    xlSheet.Cells(12, 16) = "O-RING RESULT"
    xlSheet.Cells(12, 17) = "RESTRICTOR RESULT"
    xlSheet.Cells(12, 18) = "RESTRICTOR WELDED"
    xlSheet.Cells(12, 19) = "CORRECT UNION"
    
    For i = 13 To 2013
        xlSheet.Cells(i, 1) = i - 12
    Next
    
    xlSheet.Columns("A:P").AutoFit
    xlSheet.Columns("A:P").HorizontalAlignment = xlCenter
    
    xlApp.DisplayAlerts = False
    xlBook.SaveAs FileName:=FileName, FileFormat:=56
    
    Set xlSheet = Nothing
    xlBook.Close SaveChanges:=False
    Set xlBook = Nothing
    xlApp.Quit
    Set xlApp = Nothing
    Exit Sub
    
errhandler:
    MsgBox "ERROR CREATING EXCEL FILE"
    Set xlSheet = Nothing
    If Not xlBook Is Nothing Then
        On Error Resume Next
        xlBook.Close SaveChanges:=False
        On Error GoTo 0
        Set xlBook = Nothing
    End If
    If Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.Quit
        On Error GoTo 0
        Set xlApp = Nothing
    End If
End Sub

Public Function Update25DayHoldResult() As Boolean
    Dim FilePath As String
    Dim FileName As String
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim Transducer As Long
    Dim Row As Long
    Dim InitialResult As Variant
    Dim OffDif As Double
    Dim Vout1OriginalOutputDisplay As Variant
    
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    
    On Error GoTo errhandler
    
    FilePath = ResultsPath
    FileName = FilePath & WorksOrder & ".xls"
    
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(FileName)
    Set xlSheet = xlBook.Worksheets(1)
    
    Transducer = MainForm.SensorID
    Row = Transducer + 12
    
    InitialResult = xlSheet.Cells(Row, 2).Value
    
    If InitialResult = "" Then
        MsgBox "PART HAS NOT BEEN THROUGH OFFSET 1"
        Update25DayHoldResult = False
    Else
        Update25DayHoldResult = True
        xlSheet.Cells(Row, 7) = MainForm.VOUT1OutputDisplay
        xlSheet.Cells(Row, 8) = MainForm.VOUT1OUTPUTERRORDISPLAY
        OffDif = xlSheet.Cells(Row, 8) - xlSheet.Cells(Row, 4)
        Vout1OriginalOutputDisplay = xlSheet.Cells(Row, 4)
        xlSheet.Cells(Row, 9) = OffDif
        
        MainForm.VoutDiffDisplay = Format$(OffDif, "0.000")
        MainForm.Vout1OriginalOutputDisplay = xlSheet.Cells(Row, 3)
        
        If OffDif > 0.15 Or OffDif < -0.1 Then
            xlSheet.Cells(Row, 2) = "FailedSecondOffset1"
            SensorStatus(MainForm.SensorID) = FailedSecondOffset1
            MainForm.OffDifFail.Visible = True
        End If
        
        xlApp.DisplayAlerts = False
        xlBook.SaveAs FileName:=FileName, FileFormat:=56
    End If
    
    Set xlSheet = Nothing
    xlBook.Close SaveChanges:=False
    Set xlBook = Nothing
    xlApp.Quit
    Set xlApp = Nothing
    Exit Function
    
errhandler:
    MsgBox " ERROR SAVING DATA TO FILE"
    Set xlSheet = Nothing
    If Not xlBook Is Nothing Then
        On Error Resume Next
        xlBook.Close SaveChanges:=False
        On Error GoTo 0
        Set xlBook = Nothing
    End If
    If Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.Quit
        On Error GoTo 0
        Set xlApp = Nothing
    End If
    Update25DayHoldResult = False
End Function

Public Sub UpdateExcelWithIdResults()
    Dim FilePath As String
    Dim FileName As String
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim Transducer As Long
    Dim Row As Long
    Dim QrValue As String
    Dim Response As Long
    Dim InitialResult As Variant
    Dim TotalGood As Long
    Dim TotalBad As Long
    Dim TotalParts As Long
    Dim PercentFail As String
    Dim OldReading As Variant
    Dim Column As Long
    Dim i As Integer
    
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    
    On Error GoTo errhandler
    
    FilePath = ResultsPath
    AddToHistoryLogCDrive "Open Excel"
    FileName = FilePath & WorksOrder & ".xls"
    
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(FileName)
    Set xlSheet = xlBook.Worksheets(1)
    
    AddToHistoryLogCDrive "Input Data"
    
    Transducer = MainForm.SensorID
    Row = Transducer + 12
    
    ' Milwaukee Changes
    If MainForm.PartNumber = "%PTT14043" Or MainForm.PartNumber = "%PTT14044" Then
        If MainForm.PartCount <= 3 Then
            MainForm.PartCount = MainForm.PartCount + 1
            QrValue = ""
            While QrValue = ""
                QrValue = InputBox("PLEASE SCAN THE QR CODE ON THE LABEL", "QR Scan", "", 7000, 5000)
                If QrValue = "" Then
                    Response = MsgBox("You didn't scan anything! Try again or STOP CALL WAIT", 16, "Failed to Scan")
                    QrValue = ""
                End If
                If Len(QrValue) < 70 Then
                    Response = MsgBox("Scan too short! Try again or STOP CALL WAIT", 16, "Scan too short")
                    QrValue = ""
                End If
            Wend
        End If
    End If
    
    InitialResult = xlSheet.Cells(Row, 2).Value
    xlSheet.Cells(Row, 2) = SensorStatusEnumToString(SensorStatus(MainForm.SensorID))
    
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
        If xlSheet.Cells(Row, 3) = "" Then
            xlSheet.Cells(Row, 3) = MainForm.VOUT1OutputDisplay
        Else
            OldReading = xlSheet.Cells(Row, 3)
            For i = 20 To 100
                Column = i
                If xlSheet.Cells(Row, Column) = "" Then
                    xlSheet.Cells(Row, 3) = MainForm.VOUT1OutputDisplay
                    xlSheet.Cells(Row, Column) = OldReading
                    Exit For
                End If
            Next
        End If
        
        xlSheet.Cells(1, 2) = MainForm.PassedBox
        xlSheet.Cells(2, 2) = MainForm.FailedBox
        
        xlSheet.Cells(Row, 4) = MainForm.VOUT1OUTPUTERRORDISPLAY
        xlSheet.Cells(Row, 5) = MainForm.OffsetFromCalDisplay
        xlSheet.Cells(Row, 6) = MainForm.OFFSETDIFFDISPLAY
        xlSheet.Cells(Row, 10) = MainForm.VOUT2OutputDisplay
        xlSheet.Cells(Row, 11) = MainForm.STCGNDDisplay
        xlSheet.Cells(Row, 12) = MainForm.STCVSDisplay
        xlSheet.Cells(Row, 13) = MainForm.STCVOUT1Display
        xlSheet.Cells(Row, 14) = MainForm.STCVOUT2Display
        xlSheet.Cells(Row, 15) = MainForm.CurrentDisplay
        xlSheet.Cells(Row, 16) = ORingResult
        xlSheet.Cells(Row, 17) = RestrictorResult
        xlSheet.Cells(Row, 18) = RestrictorWelded
        xlSheet.Cells(Row, 19) = CorrectUnion
        
        xlSheet.Cells(10, 2) = MainForm.PODCheck
    End If
    
    AddToHistoryLogCDrive "Close Excel"
    xlApp.DisplayAlerts = False
    xlBook.SaveAs FileName:=FileName, FileFormat:=56
    
    Set xlSheet = Nothing
    xlBook.Close SaveChanges:=False
    Set xlBook = Nothing
    xlApp.Quit
    Set xlApp = Nothing
    
    AddToHistoryLogCDrive "Excel Closed"
    AddToHistoryLogCDrive "Waiting to exit sub"
    Exit Sub
    
errhandler:
    MsgBox " ERROR SAVING DATA TO FILE"
    Set xlSheet = Nothing
    If Not xlBook Is Nothing Then
        On Error Resume Next
        xlBook.Close SaveChanges:=False
        On Error GoTo 0
        Set xlBook = Nothing
    End If
    If Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.Quit
        On Error GoTo 0
        Set xlApp = Nothing
    End If
End Sub

Public Function OpenExcelFile()
    Dim FilePath As String
    Dim FileName As String

    FilePath = ResultsPath
    FileName = FilePath & WorksOrder & ".xls"

    ShellExecute 0, vbNullString, FileName, vbNullString, vbNullString, vbNormalFocus
End Function

Public Function SensorStatusEnumToString(ByVal SensorStatus As SensorStatusEnum) As String
    Select Case SensorStatus
        Case PASSED:              SensorStatusEnumToString = "PASSED"
        Case FailedOffset1:       SensorStatusEnumToString = "FailedOffset1"
        Case FailedOffToCal:      SensorStatusEnumToString = "FailedOffToCal"
        Case FailedSecondOffset1: SensorStatusEnumToString = "FailedSecondOffset1"
        Case FailedSTC:           SensorStatusEnumToString = "FailedSTC"
        Case FailedCurrent:       SensorStatusEnumToString = "FailedCurrent"
        Case FailedSwitch:        SensorStatusEnumToString = "FailedSwitch"
        Case FailedTemp:          SensorStatusEnumToString = "FailedTemp"
        Case FailedOring:         SensorStatusEnumToString = "FailedOring"
        Case FailedRestPresent:   SensorStatusEnumToString = "FailedRestPresent"
        Case FailedRestweld:      SensorStatusEnumToString = "FailedRestWeld"
        Case FailedWrongUnion:    SensorStatusEnumToString = "FailedWrongUnion"
        Case DidNotPassCalibration: SensorStatusEnumToString = "DidNotPassCalibration"
    End Select
End Function

Public Sub PrintLabel()
    Dim FileLine As String
    Dim FileHandle As Integer
    Dim WorkOrder As String
    Dim IDNumber As Long
    Dim FileName As String
    
    On Error GoTo errhandler
   
    WorkOrder = (Mid$(MainForm.WorksOrderBarcode, 5, 15))
    IDNumber = MainForm.SensorID
    FileName = LabelPrintInputPath
    
    FileHandle = FreeFile
    Open FileName For Output As #FileHandle
    
    FileLine = WorkOrder & vbTab & IDNumber
    Print #FileHandle, FileLine
    Close #FileHandle

    Shell (LabelPrintBatchPath)
    Exit Sub
    
errhandler:
    MsgBox "Error Saving Print File"
End Sub
