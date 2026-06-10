Attribute VB_Name = "ControlPSU"
Option Explicit
Dim Driver As New Keithley2230
Public Sub InitialisePSU()
  'Initialise Power Supply Unit And Check With A Read-Back
    
    Driver.initialize "USB0::0x05E6::0x2230::" & PSU_Visa_ID & "::INSTR", False, False, ""
    Driver.System.ControlMode = Keithley2230ControlModeRemote
    SetPSU1 24

End Sub
Public Function ReceivePSUVoltage() As String
  'Read PSU voltage

    Dim Reply As String

    Reply = CStr(Driver.OutputChannels.item("OutputChannel1").Measure(Keithley2230MeasurementTypeVoltage))
    ReceivePSUVoltage = Reply
    
End Function
Public Sub SetPSU1(Supply)
'24v for vision and push button
    Driver.OutputChannels.item("OutputChannel1").VoltageLevel = Supply
    Driver.OutputChannels.item("OutputChannel1").State = True
         
End Sub
Public Sub SetPSU2(Supply)
'Transducer supply
    Driver.OutputChannels.item("OutputChannel2").CurrentLimit = 0.01
    Driver.OutputChannels.item("OutputChannel2").VoltageLevel = Supply * 0.986
    Driver.OutputChannels.item("OutputChannel2").State = True
    Sleep 100
        
End Sub
Public Sub SetPSU3(Supply)
'Switch Input

'set current to 3mA

    Driver.OutputChannels.item("OutputChannel3").CurrentLimit = 0.005
    Driver.OutputChannels.item("OutputChannel3").VoltageLevel = Supply * 0.9645
    Driver.OutputChannels.item("OutputChannel3").State = True
    Sleep 200
        
End Sub
Public Function ReadCurrent() As Double

    Dim Reply As String

    Reply = CStr(Driver.OutputChannels.item("OutputChannel2").Measure(Keithley2230MeasurementTypeCurrent))
    ReadCurrent = Val(Reply)

End Function



  

