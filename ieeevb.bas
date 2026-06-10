Attribute VB_Name = "ieeeevb"
Option Explicit

'****************************************************************************************************
'Module1  (ieeevb.bas)          Visual Basic 6.0 Module
'Last modified: never modified
'Description:
'               -This code was written externally
'
'This module contains subroutines designed to operate the features of the Keithley IEEE card
'It is a module shipped with the card, and is untouched by Senstronics
'
' CEC IEEE-488 subroutines
' for use with CEC interface cards
' Copyright (C) 1995, Capital Equipment Corporation
' Customers may use this code in their application
' programs which run with CEC interface cards.
' All other rights reserved.
' For VISUAL BASIC 4.0 and later versions
'
' revisions:
'       4/23/96 - change integer to long for 32-bit routines
'       1/28/97 - fix declaration of SetInputEOS routine for 32-bit
'       5/97    - added new routines for CEC488 v5
'
'****************************************************************************************************


Global Const IEEEListener = 0, IEEE488SD = 1, IEEEDMA = 2, IEEEIOBASE = 100, IEEETIMEOUT = 200, IEEEINPUTEOS = 201, IEEEOUTPUTEOS1 = 202, IEEEOUTPUTEOS2 = 203, IEEEBOARDSELECT = 204, IEEEDMACHANNEL = 205
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#If Win32 Then
'----------------------------------------------------------------------------
' 32-bit versions of IEEE488 routines
'----------------------------------------------------------------------------
Declare Sub initialize Lib "IEEE_32M.DLL" Alias "_ieee_initialize@8" (ByVal addr As Long, ByVal level As Long)
Declare Sub IEtrans Lib "IEEE_32M.DLL" Alias "_ieee_transmit@12" (ByVal cmd As String, ByVal l As Long, Status As Long)
Declare Sub IEreceive Lib "IEEE_32M.DLL" Alias "_ieee_receive@16" (ByVal r As String, ByVal maxlen As Long, l As Long, Status As Long)
Declare Sub IEsend Lib "IEEE_32M.DLL" Alias "_ieee_send@16" (ByVal addr As Long, ByVal s As String, ByVal l As Long, Status As Long)
Declare Sub IEenter Lib "IEEE_32M.DLL" Alias "_ieee_enter@20" (ByVal r As String, ByVal maxlen As Long, l As Long, ByVal addr As Long, Status As Long)
Declare Sub IEspoll Lib "IEEE_32M.DLL" Alias "_ieee_spoll@12" (ByVal addr As Long, poll As Long, Status As Long)
Declare Sub IEppoll Lib "IEEE_32M.DLL" Alias "_ieee_ppoll@4" (poll As Long)
Declare Sub IEtarray Lib "IEEE_32M.DLL" Alias "_ieee_tarray@16" (d As Any, ByVal Count As Long, ByVal eoi As Long, Status As Long)
Declare Sub IErarray Lib "IEEE_32M.DLL" Alias "_ieee_rarray@16" (d As Any, ByVal Count As Long, l As Long, Status As Long)
Declare Function srq Lib "IEEE_32M.DLL" Alias "_ieee_srq@0" () As Long
Declare Sub setport Lib "IEEE_32M.DLL" Alias "_ieee_setport@8" (ByVal board As Long, ByVal Port As Long)
Declare Sub boardselect Lib "IEEE_32M.DLL" Alias "_ieee_boardselect@4" (ByVal board As Long)
Declare Sub dmachannel Lib "IEEE_32M.DLL" Alias "_ieee_dmachannel@4" (ByVal chan As Long)
Declare Sub settimeout Lib "IEEE_32M.DLL" Alias "_ieee_settimeout@4" (ByVal msec As Long)
Declare Sub setoutputEOS Lib "IEEE_32M.DLL" Alias "_ieee_setoutputEOS@8" (ByVal c1 As Long, ByVal c2 As Long)
Declare Sub setinputEOS Lib "IEEE_32M.DLL" Alias "_ieee_setinputEOS@4" (ByVal c As Long)
Declare Sub Enable488EX Lib "IEEE_32M.DLL" Alias "_ieee_enable_488ex@4" (ByVal e As Long)
Declare Sub Enable488SD Lib "IEEE_32M.DLL" Alias "_ieee_enable_488sd@8" (ByVal e As Long, ByVal t As Long)
Declare Function ListenerPresent Lib "IEEE_32M.DLL" Alias "_ieee_listener_present@4" (ByVal a As Long) As Long
Declare Function GpibBoardPresent Lib "IEEE_32M.DLL" Alias "_ieee_board_present@0" () As Long
Declare Function GPIBFeature Lib "IEEE_32M.DLL" Alias "_ieee_feature@4" (ByVal f As Long) As Long
#Else
'----------------------------------------------------------------------------
' 16-bit versions of IEEE488 routines
'----------------------------------------------------------------------------
Declare Sub initialize Lib "win488.dll" Alias "IEEE488_INITIALIZE" (ByVal addr As Integer, ByVal level As Integer)
Declare Sub IEtrans Lib "win488.dll" Alias "IEEE488_TRANSMIT" (ByVal cmd As String, ByVal l As Integer, Status As Integer)
Declare Sub IEreceive Lib "win488.dll" Alias "IEEE488_RECEIVE" (ByVal r As String, ByVal maxlen As Integer, l As Integer, Status As Integer)
Declare Sub IEsend Lib "win488.dll" Alias "IEEE488_SEND" (ByVal addr As Integer, ByVal s As String, ByVal l As Integer, Status As Integer)
Declare Sub IEenter Lib "win488.dll" Alias "IEEE488_ENTER" (ByVal r As String, ByVal maxlen As Integer, l As Integer, ByVal addr As Integer, Status As Integer)
Declare Sub IEspoll Lib "win488.dll" Alias "IEEE488_SPOLL" (ByVal addr As Integer, poll As Integer, Status As Integer)
Declare Sub IEppoll Lib "win488.dll" Alias "IEEE488_PPOLL" (poll As Integer)
Declare Sub IEtarray Lib "win488.dll" Alias "IEEE488_TARRAY" (d As Any, ByVal Count As Integer, ByVal eoi As Integer, Status As Integer)
Declare Sub IErarray Lib "win488.dll" Alias "IEEE488_RARRAY" (d As Any, ByVal Count As Integer, l As Integer, Status As Integer)
Declare Function srq Lib "win488.dll" Alias "IEEE488_SRQ" () As Integer
Declare Sub setport Lib "win488.dll" Alias "IEEE488_SETPORT" (ByVal board As Integer, ByVal Port As Integer)
Declare Sub boardselect Lib "win488.dll" Alias "IEEE488_BOARDSELECT" (ByVal board As Integer)
Declare Sub dmachannel Lib "win488.dll" Alias "IEEE488_DMACHANNEL" (ByVal chan As Integer)
Declare Sub settimeout Lib "win488.dll" Alias "IEEE488_SETTIMEOUT" (ByVal msec As Integer)
Declare Sub setoutputEOS Lib "win488.dll" Alias "IEEE488_SETOUTPUTEOS" (ByVal c1 As Integer, ByVal c2 As Integer)
Declare Sub setinputEOS Lib "win488.dll" Alias "IEEE488_SETINPUTEOS" (ByVal c As Integer)
Declare Sub Enable488EX Lib "win488.dll" Alias "IEEE488_ENABLE_488EX" (ByVal e As Integer)
Declare Sub Enable488SD Lib "win488.dll" Alias "IEEE488_ENABLE_488SD" (ByVal e As Integer, ByVal t As Integer)
Declare Function ListenerPresent Lib "win488.dll" Alias "IEEE488_LISTENER_PRESENT" (ByVal a As Integer) As Integer
Declare Function GpibBoardPresent Lib "win488.dll" Alias "IEEE488_BOARD_PRESENT" () As Integer
Declare Function GPIBFeature Lib "win488.dll" Alias "IEEE488_FEATURE" (ByVal f As Integer) As Integer
#End If
'-------------------------------------------------------
Sub enter(r As String, maxlen As Integer, l As Integer, addr As Integer, Status As Integer)
#If Win32 Then
    Dim stl As Long
    Dim ll As Long
    r = Space$(maxlen)
    Call IEenter(r, maxlen, ll, addr, stl)
    l = ll
    r = Left$(r, l)
    Status = stl
#Else
    r = Space$(maxlen)
    Call IEenter(r, maxlen, l, addr, Status)
    r = Left$(r, l)
#End If
End Sub
'-------------------------------------------------------
Sub receive(r As String, maxlen As Integer, l As Integer, Status As Integer)
#If Win32 Then
    Dim stl As Long
    Dim ll As Long
    r = Space$(maxlen)
    Call IEreceive(r, maxlen, ll, stl)
    l = ll
    r = Left$(r, l)
    Status = stl
#Else
    r = Space$(maxlen)
    Call IEreceive(r, maxlen, l, Status)
    r = Left$(r, l)
#End If
End Sub
'-------------------------------------------------------
Public Sub send(addr As Integer, s As String, Status As Integer)
#If Win32 Then
    Dim stl As Long
    Call IEsend(addr, s, -1, stl)
    Status = stl
#Else
    Call IEsend(addr, s, -1, Status)
#End If
End Sub
'-------------------------------------------------------
Sub transmit(cmd As String, Status As Integer)
#If Win32 Then
    Dim stl As Long
    Call IEtrans(cmd, -1, stl)
    Status = stl
#Else
    Call IEtrans(cmd, -1, Status)
#End If
End Sub
'-------------------------------------------------------
Sub spoll(ByVal addr As Integer, poll As Integer, Status As Integer)
#If Win32 Then
    Dim stl As Long
    Dim pl As Long
    Call IEspoll(addr, pl, stl)
    poll = pl
    Status = stl
#Else
    Call IEspoll(addr, poll, Status)
#End If
End Sub
'-------------------------------------------------------
Sub ppoll(poll As Integer)
#If Win32 Then
    Dim pl As Long
    Call IEppoll(pl)
    poll = pl
#Else
    Call IEppoll(poll)
#End If
End Sub
'-------------------------------------------------------
' IMPORTANT NOTE:
'       To call tarray and rarray in VB 4, you must pass the
'       entire array, as in:
'               call tarray(d(),100,1,status%)
'       You CANNOT pass d(1), as was done in previous versions
'       of VB.  Microsoft changed the data types and argument
'       conventions, so this source code change is required.
'
' IMPORTANT NOTE: If you want to use arrays of a type other than
'       integer, you must edit these procedures and change the
'       type of the local array dd() in both tarray and rarray.
'       Also, you must change the loop limit variable yy.
'       For example, to use Byte arrays, change the dim statements to:
'               dim dd(65535) as byte
'       and the statement just before the "for xx" loops to:
'               yy = count
'       and
'               yy = l
Sub tarray(d As Variant, ByVal Count As Long, ByVal eoi As Integer, Status As Integer)
    Dim dd(32767) As Integer
    Dim xx As Long
    Dim yy As Long
#If Win32 Then
    Dim stl As Long
    If (Count And 1) = 0 Then yy = Count / 2 Else yy = Count / 2 + 1
    For xx = 1 To yy: dd(xx) = d(xx): Next xx
    Call IEtarray(dd(1), Count, eoi, stl)
    Status = stl
#Else
    If (Count And 1) = 0 Then yy = Count / 2 Else yy = Count / 2 + 1
    For xx = 1 To yy: dd(xx) = d(xx): Next xx
    Call IEtarray(dd(1), Count, eoi, Status)
#End If
End Sub
'-------------------------------------------------------
' IMPORTANT NOTE:
'       To call tarray and rarray in VB 4, you must pass the
'       entire array, as in:
'               call tarray(d(),100,1,status%)
'       You CANNOT pass d(1), as was done in previous versions
'       of VB.  Microsoft changed the data types and argument
'       conventions, so this source code change is required.
'
Sub rarray(d As Variant, ByVal Count As Long, l As Integer, Status As Integer)
    Dim dd(32767) As Integer
    Dim xx As Long
    Dim yy As Long
#If Win32 Then
    Dim stl As Long
    Dim ll As Long
    Call IErarray(dd(1), Count, ll, stl)
    l = ll
    Status = stl
#Else
    Call IErarray(dd(1), Count, l, Status)
#End If
    If (l And 1) = 0 Then yy = l / 2 Else yy = l / 2 + 1
    For xx = 1 To yy: d(xx) = dd(xx): Next xx
End Sub



