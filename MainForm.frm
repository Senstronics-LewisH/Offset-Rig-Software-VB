VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form MainForm 
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15450
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H0080FFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15450
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox PartCount 
      Height          =   375
      Left            =   240
      TabIndex        =   101
      Text            =   "1"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox PartNumber 
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   3360
      Width           =   5415
   End
   Begin VB.TextBox ThirdMCSBarcode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   2760
      Width           =   5415
   End
   Begin VB.TextBox SecondMCSBarcode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2160
      Width           =   5415
   End
   Begin VB.TextBox FirstMCSBarcode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   1560
      Width           =   5415
   End
   Begin VB.TextBox WorksOrderBarcode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   960
      Width           =   5415
   End
   Begin VB.TextBox ProgramDisplay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   97
      Top             =   9360
      Width           =   1815
   End
   Begin VB.CheckBox DisablePrinterCheck 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   12360
      TabIndex        =   93
      Top             =   120
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox PercentBox 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   92
      Text            =   "0"
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open Excel File"
      Height          =   735
      Left            =   13920
      TabIndex        =   91
      Top             =   9960
      Width           =   975
   End
   Begin VB.TextBox FailedBox 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   90
      Text            =   "0"
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox PassedBox 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   89
      Text            =   "0"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox PODCheck 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   14520
      TabIndex        =   82
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox RigDisplay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   81
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton ResetCableUsage 
      Caption         =   "Reset Cable Usage"
      Height          =   735
      Left            =   10560
      TabIndex        =   79
      Top             =   9960
      Width           =   1095
   End
   Begin VB.FileListBox FileList 
      Height          =   1260
      Left            =   600
      TabIndex        =   74
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CheckDetailsButton 
      Caption         =   "Check Details"
      Height          =   735
      Left            =   11760
      TabIndex        =   71
      Top             =   9960
      Width           =   975
   End
   Begin VB.CheckBox ScanAllCheck 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   13560
      TabIndex        =   70
      Top             =   120
      Width           =   255
   End
   Begin VB.Timer InputTimer 
      Interval        =   100
      Left            =   720
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Previous"
      Height          =   495
      Left            =   6360
      TabIndex        =   56
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   495
      Left            =   7320
      TabIndex        =   55
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox SensorID 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   53
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Test1 
      Caption         =   "Test 1"
      Height          =   495
      Left            =   360
      TabIndex        =   99
      Top             =   9600
      Width           =   615
   End
   Begin VB.CommandButton Test2 
      Caption         =   "Test 2"
      Height          =   495
      Left            =   1080
      TabIndex        =   45
      Top             =   9600
      Width           =   615
   End
   Begin VB.CommandButton ClearDownButton 
      Caption         =   "Clear Down Values"
      Height          =   735
      Left            =   12840
      TabIndex        =   0
      Top             =   9960
      Width           =   975
   End
   Begin MSCommLib.MSComm VisionComm 
      Left            =   120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   5000
      InputLen        =   500
      OutBufferSize   =   0
      RTSEnable       =   -1  'True
      EOFEnable       =   -1  'True
   End
   Begin MSCommLib.MSComm DMMComm 
      Left            =   1200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      Handshaking     =   1
      InBufferSize    =   100
      InputLen        =   100
      OutBufferSize   =   100
      RThreshold      =   100
      RTSEnable       =   -1  'True
      SThreshold      =   58
      EOFEnable       =   -1  'True
   End
   Begin VB.Label PartNumberLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SCAN PART NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   100
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Image VOUT2SWPASS 
      Height          =   480
      Left            =   7680
      Picture         =   "MainForm.frx":0000
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   480
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vision Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   98
      Top             =   9480
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Image RestrictorWeldedPass 
      Height          =   480
      Left            =   3840
      Picture         =   "MainForm.frx":2CF4
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image UnionPass 
      Height          =   480
      Left            =   5760
      Picture         =   "MainForm.frx":59E8
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image RestrictorWeldedFail 
      Height          =   480
      Left            =   3840
      Picture         =   "MainForm.frx":86DC
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   480
   End
   Begin VB.Label RestrictorWeldedLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Restrictor Welded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   96
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Label CorrectUnionLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Union"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   95
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Diasable Printer"
      Enabled         =   0   'False
      Height          =   255
      Left            =   11160
      TabIndex        =   94
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Image EnableStart 
      Height          =   480
      Left            =   600
      Picture         =   "MainForm.frx":A52D
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PASSED   FAILED   % Scrap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   88
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Image OffDifPass 
      Height          =   480
      Left            =   10320
      Picture         =   "MainForm.frx":D221
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image OffDifFail 
      Height          =   480
      Left            =   10320
      Picture         =   "MainForm.frx":FF15
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   480
   End
   Begin VB.Label VoutDiffDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   87
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label VoutDiffLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Off2-Off1 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   86
      Top             =   6360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Vout1OriginalOutputDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   85
      Top             =   5760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Vout1OriginalOutputLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Off1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   84
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image FAILED 
      Height          =   2685
      Left            =   7920
      Picture         =   "MainForm.frx":11D66
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   2715
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "POD"
      Height          =   255
      Left            =   13800
      TabIndex        =   83
      Top             =   120
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "RIG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   80
      Top             =   8880
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Uses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   78
      Top             =   8325
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label NumberOfUsesDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   77
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Label CableNumberDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   76
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cable Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   75
      Top             =   7845
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "O-Ring Colour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   73
      Top             =   7365
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label ORingColourDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   72
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Scan All"
      Height          =   255
      Left            =   12600
      TabIndex        =   69
      Top             =   120
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Label UnionCodeDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   66
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Union Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   67
      Top             =   6000
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Image OringPass 
      Height          =   480
      Left            =   3840
      Picture         =   "MainForm.frx":13BB7
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image OringFail 
      Height          =   480
      Left            =   3840
      Picture         =   "MainForm.frx":168AB
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   480
   End
   Begin VB.Label RestrictorLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Restrictor Present"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   65
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Label OringLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "O-Ring"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   64
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label ORingDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   63
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label RestrictorDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   62
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Is Restrictor Required?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   61
      Top             =   6960
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Is O-Ring Required?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   60
      Top             =   6480
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label OFFSETDIFFDISPLAY 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   59
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label OffsetFromCalDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   58
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label OFFSETFROMCALLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Offset From Calibration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   57
      Top             =   7680
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SENSOR ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   54
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label InsulationDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   52
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Label CurrentDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   50
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Label CurrentLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   49
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label VOUT2TEMPLABEL 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vout 2 Temp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11600
      TabIndex        =   48
      Top             =   3600
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label VOUT2TempDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   47
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label VOUT2OUTPUTERRORDISPLAY 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   44
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label VOUT1OUTPUTERRORDISPLAY 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   43
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label VOUT2LimitDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   41
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label ConnectorTypeDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   40
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label ConnectorLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Connector Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11160
      TabIndex        =   39
      Top             =   4500
      Width           =   1935
      WordWrap        =   -1  'True
   End
   Begin VB.Label STCTargetDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   35
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label STCOUTPUTLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "4th Wire STC Output"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   38
      Top             =   5520
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label VOUT2TargetDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   37
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label VOUT2TARGETLABEL 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vout 2 Target"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11450
      TabIndex        =   36
      Top             =   3120
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label BoardTypeDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   34
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label BoardTypeLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Board Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   33
      Top             =   5000
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelScanWOLAbel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SCAN WORKS ORDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   32
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label LimitPercentDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   31
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label VOUT1ERRORLABEL 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vout 1 Error Limit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   30
      Top             =   1680
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label SCANMCS3Label 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SCAN 3RD MCS BAR CODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label VOUT2OutputDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   28
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label VOUT1OutputDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   27
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label STCVOUT2Display 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   26
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label STCVOUT1Display 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   25
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label STCGNDDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   24
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label STCVSDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   23
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Image VOUT2PASS 
      Height          =   480
      Left            =   7080
      Picture         =   "MainForm.frx":186FC
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   480
   End
   Begin VB.Image STCGNDPASS 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":1B3F0
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image STCVOUT1PASS 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":1E0E4
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image STCVOUT2PASS 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":20DD8
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image VOUT1PASS 
      Height          =   480
      Left            =   7080
      Picture         =   "MainForm.frx":23ACC
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image STCVSPASS 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":267C0
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label VOUT2OUTPUTLABEL 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vout 2 Output"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   22
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label VOUT1OUTPUTLABEL 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vout 1 Output"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   21
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label STCVOUT2LABEL 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "STC Vout 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label STCVOUT1Label 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "STC Vout 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   19
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label STCGNDLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "STC GND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   18
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label STCVSLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "STC VS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   17
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label LoadTypeDisplay 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   12960
      TabIndex        =   16
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label LoadValueDisplay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   11880
      TabIndex        =   15
      Top             =   480
      Width           =   975
   End
   Begin VB.Label SCANMCS1Label 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SCAN 1St MCS BAR CODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label OffsetTargetDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   13080
      TabIndex        =   13
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vout 1 Target"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11450
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label UpperLimitDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label LowerLimitDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13080
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vout 1 Upper Limit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   8
      Top             =   2640
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vout 1 Lower Limit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
      WordWrap        =   -1  'True
   End
   Begin VB.Label SCANMCS2Label 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SCAN 2ND MCS BAR CODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label TitleLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "POST CALIBRATION OFFSET CHECKER"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   -600
      TabIndex        =   4
      Top             =   0
      Width           =   12015
   End
   Begin VB.Image CurrentPass 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":294B4
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   480
   End
   Begin VB.Image InsulationPass 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":2C1A8
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   480
   End
   Begin VB.Label InsulationLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Insulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   51
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label LoadLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   46
      Top             =   600
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Label VOUT2ERRORLIMITLABEL 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vout 2 Error Limit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11100
      TabIndex        =   42
      Top             =   4080
      Width           =   1935
      WordWrap        =   -1  'True
   End
   Begin VB.Image STCGNDFAIL 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":2EE9C
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image STCVOUT2FAIL 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":30CED
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image VOUT2FAIL 
      Height          =   480
      Left            =   7080
      Picture         =   "MainForm.frx":32B3E
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   480
   End
   Begin VB.Image OFFSETFROMCALPASS 
      Height          =   480
      Left            =   7080
      Picture         =   "MainForm.frx":3498F
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image CurrentFail 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":37683
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   480
   End
   Begin VB.Image RestrictorPass 
      Height          =   480
      Left            =   5760
      Picture         =   "MainForm.frx":394D4
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image STCVSFAIL 
      Height          =   465
      Left            =   5640
      Picture         =   "MainForm.frx":3C1C8
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image STCVOUT1FAIL 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":3E019
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image VOUT1FAIL 
      Height          =   480
      Left            =   7080
      Picture         =   "MainForm.frx":3FE6A
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image OFFSETFROMCALFAIL 
      Height          =   480
      Left            =   7080
      Picture         =   "MainForm.frx":41CBB
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image RestrictorFail 
      Height          =   480
      Left            =   5760
      Picture         =   "MainForm.frx":43B0C
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   480
   End
   Begin VB.Image InsulationFail 
      Height          =   480
      Left            =   5640
      Picture         =   "MainForm.frx":4595D
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   480
   End
   Begin VB.Label VoltageLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Voltage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   68
      Top             =   8280
      Width           =   975
   End
   Begin VB.Image PASSED 
      Height          =   2655
      Left            =   7920
      Picture         =   "MainForm.frx":477AE
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Image UnionFail 
      Height          =   480
      Left            =   5760
      Picture         =   "MainForm.frx":4A4A2
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image VOUT2SWFAIL 
      Height          =   480
      Left            =   7680
      Picture         =   "MainForm.frx":4C2F3
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   480
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClearDownButton_Click()
    CheckForUpdates

    Dim WshShell As Object
     
    Set WshShell = CreateObject("WScript.Shell")
     
    ClearDown
    DetailChecked = False
    WshShell.SendKeys "{Tab}"
End Sub
Private Sub Command3_Click()

    OpenExcelFile

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyF1
        
        Outputs.Show
        
    Case vbKeyF2
        If DevMode Then
            If DetailChecked = True Then
                START
            Else
                MsgBox "Please scan or enter Works Order and Barcodes first (DetailChecked must be True)."
            End If
        End If
        
    End Select

End Sub

Private Sub ScanAllCheck_Click()

    If ScanAllCheck = 1 Then
        WorksOrderBarcode.Enabled = True
        FirstMCSBarcode.Enabled = True
        SecondMCSBarcode.Enabled = True
        ThirdMCSBarcode.Enabled = True
        PartNumber.Enabled = True
        SCANMCS1Label.Enabled = True
        SCANMCS2Label.Enabled = True
        SCANMCS3Label.Enabled = True
        PartNumberLabel.Enabled = True
    End If

End Sub
Private Sub Form_KeyPress(keyascii As Integer)
    ProcessKeypress (keyascii)
End Sub
Public Sub ProcessKeypress(ByVal keyascii As Integer)


    If Len(WorksOrderBarcode.Text) >= 5 And keyascii = 13 Then
        PartCount = 1
        WorksOrderBarcode = UCase(WorksOrderBarcode)
        If ScanAllCheck = 0 Then
            CheckDetails
        Else
            CheckDetailsButton.Visible = True
            FirstMCSBarcode.Enabled = True
            SecondMCSBarcode.Enabled = True
            ThirdMCSBarcode.Enabled = True
            PartNumber.Enabled = True
            SCANMCS1Label.Enabled = True
            SCANMCS2Label.Enabled = True
            SCANMCS3Label.Enabled = True
            PartNumberLabel.Enabled = True

            Dim WshShell As Object
             
            Set WshShell = CreateObject("WScript.Shell")
             
                WshShell.SendKeys "{Tab}"
        End If
    End If

End Sub
Private Sub CheckDetailsButton_Click()
    CheckDetails
End Sub
Private Sub InputTimer_Timer()
    CheckInputs
End Sub
Private Sub Command1_Click()
Dim id As Double

    id = SensorID.Text
    id = id + 1
    SensorID.Text = id
    SensorNumber = id

End Sub
Private Sub Command2_Click()
Dim id As Double

    id = SensorID.Text
    id = id - 1
    SensorID.Text = id
    SensorNumber = id

End Sub
Private Sub Form_Load()
    LoadOffsetConfig
    
    ' Programmatically set the version caption immediately so it's correct during initialization/MsgBoxes
    TitleLabel.Caption = "POST CALIBRATION OFFSET CHECKER " & CompileVersion
    Me.Caption = TitleLabel.Caption
        
    InitialisePSU
    InitialiseVision
    InitialiseDMM
    OpenAllSwitches
        
    FindRigType
    ClearDownButton_Click
    MsgBox "Please Confirm Vision Light Is On"
 
'1= Offset with switch and printer
'2= 25 Day Hold
'3= 90 Day Hold
 
If OffsetType = 1 Then
     
    IOCard = True
    UsePrinter = True
    Initialise7250
    OffsetOnly = False
    Post25DayTest = False
    Retests = False
    TitleLabel.Caption = "POST CALIBRATION OFFSET CHECKER " & CompileVersion
    Me.Caption = TitleLabel.Caption
    Vout1OriginalOutputLabel.Visible = False
    Vout1OriginalOutputDisplay.Visible = False
    VoutDiffLabel.Visible = False
    VoutDiffDisplay.Visible = False
 
ElseIf OffsetType = 2 Then

    IOCard = True
    UsePrinter = False
    Initialise7250
    OffsetOnly = False
    Post25DayTest = True
    Retests = False
    MainForm.BackColor = &H80C0FF
    TitleLabel.Caption = "POST 25 Day HOLD CHECK " & CompileVersion
    Me.Caption = TitleLabel.Caption
    Vout1OriginalOutputLabel.Visible = True
    Vout1OriginalOutputDisplay.Visible = True
    VoutDiffLabel.Visible = True
    VoutDiffDisplay.Visible = True
    
ElseIf OffsetType = 3 Then

    IOCard = True
    UsePrinter = False
    Initialise7250
    OffsetOnly = False
    Post25DayTest = False
    Retests = False
    MainForm.BackColor = &H8080FF
    TitleLabel.Caption = "0.4% RETEST SCRAP OFFSET CHECKER " & CompileVersion
    Me.Caption = TitleLabel.Caption
    Vout1OriginalOutputLabel.Visible = False
    Vout1OriginalOutputDisplay.Visible = False
    VoutDiffLabel.Visible = False
    VoutDiffDisplay.Visible = False

End If
    
    DetailChecked = False
    ReadyFlag = True
    WorkOrderOnly = True
    FirstMCSBarcode.Enabled = False
    SecondMCSBarcode.Enabled = False
    ThirdMCSBarcode.Enabled = False
    PartNumber.Enabled = False
    SCANMCS1Label.Enabled = False
    SCANMCS2Label.Enabled = False
    SCANMCS3Label.Enabled = False
    PartNumberLabel.Enabled = False

End Sub
Private Sub Form_Unload(cancel As Integer)

    SetPSU2 0
    SetPSU3 0
    OpenAllSwitches
End Sub

Public Sub EmptyFields()

    STCVSDisplay = ""
    STCGNDDisplay = ""
    STCVOUT1Display = ""
    STCVOUT2Display = ""
    VOUT1OutputDisplay = ""
    VOUT1OUTPUTERRORDISPLAY = ""
    VOUT2OutputDisplay = ""
    VOUT2OUTPUTERRORDISPLAY = ""
    CurrentDisplay = ""
    InsulationDisplay = ""
    OffsetFromCalDisplay = ""
    Vout1OriginalOutputDisplay = ""
    VoutDiffDisplay = ""
    OFFSETDIFFDISPLAY = ""
    PASSED.Visible = False
    FAILED.Visible = False
    OFFSETFROMCALPASS.Visible = False
    OFFSETFROMCALFAIL.Visible = False
    STCVSFAIL.Visible = False
    STCGNDFAIL.Visible = False
    STCVOUT1FAIL.Visible = False
    STCVOUT2FAIL.Visible = False
    VOUT1FAIL.Visible = False
    VOUT2FAIL.Visible = False
    STCVSPASS.Visible = False
    STCGNDPASS.Visible = False
    STCVOUT1PASS.Visible = False
    STCVOUT2PASS.Visible = False
    VOUT2SWPASS.Visible = False
    VOUT2SWFAIL.Visible = False
    VOUT1PASS.Visible = False
    VOUT2PASS.Visible = False
    OffDifPass.Visible = False
    OffDifFail.Visible = False
    InsulationPass.Visible = False
    InsulationFail.Visible = False
    CurrentPass.Visible = False
    CurrentFail.Visible = False
    OFFSETFROMCALPASS.Visible = False
    OFFSETFROMCALFAIL.Visible = False
    RestrictorFail.Visible = False
    RestrictorPass.Visible = False
    OringFail.Visible = False
    OringPass.Visible = False
    RestrictorWeldedFail.Visible = False
    RestrictorWeldedPass.Visible = False
    UnionPass.Visible = False
    UnionFail.Visible = False
    
End Sub
Public Sub ClearDownForOffsetOnly()

    STCVSDisplay.Visible = False
    STCGNDDisplay.Visible = False
    STCVOUT1Display.Visible = False
    STCVOUT2Display.Visible = False
    VOUT2OUTPUTLABEL.Visible = False
    VOUT2OutputDisplay.Visible = False
    VOUT2OUTPUTERRORDISPLAY.Visible = False
    VOUT2TARGETLABEL.Visible = False
    VOUT2TargetDisplay.Visible = False
    VOUT2TempDisplay.Visible = False
    VOUT2LimitDisplay.Visible = False
    VOUT2ERRORLIMITLABEL.Visible = False
    STCTargetDisplay.Visible = False
    STCVOUT2LABEL.Visible = False
    STCVSLabel.Visible = False
    STCVOUT1Label.Visible = False
    STCGNDLabel.Visible = False
    STCOUTPUTLabel.Visible = False
End Sub
Public Sub ClearDownForPackOnly()

    LabelScanWOLAbel.Visible = False
    VOUT2TEMPLABEL.Visible = False
    LoadValueDisplay.Visible = False
    LoadTypeDisplay.Visible = False
    LoadLabel.Visible = False
    VOUT2TARGETLABEL.Visible = False
    VOUT2TargetDisplay.Visible = False
    VOUT2TempDisplay.Visible = False
    VOUT2ERRORLIMITLABEL.Visible = False
    VOUT2LimitDisplay.Visible = False
    ConnectorLabel.Visible = False
    ConnectorTypeDisplay.Visible = False
    BoardTypeLabel.Visible = False
    BoardTypeDisplay.Visible = False
    STCOUTPUTLabel.Visible = False
    STCTargetDisplay.Visible = False
    VOUT1ERRORLABEL.Visible = False
    LimitPercentDisplay.Visible = False
    VOUT1OUTPUTERRORDISPLAY.Visible = False

    
End Sub
Public Sub ClearDown()

Dim i As Integer
    
    For i = 1 To 2000       'increased from 800 by DW
        SensorStatus(i) = PASSED
    Next

    MainForm.PercentBox.BackColor = &HFFFFFF
    PASSED.Visible = False
    FAILED.Visible = False
    STCVSFAIL.Visible = False
    STCGNDFAIL.Visible = False
    STCVOUT1FAIL.Visible = False
    STCVOUT2FAIL.Visible = False
    VOUT1FAIL.Visible = False
    VOUT2FAIL.Visible = False
    STCVSPASS.Visible = False
    STCGNDPASS.Visible = False
    STCVOUT1PASS.Visible = False
    STCVOUT2PASS.Visible = False
    VOUT2SWPASS.Visible = False
    VOUT2SWFAIL.Visible = False
    OffDifPass.Visible = False
    OffDifFail.Visible = False
    VOUT1PASS.Visible = False
    VOUT2PASS.Visible = False
    CurrentFail.Visible = False
    CurrentPass.Visible = False
    InsulationPass.Visible = False
    InsulationFail.Visible = False
    OringPass.Visible = False
    OringFail.Visible = False
    RestrictorFail.Visible = False
    RestrictorPass.Visible = False
    LoadLabel.Visible = True
    ConnectorLabel.Visible = True
    ConnectorTypeDisplay.Visible = True
    BoardTypeLabel.Visible = True
    BoardTypeDisplay.Visible = True
    STCOUTPUTLabel.Visible = True
    STCVSDisplay.Visible = True
    STCGNDDisplay.Visible = True
    STCVOUT1Display.Visible = True
    STCVOUT2Display.Visible = True
    VOUT2OUTPUTLABEL.Visible = True
    VOUT2OutputDisplay.Visible = True
    VOUT2OUTPUTERRORDISPLAY.Visible = True
    VOUT2TARGETLABEL.Visible = True
    VOUT2TargetDisplay.Visible = True
    VOUT1OUTPUTLABEL.Visible = True
    VOUT1OutputDisplay.Visible = True
    VOUT2TempDisplay.Visible = True
    VOUT2LimitDisplay.Visible = True
    VOUT2ERRORLIMITLABEL.Visible = True
    STCTargetDisplay.Visible = True
    STCVOUT2LABEL.Visible = True
    STCVSLabel.Visible = True
    STCVOUT1Label.Visible = True
    STCGNDLabel.Visible = True
    STCOUTPUTLabel.Visible = True
    LabelScanWOLAbel.Visible = True
    SCANMCS1Label.Visible = True
    SCANMCS2Label.Visible = True
    SCANMCS3Label.Visible = True
    PartNumber.Visible = True
    WorksOrderBarcode.Visible = True
    FirstMCSBarcode.Visible = True
    SecondMCSBarcode.Visible = True
    ThirdMCSBarcode.Visible = True
    PartNumber.Visible = True
    VOUT1ERRORLABEL.Visible = True
    LimitPercentDisplay.Visible = True
    VOUT1OUTPUTERRORDISPLAY.Visible = True
    CurrentDisplay.Visible = False
    CurrentLabel.Visible = False
    VoltageLabel.Visible = False
    InsulationDisplay.Visible = False
    InsulationLabel.Visible = False
    OFFSETFROMCALLabel.Visible = False
    OffsetFromCalDisplay.Visible = False
    OFFSETDIFFDISPLAY.Visible = False
    OFFSETFROMCALPASS.Visible = False
    OFFSETFROMCALFAIL.Visible = False
    RestrictorFail.Visible = False
    RestrictorPass.Visible = False
    OringFail.Visible = False
    OringPass.Visible = False
    RestrictorWeldedFail.Visible = False
    RestrictorWeldedPass.Visible = False
    UnionPass.Visible = False
    UnionFail.Visible = False
    
    WorksOrderBarcode = ""
    FirstMCSBarcode = ""
    SecondMCSBarcode = ""
    ThirdMCSBarcode = ""
    PartNumber = ""
    
    LoadValueDisplay = ""
    OffsetTargetDisplay = ""
    VOUT2TargetDisplay = ""
    VOUT2TempDisplay = ""
    VOUT2LimitDisplay = ""
    LowerLimitDisplay = ""
    UpperLimitDisplay = ""
    LimitPercentDisplay = ""
    BoardTypeDisplay = ""
    ConnectorTypeDisplay = ""
    STCTargetDisplay = ""
    ORingDisplay = ""
    RestrictorDisplay = ""
    UnionCodeDisplay = ""
    OffsetFromCalDisplay = ""
    OFFSETDIFFDISPLAY = ""
    CurrentDisplay = ""
    STCVSDisplay = ""
    STCGNDDisplay = ""
    STCVOUT1Display = ""
    STCVOUT2Display = ""
    VOUT1OutputDisplay = ""
    VOUT1OUTPUTERRORDISPLAY = ""
    VOUT2OutputDisplay = ""
    VOUT2OUTPUTERRORDISPLAY = ""
    ORingColourDisplay = ""
    CableNumberDisplay = ""
    RigDisplay = ""
    LimitPercent = 100000
    Limit = 100000
    Offset = 100000
    FullScale = 100000
    ZeroPressure = 100000
    FSPressure = 100000
    BarcodeUnits = 100000
    BarcodeSpan = 100000
    BarcodeOffset = 100000
    BarcodeFullScale = 100000
    BarcodeZeroPressure = 100000
    BarcodeFSPressure = 100000
    PressureSpan = 100000
    numberofcycles = 0
    PassedBox = 0
    FailedBox = 0
    SensorID = 0
    PercentBox = 0
   

End Sub
Private Sub ResetCableUsage_Click()
    PassForm.Show
    MainForm.Hide
                 
End Sub


Private Sub Test1_Click()

    MainForm.SensorID = 1
    SensorStatus(MainForm.SensorID) = PASSED
    
    MainForm.EmptyFields
    DoEvents
    
    TestVout1and2
    TestSTC
    TestOringAndRest
    
    
    If SensorStatus(1) <> PASSED Then
        MainForm.PASSED.Visible = False
        MainForm.FAILED.Visible = True
        DoEvents
        MsgBox "CONTACT ENGINEERING"
    Else
        MainForm.PASSED.Visible = True
        MainForm.FAILED.Visible = False
        DoEvents
    End If

    
End Sub
Private Sub Test2_Click()

    MainForm.SensorID = 1
    SensorStatus(MainForm.SensorID) = PASSED
    
    MainForm.EmptyFields
    DoEvents
    
    TestSTC2
    TestOringAndRest2
    
    If SensorStatus(1) <> PASSED Then
        MainForm.PASSED.Visible = False
        MainForm.FAILED.Visible = True
        DoEvents
        MsgBox "CONTACT ENGINEERING"
    Else
        MainForm.PASSED.Visible = True
        MainForm.FAILED.Visible = False
        DoEvents
    End If
    
End Sub
