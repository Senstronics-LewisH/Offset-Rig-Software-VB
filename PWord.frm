VERSION 5.00
Begin VB.Form PassForm 
   Caption         =   "Team Leader Cable Counter Reset"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8040
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdpword 
      Caption         =   "Enter"
      Height          =   855
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox pbox 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TEAM LEADER - PLEASE ENTER PASSWORD BELOW AND PRESS ENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "PassForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdpword_Click()

Const PassWord As String = "sens666"
         
    If PassForm.pbox.Text = PassWord Then
     MainForm.NumberOfUsesDisplay = 0
     PassForm.pbox.Text = ""
     PassForm.Hide
     MainForm.Show
    End If

End Sub


