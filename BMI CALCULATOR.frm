VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMI CALCULATOR"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000080&
      Caption         =   "END"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   " CHECK"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      MaskColor       =   &H00FFC0FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox high 
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Text            =   " "
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox weigh 
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label result 
      BackColor       =   &H8000000E&
      Caption         =   " "
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "    HEIGHT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "   WEIGHT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   1410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BMI, weght, heght As Integer
Private Sub Command1_Click()
weght = Val(weigh.Text)
heght = Val(high.Text)
BMI = weght / heght ^ 2
Select Case BMI
Case Is < 18.5
result.Caption = BMI & "-UnderWeight"
Case 18.5 To 24.9
result.Caption = BMI & "-Normal"
Case 25 To 29.9
result.Caption = BMI & "-Overweight"
Case Is >= 30
result.Caption = BMI & "-Obese"
Case Else
result.Caption = "Check Input and Try again"
End Select
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
result.Caption = "Result Displays Here......."
End Sub

