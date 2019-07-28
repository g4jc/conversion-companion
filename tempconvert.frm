VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Temperature Conversion Tool"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox InTemp 
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0°F"
      Top             =   1320
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Celsius to Fahrenheit"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Fahrenheit to Celsius"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Value           =   -1  'True
      Width           =   1815
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "Slide me to set input temperature!"
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      Min             =   -500
      Max             =   500
   End
   Begin VB.TextBox OutTemp 
      Height          =   375
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "-18°C"
      ToolTipText     =   "Converted Temperature"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "="
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Converted Temperature:"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "500"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "-500"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Set Temperature"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Temperature Conversion Tool in VB6
' Licensed CC0

Option Explicit

' Setup Global Variables
Dim F2C As Boolean
Dim FTemp As Integer
Dim CTemp As Integer


Private Sub Form_Load()
    ' Set Defaults on Load
    F2C = True
    InTemp.ForeColor = vbBlue
    OutTemp.ForeColor = vbBlue
End Sub

Private Sub Option1_Click()
    F2C = True
    TempCalculate
End Sub

Private Sub Option2_Click()
    F2C = False
    TempCalculate
End Sub

Public Sub TempCalculate()
' Public Calculation Function to be called
' When user clicks Radio Buttons Or Moves Slider
    If (Not F2C) Then
        ' --- Celcius to Fahrenheit Conversion ---
        FTemp = (Slider1.Value * 9 / 5) + 32
        InTemp.Text = Slider1.Value & "°C"
        OutTemp.Text = FTemp & "°F"
        
        ' --- Extra GUI Hacks ---
        ' With Respect To Ray Bradbury
        If FTemp = 451 Then
            Form1.Caption = "It was a pleasure to burn!"
        ElseIf FTemp = -459 Then
            Form1.Caption = "Absolute Zero!"
        Else
            Form1.Caption = "Temperature Conversion Tool"
        End If
        
        ' Freeze or Boil Text Color Modifer
        If Slider1.Value <= 0 Then
            InTemp.ForeColor = vbBlue
            OutTemp.ForeColor = vbBlue
        ElseIf Slider1.Value >= 100 Then
            InTemp.ForeColor = vbRed
            OutTemp.ForeColor = vbRed
        Else
            InTemp.ForeColor = vbBlack
            OutTemp.ForeColor = vbBlack
        End If
        ' -----------------------
    Else
        ' --- Fahrenheit to Celcius ---
        CTemp = (Slider1.Value - 32) * 5 / 9
        InTemp.Text = Slider1.Value & "°F"
        OutTemp.Text = CTemp & "°C"

        ' --- Extra GUI Hacks ---
        ' With Respect To Ray Bradbury
        If CTemp = 233 Then
            Form1.Caption = "It was a pleasure to burn!"
        ElseIf CTemp = -273 Then
            Form1.Caption = "Absolute Zero!"
        Else
            Form1.Caption = "Temperature Conversion Tool"
        End If
        
        ' Freeze or Boil Text Color Modifer
        If Slider1.Value <= 32 Then
            InTemp.ForeColor = vbBlue
            OutTemp.ForeColor = vbBlue
        ElseIf Slider1.Value >= 212 Then
            InTemp.ForeColor = vbRed
            OutTemp.ForeColor = vbRed
        Else
            InTemp.ForeColor = vbBlack
            OutTemp.ForeColor = vbBlack
        End If
        ' -----------------------
    End If
End Sub

Private Sub Slider1_KeyDown(KeyCode As Integer, Shift As Integer)
    TempCalculate
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    TempCalculate
End Sub



