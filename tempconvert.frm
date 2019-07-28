VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversion Companion"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2865
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5054
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Welcome"
      TabPicture(0)   =   "tempconvert.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Liquids"
      TabPicture(1)   =   "tempconvert.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CalculateBtn"
      Tab(1).Control(1)=   "OutOz"
      Tab(1).Control(2)=   "InOz"
      Tab(1).Control(3)=   "Fluids"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Measurement"
      TabPicture(2)   =   "tempconvert.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Inches"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Temperature"
      TabPicture(3)   =   "tempconvert.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label3"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label4"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label5"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "LabelStatus"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Slider1"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "OutTemp"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Option1"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Option2"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "InTemp"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).ControlCount=   11
      Begin VB.Frame Inches 
         Caption         =   "Inches"
         Height          =   855
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   5535
      End
      Begin VB.CommandButton CalculateBtn 
         Caption         =   "Calculate"
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox OutOz 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "1.04"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox InOz 
         Height          =   375
         Left            =   -74760
         TabIndex        =   13
         Text            =   "1"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox InTemp 
         Height          =   375
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0°F"
         Top             =   1920
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Celsius to Fahrenheit"
         Height          =   255
         Left            =   -71160
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fahrenheit to Celsius"
         Height          =   255
         Left            =   -74880
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox OutTemp 
         Height          =   375
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "-18°C"
         ToolTipText     =   "Converted Temperature"
         Top             =   1920
         Width           =   735
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   -74280
         TabIndex        =   5
         ToolTipText     =   "Slide me to set input temperature!"
         Top             =   1080
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   393216
         Min             =   -500
         Max             =   500
      End
      Begin VB.Frame Fluids 
         Caption         =   "Fluid Ounces"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   5535
         Begin VB.TextBox OutmL 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "29.57"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "mL"
            Height          =   255
            Left            =   3840
            TabIndex        =   21
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "UK fl oz"
            Height          =   255
            Left            =   2400
            TabIndex        =   18
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "US fl oz  ="
            Height          =   375
            Left            =   840
            TabIndex        =   17
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Label Label6 
         Caption         =   $"tempconvert.frx":0070
         Height          =   735
         Left            =   840
         TabIndex        =   12
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label LabelStatus 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   -73080
         TabIndex        =   11
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "="
         Height          =   255
         Left            =   -72240
         TabIndex        =   10
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Converted Temperature"
         Height          =   255
         Left            =   -73080
         TabIndex        =   9
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "500"
         Height          =   255
         Left            =   -69840
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "-500"
         Height          =   255
         Left            =   -74640
         TabIndex        =   7
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Set Temperature"
         Height          =   375
         Left            =   -72840
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Companion Conversion Tool v1.0
' Written in VB6.
' Licensed CC0.

Option Explicit

' Setup Global Variables
Dim F2C As Boolean
Dim FTemp As Integer
Dim CTemp As Integer

Private Sub OuncesCalc()
OutOz.Text = InOz.Text * 1.04
OutmL.Text = InOz.Text * 29.57
End Sub

Private Sub CalculateBtn_Click()
OuncesCalc
End Sub


Private Sub Form_Load()
    ' Set Defaults on Load
    F2C = True
    InTemp.ForeColor = vbBlue
    OutTemp.ForeColor = vbBlue
    LabelStatus.Caption = "Freezing!"
End Sub

Private Sub InOz_KeyPress(KeyAscii As Integer)
    ' Force user to only input numbers in Text Feild
    Dim Keychar As String
        ' Allow Enter Key
        If KeyAscii = 13 Then
            OuncesCalc
        ' Block everything else except for Numerals
        ElseIf KeyAscii > 31 Then
        Keychar = Chr(KeyAscii)
            If Not IsNumeric(Keychar) Then
                KeyAscii = 0
        End If
    End If
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
        
        ' -----------------------
    Else
        ' --- Fahrenheit to Celcius ---
        CTemp = (Slider1.Value - 32) * 5 / 9
        InTemp.Text = Slider1.Value & "°F"
        OutTemp.Text = CTemp & "°C"

        ' -----------------------
    End If
    
        ' Freeze or Boil Text Color Modifers When F2C
        If (Slider1.Value = 451) And (F2C) Then
            LabelStatus.Caption = "It was a pleasure to burn!"
        ElseIf (Slider1.Value = -459) And (F2C) Then
            LabelStatus.Caption = "Absolute Zero!"
        ElseIf (Slider1.Value <= 32) And (F2C) Then
            InTemp.ForeColor = vbBlue
            OutTemp.ForeColor = vbBlue
            LabelStatus.Caption = "Freezing!"
        ElseIf (Slider1.Value >= 212) And (F2C) Then
            InTemp.ForeColor = vbRed
            OutTemp.ForeColor = vbRed
            LabelStatus.Caption = "Boiling!"
        ' Freeze or Boil Text Color Modifers When -NOT- F2C
        ElseIf (Slider1.Value = 233) And (Not F2C) Then
            LabelStatus.Caption = "It was a pleasure to burn!"
        ElseIf (Slider1.Value = -273) And (Not F2C) Then
            LabelStatus.Caption = "Absolute Zero!"
        ElseIf (Slider1.Value <= 0) And (Not F2C) Then
            InTemp.ForeColor = vbBlue
            OutTemp.ForeColor = vbBlue
            LabelStatus.Caption = "Freezing!"
        ElseIf (Slider1.Value >= 100) And (Not F2C) Then
            InTemp.ForeColor = vbRed
            OutTemp.ForeColor = vbRed
            LabelStatus.Caption = "Boiling!"

        Else
            InTemp.ForeColor = vbBlack
            OutTemp.ForeColor = vbBlack
            LabelStatus.Caption = ""
        End If
End Sub


Private Sub Slider1_KeyDown(KeyAscii As Integer, Shift As Integer)
    If KeyAscii = 40 Then ' Down
    Slider1.Value = Slider1.Value - 5
    TempCalculate
    ElseIf KeyAscii = 38 Then ' Up
    Slider1.Value = Slider1.Value + 5
    TempCalculate
    ElseIf KeyAscii = 37 Then ' Left
    Slider1.Value = Slider1.Value - 1
    TempCalculate
    ElseIf KeyAscii = 39 Then ' Right
    Slider1.Value = Slider1.Value + 1
    TempCalculate
    End If
End Sub


Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    TempCalculate

End Sub
