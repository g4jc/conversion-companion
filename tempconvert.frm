VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Temperature Conversion Tool"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Celsius to Fahrenheit"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Fahrenheit to Celsius"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Value           =   -1  'True
      Width           =   1815
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Slide me to set input temperature!"
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      Min             =   -451
      Max             =   451
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      ToolTipText     =   "Converted Temperature"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Converted Temperature:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "451"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "-451"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Set Temperature"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Setup Global Variables
Dim F2C As Boolean
Dim FTemp As Integer
Dim CTemp As Integer


Private Sub Form_Load()

    ' Set Defaults on Load
    F2C = True
    Text1.Text = "-18°C"

End Sub

Private Sub Option1_Click()
    F2C = True
End Sub

Private Sub Option2_Click()
    F2C = False
End Sub

Private Sub Slider1_Change()
    If (Not F2C) Then
        ' --- Celcius to Fahrenheit Conversion ---
        FTemp = (Slider1.Value * 9 / 5) + 32
        Text1.Text = FTemp & "°F"
        
        If FTemp = 451 Then
            Form1.Caption = "It was a pleasure to burn!"
        Else
            Form1.Caption = "Temperature Conversion Tool"
        End If
        
    Else
        ' --- Fahrenheit to Celcius ---
        CTemp = (Slider1.Value - 32) * 5 / 9
        Text1.Text = CTemp & "°C"
        
        If Slider1.Value = 451 Then
            Form1.Caption = "It was a pleasure to burn!"
        Else
            Form1.Caption = "Temperature Conversion Tool"
        End If
        
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    ' Force user to only input numbers in Text Feild
    Dim Keychar As String
        ' Allow -
        If KeyAscii = 45 Then
            Keychar = Chr(KeyAscii)
        ' Block everything else except for Numerals
        ElseIf KeyAscii > 31 Then
        Keychar = Chr(KeyAscii)
            If Not IsNumeric(Keychar) Then
                KeyAscii = 0
        End If
    End If
    
    ' If user presses Enter
    If KeyAscii = 13 Then
        ' --- Begin Sanitize Input --- '
        ' Check for Double Negatives and Abort
        If (InStr(Text1.Text, "--") > 0) Then
            MsgBox "Not a valid input.", vbCritical, "Error"

        ' Check for Double Negatives and Abort
        ElseIf (InStr(Text1.Text, "°") > 0) Then
            Text1.Text = Replace(Text1.Text, "°", "", 1)


        ' Check for Double Negatives and Abort
        ElseIf (InStr(Text1.Text, "F") > 0) Then
            Text1.Text = Replace(Text1.Text, "F", "", 1)

        ' Check for Double Negatives and Abort
        ElseIf (InStr(Text1.Text, "C") > 0) Then
            Text1.Text = Replace(Text1.Text, "C", "", 1)

        '' Remove Trailing -
        Do While Right(Text1.Text, 1) = "-"
            Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
        Loop

        '' TODO: Handle '-1-8' (Middle - )

        ' Check for Greater Than Limit and Abort
        ElseIf Text1.Text > 451 Then
            MsgBox "You can't calculate a number exceeding 451.", vbCritical, "Error"
            
        ' Check for Less Than Limit and Abort
        ElseIf Text1.Text < -451 Then
            MsgBox "You can't calculate a number less than -451.", vbCritical, "Error"
            
        ' --- End Sanitize Input --- '
        ' Allow user to type filtered value and calculate
        Else
            Slider1.Value = Text1.Text
        End If
    End If
End Sub

