' Simple Fahrenheit/Celcius Conversion Script in VBScript
' CC0

Option Explicit

Dim FTemp
Dim CTemp
Dim result

result = MsgBox ("Is Fahrenheit your native temperature unit?", vbYesNo + vbQuestion, "Default Temperature Check")

Select Case result
Case vbYes
	' --- Fahrenheit to Celsius ---
	FTemp=InputBox("Enter Fahrenheit temperature:", "Fahrenheit to Celsius", "451")

	If IsEmpty(FTemp) Then
		' Abort if users clicks Cancle
		WScript.Quit 1
	Else
		If FTemp = "" Then
			' Warn and abort if user has entered an empty string
			MsgBox "You Must Enter a number! Exiting.", vbCritical, "Error!"
			WScript.Quit 1
	        End If

		' The Math Forumula to Convert Fahrenheit to Celcius
		CTemp = (FTemp - 32) * 5 / 9

		If (FTemp = 451) Then
			InputBox "Your Fahrenheit temperature in Celsius is: ", "It was a pleasure to burn!", CTemp & "°C"
		Else

			InputBox "Your Fahrenheit temperature in Celsius is: ", "Fahrenheit to Celsius Converted!", CTemp & "°C"
		End If
	End If
Case vbNo
	' --- Celsius to Fahrenheit ---
	CTemp=InputBox("Enter Celsius temperature:", "Celsius to Fahrenheit", "0")

	If IsEmpty(CTemp) Then
		' Abort if users clicks Cancle
		WScript.Quit 1
	Else
		If CTemp = "" Then
			' Warn and abort if user has entered an empty string
			MsgBox "You Must Enter a number! Exiting.", vbCritical, "Error!"
			WScript.Quit 1
		End If

		' The Math Forumula to Convert Celcius to Fahrenheit
		FTemp = (CTemp * 9 / 5) + 32

		InputBox "Your Celsius temperature in Fahrenheit is: ", "Celsius to Fahrenheit Converted!", FTemp & "°F"
	End If
End Select
