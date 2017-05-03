'VBScript to write "KillPower" string value to the registry in Run
'by Philip Simonson

Public Sub Main()
	Dim Res
	Dim Input
	Input = InputBox("Enter file path to add into Registry (Run)?")
	Res = MsgBox("Write key value in registry?" & vbCrLf & "(If you choose no, deletes value)", vbYesNoCancel)
	If Res = vbYes Then
		WriteValue ".", Input
	ElseIf Res = vbNo Then
		DeleteValue "."
	End If
End Sub

Const HKEY_CURRENT_USER = &H8000001

Private Sub WriteValue(ByVal strComputer, ByVal strInput)
	'Write key value to registry
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Windows\CurrentVersion\Run"
	strValueName = "KillPower"
	objRegistry.SetStringValue HKEY_CURRENT_USER, strKeyPath, strValueName, strInput
End Sub

Private Sub DeleteValue(ByVal strComputer)
	'Delete key value in registry
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Windows\CurrentVersion\Run"
	strValueName = "KillPower"
	objRegistry.DeleteValue HKEY_CURRENT_USER, strKeyPath, strValueName
End Sub

Call Main