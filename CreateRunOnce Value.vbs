'VBScript to write a key to the registry in RunOnce
'by Philip Simonson

Public Sub Main()
	Dim Res
	Res = MsgBox("Write key value in registry?" & vbCrLf & "(If you choose no, deletes value)", vbYesNoCancel)
	If Res = vbYes Then
		WriteValue(".")
	ElseIf Res = vbNo Then
		DeleteValue(".")
	End If
End Sub

Const HKEY_CURRENT_USER = &H8000001

Private Sub WriteValue(ByVal strComputer)
	'Write key value to registry
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Windows\CurrentVersion\RunOnce"
	strValueName = "KillPower"
	strValue = "C:\Users\phili\Desktop\KillPower.exe"
	objRegistry.SetStringValue HKEY_CURRENT_USER, strKeyPath, strValueName, strValue
End Sub

Private Sub DeleteValue(ByVal strComputer)
	'Delete key value in registry
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Windows\CurrentVersion\RunOnce"
	strValueName = "KillPower"
	objRegistry.DeleteValue HKEY_CURRENT_USER, strKeyPath, strValueName
End Sub

Call Main