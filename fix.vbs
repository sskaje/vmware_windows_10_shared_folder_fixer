'
' VMware Fusion/Workstation Shared Folder Fixer
' 
' Author: sskaje
'
'

' Ask for Administrator Access
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

' Registry related variables
Dim objShell, regNetworkProviderOrder, regValueType, currentValue, vmhgfs, comma
regNetworkProviderOrder = "HKLM\SYSTEM\CurrentControlSet\Control\NetworkProvider\Order\ProviderOrder"
regValueType = "REG_SZ"
vmhgfs = "vmhgfs"
comma = ","

Set objShell = WScript.CreateObject("WScript.Shell")
currentValue = objShell.RegRead(regNetworkProviderOrder)

Dim searchPos, needWriteBack
searchPos = InStr(currentValue, vmhgfs)
needWriteBack = False

If searchPos = 0 Then 
	' If vmhgfs is not found in registry value, prepend 
	currentValue = vmhgfs + comma + currentValue
	needWriteBack = True
ElseIf searchPos <> 1 Then 
	' If vmhgfs is not at the beginning, move it there
	needWriteBack = True
	
	Dim providerSets, x
	providerSets = Split(currentValue, comma)
	currentValue = vmhgfs
	For Each x in providerSets
		Do
			If x = vmhgfs Then Exit Do					' Skip if current is vmhgfs
			If x = "" Then Exit Do						' Skip if current is empty string
			currentValue = currentValue + comma + x		' Append to new value
		Loop While False
	Next
' Else ' If vmhgfs is at the beginning, skip 
End If


If needWriteBack = True Then
	' Write Registry
	objShell.RegWrite regNetworkProviderOrder, currentValue, regValueType
End If

WScript.Echo "Registry has been fixed. Please try Shared Folder."
