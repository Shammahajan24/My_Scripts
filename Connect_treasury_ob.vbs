Option Explicit

Dim objNetwork, strDriveLetter, strRemotePath, strUserName, strPassword

' Drive letter you want to assign to the network drive
strDriveLetter = "Z:"

' UNC path to the network shared drive
strRemotePath = "\\172.16.5.165\c$\Treasury"

' Username and password for the network share (optional)
strUserName = "TATACAPITAL\88031001"
strPassword = "Welcome@789"

' Create a network object
Set objNetwork = CreateObject("WScript.Network")

' Check if the drive is already mapped, and if so, disconnect it
On Error Resume Next
objNetwork.RemoveNetworkDrive strDriveLetter, True, True
On Error GoTo 0

' Map the network drive
If strUserName = "" Then
    objNetwork.MapNetworkDrive strDriveLetter, strRemotePath
Else
    objNetwork.MapNetworkDrive strDriveLetter, strRemotePath, False, strUserName, strPassword
End If