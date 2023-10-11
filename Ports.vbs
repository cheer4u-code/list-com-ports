strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity WHERE PNPClass = 'Ports'")
Set ports = CreateObject("System.Collections.ArrayList")
For Each objItem in colItems
    ports.Add(objItem.Name)
Next
If ports.Count = 0 Then
    msg = "The port was not found."
Else
    msg = Join(ports.ToArray(), vbCrLf)
End If
Wscript.Echo msg
