Sub include(fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

include "QualysAPI.vbs"

'Proxy to reach Qualys API
Proxy = ""

'User for Qualys API
quser = ""
qpass = ""

'Prefix to save the report
pathprefix="."

'-------------------- Hosts in Qualys Subscription ever scanned ------------------------
dim h, assets, fecha, filename, fso, outfile
set assets = listAssetsHosts()
fecha = replace(toQualysTime(now()),":","-")
filename = pathprefix & "\Assets_" & fecha & ".csv"

Const ForWriting = 2
Set fso = CreateObject("Scripting.FileSystemObject")
Set outfile = fso.OpenTextFile( filename, ForWriting, True)

outfile.WriteLine "ASSETID;IP;DNS;NETBIOS;OS;LASTVMSCAN;LASTPCSCAN;COMMENTS;TRACKINGMETHOD"
for each h in assets
	outfile.WriteLine h.assetid & ";" & h.ip & ";" & h.name & ";" & h.netbios & ";" & h.os & ";" & h.lastVMScan & ";" & h.lastPCScan & ";" & h.comments & ";" & h.trackmethod
next

outfile.Close
