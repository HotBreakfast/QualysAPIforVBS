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

'------------------ Weekly Detected Hosts by Map scans -----------------------
dim DetectedHosts, fecha, fso, outfile
set DetectedHosts = getLastWeekMapDetectedHosts()
fecha = replace(toQualysTime(now()),":","-")
filename = pathprefix & "\Discovery_" & fecha & ".csv"

Const ForWriting = 2
Set fso = CreateObject("Scripting.FileSystemObject")
Set outfile = fso.OpenTextFile( filename, ForWriting, True)

outfile.WriteLine "IP;DNS;NETBIOS;OS;DISCOVERY"

for each h in DetectedHosts
	dim methods
	methods = ""
	for each m in h.discoveryMethods
		if methods <> "" then
			methods = methods & ", "
		end if
		methods = methods & m
	next
	outfile.WriteLine h.ip & ";" & h.name & ";" & h.netbios & ";" & h.os & ";" & methods
next

outfile.Close
