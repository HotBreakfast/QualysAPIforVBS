Sub include(fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

include "QualysAPI.vbs"

'Proxy to reach Qualys API
Proxy = ""

'User for Qualys API
quser = ""
qpass = ""



function oneline(txt)
	txt = Replace(txt,chr(13),"")
	txt = Replace(txt,chr(10),"")
	oneline = txt
end function


function getMHostsVulnsCustReport(IPs)
	if TypeName(IPs) <> "ArrayList" then
		err.raise 8, "getMHostsVulnsCustReport", "The argument must be an ArrayList of numbers. A '" & TypeName(IPs) & "' was found instead."
	end if
	dim vulnhosts, vulnsinfo, vulnsinfoquery
	set vulnhosts = getMHostsVMDetections(IPs)
	set vulnsinfoquery = CreateObject("System.collections.ArrayList")
	for each host in vulnhosts
		set hvulns = vulnhosts.item(host).vulns
		for each vuln in hvulns 
			vulnsinfoquery.add vuln.qid
		next
	next
	vulnsinfoquery.sort()
	set vulnsinfoquery = unique(vulnsinfoquery)
	set vulnsinfo = getMVulnsInfo(vulnsinfoquery)

	'HTML header
	getMHostsVulnsCustReport = "<html>"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<head>"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & " <style type=""text/css"">"
	
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl1c { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl1cIcon)) & "); } " 
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl2c { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl2cIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl3c { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl3cIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl4c { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl4cIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl5c { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl5cIcon)) & "); } "
	
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl1p { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl1pIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl2p { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl2pIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl3p { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl3pIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl4p { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl4pIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl5p { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl5pIcon)) & "); } "
	
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl1i { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl1iIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl2i { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl2iIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl3i { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl3iIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl4i { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl4iIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .lvl5i { background-repeat: no-repeat; height: 10px; width: 50px; background-image: "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "url(data:image/svg+xml;base64," & oneline(Base64Encode(lvl5iIcon)) & "); } "
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .HostVulnsDetail:hover { background-color: lavender; }"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .VulnDetails:hover { background-color: snow; }"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .HostVulnsSection { padding-top: 10; padding-left: 10; }"	
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .VulnDiagnosis { padding-top: 10; padding-left: 10; }"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .VulnSolution { padding-top: 10; padding-left: 10; }"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .VulnExploits { padding-top: 10; padding-left: 10; }"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  .VulnMalware { padding-top: 10; padding-left: 10; }"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & " </style>"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</head>"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<body>"

	'Hosts and vulnerabilities
	for each IP in vulnhosts
		'Host Information
		set host = vulnhosts.item(IP)
		getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""HostSection"" ><a id=""host_" & IP & """ ><h2>Host " & IP & "</h2></a>"
		getMHostsVulnsCustReport = getMHostsVulnsCustReport & " <div class=""HostInfoSection"" >"
		getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  IP Address: """ & host.ip & """, Hostname: """ & host.name & """, Netbios: """ & host.netbios & """<br />"
		getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  Last Scan: """ & left(host.lastVMScan,10) & """, Operating System: """ & host.os & """"
		getMHostsVulnsCustReport = getMHostsVulnsCustReport & " </div>"
		getMHostsVulnsCustReport = getMHostsVulnsCustReport & " <div class=""HostVulnsSection"">"
		
		'Vulnerabilities
		'Confirmed
		for severity = 5 to 1 step -1
		for each vuln in host.vulns
			if vuln.dtype = "Confirmed" and vuln.severity = severity then
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  <div class=""HostVulnsDetail"">"
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<img class=""lvl" & severity & "c"" alt=""Severity "& severity & " Confirmed"" /> "
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<a href=""#QID_" & vuln.qid & "_info"">QID_" & vuln.qid & "</a>"
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "&nbsp;&nbsp;&nbsp;<b>" & vulnsinfo.item(vuln.qid).title & "</b>&nbsp;&nbsp;&nbsp;" & "Last Detected: " & left(vuln.lastFound,10)
				if vuln.port <> "" then
					getMHostsVulnsCustReport = getMHostsVulnsCustReport & " Port: " & vuln.port	& vuln.proto
					if vuln.ssl then
						getMHostsVulnsCustReport = getMHostsVulnsCustReport & " SSL"
					end if
				end if
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  </div>"
			end if
		next
		next

		'Potential 
		for severity = 5 to 1 step -1
		for each vuln in host.vulns
			if vuln.dtype = "Potential" and vuln.severity = severity then
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  <div class=""HostVulnsDetail"">"
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<img class=""lvl" & severity & "p"" alt=""Severity " & severity & " Potential"" /> "
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<a href=""#QID_" & vuln.qid & "_info"">QID_" & vuln.qid & "</a>"
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "&nbsp;&nbsp;&nbsp;<b>" & vulnsinfo.item(vuln.qid).title & "</b>&nbsp;&nbsp;&nbsp;" & "Last Detected: " & left(vuln.lastFound,10)
				if vuln.port <> "" then
					getMHostsVulnsCustReport = getMHostsVulnsCustReport & " Port: " & vuln.port	& vuln.proto
					if vuln.ssl then
						getMHostsVulnsCustReport = getMHostsVulnsCustReport & " SSL"
					end if
				end if
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "  </div>"
			end if
		next
		next

		getMHostsVulnsCustReport = getMHostsVulnsCustReport & " </div>"
		getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
		getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<hr />"
	next
	
	'Vuln info section confirmed
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulninfoSection"" ><a id=""VulninfoSection""><h2>Vulnerability Information</h2></a>"
	for severity = 5 to 1 step -1
	for each qid in vulnsinfo
		set vuln = vulnsinfo.item(qid)
		if vuln.vulntype = "Vulnerability" and vuln.severity = severity then
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnDetails"">"
			'Vuln title
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnHeader"">"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<h3>"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<img class=""lvl" & severity & "c"" alt=""Severity " & severity & " 5"" /> "
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<a id=""QID_" & vuln.qid & "_info"">&nbsp;&nbsp;&nbsp;" & " " & vuln.title & "</a>"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</h3>"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			'Vuln Summary
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnSummary"">"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "QID: " & vuln.qid
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "&nbsp;&nbsp;&nbsp;Published: " & replace( replace(vuln.published,"T"," "), "Z"," GMT" )
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "&nbsp;&nbsp;&nbsp;Type: " & vuln.vulntype
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "&nbsp;&nbsp;&nbsp;Category: " & vuln.category
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<br />"
			dim cvelist
			cvelist = ""
			for each cve in vuln.cvelist
				if cvelist <> "" then
					cvelist = cvelist & ", "
				end if
				cvelist = cvelist & cve
			next
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "CVE Reference: " & cvelist & "<br />"
			dim references
			references = ""
			for each ref in vuln.ref
				if references <> "" then
					references = references & ", "
				end if
				references = references & ref
			next
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "Vendor Reference: " & references
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			'Vuln Diagnosis
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnDiagnosis""><h4>Diagnosis</h4>"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & vuln.diagnosis
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			'Vuln Solution
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnSolution""><h4>Solution</h4>"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & vuln.solution
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			'Vuln Exploits
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnExploits""><h4>Exploits</h4>"
			if vuln.exploits.count > 0 then
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<ul class=""exploitslist"">"
				for each exploit in vuln.exploits
					getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<li>" & exploit & "</li>"
				next
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</ul>"
			else
					getMHostsVulnsCustReport = getMHostsVulnsCustReport & "No exploits are available for this vulnerability."
			end if
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			'Vuln Exploits
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnMalware""><h4>Malware</h4>"
			if vuln.malware.count > 0 then
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<ul class=""Malwarelist"">"
				for each malware in vuln.malware
					getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<li>" & malware & "</li>"
				next
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</ul>"
			else
					getMHostsVulnsCustReport = getMHostsVulnsCustReport & "No malware exploits this vulnerability."
			end if
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
		end if
	next
	next
	
	'Vuln info section potential 
	'TODO: not working, check if statement
	for severity = 5 to 1 step -1
	for each qid in vulnsinfo
		set vuln = vulnsinfo.item(qid)
		if (vuln.vulntype = "Potential Vulnerability" or vuln.vulntype = "Vulnerability or Potential Vulnerability") and vuln.severity = severity then
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnDetails"">"
			'Vuln title
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnHeader"">"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<h3>"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<img class=""lvl" & severity & "p"" alt=""Severity " & severity & " 5"" /> "
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<a id=""QID_" & vuln.qid & "_info"">&nbsp;&nbsp;&nbsp;" & " " & vuln.title & "</a>"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</h3>"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			'Vuln Summary
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnSummary"">"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "QID: " & vuln.qid
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "&nbsp;&nbsp;&nbsp;Published: " & replace( replace(vuln.published,"T"," "), "Z"," GMT" )
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "&nbsp;&nbsp;&nbsp;Type: " & vuln.vulntype
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "&nbsp;&nbsp;&nbsp;Category: " & vuln.category
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<br />"
			cvelist = ""
			for each cve in vuln.cvelist
				if cvelist <> "" then
					cvelist = cvelist & ", "
				end if
				cvelist = cvelist & cve
			next
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "CVE Reference: " & cvelist & "<br />"
			references = ""
			for each ref in vuln.ref
				if references <> "" then
					references = references & ", "
				end if
				references = references & ref
			next
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "Vendor Reference: " & references
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			'Vuln Diagnosis
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnDiagnosis""><h4>Diagnosis</h4>"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & vuln.diagnosis
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			'Vuln Solution
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnSolution""><h4>Solution</h4>"
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & vuln.solution
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			'Vuln Exploits
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnExploits""><h4>Exploits</h4>"
			if vuln.exploits.count > 0 then
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<ul class=""exploitslist"">"
				for each exploit in vuln.exploits
					getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<li>" & exploit & "</li>"
				next
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</ul>"
			else
					getMHostsVulnsCustReport = getMHostsVulnsCustReport & "No exploits are available for this vulnerability."
			end if
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			'Vuln Exploits
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<div class=""VulnMalware""><h4>Malware</h4>"
			if vuln.malware.count > 0 then
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<ul class=""Malwarelist"">"
				for each malware in vuln.malware
					getMHostsVulnsCustReport = getMHostsVulnsCustReport & "<li>" & malware & "</li>"
				next
				getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</ul>"
			else
					getMHostsVulnsCustReport = getMHostsVulnsCustReport & "No malware exploits this vulnerability."
			end if
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
			
			getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
		end if
	next
	next
	'TODO: Information Gathered
	
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</div>"
	
	'HTML end
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</body>"
	getMHostsVulnsCustReport = getMHostsVulnsCustReport & "</html>"
end function

'---------------------- Custom Vulnerability report ------------------------

' Run on a command line with "cscript /nologo CustomReport.vbs > Report.html"

dim qhosts, report

set qhosts = CreateObject("System.collections.ArrayList")
qhosts.add "192.168.0.100"
qhosts.add "192.168.0.101"

report = getMHostsVulnsCustReport(qhosts)

Wscript.echo report
