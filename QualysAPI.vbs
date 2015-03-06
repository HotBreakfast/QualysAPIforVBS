dim Proxy
Proxy = ""

dim quser,qpass
quser=""
qpass=""

function unique(a)
	set unique = CreateObject("System.Collections.ArrayList")
	for each elem in a
		if lastelem <> elem then
			unique.add elem
		end if
		lastelem = elem
	next
end function

function regMatch(pattern, subject, ignorecase)
	dim rge
	set rge = New RegExp 
	rge.Ignorecase = ignorecase
	rge.global = true 
	rge.pattern = pattern 
	set regMatch = rge.Execute(subject)
	set rge = nothing
end function

Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue =Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function Base64Decode(ByVal vCode)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.text = vCode
    Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
    Set oNode = Nothing
    Set oXML = Nothing
End Function

'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string 
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To text/string
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the output text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And get text/string data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function

Function xmlGET(URL, httpuser, httppass)
	Dim oHTTPreq
        Set oHTTPreq = CreateObject("WinHttp.WinHttpRequest.5.1")
	oHTTPreq.SetTimeouts 10000, 10000, 10000, 10000

	oHTTPreq.Open "GET", URL, false
	if Proxy <> "" then
		oHTTPreq.setProxy 2, Proxy
	end if
	'oHTTPreq.SetCredentials httpuser, httppass, 0
	oHTTPreq.Option(WinHttpRequestOption_UserAgentString) = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
	oHTTPreq.setRequestHeader "Content-type", "text/xml"
	oHTTPreq.setRequestHeader "Translate", "f"
	oHTTPreq.setRequestHeader "X-Requested-With", "VBScript"
	if httpuser <> "" then
		oHTTPreq.setRequestHeader "Authorization", "Basic " & base64encode(httpuser & ":" & httppass)
	end if
	oHTTPreq.send
	
	 'wscript.echo oHTTPreq.status
	xmlGET = oHTTPreq.responseText
	Set oHTTPreq = Nothing
End Function

Class QualysMapReport
	dim username, company, date, title, target, duration, scanner, status
	dim optionprofile
	dim IPs
end Class

Class QualysHost
	dim assetID, name, os, netbios, ip, trackmethod, comments, lastVMScan, lastPCScan
end class

function listAssetsIPs()
	dim result
	set listAssetsIPs = CreateObject("System.Collections.ArrayList")
	result = xmlGET("https://qualysapi.qualys.com/api/2.0/fo/asset/host/?action=list&truncation_limit=10000",quser,qpass)
	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	objXMLDoc.async = False 
	objXMLDoc.loadXML result
	Set Root = objXMLDoc.documentElement 
	Set NodeList = Root.getElementsByTagName("IP")
	for each IP in NodeList
		listAssetsIPs.add IP.text
	next
end function

function listAssetsHosts()
	dim result
	set listAssetsHosts = CreateObject("System.Collections.ArrayList")
	result = xmlGET("https://qualysapi.qualys.com/api/2.0/fo/asset/host/?action=list&details=All&truncation_limit=10000",quser,qpass)
	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	objXMLDoc.async = False 
	objXMLDoc.loadXML result
	Set Root = objXMLDoc.documentElement 
	Set NodeList = Root.getElementsByTagName("HOST")
	for each host in NodeList
		dim h
		set h = new QualysHost
		on error resume next
		h.assetID = host.getElementsByTagName("ID").item(0).text
		h.ip = host.getElementsByTagName("IP").item(0).text
		h.name = host.getElementsByTagName("DNS").item(0).text
		h.os = host.getElementsByTagName("OS").item(0).text
		h.trackmethod = host.getElementsByTagName("TRACKING_METHOD").item(0).text
		h.netbios = host.getElementsByTagName("NETBIOS").item(0).text
		h.lastVMScan = host.getElementsByTagName("LAST_VULN_SCAN_DATETIME").item(0).text
		h.lastPCScan = host.getElementsByTagName("LAST_COMPLIANCE_SCAN_DATETIME").item(0).text
		h.comments = host.getElementsByTagName("COMMENTS").item(0).text
		h.comments = Replace(h.comments,chr(13),"")
		h.comments = Replace(h.comments,chr(10)," ")
		h.comments = Replace(h.comments,";",",")
		on error goto 0
		listAssetsHosts.add h
	next
end function

function listScans(fromdate,todate)
	dim result
	set listScans = CreateObject("System.Collections.ArrayList")
	result = xmlGET("https://qualysapi.qualys.com/api/2.0/fo/scan/?action=list&launched_after_datetime=" & fromdate & "&launched_before_datetime=" & todate, quser, qpass)
	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	objXMLDoc.async = False 
	objXMLDoc.loadXML result
	Set Root = objXMLDoc.documentElement 
	Set NodeList = Root.getElementsByTagName("REF")
	for each elem in NodeList
		listScans.add elem.text
	next
end function

function listMaps(fromdate, todate)
	dim result
	set listMaps = CreateObject("System.Collections.ArrayList")
	result = xmlGET("https://qualysapi.qualys.com/msp/action_log_report.php?date_from=" & fromdate & "&date_to=" & todate,quser,qpass)
	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	objXMLDoc.async = False 
	objXMLDoc.loadXML result
	Set Root = objXMLDoc.documentElement
	Set NodeList = Root.getElementsByTagName("ACTION_LOG")
	for each elem in NodeList
		if elem.getElementsByTagName("MODULE")(0).text = "map" then
			dim date, ref, matchlist
			  'date = elem.getElementsByTagName("DATE")(0).text
			set matchlist = regMatch("\(ref: (map\/[0-9\.]+)\)", elem.getElementsByTagName("DETAILS")(0).text, true)
			if matchlist.Count <> 0 then
				listMaps.add matchlist(0).SubMatches(0)
			end if
		end if
	next
end function

function getMapReport(ref)
	getMapReport = xmlGET("https://qualysapi.qualys.com/msp/map_report.php?ref=" & ref,quser,qpass)
end function

function getMapDetectedIPs(ref)
	dim mapreport
	set getMapDetectedIPs = CreateObject("System.Collections.ArrayList")
	mapreport = getMapReport(ref)
	Set xmlreport = CreateObject("Microsoft.XMLDOM") 
	xmlreport.async = False 
	xmlreport.loadXML mapreport
	Set IPlist = xmlreport.getElementsByTagName("IP")
	for each IP in IPlist
		getMapDetectedIPs.add IP.getAttribute("value")
	next
end function

function getMapResults(ref)
	set getMapResults = new QualysMapReport
	dim mapreport
	mapreport = getMapReport(ref)
	Set xmlreport = CreateObject("Microsoft.XMLDOM") 
	xmlreport.async = False 
	xmlreport.loadXML mapreport
	Set KEYlist = xmlreport.getElementsByTagName("KEY")
	for each KEY in KEYlist
		attrvalue =  KEY.getAttribute("value")
		if attrvalue = "STATUS" then
			getMapResults.status = KEY.text
		end if
		if attrvalue = "DATE" then
			getMapResults.date = KEY.text
		end if
		if attrvalue = "TITLE" then
			getMapResults.title = KEY.text
		end if
		if attrvalue = "SCAN_HOST" then
			getMapResults.scanner = KEY.text
		end if
		if attrvalue = "TARGET" then
			getMapResults.target = KEY.text
		end if
	next
end function

function getMapDetectedLiveIPs(ref)
	dim mapreport
	set getMapDetectedLiveIPs = CreateObject("System.Collections.ArrayList")
	mapreport = getMapReport(ref)
	Set xmlreport = CreateObject("Microsoft.XMLDOM") 
	xmlreport.async = False 
	xmlreport.loadXML mapreport
	Set IPlist = xmlreport.getElementsByTagName("IP")
	for each IP in IPlist
		set discoverylist = IP.getElementsByTagName("DISCOVERY")
		for each discovery in discoverylist
			live = false
			method = discovery.getAttribute("method")
			if method = "ICMP" then
				live = true
			end if
			if instr(1,method,"TCP ") > 0 then
				live = true
			end if
			if live = true then
				getMapDetectedLiveIPs.add IP.getAttribute("value")
				exit for
			end if
		next
	next
end function

function getMapDetectedLiveHosts(ref)
	dim mapreport
	set getMapDetectedLiveHosts = CreateObject("System.Collections.ArrayList")
	mapreport = getMapReport(ref)
	Set xmlreport = CreateObject("Microsoft.XMLDOM") 
	xmlreport.async = False 
	xmlreport.loadXML mapreport
	Set IPlist = xmlreport.getElementsByTagName("IP")
	for each IP in IPlist
		set discoverylist = IP.getElementsByTagName("DISCOVERY")
		for each discovery in discoverylist
			live = false
			method = discovery.getAttribute("method")
			if method = "ICMP" then
				live = true
			end if
			if instr(1,method,"TCP ") > 0 then
				live = true
			end if
			if live = true then
				dim h
				set h = new QualysHost
				h.ip = IP.getAttribute("value")
				h.os = IP.getAttribute("os")
				h.name = IP.getAttribute("name")
				h.netbios = IP.getAttribute("netbios")
				getMapDetectedLiveHosts.add h
				exit for
			end if
		next
	next
end function

function getMapStatus(ref)
	dim mapreport
	mapreport = getMapReport(ref)
	Set xmlreport = CreateObject("Microsoft.XMLDOM") 
	xmlreport.async = False 
	xmlreport.loadXML mapreport
	Set KEYlist = xmlreport.getElementsByTagName("KEY")
	for each KEY in KEYlist
		attrvalue =  KEY.getAttribute("value")
		if attrvalue = "STATUS" then
			getMapStatus = KEY.text
		end if
	next
end function

function toQualysTime(oDate)
	qYear = Year(oDate)
	qMonth = Month(oDate)
	if qMonth < 10 then
		qMonth= "0" & qMonth
	end if
	qDay = Day(oDate)
	if qDay < 10 then
		qDay = "0" & qDay
	end if

	qHour = Hour(oDate)
	if qHour < 10 then
		qHour = "0" & qHour
	end if
	qMinute = Minute(oDate)
	if qMinute < 10 then
		qMinute = "0" & qMinute
	end if
	qSecond = Second(oDate)
	if qSecond < 10 then
		qSecond = "0" & qSecond
	end if

	toQualysTime = "" & qYear & "-" & qMonth & "-" & qDay
	toQualysTime = toQualysTime & "T"
	toQualysTime = toQualysTime & qHour & ":" & qMinute & ":" & qSecond
end function

function getLastWeekMapDetectedIps()
	set detectedIPs = CreateObject("System.Collections.ArrayList")
	qDateToday = toQualysTime(Date())
	qDate7days = toQualysTime(DateAdd("d",Date(),-7))
	set maplist = listMaps(qDate7days,qDateToday)
	'getLastWeekMapDetectedIps = 0
	for each map in maplist
		detectedIPs.addRange(getMapDetectedLiveIPs(map))
	next
	detectedIPs.sort()
	set getLastWeekMapdetectedIps = unique(detectedIPs)
end function

function getLastWeekMapDetectedHosts()
	set detectedIPs = CreateObject("System.Collections.ArrayList")
	qDateToday = toQualysTime(Date())
	qDate7days = toQualysTime(DateAdd("d",Date(),-7))
	set maplist = listMaps(qDate7days,qDateToday)
	'getLastWeekMapDetectedIps = 0
	for each map in maplist
		detectedIPs.addRange(getMapDetectedLiveHosts(map))
	next
	'detectedIPs.sort()
	'set getLastWeekMapdetectedHosts = unique(detectedIPs)
	set getLastWeekMapDetectedHosts = detectedIPs
end function

Class QualysVulInfo
	dim qid, vulntype, severity, title, category, published, patchable, cvelist, diagnosis, consequence, solution, exploits, malware, ref
end class

function getVulnInfo(qid)
	dim response,xml,vulnlist,vuln,info
	response = xmlGET("https://qualysapi.qualys.com/api/2.0/fo/knowledge_base/vuln/?action=list&details=All&ids=" & qid,quser,qpass)
	Set xml = CreateObject("Microsoft.XMLDOM") 
	xml.async = False 
	xml.loadXML response
	Set vulnlist = xml.getElementsByTagName("VULN")
	set vuln = vulnlist.item(0)
	set info = new QualysVulInfo
	On error resume next
	info.qid = Clng(vuln.getElementsByTagName("QID").item(0).text)
	info.vulntype    = vuln.getElementsByTagName("VULN_TYPE").item(0).text
	info.severity    = Clng(vuln.getElementsByTagName("SEVERITY_LEVEL").item(0).text)
	info.title       = vuln.getElementsByTagName("TITLE").item(0).text
	info.published   = vuln.getElementsByTagName("PUBLISHED_DATETIME").item(0).text
	info.category    = vuln.getElementsByTagName("CATEGORY").item(0).text
	info.diagnosis   = vuln.getElementsByTagName("DIAGNOSIS").item(0).text
	info.consequence = vuln.getElementsByTagName("CONSEQUENCE").item(0).text
	info.solution    = vuln.getElementsByTagName("SOLUTION").item(0).text
	info.patchable   = vuln.getElementsByTagName("PATCHABLE").item(0).text
	if info.patchable = "1" then
		info.patchable = true
	else
		info.patchable = false
	end if
	dim cvelist, cve
	set info.cvelist = CreateObject("System.Collections.ArrayList")
	set cvelist = vuln.getElementsByTagName("CVE")
	for each cve in cvelist
		info.cvelist.add cve.getElementsByTagName("ID").item(0).text
	next
	dim explist, expl
	set info.exploits = CreateObject("System.Collections.ArrayList")
	set explist = vuln.getElementsByTagName("EXPLT")
	for each expl in explist
		info.exploits.add expl.getElementsByTagName("DESC").item(0).text
	next
	dim mlwlist, mlw
	set info.malware = CreateObject("System.Collections.ArrayList")
	set mlwlist = vuln.getElementsByTagName("MW_INFO")
	for each mlw in mlwlist
		info.malware.add mlw.getElementsByTagName("MW_ID").item(0).text
	next
	dim reflist, ref
	set info.ref = CreateObject("System.Collections.ArrayList")
	set reflist = vuln.getElementsByTagName("VENDOR_REFERENCE")
	for each ref in reflist
		info.ref.add reflist.getElementsByTagName("ID").item(0).text
	next
	On error goto 0
	set getVulnInfo = info
end function

function getMVulnsInfo(qids)
	if TypeName(qids) <> "ArrayList" then
		err.raise 8, "getMVulnsInfo", "The argument must be an ArrayList of numbers. A '" & TypeName(qids) & "' was found instead."
	end if
	set getMVulnsInfo = CreateObject("Scripting.Dictionary")
	dim response,xml,query,qid,vulnlist,vuln
	query = ""
	for each qid in qids
		if query = "" then
			query = query & qid
		else
			query = query & "," & qid
		end if
	next
	 'wscript.echo "querying: " & query
	response = xmlGET("https://qualysapi.qualys.com/api/2.0/fo/knowledge_base/vuln/?action=list&details=All&ids=" & query,quser,qpass)
	 'wscript.echo response
	Set xml = CreateObject("Microsoft.XMLDOM") 
	xml.async = False 
	xml.loadXML response
	Set vulnlist = xml.getElementsByTagName("VULN")
	for each vuln in vulnlist
		dim info
		set info = new QualysVulInfo
		On error resume next
		info.qid         = Clng(vuln.getElementsByTagName("QID").item(0).text)
		info.vulntype    = vuln.getElementsByTagName("VULN_TYPE").item(0).text
		info.severity    = Clng(vuln.getElementsByTagName("SEVERITY_LEVEL").item(0).text)
		info.title       = vuln.getElementsByTagName("TITLE").item(0).text
		info.published   = vuln.getElementsByTagName("PUBLISHED_DATETIME").item(0).text
		info.category    = vuln.getElementsByTagName("CATEGORY").item(0).text
		info.diagnosis   = vuln.getElementsByTagName("DIAGNOSIS").item(0).text
		info.consequence = vuln.getElementsByTagName("CONSEQUENCE").item(0).text
		info.solution    = vuln.getElementsByTagName("SOLUTION").item(0).text
		info.patchable   = vuln.getElementsByTagName("PATCHABLE").item(0).text
		if info.patchable = "1" then
			info.patchable = true
		else
			info.patchable = false
		end if
		dim cvelist, cve
		set info.cvelist = CreateObject("System.Collections.ArrayList")
		set cvelist = vuln.getElementsByTagName("CVE")
		for each cve in cvelist
			info.cvelist.add cve.getElementsByTagName("ID").item(0).text
		next
		dim explist, expl
		set info.exploits = CreateObject("System.Collections.ArrayList")
		set explist = vuln.getElementsByTagName("EXPLT")
		for each expl in explist
			info.exploits.add expl.getElementsByTagName("DESC").item(0).text
		next
		dim mlwlist, mlw
		set info.malware = CreateObject("System.Collections.ArrayList")
		set mlwlist = vuln.getElementsByTagName("MW_INFO")
		for each mlw in mlwlist
			info.malware.add mlw.getElementsByTagName("MW_ID").item(0).text
		next
		dim reflist, ref
		set info.ref = CreateObject("System.Collections.ArrayList")
		set reflist = vuln.getElementsByTagName("VENDOR_REFERENCE")
		for each ref in reflist
			info.ref.add reflist.getElementsByTagName("ID").item(0).text
		next
		On error goto 0
		Wscript.echo info.qid & " " & TypeName(info.qid)
		if TypeName(info.qid) = "Long" and not getMVulnsInfo.exists(info.qid) then
			getMVulnsInfo.add info.qid, info
		end if
	next
end function
