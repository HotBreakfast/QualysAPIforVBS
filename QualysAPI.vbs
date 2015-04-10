'Qualys vbscript library by Chema Sanchez
'April 2015

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

function urlDecode(str)
	dim ret,val,i
	ret = replace(str,"+"," ")
	for i = 0 to 128
		if i < &H10 then
			val = "%0" & hex(i)
		else
			val = "%" & hex(i)
		end if
		if i <> &H25 then
			ret = replace (ret,val,chr(i))
		end if
	next
	val = "%25"
	i = &h25
	ret = replace (ret,val,chr(i))
	urlDecode = ret
end function

function urlEncode(str)
	dim ret
	ret = ""
	for i = 1 to len(str)
		c = Mid(str,i,1)
		ic = asc(c)
		if ic >= &h41 and ic <= &h7a then
			'A-z
			ret = ret & c
		elseif ic >= &h30 and ic <= &h39 then
			'0-9
			ret = ret & c
		elseif ic = &h2d or ic = &h2e or ic = &h5f or ic = &h7e then
			' - . _ ~
			ret = ret & c
		elseif ic = &h20 then
			' space
			ret = ret & "+"
		elseif ic < &h10 then
			ret = ret & "%0" & hex(ic)
		else
			ret = ret & "%" & hex(ic)
		end if
	next
	urlEncode = ret
end function

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

Function xmlPost(URL, data, httpuser, httppass)
	Dim oHTTPreq
        Set oHTTPreq = CreateObject("WinHttp.WinHttpRequest.5.1")
	oHTTPreq.SetTimeouts 10000, 10000, 10000, 10000

	oHTTPreq.Open "POST", URL, false
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
	oHTTPreq.send data
	
	 'wscript.echo oHTTPreq.status
	xmlPost = oHTTPreq.responseText
	Set oHTTPreq = Nothing
End Function

Class QualysMapReport
	dim username, company, date, title, target, duration, scanner, status
	dim optionprofile
	dim IPs
end Class

Class QualysHost
	dim assetID, name, os, netbios, ip, trackmethod, comments, lastVMScan, lastPCScan, vulns, discoveryMethods
	
	function vulnCount()
		'TODO: vulnerability Count
	end function
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
				set h.discoveryMethods = CreateObject("System.Collections.ArrayList")
				for each d in discoverylist
					method = d.getAttribute("method")
					h.discoveryMethods.add method
				next
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
	qDateToday = toQualysTime(DateAdd("d",Date(),1))
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

dim lvl1cIcon, lvl2cIcon, lvl3cIcon, lvl4cIcon, lvl5cIcon
dim lvl1pIcon, lvl2pIcon, lvl3pIcon, lvl4pIcon, lvl5pIcon
dim lvl1iIcon, lvl2iIcon, lvl3iIcon, lvl4iIcon, lvl5iIcon

lvl1cIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""10"" height=""10"" fill=""red"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl2cIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""20"" height=""10"" fill=""red"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl3cIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""30"" height=""10"" fill=""red"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl4cIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""40"" height=""10"" fill=""red"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl5cIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""red"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"


lvl1pIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""10"" height=""10"" fill=""yellow"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl2pIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""20"" height=""10"" fill=""yellow"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl3pIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""30"" height=""10"" fill=""yellow"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl4pIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""40"" height=""10"" fill=""yellow"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl5pIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""yellow"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"


lvl1iIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""10"" height=""10"" fill=""cyan"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl2iIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""20"" height=""10"" fill=""cyan"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl3iIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""30"" height=""10"" fill=""cyan"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl4iIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""40"" height=""10"" fill=""cyan"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"

lvl5iIcon = "" &_
"<svg xmlns=""http://www.w3.org/2000/svg"" width=""50"" height=""10""> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""cyan"" /> " & VbCrLf   &_
" <rect x=""0"" y=""0"" width=""50"" height=""10"" fill=""none"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""10"" x2=""10"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""20"" x2=""20"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""30"" x2=""30"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""40"" x2=""40"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
" <line x1=""50"" x2=""50"" y=""0"" y2=""10"" stroke=""black"" stroke-width=""1"" /> " & VbCrLf   &_
"</svg>"


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
	dim reflist, refe
	set info.ref = CreateObject("System.Collections.ArrayList")
	set reflist = vuln.getElementsByTagName("VENDOR_REFERENCE")
	for each refe in reflist
		info.ref.add refe.getElementsByTagName("ID").item(0).text
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
		 'Wscript.echo info.qid & " " & TypeName(info.qid)
		if TypeName(info.qid) = "Long" and not getMVulnsInfo.exists(info.qid) then
			getMVulnsInfo.add info.qid, info
		end if
	next
end function

class QualysVMDetection
	dim qid, dtype, severity, ssl, port, proto, status, results, firstFound, lastFound, lastTest, lastUpdate
	
	function getIcon()
		dim icon
		select case severity
		case 1
			if dtype = "Confirmed" then
				icon = lvl1cIcon
			else 
				if dtype= "Potential" then
					icon = lvl1pIcon
				else
					icon = lvl1iIcon
				end if
			end if
		case 2
			if dtype = "Confirmed" then
				icon = lvl2cIcon
			else
				if dtype= "Potential" then
					icon = lvl2pIcon
				else
					icon = lvl2iIcon
				end if
			end if
		case 3
			if dtype = "Confirmed" then
				icon = lvl3cIcon
			else
				if dtype= "Potential" then
					icon = lvl3pIcon
				else
					icon = lvl3iIcon
				end if
			end if
		case 4
			if dtype = "Confirmed" then
				icon = lvl4cIcon
			else
				if dtype= "Potential" then
					icon = lvl4pIcon
				else
					icon = lvl4iIcon
				end if
			end if
		case 5
			if dtype = "Confirmed" then
				icon = lvl5cIcon
			else
				if dtype= "Potential" then
					icon = lvl5pIcon
				else
					icon = lvl5iIcon
				end if
			end if
		case else
			if dtype = "Confirmed" then
				icon = lvl1cIcon
			else
				if dtype= "Potential" then
					icon = lvl1pIcon
				else
					icon = lvl1iIcon
				end if
			end if
		end select
		getIcon = icon
	end function
end class

function getHostVMDetections(host)
	dim response,xml, detectlist, detection
	response = xmlGET("https://qualysapi.qualys.com/api/2.0/fo/asset/host/vm/detection/?action=list&show_igs=1&ips=" & host,quser,qpass)
	Set xml = CreateObject("Microsoft.XMLDOM") 
	xml.async = False 
	xml.loadXML response
	set getHostVMDetections = CreateObject("System.Collections.ArrayList")
	Set detectlist = xml.getElementsByTagName("DETECTION")
	for each detection in detectlist
		dim vuln
		set vuln = new QualysVMDetection
		on error resume next
		vuln.qid = Clng(detection.getElementsByTagName("QID").item(0).text)
		vuln.dtype = detection.getElementsByTagName("TYPE").item(0).text
		vuln.severity = Clng(detection.getElementsByTagName("SEVERITY").item(0).text)
		vuln.port = Clng(detection.getElementsByTagName("PORT").item(0).text)
		vuln.proto = detection.getElementsByTagName("PROTOCOL").item(0).text
		vuln.ssl   = detection.getElementsByTagName("SSL").item(0).text
		if detection.ssl = "1" then
			detection.ssl = true
		else
			detection.ssl = false
		end if
		vuln.results = detection.getElementsByTagName("RESULTS").item(0).text
		vuln.status = detection.getElementsByTagName("STATUS").item(0).text
		vuln.firstFound = detection.getElementsByTagName("FIRST_FOUND_DATETIME").item(0).text
		vuln.lastFound = detection.getElementsByTagName("LAST_FOUND_DATETIME").item(0).text
		vuln.lastTest = detection.getElementsByTagName("LAST_TEST_DATETIME").item(0).text
		vuln.lastUpdate = detection.getElementsByTagName("LAST_UPDATE_DATETIME").item(0).text
		on error goto 0
			if TypeName(vuln.qid) = "Long" then
			getHostVMDetections.add vuln
		end if
	next
end function

function getMHostsVMDetections(IPs)
	if TypeName(IPs) <> "ArrayList" then
		err.raise 8, "getMHostsVMDetections", "The argument must be an ArrayList of numbers. A '" & TypeName(IPs) & "' was found instead."
	end if
	set getMHostsVMDetections = CreateObject("Scripting.Dictionary")
	dim response, xml, hostlist, host, detectlist, detection, query
	query = ""
	for each IP in IPs
		if query = "" then
			query = query & IP
		else
			query = query & "," & IP
		end if
	next
	response = xmlGET("https://qualysapi.qualys.com/api/2.0/fo/asset/host/vm/detection/?action=list&show_igs=1&ips=" & query,quser,qpass)
	Set xml = CreateObject("Microsoft.XMLDOM") 
	xml.async = False 
	xml.loadXML response
	Set hostlist = xml.getElementsByTagName("HOST")
	for each host in hostlist
		dim h
		set h = new QualysHost
		On error resume next
		h.assetID = host.getElementsByTagName("ID").item(0).text
		h.ip = host.getElementsByTagName("IP").item(0).text
		h.name = host.getElementsByTagName("DNS").item(0).text
		h.os = host.getElementsByTagName("OS").item(0).text
		h.trackmethod = host.getElementsByTagName("TRACKING_METHOD").item(0).text
		h.netbios = host.getElementsByTagName("NETBIOS").item(0).text
		h.lastVMScan = host.getElementsByTagName("LAST_SCAN_DATETIME").item(0).text
		On error goto 0
		set h.vulns = CreateObject("System.collections.ArrayList")
		Set detectlist = host.getElementsByTagName("DETECTION")
		for each detection in detectlist
			dim vuln
			set vuln = new QualysVMDetection
			on error resume next
			vuln.qid = Clng(detection.getElementsByTagName("QID").item(0).text)
			vuln.dtype = detection.getElementsByTagName("TYPE").item(0).text
			vuln.severity = Clng(detection.getElementsByTagName("SEVERITY").item(0).text)
			vuln.port = Clng(detection.getElementsByTagName("PORT").item(0).text)
			vuln.proto = detection.getElementsByTagName("PROTOCOL").item(0).text
			vuln.ssl   = detection.getElementsByTagName("SSL").item(0).text
			if detection.ssl = "1" then
				detection.ssl = true
			else
				detection.ssl = false
			end if
			vuln.results = detection.getElementsByTagName("RESULTS").item(0).text
			vuln.status = detection.getElementsByTagName("STATUS").item(0).text
			vuln.firstFound = detection.getElementsByTagName("FIRST_FOUND_DATETIME").item(0).text
			vuln.lastFound = detection.getElementsByTagName("LAST_FOUND_DATETIME").item(0).text
			vuln.lastTest = detection.getElementsByTagName("LAST_TEST_DATETIME").item(0).text
			vuln.lastUpdate = detection.getElementsByTagName("LAST_UPDATE_DATETIME").item(0).text
			on error goto 0
				if TypeName(vuln.qid) = "Long" then
				h.vulns.add vuln
			end if
		next
		if not getMHostsVMDetections.exists(h.ip) then
			getMHostsVMDetections.add h.ip, h
		end if
	next
end function

function launchScoreCardReport(name,format,title)
	dim ret,response
	name = urlEncode(name)
	title = urlEncode(title)
	response = xmlPost("https://qualysapi.qualys.com/api/2.0/fo/report/scorecard/?action=launch&output_format=" & format & "&name=" & name & "&source=asset_groups&asset_groups=All&report_title=" & title,"",quser,qpass)
	Set xml = CreateObject("Microsoft.XMLDOM") 
	xml.async = False 
	xml.loadXML response
	dim txt, val
	set txt = xml.getElementsByTagName("TEXT")
	set val = xml.getElementsByTagName("VALUE")
	if txt.length > 0 and txt.item(0).text = "New scorecard launched" and val.length > 0 then
		ret = CDbl(val.item(0).text)
	elseif txt.length > 0 then
		err.raise 8, "launchScoreCardReport", txt.item(0).text
	else
		err.raise 8, "launchScoreCardReport", "Error retrieving response text from QualysGuard"
	end if
	launchScoreCardReport = ret
end function

function fetchReport(id)
	dim response
	response = xmlPost("https://qualysapi.qualys.com/api/2.0/fo/report/?action=fetch&id=" & id,"",quser,qpass)
	fetchReport = response
end function
