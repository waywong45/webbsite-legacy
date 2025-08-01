﻿<%Option Explicit%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--#include virtual="dbpub/functions1.asp"-->
<!--#include virtual="dbpub/navbars.asp"-->
<%Call login
Dim atDate,latestDate,issueID,ob,title,d1,d2,etf,sqletf,mailcon,fromDate,toDate,ID,sort,x,con,rs,URL
Call openEnigmaRs(con,rs)

ID=session("ID")
Call openMailDB(mailcon)
sort=Request("sort")

'whether to include unit ETFs. default no
etf=getBool("etf")
If Not etf Then sqletf=" AND orgType<>4"

d1=Request("d1")
d2=Request("d")
latestdate=Cdate(GetLog("CCASSdateDone"))
If isDate(d2) Then d2=Cdate(d2) Else d2=latestDate
If d2>latestDate Then d2=latestDate
If d2<#27-Jun-2007# Then d2=#27-Jun-2007#
d2=con.Execute("SELECT MAX(settleDate) FROM ccass.calendar WHERE settleDate<='" & MSdate(d2) & "'").Fields(0)
If isDate(d1) Then d1=Cdate(d1) Else d1=d2-1
If d1>=d2 Then d1=d2-1
d1=con.Execute("SELECT MAX(settleDate) FROM ccass.calendar WHERE settleDate<='" & MSdate(d1) & "'").Fields(0)
toDate=MSdate(d2)
fromDate=MSdate(d1)

Select Case sort
	Case "nameup" ob="Name1,stkchg DESC"
	Case "namedn" ob="Name1 DESC,stkchg"
	Case "partup" ob="partName,stkchg DESC"
	Case "partdn" ob="partName DESC,stkchg"
	Case "chgdn" ob="stkchg DESC,name1,partName"
	Case "chgup" ob="stkchg,name1,partName"
	Case "scup" ob="stockCode,stkchg DESC"
	Case "scdn" ob="stockCode DESC,stkchg"
	Case "datup" ob="atDate,name1,stkchg,partName"
	Case Else
		sort="datdn"
		ob="atDate DESC,name1,stkchg DESC,partName"
End Select
rs.Open "SELECT b.issueID,b.partID,stkchg,b.atDate,prevDate,name1,partName,enigma.lastCode(b.issueID) AS sc,typeShort "&_
	"FROM ccass.bigchanges b JOIN (mystocks m,enigma.issue i,enigma.organisations o,ccass.participants p,enigma.secTypes s) "&_
	"ON b.issueID=m.issueID AND m.user="&ID&" AND m.issueID=i.ID1 AND b.partID=p.partID AND i.issuer=o.personID AND i.typeID=s.typeID "&_
	"WHERE b.atDate>'" & fromDate & "' AND b.atDate<='" & toDate & "'" & sqletf & " ORDER BY "&ob,mailcon
URL=Request.ServerVariables("URL")&"?d1="&fromDate&"&amp;d="&toDate
title="Big CCASS changes in my stocks from "&fromDate&" To "&toDate%>
<title><%=title%></title>
<link rel="stylesheet" type="text/css" href="../templates/main.css">
</head>
<body>
<!--#include file="../templates/cotopdb.asp"-->
<%Call userBar(9)%>
<ul class="navlist">
	<li><a href="/ccass/CCASSnotes.asp">Notes</a></li>
	<li><a href="/ccass/bigchanges.asp">Whole market</a></li>
</ul>
<div class="clear"></div>

<h2><%=title%></h2>
<form method="get" action="mybigchanges.asp">
	<input type="hidden" name="sort" value="<%=sort%>">
	<div class="inputs">
		<input type="checkbox" name="etf" value="1" <%=checked(etf)%> onchange="this.form.submit()">include ETFs
	</div>
	<div class="inputs">
		From <input type="date" name="d1" id="d1" value="<%=fromDate%>" onchange="this.form.submit()">
	</div>
	<div class="inputs">
		to <input type="date" name="d" id="d2" value="<%=toDate%>" onchange="this.form.submit()">
	</div>
	<div class="inputs">
		<input type="submit" value="Go">
		<input type="submit" value="Clear" onclick="document.getElementById('d1').value='';document.getElementById('d2').value='';">
	</div>
	<div class="clear"></div>
</form>
<p>This table shows daily movements greater than 0.25% of outstanding shares in 
any of your stocks in the chosen period. Click the date to see all movements in 
that stock on that date. Click the issue to see the history of big changes for 
that stock. Click the participant to see its history in that stock. Click on 
column headings to sort.</p>
<%=mobile(1)%>
<table class="numtable">
	<tr>
		<th class="colHide1">Row</th>
		<th class="colHide3"><%SL "Stock<br>code","scup","scdn"%></th>
		<th class="left"><%SL "Issue","nameup","namedn"%></th>
		<th class="left"><%SL "Participant","partup","partdn"%></th>
		<th><%SL "Change","chgdn","chgup"%></th>
		<th class="colHide2"><%SL "Date","datdn","datup"%></th>
		<th class="colHide1">Previous<br>change</th>
	</tr>
	<%x=0
	Do Until rs.EOF
		x=x+1
		issueID=rs("issueID")
		atDate=MSdate(rs("atDate"))
		%>
		<tr>
			<td class="colHide1"><%=x%></td>
			<td class="colHide3"><%=rs("sc")%></td>
			<td class="left"><a href="../ccass/bigchangesissue.asp?i=<%=issueID%>"><%=rs("Name1")&":"&rs("typeShort")%></a></td>
			<td class="left"><a href="../ccass/chistory.asp?i=<%=rs("issueID")%>&amp;part=<%=rs("partID")%>"><%=rs("PartName")%></a></td>
			<td><%=FormatNumber(cdbl(rs("stkchg"))*100,2)%></td>
			<td class="colHide2 nowrap"><a href="../ccass/chldchg.asp?i=<%=issueID%>&amp;d=<%=atDate%>"><%=MSSdate(rs("atDate"))%></a></td>
			<td class="colHide1 nowrap"><%=MSSdate(rs("prevDate"))%></td>
		</tr>
		<%rs.MoveNext
	Loop%>
</table>
<%Call CloseConRs(con,rs)
Call CloseCon(mailCon)
If x=0 Then%>
	<p>None found.</p>
<%End If%>
<!--#include virtual="/templates/footerdb.asp"-->
</body>
</html>