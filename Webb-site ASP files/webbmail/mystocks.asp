﻿<%Option Explicit%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--#include virtual="dbpub/functions1.asp"-->
<!--#include virtual="dbpub/navbars.asp"-->
<%Call login
Dim issue,iName,person,sc,mailcon,hint,ID,submit,title,delRec,sort,ob,bln,adoCon,rs,URL
URL=Request.ServerVariables("URL")
Call openMailrs(mailcon,rs)
Call openEnigma(adoCon)
sort=Request("sort")
Select Case sort
	Case "namup" ob="name1,sc"
	Case "namdn" ob="name1 DESC,sc"
	Case "coddn" ob="sc DESC"
	Case Else
		sort="codup"
		ob="sc"
End Select
ID=session("ID")
submit=Request("submit")
delRec=Request.Form("delRec")
issue=0
If isNumeric(delRec) and delRec<>"" Then
	mailcon.Execute "DELETE FROM mystocks WHERE user="&ID&" AND issueID="&delRec
	hint="Stock deleted. "
Else
	sc=Request("sc")
	If sc<>"" And isNumeric(sc) Then
		rs.Open "SELECT issueID FROM enigma.stockListings WHERE stockExID IN(1,20,22,23,38,71) AND stockCode="&sc&" ORDER BY firstTradeDate DESC",adoCon
		If rs.EOF Then
			hint="Stock not found. "
		Else
			issue=rs("issueID")
			hint="Stock "&sc&" added. "
		End If
		rs.Close
	Else
		issue=Request.QueryString("i")
		If isNumeric(issue) and issue<>"" Then
			If Not CBool(adoCon.Execute("SELECT EXISTS(SELECT 1 FROM issue WHERE ID1="&issue&")").Fields(0)) Then
				hint="Stock not found. "
				issue=0
			End If
		End If
	End If
	If issue>0 Then
		If mailcon.Execute ("SELECT EXISTS(SELECT 1 FROM mystocks WHERE user="&ID&" AND issueID="&issue&")").Fields(0) Then
			hint="That stock is already in your list. "
		Else
			mailcon.Execute "INSERT INTO mystocks (user,issueID) VALUES("&ID&","&issue&")"
			hint="Stock added. "
		End If
	End If
End If
title="My stocks dashboard"%>
<title><%=title%></title>
<link rel="stylesheet" type="text/css" href="../templates/main.css">
</head>
<body>
<!--#include file="../templates/cotopdb.asp"-->
<%Call userBar(8)%>
<h2><%=title%></h2>
<p>Welcome to the user zone. Build a list of stocks that you wish to track, then click the navigation bar 
above to see big changes in CCASS positions and total returns on the stocks, or 
use the dashboard below to jump to information on individual stocks. Hit "X" to 
delete a stock. Click the issue name to jump to the "Key Data" page of the 
issuer, where you can enter a Governance Rating.</p>
<form method="get" action="mystocks.asp">
	<div class="inputs">
		HK stock code: <input type="text" name="sc" size="5" value="">
		<input type="submit" name="submit" value="Add">
	</div>
	<div class="clear"></div>
</form>
<%If hint<>"" Then%>
	<h3><%=hint%></h3>
<%End If%>
<table class="txtable">
	<tr>
		<th><%SL "Stock<br>code","codup","coddn"%></th>
		<th></th>
		<th><%SL "Issue name","namup","namdn"%></th>
		<th></th>
	</tr>
	<%rs.Open "SELECT name1,typeShort,expmat,expAcc,personID,enigma.lastCode(i.ID1) AS sc,issueID "&_
		"FROM mystocks m JOIN(enigma.issue i,enigma.organisations o,enigma.sectypes st) "&_
		"ON m.issueID=i.ID1 AND i.issuer=o.personID AND i.typeID=st.typeID WHERE user="&ID& " ORDER BY "&ob,mailcon
	Do Until rs.EOF
		iName=rs("Name1")&": "&rs("typeShort")&" "&DateStr(rs("expmat"),rs("expAcc"))
		issue=rs("issueID")
		person=rs("personID")
		%>
		<tr>
			<td><%=rs("sc")%></td>
			<td>
				<form method="post" action="mystocks.asp">
					<input type="submit" name="submit" value="X">
					<input type="hidden" name="delRec" value="<%=issue%>">
				</form>
			</td>
			<td><a href="/dbpub/orgdata.asp?p=<%=person%>"><%=iName%></a></td>
			<td><%Call stockbar(issue,0)%></td>
		</tr>
		<%
		rs.MoveNext
	Loop%>
</table>
<%Call CloseConRs(mailcon,rs)
Call CloseCon(adoCon)%>
<!--#include virtual="/templates/footerdb.asp"-->
</body>
</html>