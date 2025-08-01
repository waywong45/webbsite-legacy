﻿<%Option Explicit%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--#include virtual="dbpub/functions1.asp"-->
<!--#include virtual="dbpub/navbars.asp"-->
<%Call login
Dim atDate,latestDate,issueID,ob,title,d1,d2,mailcon,fromDate,toDate,ID,ds,sort,URL,x,sql,sql2,col1
Dim ytdf,y1f,y2f,y5f,y10f,ytd,y1,y2,y5,y10,con,rs
Call openEnigmaRs(con,rs)
Call openMailDB(mailcon)
ID=session("ID")
latestdate=GetLog("MBquotesDate")
sort=Request("sort")
'initialise reference dates. If null, then don't calculate or show columns Y1,Y2,Y5,Y10
y1f=Null: y2f=Null: y5f=Null: y10f=Null

d1=Request("d1")
d2=Request("d2")
If isDate(d2) Then
	d2=Cdate(d2)
	If d2>latestDate Then d2=latestDate
	If d2<#4-Jan-1994# Then d2=#4-Jan-1994#
End If
If isDate(d1) Then
	d1=CDate(d1)
	If d1>=latestDate Then d1=latestDate-1
	If d1<#3-Jan-1994# Then d1=#3-Jan-1994#
End If

If isDate(d1) AND isDate(d2) Then
	col1="Period"
	d2=con.Execute("SELECT MAX(tradeDate) FROM ccass.calendar WHERE tradeDate<='" & MSdate(d2) & "'").Fields(0)
	If d1>=d2 Then d1=d2-1
	d1=con.Execute("SELECT MAX(tradeDate) FROM ccass.calendar WHERE tradeDate<='" & MSdate(d1) & "'").Fields(0)
	title="Total returns on my stocks from "&MSdate(d1)&" to "&MSdate(d2)
	sql="100*(enigma.totret(issueID,'"&MSdate(d1)&"','"&MSdate(d2)&"')-1) AS ytd"
ElseIf isDate(d2) Or Not isDate(d1) Then
	col1="YTD"
	If Not isDate(d2) Then d2=latestdate
	d2=con.Execute("SELECT MAX(tradeDate) FROM ccass.calendar WHERE tradeDate<='" & MSdate(d2) & "'").Fields(0)
	ds=MSdate(d2)
	title="Total returns on my stocks until "&ds
	ytdf=MSdate(con.Execute("SELECT Max(tradeDate) FROM ccass.calendar WHERE tradeDate<'"&year(d2)&"-01-01'").Fields(0))
	If isNull(ytdf) Then ytdf="1994-01-03"
	sql2="SELECT Max(tradeDate) FROM ccass.calendar WHERE tradeDate<=DATE_SUB('"&ds&"', INTERVAL "
	sql="100*(enigma.totret(issueID,'"&ytdf&"','"&ds&"')-1) AS ytd"
	'look back 1,2,5,10 years but not beyond the start of our records
	If DateAdd("yyyy",-1,d2)>=#3-Jan-1994# Then
		y1f=MSdate(con.Execute(sql2&"1 YEAR)").Fields(0))
		sql=sql & ",100*(enigma.totret(issueID,'"&y1f&"','"&ds&"')-1) AS y1"
		If DateAdd("yyyy",-2,d2)>=#3-Jan-1994# Then
			y2f=MSdate(con.Execute(sql2&"2 YEAR)").Fields(0))
			sql=sql & ",100*(enigma.totret(issueID,'"&y2f&"','"&ds&"')-1) AS y2"
			If DateAdd("yyyy",-5,d2)>=#3-Jan-1994# Then
				y5f=MSdate(con.Execute(sql2&"5 YEAR)").Fields(0))
				sql=sql & ",100*(enigma.totret(issueID,'"&y5f&"','"&ds&"')-1) AS y5"
				If DateAdd("yyyy",-10,d2)>=#3-Jan-1994# Then
					y10f=MSdate(con.Execute(sql2&"10 YEAR)").Fields(0))
					sql=sql & ",100*(enigma.totret(issueID,'"&y10f&"','"&ds&"')-1) AS y10"
				End If
			End if
		End if
	End If
Else
	'show returns since d1
	col1="To-date"
	d1=con.Execute("SELECT MAX(tradeDate) FROM ccass.calendar WHERE tradeDate<='" & MSdate(d1) & "'").Fields(0)
	ds=MSdate(d1)
	title="Total returns on my stocks since "&ds
	ytdf=MSdate(latestDate)
	sql="100*(enigma.totret(issueID,'"&ds&"','"&ytdf&"')-1) AS ytd"
	'look forward 1,2,5,10 years but not beyond latest date
	If DateAdd("yyyy",1,d1)<=latestDate Then
		sql2="SELECT Max(tradeDate) FROM calendar WHERE tradeDate<=DATE_ADD('"&ds&"', INTERVAL "
		y1f=MSdate(con.Execute(sql2&"1 YEAR)").Fields(0))
		sql=sql & ",100*(enigma.totret(issueID,'"&ds&"','"&y1f&"')-1) AS y1"
		If DateAdd("yyyy",2,d1)<=latestDate Then
			y2f=MSdate(con.Execute(sql2&"2 YEAR)").Fields(0))
			sql=sql & ",100*(enigma.totret(issueID,'"&ds&"','"&y2f&"')-1) AS y2"
			If DateAdd("yyyy",5,d1)<=latestDate Then
				y5f=MSdate(con.Execute(sql2&"5 YEAR)").Fields(0))
				sql=sql & ",100*(enigma.totret(issueID,'"&ds&"','"&y5f&"')-1) AS y5"
				If DateAdd("yyyy",10,d1)<=latestDate Then
					y10f=MSdate(con.Execute(sql2&"10 YEAR)").Fields(0))
					sql=sql & ",100*(enigma.totret(issueID,'"&ds&"','"&y10f&"')-1) AS y10"
				End If
			End If
		End If
	End If
End If
fromDate=MSdate(d1)
toDate=MSdate(d2)
'restrict the sort order based on calculated columns
If isNull(y10f) Then
	If sort="10ydn" Then sort="5ydn"
	If sort="10yup" Then sort="5yup"
	If isNull(y5f) Then
		If sort="5ydn" Then sort="2ydn"
		If sort="5yup" Then sort="2yup"
		If isNull(y2f) Then
			If sort="2ydn" Then sort="1ydn"
			If sort="2yup" Then sort="1yup"
			If isNull(y1f) Then
				If sort="1ydn" Then sort="ytddn"
				If sort="1yup" Then sort="ytdup"
			End If
		End If
	End If
End If

Select Case sort
	Case "nameup" ob="Name1,typeShort"
	Case "namedn" ob="Name1 DESC,typeShort DESC"
	Case "scup" ob="sc"
	Case "scdn" ob="sc DESC"
	Case "ytddn" ob="ytd DESC"
	Case "ytdup" ob="ytd"
	Case "1ydn" ob="y1 DESC"
	Case "1yup" ob="y1"
	Case "2ydn" ob="y2 DESC"
	Case "2yup" ob="y2"
	Case "5ydn" ob="y5 DESC"
	Case "5yup" ob="y5"
	Case "10ydn" ob="y10 DESC"
	Case "10yup" ob="y10"
	Case Else
		sort="scup"
		ob="sc"
End Select
sql="SELECT issueID,name1,typeShort,enigma.lastcode(issueID) AS sc,"&sql
rs.Open sql& " FROM mailvote.mystocks m JOIN (enigma.issue i,enigma.organisations o,enigma.sectypes st) "&_
	"ON user="&ID&" AND m.issueID=i.ID1 AND i.issuer=o.personID AND i.typeID=st.typeID "&_
	"ORDER BY "&ob,mailcon
URL=Request.ServerVariables("URL")&"?d1="&fromDate&"&amp;d2="&toDate%>
<title><%=title%></title>
<link rel="stylesheet" type="text/css" href="../templates/main.css">
</head>
<body>
<!--#include file="../templates/cotopdb.asp"-->
<%Call userBar(10)%>
<ul class="navlist">
	<li><a href="/dbpub/TRnotes.asp">Notes</a></li>
	<li><a href="/dbpub/alltotrets.asp">Whole market</a></li>
</ul>
<div class="clear"></div>
<h2><%=title%></h2>
<form method="get" action="mytotrets.asp">
	<input type="hidden" name="sort" value="<%=sort%>">
	<div class="inputs">
		From <input type="date" name="d1" id="d1" value="<%=fromDate%>">
	</div>
	<div class="inputs">
		to <input type="date" name="d2" id="d2" value="<%=toDate%>">
	</div>
	<div class="inputs">
		<input type="submit" value="Go">
		<input type="button" value="Clear" onclick="document.getElementById('d1').value='';document.getElementById('d2').value='';">
	</div>
	<div class="clear"></div>
</form>
<p>This table shows Webb-site Total Returns on your stock list. Pick an ending 
date to look backwards or a starting date to look forwards, or both to show 
returns over a specific period. The data start on 3-Jan-1994. Returns for each 
stock are from the first day of trading in the period to the last day of trading 
in the period. "-" means the stock was not yet listed or was suspended 
throughout the period. Click the issue name to see the total return graph in 
that stock, or click the figure to see the graph over that period. Click a 
column-heading to sort.</p>
<%=mobile(2)%>
<table class="numtable">
	<tr>
		<th class="colHide1">Row</th>
		<th class="colHide2"><%SL "Stock<br>code","scup","scdn"%></th>
		<th class="left"><%SL "Issue","nameup","namedn"%></th>
		<th><%SL col1&"<br>%","ytddn","ytdup"%></th>
		<%If Not isNull(y1f) Then%>
			<th><%SL "1Y<br>%","1ydn","1yup"%></th>
			<%If Not isNull(y2f) Then%>
				<th><%SL "2Y<br>%","2ydn","2yup"%></th>
				<%If Not isNull(y5f) Then%>
					<th><%SL "5Y<br>%","5ydn","5yup"%></th>
					<%If Not isNull(y10f) Then%>
						<th class="colHide3"><%SL "10Y<br>%","10ydn","10yup"%></th>
					<%End if
				End If
			End If
		End If%>
	</tr>
	<%x=0
	Do Until rs.EOF
		x=x+1
		issueID=rs("issueID")
		If isnull(rs("ytd")) Then ytd="-" Else ytd=FormatNumber(rs("ytd"),2)%>
		<tr>
			<td class="colHide1"><%=x%></td>
			<td class="colHide2"><%=rs("sc")%></td>
			<td class="left"><a href="../dbpub/str.asp?i=<%=issueID%>"><%=rs("Name1")&":"&rs("typeShort")%></a></td>
			<td><a href="../dbpub/ctr.asp?i1=<%=issueID%>&amp;d1=<%=ytdf%>"><%=ytd%></a></td>
			<%If Not isNull(y1f) Then
				If isnull(rs("y1")) Then y1="-" Else y1=FormatNumber(rs("y1"),2)%>
				<td><a href="../dbpub/ctr.asp?i1=<%=issueID%>&amp;d1=<%=y1f%>"><%=y1%></a></td>
				<%If Not isNull(y2f) Then
					If isnull(rs("y2")) Then y2="-" Else y2=FormatNumber(rs("y2"),2)%>
					<td><a href="../dbpub/ctr.asp?i1=<%=issueID%>&amp;d1=<%=y2f%>"><%=y2%></a></td>
					<%If Not isNull(y5f) Then
						If isnull(rs("y5")) Then y5="-" Else y5=FormatNumber(rs("y5"),2)%>
						<td><a href="../dbpub/ctr.asp?i1=<%=issueID%>&amp;d1=<%=y5f%>"><%=y5%></a></td>
						<%If Not isNull(y10f) Then
							If isnull(rs("y10")) Then y10="-" Else y10=FormatNumber(rs("y10"),2)%>
							<td class="colHide3"><a href="../dbpub/ctr.asp?i1=<%=issueID%>&amp;d1=<%=y10f%>"><%=y10%></a></td>
						<%End If
					End If
				End If
			End If%>
		</tr>
		<%rs.MoveNext
	Loop
	Call CloseConRs(con,rs)
	Call CloseCon(mailcon)%>
</table>
<%If col1<>"Period" Then%>
	<%If col1="YTD" Then%>
		<h4>Starting dates</h4>
	<%Else%>
		<h4>Ending dates</h4>
	<%End If%>
	<table class="txtable">
		<tr><td>
		<%If col1="YTD" Then Response.Write "Year-to-date" Else Response.Write "To-date"%></td>
		<td><%=ytdf%></td></tr>
		<tr><td>1 year</td><td><%=y1f%></td></tr>
		<tr><td>2 years</td><td><%=y2f%></td></tr>
		<tr><td>5 years</td><td><%=y5f%></td></tr>
		<tr><td>10 years</td><td><%=y10f%></td></tr>
	</table>
<%End If
If x=0 Then%>
	<p>None found.</p>
<%End If%>
<!--#include virtual="/templates/footerdb.asp"-->
</body>
</html>