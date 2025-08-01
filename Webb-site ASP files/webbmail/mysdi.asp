﻿<%Option Explicit%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--#include virtual="dbpub/functions1.asp"-->
<!--#include virtual="dbpub/navbars.asp"-->
<%Call login
Dim orderBy,sort,title,issue,stockName,person,posType,posLabel,shsInv,holding,stake,stkChg,price,avCon,value,relDate,settleDate,stockCode,ID,mailcon,rs,URL
URL=Request.ServerVariables("URL")
ID=session("ID")
Call openMailrs(mailcon,rs)
sort=Request("sort")
Select Case sort
	Case "codeup" orderBy="stockCode,pName,relDate DESC"
	Case "codedn" orderBy="stockCode DESC,pName,relDate DESC"
	Case "nameup" orderBy="pName,oName,relDate DESC"
	Case "namedn" orderBy="pName DESC,oName,relDate DESC"
	Case "stokup" orderBy="oName,pName,relDate DESC"
	Case "stokdn" orderBy="oName DESC,pName,relDate DESC"
	Case "rsnup" orderBy="rsnSht,oName,pName,relDate DESC"
	Case "rsndn" orderBy="rsnSht DESC,oName,pName,relDate DESC"
	Case "lvalup" orderBy="value,oName,pName"
	Case "lvaldn" orderBy="value DESC,oName,pName"
	Case "reldup" orderBy="relDate,oName,pName"
	Case Else orderBy="relDate DESC,oName,pName":sort="relddn"
End Select
title="My director & CEO dealings"%>
<title><%=title%></title>
<link rel="stylesheet" type="text/css" href="../templates/main.css"/>
</head>
<body>
<!--#include file="../templates/cotopdb.asp"-->
<%Call userBar(12)%>
<h2><%=title%></h2>
<ul class="navlist">
	<li><a href="/dbpub/sdiNotes.asp">Notes</a></li>
	<li><a href="/dbpub/sdilatest.asp">Whole market</a></li>
</ul>
<div class="clear"></div>
<%
rs.Open "SELECT t2.id,filing,relDate,issueID,enigma.lastCode(issueID) as stockCode,posType,rsnSht,rsnLng,dir,shsInv,longShs1,longShs2,shortShs1,shortShs2,price,avCon,currency,"&_
	"longStk2,shortStk2,lngStkChg,shtStkChg,CONCAT(p.name1,', ',IFNULL(p.name2,''),' ',IFNULL(p.cname,'')) AS pName,p.personID as holderID,"&_
	"IFNULL(price,avCon)*shsInv AS value,CONCAT(o.Name1,':',typeShort) as oName,settleDate FROM "&_
	"(SELECT id,curr,filing,relDate,s.issueID,dir,longShs1,longShs2,shortShs1,shortShs2,IFNULL(avPrice,hiPrice) as price,	avCon,longStk2,shortStk2,"&_
	"longStk2-longStk1 AS lngStkChg,shortStk2-IFNULL(shortStk1,0) AS shtStkChg FROM enigma.sdi s JOIN mystocks m ON s.issueID=m.issueID AND user="&ID&_
	" WHERE isNull(serNoSuper) ORDER BY relDate DESC LIMIT 200) as t2 "&_
	"JOIN (enigma.people p,enigma.sdievent,enigma.sdireason r,enigma.issue i,enigma.organisations o,enigma.secTypes st) "&_
	"ON t2.dir=p.personID AND t2.ID=sdiID AND reason=r.id AND issueID=i.ID1 AND i.issuer=o.personID AND i.typeID=st.typeID "&_
	"LEFT JOIN enigma.currencies c ON curr=c.id LEFT JOIN ccass.calendar c3 ON relDate=c3.tradeDate ORDER BY " & orderBy,mailcon
If not rs.EOF Then%>
	<p>The latest 200 filings in your stocks are shown. Click the date to see more details. Click on a stock to see all 
	filings in that stock. Click on a name to see 
	all filings by that person in that stock. L=Long, S=Short. Click the on-exchange price to 
	see the CCASS movements on the settlement date corresponding to the 
	relevant event.</p>
	<%=mobile(1)%>
	<table class="numtable" style="font-size:9pt">
		<tr>
			<th class="colHide3"><%SL "Relevant<br>date<br>Y-M-D","relddn","reldup"%></th>
			<th class="colHide2"><%SL "Stock<br>code","codeup","codedn"%></th>
			<th class="left"><%SL "Stock","stokup","stokdn"%></th>
			<th class="left"><%SL "Name","nameup","namedn"%></th>
			<th class="colHide3 left"><%SL "Stated<br>reason","rsnup","rsndn"%></th>
			<th>L<br/>S</th>
			<th>Shares<br/>involved</th>
			<th class="colHide1">Curr</th>
			<th>OnEx<br/>Price</th>
			<th class="colHide2">OffEx<br/>Price</th>
			<th class="colHide1"><%SL "Value","lvaldn","lvalup"%></th>
			<th class="colHide1" colspan="2" style="text-align:center">Stake,<br/>Change<br/>%</th>
		</tr>
	<%Do Until rs.EOF
		posType=rs("posType")
		shsInv=rs("shsInv")
		price=rs("price")
		avCon=rs("avCon")
		value=rs("value")
		issue=rs("issueID")
		relDate=rs("relDate")
		settleDate=rs("settleDate")
		If posType=1 then posLabel="L" Else posLabel="S"
		If posType=1 Then
			holding=rs("longShs2")
			stake=rs("longStk2")
			stkChg=rs("lngStkChg")
			If rs("longShs2")<rs("longShs1") and rs("rsnSht")<>"Acquire" Then
				shsInv=-shsInv
				value=-value
			End If
		Else
			holding=-rs("shortShs2")
			stake=-rs("shortStk2")
			stkChg=-rs("shtStkChg")
			If rs("shortShs2")>rs("shortShs1") or isNull(rs("shortShs1")) Then shsInv=-shsInv
		End If
		If not isNull(shsInv) Then shsInv=FormatNumber(shsInv,0)
		If not isnull(holding) Then holding=FormatNumber(holding,0)
		If not isNull(stake) Then stake=FormatNumber(stake,2)
		If not isNull(stkChg) Then stkChg=FormatNumber(stkChg,2)
		If not isNull(value) Then value=FormatNumber(value,0)
		If not isNull(price) Then price=formatNumber(price,3)
		If not isNull(avCon) Then avCon=formatNumber(avCon,3)
		%>
		<tr>
			<td class="colHide3"><a href='../dbpub/sdicap.asp?r=<%=rs("id")%>'><%=MSsdate(relDate)%></a></td>
			<td class="colHide2"><%=rs("stockCode")%></td>
			<td class="left"><a href="../dbpub/sdiissue.asp?i=<%=issue%>"><%=rs("oName")%></a></td>
			<td class="left"><a href='../dbpub/sdidirco.asp?p=<%=rs("holderID")%>&amp;i=<%=issue%>'><%=rs("pName")%></a></td>
			<td class="colHide3 left"><span class="info"><%=rs("rsnSht")%><span><%=rs("rsnLng")%></span></span></td>
			<td><%=posLabel%></td>
			<td><%=shsInv%></td>
			<%If posType=1 Then%>
				<td class="colHide1"><%=rs("currency")%></td>
				<%if not isNull(price) and relDate>#22-Jun-2007# and settleDate+1.2<Now Then%>
					<td><a href="../ccass/chldchg.asp?i=<%=issue%>&d=<%=MSdate(settleDate)%>"><%=price%></a></td>
				<%Else%>
					<td><%=price%></td>		
				<%End If%>
				<td class="colHide2"><%=avCon%></td>
				<td class="colHide1"><%=value%></td>
			<%Else%>
				<td class="colHide1" style="background-color:gray"></td>
				<td style="background-color:gray"></td>
				<td class="colHide2" style="background-color:gray"></td>
				<td class="colHide1" style="background-color:gray"></td>
			<%End If%>
			<td class="colHide1"><%=stake%></td>
			<td class="colHide1"><%=stkChg%></td>
		</tr>
		<%rs.MoveNext
	Loop%>
	</table>
<%Else%>
	<p><b>None found.</b></p>
<%End If
Call CloseConRs(mailcon,rs)%>
<!--#include virtual="/templates/footerdb.asp"-->
</body>
</html>