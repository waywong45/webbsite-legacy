<%Option Explicit%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--#include file="../dbeditor/prepMaster.inc"-->
<!--#include file="../dbpub/functions1.asp"-->
<%
Call requireRoleExec
Dim referer,tv,n,x,blnFnd,title,rs,org,orgName,addorg,hint,p,st,m,sql
Call prepMasterRs(conMaster,rs)	

blnFnd=False
Call getReferer(referer,tv)

Const limit=50
addorg=getLng("addorg",0)
If addorg>0 Then
	conMaster.Execute("INSERT IGNORE INTO advisers VALUES("&addorg&")")
	hint=hint&"Adviser with ID "&addorg&" added to table. "
End If
n=RemSpace(Request("n"))
If lcase(left(n,4))="the " Then	n=Right(n,len(n)-4)
st=Request("st")
title="Search advisers"
%>
<title><%=title%></title>
<link rel="stylesheet" type="text/css" href="../templates/main.css"/>
</head>
<body>
<!--#include file="cotop.inc"-->
<h2><%=title%></h2>
<p><b><%=hint%></b></p>
<form method="post" action="searchadvisers.asp">
	<h2><input type="text" style="font-size:medium" name="n" size="40" value="<%=n%>">
	<%=makeSelect("st",st,"l,Left match,a,Any match",True)%>
	<input type="submit" style="font-size:medium" value="search" name="search"></h2>
</form>
<%
If n<>"" Then
	If st="a" Then
		m = " AGAINST('+" & apos(join(split(n),"+")) & "' IN BOOLEAN MODE)"
		sql="MATCH name1" & m
	Else
		m= " LIKE '" & apos(n) & "%'"
		sql="name1" & m
	End If
	rs.Open "SELECT a.personID,name1,orgType,domicile,incID,shortName,incDate,disDate,SFCID FROM "&_
		"advisers a JOIN organisations o ON a.personID=o.personID LEFT JOIN domiciles d ON domicile=d.ID WHERE "&_
		sql & " ORDER BY Name1 LIMIT "&limit,conMaster
	%>
	<p>"*"=listed company</p>
	<h3>Matches in current names</h3>
	<form method="post" action="<%=referer%>">
	<%
	If rs.EOF then
		Response.write "None"
	Else
		blnFnd=True%>
		<table class="txtable">
			<tr>
				<th colspan="4"></th>
				<th>Name</th>
				<th>Established</th>
				<th>Dissolved</th>
				<th></th>
				<%If tv>"" Then%><th></th><%End If%>
			</tr>
			<%x=0
			Do Until rs.EOF
				x=x+1
				p=rs("personID")%>
				<tr>
					<td><%=x%></td>
					<td><%If rs("orgType")=22 Then Response.Write "*" Else Response.Write "&nbsp;"%></td>
					<td><input type="radio" name="<%=tv%>" value="<%=p%>"/></td>
					<td><%=rs("shortName")%></td>
					<td><a href='https://webb-site.com/dbpub/orgdata.asp?p=<%=p%>' target="_blank"><%=rs("Name1")%></a></td>
					<td><%=MSdate(rs("incDate"))%></td>
					<td><%=MSdate(rs("disDate"))%></td>
					<td><a href="org.asp?p=<%=p%>">Edit</a></td>
					<%If tv>"" Then%><td><a href="<%=referer%>?<%=tv%>=<%=p%>">Use</a></td><%End If%>
				</tr>
				<%
				rs.MoveNext
			Loop
			%>
		</table>
		<%If x=limit Then%>
			<p><b>Only the first <%=limit%> matches are displayed. Please narrow your 
			search.</b></p>
		<%End If
	End if
	rs.Close
	'search old names
	rs.Open "SELECT n.PersonID,OldName,orgType,dateChanged,name1,shortName,domicile,incID,SFCID "&_
		"FROM (NameChanges n JOIN (advisers a,organisations o) ON n.PersonID=a.PersonID AND a.personID=o.personID) "&_
		"LEFT JOIN domiciles d ON o.domicile=d.ID WHERE "&_
		sql & " ORDER BY OldName LIMIT "&limit,conMaster
	%>
	<h3>Matches in old names</h3>
	<%If rs.EOF then
		Response.write "None"
	Else
		blnFnd=True%>
		<table class="txtable">
		<tr>
			<th colspan="4"></th>
			<th>Name</th>
			<th>Until</th>
			<th>Current name</th>
			<th></th>
			<%If tv>"" Then%><th></th><%End If%>
		</tr>
		<%x=0
		Do Until rs.EOF
			x=x+1
			p=rs("personID")%>
			<tr>
				<td><%=x%></td>
				<td><%If rs("orgType")=22 Then Response.Write "*" Else Response.Write "&nbsp;"%></td>
				<td><input type="radio" name="<%=tv%>" value="<%=rs("PersonID")%>"></td>
				<td><%=rs("shortName")%></td>
				<td><a href='https://webb-site.com/dbpub/orgdata.asp?p=<%=p%>' target="_blank"><%=rs("OldName")%></a></td>
				<td><%=MSdate(rs("dateChanged"))%></td>		
				<td><%=rs("name1")%></td>
				<td><a href="org.asp?p=<%=p%>">Edit</a></td>
				<%If tv>"" Then%><td><a href="<%=referer%>?<%=tv%>=<%=p%>">Use</a></td><%End If%>
			</tr>
			<%rs.MoveNext
		Loop
		%>
		</table>
		<%If x=limit Then%>
			<p><b>Only the first <%=limit%> matches are displayed. Please narrow your 
			search.</b></p>
		<%End If
	End if
	rs.Close%>
	<p><b><%=hint%></b></p>
	<%If referer>"" And blnFnd And tv>"" Then%>
		<p><input type="submit" name="submitBtn" value="Use selected organisation"></p>
	<%End If%>
	</form>
	<h3>Not what you are looking for?</h3>
	<p><a href="searchorgs.asp?tv=org&amp;n=<%=n%>">Find or add an organisation</a></p>
	<%
End if

'have we returned from trying to add an org to use as an adviser?
org=Request("org")
If org>"" Then
	orgName=conMaster.Execute("SELECT name1 FROM organisations WHERE personID="&org).Fields(0)
	%>
	<form action="searchadvisers.asp">
		<input type="hidden" name="addorg" value="<%=org%>">
		<input type="hidden" name="n" value="<%=orgName%>">
		<table class="txtable">
			<tr>
				<th>ID</th>
				<th>Name</th>
				<th></th>
			</tr>
			<tr>
				<td><%=org%></td>
				<td><a target="_blank" href="https://webb-site.com/dbpub/orgdata.asp?p=<%=org%>"><%=orgName%></a></td>
				<td><input type="submit" name="submitBtn" value="Add to advisers"></td>
			</tr>
		</table>
	</form>
<%End If
Call closeConRs(conMaster,rs)%>
<!--#include file="cofooter.asp"-->
</body>
</html>
