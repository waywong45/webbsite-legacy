<%If month(Date)=6 and day(Date)=4 Then%>
	<style>
	body {
		  filter: grayscale(1);
	}
	</style>
<%End If
Call cookiechk%>
<!--#include virtual="templates/cookiechk.asp"-->
<div id="banner" style="background-color:maroon">
	<div class="box1">
		<a href="/dbpub/" class="nodec">
		<span style="font-size:1.6em;margin:0"><b>Webb-site Database</b></span><br>
		<span style="font-size:0.9em"><b>Scientia potentia est</b></span><br></a>
		<div id="rss" style="float:left;height:30px;padding:2px;margin-top:4px;">
			<a type="application/rss+xml" href="/rss.asp"><img alt="RSS feed" src="/images/RSS28x28.png"></a>
			<div id="social" style="float:right;margin-left:2px">
				<a href="https://x.com/webbhk"><img alt="Follow us on X" src="/images/x27x28.png" style="background-color:black;margin-left:2px"></a>
				<a href="https://www.facebook.com/webbfb"><img alt="Follow us on Facebook" src="/images/facebook28x28.png" style="margin-left:2px"></a>
			</div>
		</div>
		<label for="menuchk" id="menubtn">Menu</label>
		<div id="loginbtn">
			<%If Session("e")="" Then%>
				<a href="/webbmail/login.asp" class="nodec">log in</a>
			<%Else%>
				<a href="/webbmail/myratings.asp" class="nodec">logged in</a>
			<%End If%>
		</div>
		<div class="clear"></div>
		<div id="volunteer">
			<a href="/webbmail/username.asp" class="nodec"><b>Volunteer to edit the database</b></a>
		</div>
		<label for="srchchk" id="srchbtn">search</label>
	</div>
	<input type="checkbox" id="srchchk" style="display:none">
	<div id="srchblk" style="background-color:inherit;">
		<div class="box4">
			<!-- SiteSearch Google -->
			<form class="box4a" method="get" action="https://www.google.com/search">
				<input type="hidden" name="ie" value="UTF-8">
				<input type="hidden" name="oe" value="UTF-8">
				<input type="hidden" name="domains" value="Webb-site.com">
				<input type="hidden" name="sitesearch" value="Webb-site.com">
				<input type="text" class="inptxt searchws" name="q" maxlength="255" value="search Webb-site.com" onclick="value=''">
				<input type="submit" class="btnFont" name="btnG" value="search">
			</form>
			<form class="box4b" method="post" action="/webbmail/join.asp">
				<input type="text" class="inptxt signup" name="e" value="email address" onclick="value=''">
				<input type="submit" class="btnFont" value="sign up">
				<input type="hidden" name="R1" value="join">
			</form>
		</div>
		<div class="group1">
			<div class="box3">
				<form class="box3a" method="post" action="/dbpub/searchorgs.asp" style="margin-bottom:5px">
					<input type="text" class="inptxt orgsearch" name="n" maxlength="255" value="Organisation" onclick="value=''">
					<input type="submit" class="btnFont" name="btnG" value="search organisations">
				</form>
				<form class="box3b" method="post" action="/dbpub/searchpeople.asp">
					<input type="text" class="inptxt famsearch" name="n1" maxlength="255" value="Family name" onclick="value=''">
					<input type="text" class="inptxt famsearch" name="n2" maxlength="255" value="First name" onclick="value=''">
					<input type="submit" class="btnFont" name="btnG" value="search people">
				</form>
			</div>
			<form class="stockbox" action="/dbpub/orgdata.asp" method="get" name="f1">
				<p>Stock code</p>
				<input type="number" class="inptxt stockcode" name="code" min="1" max="99999" pattern="[0-9]*" onclick="value=''"><br>
				<input type="submit" class="btnFont" name="Submit" value="current" onclick="f1.action='/dbpub/orgdata.asp'">
				<input type="submit" class="btnFont" value="past" onclick="f1.action='/dbpub/code.asp'">
			</form>
		</div>
	</div>
</div>
<div id="menubar" style="background-color:maroon;">
	<div class="hnav">
		<input type="checkbox" id="menuchk" style="display:none">
		<ul>
			<li><a href="/dbpub/">Home</a></li>
			<li><a href="/">Webb-site Reports</a></li>
			<li><a href="/webbmail/login.asp">User</a>
				<ul>
					<%If session("editor") Then%>
						<li><a href="/dbeditor/">Edit database</a></li>
					<%End If%>
					<%If session("e")<>"" Then%>
						<li><a href="/webbmail/myratings.asp">My ratings</a></li>
						<li><a href="/webbmail/mystocks.asp">My stocks</a></li>
						<li><a href="/webbmail/mybigchanges.asp">Big CCASS changes</a></li>
						<li><a href="/webbmail/mytotrets.asp">Total returns</a></li>
						<li><a href="/webbmail/mysdi.asp">Dealings</a></li>
						<li><a href="/webbmail/mailpref.asp">Mail On/Off</a></li>
						<li><a href="/webbmail/changeaddr.asp">Change address</a></li>
						<li><a href="/webbmail/username.asp">Username/Volunteer!</a></li>
						<li><a href="/webbmail/reset.asp">Change password</a></li>
					<%Else%>
						<li><a href="/webbmail/login.asp">Log in</a></li>
					<%End If%>
					<li><a href="/webbmail/join.asp">Sign up</a></li>
					<li><a href="/webbmail/forgot.asp">Forgot password</a></li>
					<%If session("e")<>"" Then%>
						<li><a href="/webbmail/login.asp?b=1">Log out</a></li>
					<%End If%>
				</ul>
			</li>
			<li><a href="/contact/">Contact</a></li>
			<li><a href="/pages/refer.asp">Tell a Friend!</a></li>
		</ul>
	</div>
</div>
<div class="clear"></div>
<div class="mainbody">
