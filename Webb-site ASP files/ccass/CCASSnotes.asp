<%option explicit%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--#include file="../dbpub/functions1.asp"-->
<!--#include file="../dbpub/navbars.asp"-->
<link rel="stylesheet" type="text/css" href="../templates/main.css">
<%Dim title
title="About the Webb-site CCASS Analysis System"%>
<title><%=title%></title>
</head>
<body>
<!--#include file="../templates/cotopdb.asp"-->
<h2><%=title%></h2>
<%Call ccassallbar("",5)%>
<ol>
	<li>CCASS stands for Central Clearing and Automated Settlement System.</li>
	<li>CCASS is operated by Hong Kong Securities Clearing Company Limited (<strong>HKSCC</strong>), 
	wholly-owned by Hong Kong Exchanges and Clearing Limited (<strong>HKEx</strong>). 
	In order to settle a trade on the Stock Exchange of Hong Kong Ltd (<strong>SEHK</strong>, 
	also wholly-owned by HKEx), securities must be deposited with CCASS. 
	HKSCC then registers them in the name of its 100% subsidiary, HKSCC Nominees Limited (<strong>HKSCCN</strong>), 
	on the issuer's share register. Hong Kong is still an immobilized but not 
	scripless system.</li>
	<li>Consequently, listed share registers tend to have a huge holding in the 
	name of HKSCCN, and not much else, besides controlling shareholders and a 
	few employee holders. For this reason,
	<a target="_blank" href="http://www.hkex.com.hk/eng/newsconsul/hkexnews/2008/0804232news.htm">
	since</a> 28-Apr-08, HKEx has been
	<a target="_blank" href="http://www.hkexnews.hk/sdw/search/search_sdw.asp">
	disclosing</a> the list of CCASS Participant holdings on its web site, up to 
	1 year back.</li>
	<li>We capture, preserve and analyze the data. Analysis allows you to see 
	the date the holding last changed, the holdings of each participant across 
	the market, and 
	the net daily movements of holdings in each stock and of holdings by each 
	participant, as well as the time series of holdings by one participant in 
	one stock.</li>
	<li>&nbsp;In the 
	case of brokers, this will give you a clearer idea of what stocks 
	they deal in most, and if you have a margin account, what the pool of 
	collateral might include. This collateral pool is often pledged to lenders, 
	and if its value falls suddenly then it can trigger a brokerage collapse. By 
	looking at recent trends you can have some idea, up to 2 days ago, whether a 
	broker was selling or buying a line of stock. For institutional brokers, it 
	is harder to tell, as most of their clients hold stock through the major 
	custodians. If custodian holdings are increasing, then institutions are 
	probably net 
	buyers.</li>
	<li>Our records for shares and subscription warrants begin on 26-Jun-07, so any holding you see dated 
	26-Jun-2007 may not have changed on that date. For REITs, our records begin on 
	1-Mar-2011, because HKEx did not make them available before that date - the 
	records from that date onwards were released in 2012 on our request. We only 
	record changed data, but even then, we now (27-Nov-2015) have over 100 
	million records in over 2000 issues.</li>
	<li>Types of CCASS Participants are brokers, custodians, pledgees, clearing 
	houses and Investor Participants (<strong>IPs</strong>). The CCASS IDs 
of brokers are prefixed &quot;B&quot;, 
custodians are prefixed &quot;C&quot;, pledgee participants are prefixed &quot;P&quot; and 
	clearing houses are prefixed &quot;A&quot;. Broker participants may also be pledgees, 
	and custodians may hold pledged stock.</li>
	<li>With the exception of IPs, CCASS Participants may or may not have 
	beneficial interests in the shares they hold in CCASS, so don't use the 
	data as a guide to beneficial ownership.</li>
	<li>Holdings of IPs are not disclosed individually unless they have 
	consented. This allows them the same level of privacy that they would enjoy 
	if they held the shares through a broker.</li>
	<li>IP CCASS IDs are not published, so if they change their names, we have 
	to regard them as new participants.</li>
	<li>The domicile of corporate participants of all kinds is not disclosed, so the 
names may not uniquely identify them either.</li>
	<li>Issued shares are approximate, as issuers are not required to disclose 
	the figure whenever it changes. Consequently percentage stakes may be wrong 
	if the figure is outdated. We use the disclosed issued shares as at the 
	latest date on or prior to the date you are looking at.</li>
	<li>When a stock split or consolidation occurs, HKEx assigns a temporary 
	stock code to the issue, for trading in the "old" share certificates, 
	typically for 2 weeks, after which the original stock code reopens for the 
	"new" certificates. The old certificates then continue on the temporary 
	counter for another 3 weeks, known as "parallel trading" although in 
	practice there is very little activity. Formerly we did not track the 
	holdings on temporary stock codes, so during these periods percentage stakes may be wrong and our record 
of holdings may be frozen. Temporary counters and parallel trading have been an anachronism since CCASS 
	was introduced in 1992, and was due to be
	<a target="_blank" href="http://www.hkex.com.hk/eng/newsconsul/hkexnews/2008/080422news.htm">abolished</a> on 3-Nov-08, but abolition has been
	<a target="_blank" href="http://www.hkex.com.hk/eng/newsconsul/hkexnews/2008/080723news.htm">delayed</a> 
	indefinitely. In late 2015, we gave up waiting for HKEx to do the right 
	thing and got around to coding to collect the temporary counter holdings to provide a seamless 
	record while the original counter is closed, but as HKEx deletes information from its web site after 1 year, our 
	records for temporary counters begin on 27-Nov-2014. During parallel 
	trading, we collect the normal counter, not the temporary counter.</li>
	<li>Changes in holdings of an issue over a period show the actual holdings 
	at the end of the period. The change in shares held is calculated by 
	subtracting the holding at the start of the period, adjusted for stock 
	splits or bonus issues during the period. The change in the percentage stake 
	is the difference between the stake at the end of the period and at the 
	start of the period, using the respective outstanding shares at those 2 
	dates. The same applies to changes in a portfolio.</li>
	<li>When an issue begins trading "ex-bonus", there is a period of several 
	days before the bonus shares are actually issued. During this period, the 
	value of the ex-bonus shares will drop, all other things being equal. On an 
	adjusted basis, the number of shares held will also appear to drop as 
	holdings prior to the ex-date are adjusted upwards by the bonus factor.</li>
	<li>Sort the holdings in descending order or in date order to see former 
	holders with zero balances. To keep the list readable, we don't show zero 
	holdings when sorted by name or by ascending holding.</li>
	<li>Trades are normally settled on T+2, so the holdings you are 
	looking at, if they are due to market trades, relate to trades 2 clearing 
	days earlier. On certain half-days before a public holiday or during a 
	typhoon, there is no clearing, in which case trades from this day and the 
	previous day will be settled together. For example, trades on 21-Dec-2007 
	and 24-Dec-2007 were both settled on 28-Dec-2007, the second trading day 
	after Christmas.</li>
	<li>If a Participant ceases to exist, we do not remove its records, so the 
	list of Participants includes those who now have zero holdings in CCASS.</li>
	<li>Some movements are simply due to deposits or withdrawals of 
	securities not related to a market trade.</li>
	<li>We estimate the number of &quot;securities not in CCASS&quot; by deducting those 
	which are from the latest number of issued shares on or prior to that date, 
	as captured from the HKEx web site. Therefore if the number of issued 
	securities is wrong, then the number of securities not in CCASS is wrong.</li>
	<li>We disclaim any liability for any reliance on the data and for any 
	errors or omissions.</li>
</ol>
<!--#include virtual="/templates/footerdb.asp"-->
</body>
</html>
