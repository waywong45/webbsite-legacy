<%Option Explicit%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" type="text/css" href="/templates/main.css">
<title>Notes on short positions</title>
</head>
<body>
<!--#include file="../templates/cotopdb.asp"-->
<h2>Notes on short positions</h2>
<p>The figures you see result from the
<a href="http://www.hklii.hk/eng/hk/legis/reg/571AJ/" target="_blank">Securities and Futures (Short Position Reporting) Rules</a> 
(<strong>Rules</strong>), a piece of legislation produced as a knee-jerk 
response to the Global Financial Crisis. Because there was so little 
short-selling in HK, the Government <a href="../articles/soldshort.asp">set the 
threshold</a> low enough to detect it, so that the SFC would have something to 
report.</p>
<ol>
	<li>Reportable net short positions in "specified shares" at the end of each 
	week are filed with the SFC confidentially.</li>
	<li>The positions are aggregated and published by the 
		SFC as at the end of each week since 31-Aug-2012.</li>
	<li>Figures are published 
		1 week after the reference date. We combine these with our database of 
	outstanding share numbers in that class of shares (as disclosed to SEHK) resulting in percentage 
	stakes.</li>
	<li>Note that H-shares are usually not the only class of equity in PRC 
	issuers, and the controlling shareholder usually does not hold much of the 
	H-share class, so their free floats are larger, and hence the aggregate 
	short positions tend to be larger as a percentage of the class.</li>
	<li>Under the Rules, net short positions of HK$30m or 0.02% of the class of 
	shares (whichever is lower) in "specified shares" must be notified to SFC, but only if 
		the short sale was made on SEHK, not on an overseas exchange.</li>
	<li>Derivatives are not counted.</li>
	<li>SEHK has a
	<a href="http://www.hkex.com.hk/eng/market/sec_tradinfo/stkcdorder.htm" target="_blank">list of "designated securities"</a> for which (covered) short selling is 
		legal. Prior to 15-Mar-2017, under
		<a href="http://www.hklii.hk/eng/hk/legis/reg/571AJ/sch1.html" target="_blank">Schedule 1</a> of the Rules, the position 
	was only reportable if the 
		stock was in the Hang Seng Index or the Hang Seng China Enterprises 
		Index, or was classified by 
	<a href="http://www.hsi.com.hk/" target="_blank">Hang Seng Indexes Company 
	Ltd</a> as a "financial" stock. Since 15-Mar-2017, all stocks are covered, 
	which accounts for the jump in the total market short position on 
	17-Mar-2017, from HK$207bn to 340bn, and from 137 stocks to 937 stocks.</li>
	<li>We do not adjust the history for stock splits, consolidations or bonus issues.</li>
</ol>
<!--#include virtual="/templates/footerdb.asp"-->
</body>
</html>
