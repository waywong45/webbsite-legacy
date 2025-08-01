Option Explicit On
Option Compare Text
Imports ScraperKit
Imports JSONkit
Module Quotes
    Sub Main()
        Call GetHSITRI()
        Call GetShorts()
        Call CheckAllEnts()
        Call GetQuotesUpdate()
        Call UpdateFX()
        Call ForeignDivs()
        Call SetAdjAll()
        'Console.ReadKey()
    End Sub
    Sub ForeignDivs()
        'convert any dividends which are not in the trading currency of the stock and have passed the ex-date, then recalculate adjustments
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset,
            d As Date, e, i As Integer
        Call OpenEnigma(con)
        rs.Open("SELECT * from events e JOIN (issue i,capchangetypes ct)  on e.issueID=i.ID1 AND e.eventType=ct.CapChangeType " &
                "WHERE (Not isNull(currID)) AND isNull(priceHKD) AND isNull(cancelDate) AND (isNull(i.SEHKcurr) or i.SEHKcurr<>e.currID) AND ct.dist=True " &
                "AND exDate BETWEEN '1994-01-01' and CURRENT_DATE()", con)
        Do Until rs.EOF
            e = CInt(rs("eventID").Value)
            i = CInt(rs("issueID").Value)
            If IsDBNull(rs("cumDate").Value) Then
                d = CumDate(CInt(rs("issueID").Value), CDate(rs("exDate").Value))
                If d <> Nothing Then con.Execute("UPDATE events SET cumDate='" & MSdate(d) & "' WHERE eventID=" & e)
            Else
                d = CDate(rs("cumDate").Value)
            End If
            If d <> Nothing Then
                Call ConvDistE(e)
                Call SetAdj(i)
            End If
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
        con = Nothing
    End Sub

    Sub ConvDistE(e As Integer)
        'convert value of a single distribution of an event, based on FX rate on last cumDate
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset,
            currID, SEHKcurr, currPair As Integer
        Call OpenEnigma(con)
        rs.Open("SELECT i.SEHKcurr,e.currID FROM events e JOIN issue i ON e.issueID=i.ID1 WHERE eventID=" & e, con)
        currID = DBint(rs("currID"))
        SEHKcurr = DBint(rs("SEHKcurr"))
        rs.Close()
        If currID > 0 Or SEHKcurr > 0 Then
            If SEHKcurr = currID Then
                'distribution is in same non-HKD currency as quote
                con.Execute("UPDATE events SET priceHKD=price WHERE price<>0 AND eventID=" & e)
            Else
                'currency mismatch between distribution and quote
                rs.Open("SELECT ID FROM currpair WHERE curr1=" & currID & " AND curr2=" & SEHKcurr, con)
                If rs.EOF Then
                    'no such FX series, so try inverting the pair
                    rs.Close()
                    rs.Open("SELECT ID FROM currpair WHERE curr1=" & SEHKcurr & " AND curr2=" & currID, con)
                    If rs.EOF Then
                        'No pair, so calculate cross-rate via HKD if possible, e.g. USD div to CNY for a CNY-quoted stock is USDHKD/CNYHKD
                        rs.Close()
                        rs.Open("SELECT ID from currpair WHERE curr1=" & currID & " AND curr2=0", con)
                        If Not rs.EOF Then
                            currPair = CInt(rs("ID").Value)
                            rs.Close()
                            rs.Open("SELECT ID from currpair WHERE curr1=" & SEHKcurr & " AND curr2=0", con)
                            If Not rs.EOF Then
                                con.Execute("UPDATE events SET priceHKD = price * ROUND((SELECT rate FROM forexrates WHERE currPair=" & currPair & " AND atDate=cumDate)/" &
                                "(SELECT rate FROM forexrates WHERE currPair=" & CInt(rs("ID").Value) & " AND atDate=cumDate),4)," &
                                "FXdate=cumDate WHERE price<>0 AND (Not isNull(cumDate)) AND eventID=" & e)
                            End If
                        End If
                    Else
                        con.Execute("UPDATE events SET priceHKD=ROUND(price/(SELECT rate FROM forexrates WHERE currPair=" & CInt(rs("ID").Value) & " AND atDate=cumDate),4)," &
                        "FXdate=cumDate WHERE price<>0 AND (Not isNull(cumDate)) AND eventID=" & e)
                    End If
                Else
                    con.Execute("UPDATE events SET priceHKD=price*ROUND((SELECT rate FROM forexrates WHERE currPair=" & CInt(rs("ID").Value) & " AND atDate=cumDate),4)," &
                                "FXdate=cumDate WHERE price<>0 AND (Not isNull(cumDate)) AND eventID=" & e)
                End If
                rs.Close()
            End If
        End If
        con.Close()
        con = Nothing
    End Sub

    Sub UpdateFX()
        'new version 2024-09-06 after Yahoo ended free API
        'update the FX rates up to the last whole day
        Dim d As Date, x, currpair, dateCol, closeCol As Integer,
            sym, arr(,), rate As String,
        con As New ADODB.Connection, rs As New ADODB.Recordset
        Call OpenEnigma(con)
        rs.Open("SELECT cp.ID,concat(c1.currency,c2.currency) sym FROM currpair cp JOIN(currencies c1,currencies c2) ON curr1=c1.ID AND curr2=c2.ID", con)
        Do Until rs.EOF
            currpair = CInt(rs("ID").Value)
            d = CDate(con.Execute("SELECT Max(atDate) FROM forexrates WHERE currpair=" & currpair).Fields(0).Value)
            If d < Today.AddDays(-1) Then
                sym = rs("sym").Value.ToString
                arr = GetYahoo(sym, d.AddDays(1), Today.AddDays(-1))
                dateCol = 0
                closeCol = 4
                For x = 1 To UBound(arr, 2)
                    rate = arr(closeCol, x)
                    If rate <> "null" Then
                        con.Execute("INSERT IGNORE INTO  forexrates (currpair,atDate,rate) VALUES (" & currpair & ",'" & arr(dateCol, x) & "'," & rate & ")")
                        Console.WriteLine(sym & vbTab & arr(dateCol, x) & vbTab & rate)
                    End If
                Next
            End If
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
        con = Nothing
    End Sub
    Sub FillFX(p As Integer, d1 As Date)
        'new version 2024-09-06 scrapes Yahoo site after free API died
        'NOT YET TESTED
        'backfill an FX rate for currency pair p from earliest available date d on Yahoo
        Dim x, dateCol, closeCol As Integer,
            sym, arr(,), rate As String,
            d2 As Date,
            con As New ADODB.Connection, rs As New ADODB.Recordset
        Call OpenEnigma(con)
        sym = con.Execute("SELECT concat(c1.currency,c2.currency) sym FROM currpair cp JOIN(currencies c1,currencies c2) " &
                          "ON curr1=c1.ID And curr2=c2.ID WHERE cp.ID=" & p).Fields(0).Value.ToString
        d2 = CDate(con.Execute("SELECT Min(atDate) FROM forexrates WHERE currpair=" & p).Fields(0).Value).AddDays(-1)
        arr = GetYahoo(sym, d1, d2)
        dateCol = 0
        closeCol = 4
        For x = 1 To UBound(arr, 2)
            rate = arr(closeCol, x)
            If rate <> "null" Then
                'con.Execute("INSERT IGNORE INTO  forexrates (currpair,atDate,rate) VALUES (" & p & ",'" & arr(dateCol, x) & "'," & rate & ")")
            End If
            Console.WriteLine(sym & vbTab & arr(dateCol, x) & vbTab & rate)
        Next
        con.Close()
        con = Nothing
    End Sub
    Function GetYahoo(ByVal symbol As String, ByVal d1 As Date, ByVal d2 As Date) As String(,)
        Dim u1, u2 As Long,
            URL, r, rows, row, cell, shortmonth(), arr(,) As String,
            x, y, rowcnt, col As Integer
        u1 = DateDiff(DateInterval.Second, #1970-01-01#, d1)
        u2 = DateDiff(DateInterval.Second, #1970-01-01#, d2)
        URL = "https://finance.yahoo.com/quote/" & symbol & "=X/history/?period1=" & u1 & "&period2=" & u2
        r = GetWeb(URL)
        x = InStr(r, "adjusted for splits and dividend") 'a tooltip note near start of table
        If x = 0 Then x = 1
        rows = ""
        Call TagCont(x, r, "tbody", rows)
        ReDim arr(6, 0)
        'deliberately don't use the first row, so Ubound detects results
        'now read the rows
        y = 1
        row = ""
        shortmonth = Split("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec")
        Do Until y > Len(rows)
            Call TagCont(y, rows, "tr", row)
            If row > "" Then
                rowcnt += 1
                ReDim Preserve arr(6, rowcnt)
                x = 1
                For col = 0 To 6
                    cell = ""
                    Call TagCont(x, row, "td", cell)
                    cell = Trim(cell)
                    If col = 0 Then
                        Console.Write(cell & " ")
                        'decode date format "mmm d, yyy" to Unix
                        cell = Right(cell, 4) & "-" & Right("0" & (1 + Array.IndexOf(shortmonth, Left(cell, 3))).ToString, 2) & "-" &
                            Right("0" & Mid(cell, 5, InStr(cell, ",") - 5), 2)
                        Console.Write(cell & " ")
                    End If
                    arr(col, rowcnt) = Trim(cell)
                    Console.Write(arr(col, rowcnt) & " ")
                Next
                Console.WriteLine()
            End If
        Loop
        Return arr
    End Function
    Function SplitDet(ByVal s As String) As String()
        'convert the details of an entitlement on HKEX quote page to a 6-element array
        'some details are malformed, less than 40 characters per column, so we have to allow for that
        Dim det(5) As String, x, y As Integer
        y = 0
        Do Until s = "" Or y > 5
            x = InStr(s, "   ") 'take 3 spaces to indicate end of column
            If x > 0 And x < 38 Then
                Do Until Mid(s, x, 1) <> " "
                    x += 1
                Loop
            Else
                x = 41
            End If
            'we found double spaces in some details
            det(y) = StripSpace(Replace(Left(s, x - 1), "&amp;", "&"))
            s = Mid(s, x)
            y += 1
        Loop
        Return det
    End Function
    Sub GetQuotesUpdate()
        'bring quotes up to date
        On Error GoTo RepErr
        Dim atDate As Date, r As String, target As Date, con As New ADODB.Connection
        If Now.ToString("HH:mm") > "21:00" Then
            target = Today.Date
        Else
            target = Today.Date.AddDays(-1)
        End If
        Call OpenEnigma(con)
        atDate = CDate(GetLog("MBquotesDate"))
        'first the Main Board
        Do Until atDate = target
            atDate = atDate.AddDays(1)
            If NotHol(atDate) Then
                Console.WriteLine("Fetching Main Board quotes for " & atDate)
                r = GetMBDQS(atDate)
                If r <> "" Then
                    Console.WriteLine("Parsing MB shortnames for " & atDate)
                    Call ParseShortNames(r, atDate, 1)
                    Call ParseQuotes(r, atDate, 1)
                    Call PutLog("MBquotesDate", MSdate(atDate))
                    Console.WriteLine("MB quotes done: " & atDate)
                Else
                    Console.WriteLine("Main Board quotes not found for " & atDate)
                    Call SendMail("Main Board quotes not found for " & atDate)
                End If
            End If
        Loop
        'now do GEM
        atDate = CDate(GetLog("GEMquotesDate"))
        Do Until atDate = target
            atDate = atDate.AddDays(1)
            If NotHol(atDate) Then
                Console.WriteLine("Fetching GEM quotes for " & atDate)
                r = GetGEMDQS(atDate)
                If r <> "" Then
                    Console.WriteLine("Parsing GEM shortnames for " & atDate)
                    Call ParseShortNames(r, atDate, 20)
                    Call ParseQuotes(r, atDate, 20)
                    Call PutLog("GEMquotesDate", MSdate(atDate))
                    Console.WriteLine("GEM quotes done: " & atDate)
                Else
                    Console.WriteLine("GEM quotes not found for " & atDate)
                    Call SendMail("GEM quotes not found for " & atDate)
                End If
            End If
        Loop
        con.Close()
        con = Nothing
        Exit Sub
RepErr:
        Call ErrMail("GetQuotesUpdate failed:" & MSdate(atDate), Err)
    End Sub
    Function GetDQS(root As String, fileName As String) As String
        'get contents of one daily quotation sheet from file, downloading it first if it doesn't exist
        Dim path As String
        path = GetLog("QuotesFolder") & fileName
        If Not FileIO.FileSystem.FileExists(path) Then Call Download(root & fileName, path,,, False)
        If FileIO.FileSystem.FileExists(path) Then Return My.Computer.FileSystem.ReadAllText(path) Else Return Nothing
    End Function
    Function GetGEMDQS(ByVal atDate As Date) As String
        'download a GEM Daily Quotation Sheet and return its contents
        Return GetDQS(GetLog("GEMquotesSource"), "e_G" & Format(atDate, "yyMMdd") & ".htm")
    End Function
    Function GetMBDQS(ByVal atDate As Date) As String
        'download a Main Board Daily Quotation Sheet
        GetMBDQS = GetDQS(GetLog("MBquotesSource"), "d" & Format(atDate, "yyMMdd") & "e.htm")
    End Function
    Sub GetHSITRI()
        'fetch the daily Hang Seng Index total returns sheet and save it. Not used on Webb-site.
        Dim d As Date, err As String
        d = CDate(GetLog("HSITRIdone"))
        Do Until d = Today.Date
            err = ""
            d = d.AddDays(1)
            If Weekday(d, vbMonday) < 6 Then
                Call HSITRIone(d, err)
                If err = "" Then Call PutLog("HSITRIdone", MSdate(d))
            End If
        Loop
    End Sub
    Sub HSITRIone(d As Date, Optional err As String = "")
        Call Download(GetLog("HSITRIsource") & "idx_" & Format(d, "dMMMyy") & ".csv", GetLog("HSITRIfolder") & MSdate(d) & ".csv", err)
        If err = "" Then Console.WriteLine("Downloaded HSITRI for " & d.ToString) Else Console.WriteLine("Error while downloading HSITRI for " & d.ToString & " " & err)
    End Sub
    Sub CheckAllEnts()
        'check all the entitlements shown on HKEX quote pages
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset, t As String, changes As Boolean
        Call OpenEnigma(con)
        t = HKEXtoken()
        'NASDAQ pilot scheme entitlements are not available
        'No entitlements for Rights,Bonds, CBonds, ExBonds
        rs.Open("SELECT stockCode,issueID FROM stocklistings JOIN issue " &
                "ON stocklistings.issueID=issue.ID1 " &
                "WHERE stockExID IN (1, 20, 22, 23, 38) " &
                "And typeID Not In (1,2,40,41,46) " &
                "And (isNull(firstTradeDate) Or firstTradeDate<=NOW()) " &
                "And (isNull(deListDate) Or delistDate>NOW()) " &
                "ORDER BY stockCode", con)
        Do Until rs.EOF
            Call CheckEnts(rs("stockCode").Value.ToString, CInt(rs("issueID").Value), t, changes)
            rs.MoveNext()
            'Try slowing down to avoid throttle
            Call WaitNSec(0.5)
        Loop
        rs.Close()
        con.Close()
        con = Nothing
        Call PutLog("AllEntsChecked", MSdateTime(Now()))
    End Sub
    Sub ProcNewEnts()
        On Error GoTo repErr
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset
        Call OpenEnigma(con)
        rs.Open("SELECT importID FROM entitlements WHERE NOT autoprocess AND NOT ignoreRow", con)
        Do Until rs.EOF
            Call ProcEnt(CInt(rs("importID").Value))
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
        Exit Sub
repErr:
        ErrMail("ProcNewEnts failed", Err)
    End Sub
    Sub CheckEnts1stock(sc As String, issueID As Integer)
        Dim t As String, changes As Boolean
        t = HKEXtoken()
        Console.WriteLine("HKEXtoken: " & t)
        Call CheckEnts(sc, issueID, t, changes)
    End Sub
    Sub CheckEnts(ByVal sc As String, ByVal issueID As Integer, ByVal token As String, ByRef changes As Boolean)
        'check the entitlements of one stock
        On Error GoTo RepErr
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset, rsev As New ADODB.Recordset,
        det(), details, r, e, ents(), s, announced, yearEnd, exDate, bcFrom, bcTo, payDate, sql As String,
            y, cct, eventID, importID As Integer,
            detChng, dateChng, xdChng As Boolean,
            d As Date
        Console.WriteLine("Checking ents for stock:" & sc & " issueID:" & issueID)
        changes = False
        Call OpenEnigma(con)
        r = GetEnts(token, sc)
        r = GetItem(r, "data.entitlementlist")
        If r = "" Or r = "[]" Then
            con.Close()
            Exit Sub
        End If
        ents = ReadArray(r)
        For Each e In ents
            xdChng = False
            detChng = False
            dateChng = False
            announced = "'" & MSdate(CDate(GetVal(e, "announcement_date"))) & "'"
            s = GetVal(e, "fyear_end")
            If IsDate(s) Then yearEnd = "'" & MSdate(CDate(s)) & "'" Else yearEnd = "NULL"
            exDate = GetVal(e, "ex_date")
            bcFrom = GetVal(e, "book_closed_from")
            bcTo = GetVal(e, "book_closed_to")
            payDate = GetVal(e, "payment_date")
            det = SplitDet(GetVal(e, "detail"))
            details = Trim(Join(det, " "))
            s = sc & " " & announced & " " & yearEnd & " " & details
            sql = "SELECT * FROM entitlements WHERE issueID=" & issueID & " AND announced=" & announced
            If yearEnd = "NULL" Then sql &= " AND isNull(yearEnd)" Else sql &= " AND YearEnd=" & yearEnd
            'HKEX has split some entitlements into two records, so we need to check details_1, cannot just rely on NULL
            'NB there's a risk of creating duplicate entitlements if details_1 changes
            sql &= " AND details_1='" & Apos(det(0)) & "'"
            con.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs.Open(sql, con)
            If rs.EOF Then
                changes = True
                If IsDate(exDate) Then exDate = "'" & MSdate(CDate(exDate)) & "'" Else exDate = "NULL"
                If IsDate(bcFrom) Then bcFrom = "'" & MSdate(CDate(bcFrom)) & "'" Else bcFrom = "NULL"
                If IsDate(bcTo) Then bcTo = "'" & MSdate(CDate(bcTo)) & "'" Else bcTo = "NULL"
                If IsDate(payDate) Then payDate = "'" & MSdate(CDate(payDate)) & "'" Else payDate = "NULL"
                sql = "INSERT INTO entitlements (daily,STOCK,IssueID,Announced,YearEnd,exDate,BookCloseFr,BookCloseTo," &
                    "PayDate,DETAILS_1,DETAILS_2,DETAILS_3,DETAILS_4,DETAILS_5,DETAILS_6) VALUES (True,'" & sc & "'," &
                issueID & "," & announced & "," & yearEnd & "," & exDate & "," & bcFrom & "," & bcTo & "," & payDate & ",'" &
                det(0) & "','" & det(1) & "','" & det(2) & "','" & det(3) & "','" & det(4) & "','" & det(5) & "')"
                con.Execute(sql)
                y = LastID(con)
                Console.WriteLine("New entitlement: " & y & vbTab & s)
                'fill the event columns if possible
                rs.Close()
                Call ProcEnt(y) 'process the new entitlement and create events
            Else
                If rs.RecordCount = 1 Then
                    s = "Unique match on details_1 : " & s
                Else
                    'more than 1 match
                    rs.Close()
                    sql &= " AND details_2='" & Apos(det(1)) & "'"
                    rs.Open(sql, con)
                    If rs.RecordCount = 1 Then
                        s = "Unique match on details_1&2 " & s
                    Else
                        'this only happened once, Glencore plc on 2017-02-03, now delisted
                        rs.Close()
                        sql &= " AND details_3='" & Apos(det(2)) & "'"
                        rs.Open(sql, con)
                        If rs.RecordCount = 1 Then
                            s = "Unique match on details_1&2&3 " & s
                        End If
                    End If
                End If
                If rs.EOF Then
                    rs.Close()
                Else
                    'now we have a unique entitlement to work on
                    importID = CInt(rs("importID").Value)
                    rs.Close()
                    Console.WriteLine(importID & vbTab & s)
                    rs.Open("SELECT * FROM entitlements WHERE importID=" & importID, con, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                    For y = 1 To 6
                        If det(y - 1) <> CStr(rs("DETAILS_" & y).Value) Then
                            detChng = True
                            rs("DETAILS_" & y).Value = det(y - 1)
                            rs("col" & y).Value = False 'flag column as not processed
                            rs("AutoProcess").Value = False 'resurface
                        End If
                    Next
                    If detChng Then Console.WriteLine("Detail change on stock:" & sc & vbTab & announced & " " & yearEnd & " " & Join(det, " "))
                    'prepare dates for comparison
                    If IsDate(exDate) Then exDate = MSdate(CDate(exDate))
                    If IsDate(bcFrom) Then
                        bcFrom = MSdate(CDate(bcFrom))
                    Else
                        'record date (instantaneous book close) may be in the details
                        If InStr(details, "Record date") > 0 Then
                            d = FindDate(details, "d Date")
                            If d <> Nothing Then bcFrom = MSdate(d)
                        End If
                    End If
                    If IsDate(bcTo) Then bcTo = MSdate(CDate(bcTo))
                    If IsDate(payDate) Then payDate = MSdate(CDate(payDate))
                    'compare dates with existing values
                    s = MSdate(DBdate(rs("ExDate")))
                    If s <> exDate Then
                        xdChng = True
                        dateChng = True
                        Console.WriteLine("exDate changed")
                        If exDate = "" Then rs("ExDate").Value = DBNull.Value Else rs("exDate").Value = exDate
                    End If
                    s = MSdate(DBdate(rs("BookCloseFr")))
                    If s <> bcFrom Then
                        dateChng = True
                        Console.WriteLine("bcFrom changed")
                        If bcFrom = "" Then rs("BookCloseFr").Value = DBNull.Value Else rs("BookCloseFr").Value = bcFrom
                    End If
                    s = MSdate(DBdate(rs("BookCloseTo")))
                    If s <> bcTo Then
                        dateChng = True
                        Console.WriteLine("bcTo changed")
                        If bcTo = "" Then rs("BookCloseTo").Value = DBNull.Value Else rs("BookCloseTo").Value = bcTo
                    End If
                    s = MSdate(DBdate(rs("PayDate")))
                    If s <> payDate Then
                        dateChng = True
                        Console.WriteLine("PayDate changed")
                        If payDate = "" Then rs("PayDate").Value = DBNull.Value Else rs("PayDate").Value = payDate
                    End If
                    rs.Update()
                    rs.Close()
                    If detChng Or dateChng Then changes = True
                    If detChng Then
                        'redo all the events
                        Call ProcEnt(importID)
                    ElseIf dateChng Then
                        'only the tabulated dates have changed
                        rs.Open("SELECT * FROM entitlements WHERE importID=" & importID, con)
                        'mirror the datechanges to events. Relevant fields depend on the event type
                        Console.WriteLine("Date change on stock:" & sc & vbTab & announced & " " & yearEnd & " " & Join(det, " "))
                        For y = 1 To 4
                            If Not IsDBNull(rs("Event" & y).Value) Then
                                eventID = CInt(rs("Event" & y).Value)
                                'there should be an event if we haven't deleted it
                                cct = CInt(rs("CapChangeType" & y).Value)
                                rsev.Open("SELECT * From events WHERE eventID=" & eventID, con, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                                If Not rsev.EOF Then
                                    If cct <> 4 And cct <> 48 Then
                                        '4=split/consol, 48=share exchange
                                        rsev("exDate").Value = rs("exDate").Value
                                        rsev("bookCloseFr").Value = rs("bookCloseFr").Value
                                        rsev("bookCloseTo").Value = rs("bookCloseTo").Value
                                    End If
                                    'don't blank the payDate if we know it
                                    If Not {2, 4, 8, 45, 46, 47, 48, 49, 52, 54}.Contains(cct) And payDate <> "" Then rsev("distDate").Value = payDate
                                    rsev.Update()
                                End If
                                rsev.Close()
                                If cct <> 4 And cct <> 48 And xdChng Then Call SetCumDateEv(eventID)
                            End If
                        Next
                        rs.Close()
                    End If
                End If
            End If
        Next
        con.Close()
        con = Nothing
        Exit Sub
RepErr:
        Call ErrMail("CheckEnts failed for stock code " & sc & ", issueID " & issueID, Err)
    End Sub
    Function GEMtransfer(ByVal issueID As Integer, ByVal s As String) As Boolean
        'check the prospective transfer from GEM to Main Board and enter if missing
        Dim sc As Integer, d As Date, scStr As String
        If InStr(s, "Transfer of Listing from GEM") = 0 Then Return False
        sc = FindInt(s, "code")
        If sc = 0 Then Return False
        'don't go further if the transfer has already happened - then the stock code in the details is the old code
        If sc >= 8000 And sc < 9000 Then Return True
        d = FindDate(s, "on")
        If d = Nothing Then Return False
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset
        Call OpenEnigma(con)
        'delist the GEM stock if not already delisted
        rs.Open("SELECT * FROM stocklistings WHERE stockExID=20 AND issueID=" & issueID, con, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If Not rs.EOF Then
            rs("DelistDate").Value = d
            rs("ReasonID").Value = 2 'To main board
            rs.Update()
        End If
        rs.Close()
        'list on the main board if not already listed
        scStr = CStr(sc)
        '5-digits are possible, but otherwise we record a 4-digit string
        If Len(scStr) < 4 Then scStr = Right("000" & scStr, 4)
        rs.Open("SELECT * FROM stocklistings WHERE stockExID=1 AND issueID=" & issueID & " AND stockCode=" & sc,
                con, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If rs.EOF Then
            rs.AddNew()
            rs("issueID").Value = issueID
            rs("StockExID").Value = 1
            rs("stockCode").Value = scStr
        End If
        rs("FirstTradeDate").Value = d
        rs.Update()
        rs.Close()
        con.Close()
        Return True 'the transfer is done
    End Function
    Sub ProcEnt(ByVal e As Integer)
        'process the details of an entitlement and create or update events if possible
        'NEW VERSION 2021-11-14 to handle "No Dividend for the period ended" with a YYYY/MM/DD date in the next field 
        On Error GoTo RepErr
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset, re As New ADODB.Recordset,
        col, cct, currID, eventID, issueID, fdists, x, issuer, yeMonth, periodMonth As Integer,
        d, cancelDate As Date, done, blnDist, addNew, nextColDone, next2ColDone As Boolean,
        cctCol, colName, currCol, currStr, dateCol, eventCol, newCol, oldCol, priceCol, priceHKDcol, likes, s(), ss, t, u, curr(1, 0),
        dist(2, 0), sql, details As String,
        price, priceHKD, rNew, rOld As Double
        Call OpenEnigma(con)
        sql = "SELECT * FROM entitlements WHERE importID=" & e
        re.Open(sql, con, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If re.EOF Then
            'entitlement doesn't exist
            re.Close()
            con.Close()
            con = Nothing
            Exit Sub
        End If
        issueID = CInt(re("issueID").Value)
        issuer = CInt(con.Execute("SELECT issuer FROM issue WHERE ID1=" & issueID).Fields(0).Value)
        rs.Open("SELECT * FROM documents WHERE orgID=" & issuer & " AND docTypeID=0 ORDER BY RecordDate DESC", con)
        If rs.EOF Then
            'no annual report yet, so use current year-end of distribution (if any)
            If Not IsDBNull(re("YearEnd").Value) Then
                yeMonth = Month(CDate(re("YearEnd").Value))
            End If
        Else
            yeMonth = Month(CDate(rs("RecordDate").Value))
        End If
        rs.Close()
        details = con.Execute("SELECT details FROM entall WHERE importID=" & e).Fields(0).Value.ToString
        Console.WriteLine("Stock code:" & re("STOCK").Value.ToString & vbTab & "issueID:" & issueID & vbTab & "importID:" & e &
                          vbTab & "announced:" & MSdate(CDate(re("Announced").Value)))
        If GEMtransfer(issueID, details) Then
            Console.WriteLine(details)
            re("IgnoreRow").Value = True
            re.Update()
            re.Close()
            con.Close()
            con = Nothing
            Exit Sub
        End If
        'set record date if contained in the details
        If InStr(details, "Record date") > 0 Then
            d = FindDate(details, "d Date")
            If d <> Nothing Then
                re("BookCloseFr").Value = d
                re.Update()
            End If
        End If
        For Each u In {"cancelled", "not to proceed", "terminated", "not approved", "not passed", "not to conduct", "lapsed", "withdrew", "withdraw"}
            If InStr(details, u) > 0 Then
                cancelDate = FindDate(details, u) 'use on events
                Exit For
            End If
        Next
        re.Close()
        'fetch array of currency codes
        rs.Open("SELECT ID,HKEXcurr FROM currencies WHERE NOT ISNULL(HKEXcurr)", con)
        'load array and avoid late bound resolution warnings from getrows
        x = 0
        Do Until rs.EOF
            ReDim Preserve curr(1, x)
            curr(0, x) = rs("ID").Value.ToString
            curr(1, x) = rs("HKEXcurr").Value.ToString
            rs.MoveNext()
            x += 1
        Loop
        rs.Close()
        'fetch array of distributions with "LIKE" alternates
        rs.Open("SELECT capChangeType AS cct,likestr,`Change` FROM capchangetypes WHERE dist AND NOT ISNULL(likestr)", con)
        x = 0
        Do Until rs.EOF
            ReDim Preserve dist(2, x)
            dist(0, x) = rs("cct").Value.ToString
            dist(1, x) = rs("likestr").Value.ToString
            dist(2, x) = rs("Change").Value.ToString
            rs.MoveNext()
            x += 1
        Loop
        rs.Close()
        fdists = 1 'when searching for HKD equivalent distribution
        nextColDone = False
        next2ColDone = False
        For col = 1 To 6
            eventCol = "event" & col
            cctCol = "CapChangeType" & col
            priceCol = "Price" & col
            priceHKDcol = "PriceHKD" & col
            oldCol = "Old" & col
            newCol = "New" & col
            dateCol = "Date" & col
            currCol = "Curr" & col
            done = False
            re.Open(sql, con, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            'there is only 1 event per column, max
            colName = "DETAILS_" & col
            t = re(colName).Value.ToString
            If t = "" Or nextColDone Then
                'shift the flags
                done = True
                nextColDone = next2ColDone
                next2ColDone = False
                If col < 5 Then
                    re(eventCol).Value = DBNull.Value
                    re(cctCol).Value = DBNull.Value
                    re(currCol).Value = DBNull.Value
                    re(priceCol).Value = DBNull.Value
                    re(priceHKDcol).Value = DBNull.Value
                    re(oldCol).Value = DBNull.Value
                    re(newCol).Value = DBNull.Value
                    re(dateCol).Value = DBNull.Value
                End If
            Else
                'check for terms which are either ignored or dealt with elsewhere
                For Each u In {"(Record", "(cancelled", "(lapsed", "(not to proceed", "(terminated", "(not approved", "(not passed", "(not to conduct",
                    "(withdrew", "(equivalent", "(with scrip", "(in scrip form with",
                    "(with option", "payable", "(change of", "(approx", "w.e.f", "(AUD", "(GBP", "(HKD", "(JPY", "(MYR", "(RMB", "(USD", "(in lieu", "(Detail",
                    "(for six", "(for twelve", "(gross", "(including", "(with GBP", "(with HKD", "(with JPY", "(with RMB", "(with SGD", "(with USD"}
                    If Left(t, Len(u)) = u Then
                        done = True
                        If col < 5 Then
                            re(eventCol).Value = DBNull.Value
                            re(cctCol).Value = DBNull.Value
                            re(currCol).Value = DBNull.Value
                            re(priceCol).Value = DBNull.Value
                            re(priceHKDcol).Value = DBNull.Value
                            re(oldCol).Value = DBNull.Value
                            re(newCol).Value = DBNull.Value
                            re(dateCol).Value = DBNull.Value
                        End If
                        If Left(u, 1) = "(" And InStr(t, ")") = 0 Then nextColDone = True 'extends to the next column, treat that as done
                        Exit For
                    End If
                Next
            End If
            'process each of the first four DETAILS_N columns - we assume no events start in columns 5 & 6
            If (Not done) And col < 5 And InStr(t, "by way") + InStr(t, "in form") + InStr(t, "in specie") = 0 Then
                'try to process the entitlement
                currStr = ""
                price = 0
                priceHKD = 0
                d = Nothing
                'find distributions currency and amount
                done = False
                blnDist = False
                'NEW 2021-11-14
                If t = "No Dividend for the period ended" Or t = "No Distribution for the period ended" Then
                    periodMonth = Month(CDate(re("DETAILS_" & col + 1).Value))
                    nextColDone = True 'skip next column
                    If periodMonth < yeMonth Then periodMonth += 12
                    If t = "No Dividend for the period ended" Then
                        Select Case periodMonth - yeMonth
                            Case 0 : cct = 33 'Final dividend - but what if the year-end has changed?
                            Case 3 : cct = 27 'Q1 dividend
                            Case 6 : cct = 34 'Interim dividend
                            Case 9 : cct = 31 'Q3 dividend
                            Case Else : cct = 7 'Dividend
                        End Select
                    Else
                        'No distribution
                        Select Case periodMonth - yeMonth
                            Case 0 : cct = 44 'Final distribution - but what if the year-end has changed?
                            Case 3 : cct = 58 'Q1 distribution
                            Case 6 : cct = 40 'Interim distribution
                            Case 9 : cct = 59 'Q3 distribution
                            Case Else : cct = 37 'Distribution
                        End Select
                    End If
                    re(priceCol).Value = price
                    re(newCol).Value = DBNull.Value
                    re(oldCol).Value = DBNull.Value
                    re(dateCol).Value = DBNull.Value
                    done = True
                    blnDist = True
                Else
                    'END NEW 2021-11-14
                    cct = -1
                    For x = 0 To UBound(curr, 2)
                        currStr = curr(1, x)
                        s = Split("div, dist, cash bonus", ",")
                        For Each ss In s
                            If InStr(t, ss & " " & currStr) + InStr(t, ss & " approx. " & currStr) > 0 Then
                                price = FindDbl(t, ss)
                                currID = CInt(curr(0, x))
                                re(currCol).Value = currID
                                re(priceCol).Value = price
                                done = True
                                blnDist = True
                                Exit For
                            End If
                        Next
                        If done Then Exit For
                    Next
                End If

                If Not done Then currStr = ""
                If done And currID > 0 Then
                    'distribution in a non-HKD currency, look for equivalent in HKD
                    'there may be more than one foreign div, so find the nth one
                    If InStr(details, "approx. HKD") + InStr(details, "Equivalent to HKD") + InStr(details, "(HKD") > 0 Then
                        priceHKD = FindDbl(details, "HKD", fdists)
                        If priceHKD > 0 Then
                            fdists += 1
                            Console.WriteLine("(HKD " & priceHKD & ")")
                            re(priceHKDcol).Value = priceHKD
                        End If
                    End If
                End If
                If cct = -1 Then
                    done = False 'reset because the event type is not yet known
                    'now try to find a distribution type - even if there is no payout
                    For x = 0 To UBound(dist, 2)
                        likes = colName & " Like'" & Join(Split(dist(1, x), ","), "' OR " & colName & " LIKE '") & "'"
                        rs.Open(sql & " AND (" & likes & ")", con)
                        If Not rs.EOF Then
                            Console.WriteLine(col & vbTab & t & vbTab & dist(2, x) & vbTab & currStr & vbTab & price)
                            cct = CInt(dist(0, x))
                            re(priceCol).Value = price
                            re(newCol).Value = DBNull.Value
                            re(oldCol).Value = DBNull.Value
                            re(dateCol).Value = DBNull.Value
                            rs.Close()
                            blnDist = True
                            done = True
                            Exit For
                        End If
                        rs.Close()
                    Next
                End If
                'Not a distribution or incomplete details so try other events
                If Not done Then
                    'try bonus issues
                    rs.Open(sql & " AND (" & colName & " RLIKE '^Bonus [0-9]* for [0-9]*' OR " &
                            colName & " RLIKE '^Capitalisation issue [0-9]* for [0-9]*' OR " &
                            colName & " RLIKE '^Capitalization issue [0-9]* for [0-9]*')", con)
                    If Not rs.EOF Then
                        rNew = StripInt(t)
                        rOld = FindInt(t, "for")
                        Console.WriteLine(t & vbTab & rNew & vbTab & rOld)
                        done = True
                    Else
                        'if there's a bonus then it spans more than 1 column
                        rs.Close()
                        rs.Open("SELECT * FROM entall WHERE importID=" & e & " AND (" & colName & " LIKE 'Capitalization issue%' OR " &
                                colName & " LIKE 'Capitalisation issue%' OR " &
                                colName & " LIKE 'Bonus issue%') And Details RLIKE '[0-9]* for [0-9]*'", con)
                        If Not rs.EOF Then
                            u = details
                            x = FindStr(u, "Capitalization issue") + FindStr(u, "Capitalisation issue") + FindStr(u, "Bonus issue")
                            u = Mid(u, x)
                            rNew = StripInt(u)
                            rOld = FindInt(u, "for")
                            Console.WriteLine(details & vbTab & rNew & vbTab & rOld)
                            done = True
                            nextColDone = True
                            If InStr(re("DETAILS_" & col + 1).Value.ToString, " for ") = 0 Then next2ColDone = True 'event covers 3 columns
                        End If
                    End If
                    rs.Close()
                    If done Then
                        cct = 5 'bonus issue
                    Else
                        'try splits and consols
                        rs.Open(sql & " AND " & colName & " RLIKE '^(Split|consolidation) [0-9]* into [0-9]*'", con)
                        If Not rs.EOF Then
                            cct = 4
                            rOld = StripInt(t)
                            rNew = FindInt(t, "into")
                            d = FindDate(details, "w.e.f.")
                            Console.WriteLine(t & vbTab & rOld & vbTab & rNew & vbTab & d)
                            done = True
                        Else
                            rs.Close()
                            'try rights and open offers
                            rs.Open(sql & " AND " & colName & " RLIKE '^(Rts|Open offer) [.0-9]* for [.0-9]*'", con)
                            If Not rs.EOF Then
                                If Left(t, 3) = "Rts" Then cct = 2 Else cct = 8
                                rNew = FindNum(t)
                                rOld = FindDbl(t, "for")
                                price = FindDbl(t, "HKD")
                                If price = 0 And col < 6 Then
                                    'sometimes it's x for y consolidated sh @HKD.. and it spans 2 columns
                                    price = FindDbl(t & rs("details_" & col + 1).Value.ToString, "HKD")
                                    If price > 0 Then nextColDone = True
                                End If
                                If price > 0 Then
                                    re(currCol).Value = 0 'HKD
                                    re(priceCol).Value = price
                                End If
                                d = FindDate(details, "payable")
                                Console.WriteLine(t & vbTab & rNew & vbTab & rOld & vbTab & price & vbTab & d)
                                done = True
                            End If
                        End If
                        rs.Close()
                    End If
                    If done Then
                        re(newCol).Value = rNew
                        re(oldCol).Value = rOld
                        If d = Nothing Then re(dateCol).Value = DBNull.Value Else re(dateCol).Value = d 'split date or acceptance date
                    End If
                End If
                'now process the event
                If done Then
                    re(cctCol).Value = cct
                    re.Update()
                    re.Close()
                    'mirror the values into the event, including nulls, except for certain event types
                    re.Open(sql, con, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                    'allow for possible shift of columns due to insertion of HKD values, e.g. an entitlement with a CNY dividend and a bonus issue subsequently gets an HKD value added
                    'also possible left-shift e.g. if an effective split date is postponed
                    'assume events in an entitlement have unique eventTypes
                    rs.Open("SELECT * FROM events WHERE importID=" & e & " AND eventType=" & cct, con, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                    'rs.Open("SELECT * FROM events WHERE importID=" & e & " AND subEnt=" & col, con, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                    addNew = False
                    If rs.EOF Then
                        'assume new event - otherwise we risk overwriting, for example, when there are two special divs declared on the same date in different entitlements
                        addNew = True
                        rs.AddNew()
                        rs("issueID").Value = issueID
                        rs("importID").Value = e
                    Else
                        eventID = CInt(rs("eventID").Value)
                    End If
                    'if the events have shifted to the right then this could cause a duplicate key for ImportID-SubEnt, but it should resolve after the ent is processed, e.g 2-3 and 3-4
                    rs("subEnt").Value = col
                    rs("announced").Value = re("announced").Value
                    rs("eventType").Value = cct
                    If blnDist Then rs("yearEnd").Value = re("yearEnd").Value
                    If price > 0 Or {2, 4, 5, 8}.Contains(cct) Then
                        If cct = 4 Then 'split/consol
                            rs("exDate").Value = re(dateCol).Value
                            If IsDBNull(re(dateCol).Value) Then
                                rs("cumDate").Value = DBNull.Value
                            Else
                                d = CumDate(issueID, CDate(re(dateCol).Value))
                                If d = Nothing Then rs("cumDate").Value = DBNull.Value Else rs("cumDate").Value = d
                            End If
                        Else
                            rs("exDate").Value = re("exDate").Value
                            If IsDBNull(re("exDate").Value) Then
                                rs("cumDate").Value = DBNull.Value
                            Else
                                d = CumDate(issueID, CDate(re("exDate").Value))
                                If d = Nothing Then rs("cumDate").Value = DBNull.Value Else rs("cumDate").Value = d
                            End If
                        End If
                        If cct = 2 Or cct = 8 Then 'rights/open offer
                            rs("acceptDate").Value = re(dateCol).Value
                        End If
                        If cct <> 4 Then
                            'distribution, rights, open offer, bonus
                            rs("bookCloseFr").Value = re("BookCloseFr").Value
                            rs("bookCloseTo").Value = re("BookCloseTo").Value
                            'don't blank the PayDate as HKEX may miss it
                            If Not IsDBNull(re("PayDate").Value) Then rs("distDate").Value = re("PayDate").Value
                            If priceHKD > 0 Then
                                rs("priceHKD").Value = priceHKD
                                rs("FXdate").Value = DBNull.Value
                            End If
                        End If
                    End If
                    If cancelDate = Nothing Then rs("cancelDate").Value = DBNull.Value Else rs("cancelDate").Value = cancelDate
                    rs("New").Value = re(newCol).Value
                    rs("old").Value = re(oldCol).Value
                    rs("currID").Value = re(currCol).Value
                    rs("price").Value = re(priceCol).Value
                    rs.Update()
                    rs.Close()
                    If addNew Then
                        eventID = LastID(con)
                        Console.WriteLine("Event added:" & eventID)
                    End If
                    re(eventCol).Value = eventID
                    Call SetCumDateEv(eventID)
                End If
            End If
            re("col" & col).Value = done
            re.Update()
            re.Close()
        Next
        con.Execute("UPDATE entitlements SET autoprocess=col1*col2*col3*col4*col5*col6 WHERE importID=" & e)
        con.Close()
        con = Nothing
        Exit Sub
RepErr:
        Call ErrMail("ProcEnt failed for eventID:" & e, Err)
    End Sub
    Function HKEXtoken() As String
        'fetch the token from HKEX to use with JSON calls
        Dim x As Integer, y As Integer, r As String
        r = GetWeb(GetLog("HKEXtokenPage"))
        x = InStr(r, "Base64")
        x = InStr(x, r, "return") + 8
        y = InStr(x, r, """")
        Return Mid(r, x, y - x)
    End Function
    Function FindDbl(ByVal s As String, ByVal p As String, Optional ByVal n As Integer = 1) As Double
        'p is a CSV string
        'searches s for any of the values in p, and then finds the number after that
        'if n>1 then p is a string, finds the nth occurence of p and then the number after that
        'example: FindDbl("Search String Div HKD 0.20","Div,Cash Bonus,Dist")
        'example: FInd DBl("Search String Fin Div HKD 0.20 Sp Div HKD 0.30,"HKD")
        Dim c As Integer
        If n > 1 Then
            c = 1
            For n = 1 To n
                c = InStr(c, s, p)
                If c = 0 Then Exit For
                c += Len(p)
            Next
        Else
            For Each p In Split(p, ",")
                c = InStr(s, p)
                If c > 0 Then
                    c += Len(p)
                    Exit For
                End If
            Next
        End If
        If c = 0 Then Return Nothing Else Return StripDbl(Mid(s, c))
    End Function
    Function FindInt(ByVal s As String, ByVal p As String) As Integer
        'searches a string for any of the pretexts in the CSV string, and then finds the first number after that
        'example: FindInt([Details_1],"Div,Cash Bonus,Dist")
        Dim c As Integer
        For Each p In Split(p, ",")
            c = InStr(s, p)
            If c > 0 Then Exit For
        Next
        If c = 0 Then Return Nothing
        FindInt = StripInt(Mid(s, c + Len(p)))
    End Function
    Function FindNum(ByVal s As String) As Double
        'returns the first number in the string, or if none is found Nothing (0)
        Dim c As Integer, t As String
        s = Replace(s, ",", "")
        For c = 1 To Len(s)
            If IsNumeric(Mid(s, c, 1)) Then Exit For
        Next
        s = Mid(s, c)
        For c = 1 To Len(s)
            t = Mid(s, c, 1)
            If Not (t = "." Or IsNumeric(t)) Then Exit For
        Next
        s = Left(s, c - 1)
        If IsNumeric(s) Then Return CDbl(s) Else Return Nothing
    End Function
    Sub SetCumDateEv(e As Integer)
        'update the cumDate of an event after setting new exDate
        'run before calculating adjustments
        'exclude cancelled events unless they went ex-entitlement before being cancelled
        'exclude event if stock is not yet at the exDate.
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset,
            maxDate, maxGEMdate As Date
        Call OpenEnigma(con)
        maxDate = CDate(GetLog("MBquotesDate"))
        maxGEMdate = CDate(GetLog("GEMquotesDate"))
        If maxDate > maxGEMdate Then maxDate = maxGEMdate
        maxDate = NextTradingDay(maxDate)
        con.Execute("UPDATE events SET cumDate=(SELECT Max(atDate) FROM ccass.quotes WHERE noclose=False AND issueID=events.issueID AND atDate<exDate)" &
            " WHERE eventID=" & e & " AND (isNull(cancelDate) OR exDate<cancelDate) AND exDate<='" & MSdate(maxDate) & "'")
        con.Execute("UPDATE events SET cumDate=Null,cumPrice=Null,adjust=Null WHERE eventID=" & e &
                    " AND (isNull(exDate) or exDate>'" & MSdate(maxDate) & "')")
        con.Close()
        con = Nothing
    End Sub
    Sub SetCumDateIssue(IssueID As Long)
        'update the cumDates of an issue after setting new exDates or adding events
        'run before calculating adjustments
        'exclude cancelled events unless they went ex-entitlement before being cancelled
        'exclude event if stock is not yet at the exDate.
        Console.WriteLine("Setting cumDates for issueID: " & IssueID)
        Dim con As New ADODB.Connection, maxDate, maxGEMdate As Date
        Call OpenEnigma(con)
        maxDate = CDate(GetLog("MBquotesDate"))
        maxGEMdate = CDate(GetLog("GEMquotesDate"))
        If maxDate > maxGEMdate Then maxDate = maxGEMdate
        maxDate = NextTradingDay(maxDate)
        con.Execute("UPDATE enigma.Events SET cumDate=(SELECT Max(atDate) FROM ccass.quotes WHERE issueID=" & IssueID &
                    " AND atDate<exDate and noclose=false) WHERE issueID=" & IssueID &
                    " AND (isNull(cancelDate) OR exDate<cancelDate) AND exDate<='" & MSdate(maxDate) & "'")
        con.Execute("UPDATE enigma.events SET cumDate=Null,cumPrice=Null,adjust=Null WHERE issueID=" & IssueID &
                     " AND (isNull(exDate) or exDate>'" & MSdate(maxDate) & "')")
        con.Close()
        con = Nothing
        Console.WriteLine("cumDates done for issueID: " & IssueID)
    End Sub
    Function CumDate(issueID As Integer, xd As Date) As Date
        'compute the cumDate disregarding the actual event
        'return nothing if we haven't finished the day before xd date, because stock could still be suspended
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset, maxDate, maxGEMdate As Date, d As Date
        Call OpenEnigma(con)
        maxDate = CDate(GetLog("MBquotesDate"))
        maxGEMdate = CDate(GetLog("GEMquotesDate"))
        If maxDate > maxGEMdate Then maxDate = maxGEMdate
        maxDate = NextTradingDay(maxDate)
        If maxDate < xd Then
            'still more than a day before xd
            d = Nothing
        Else
            rs.Open("SELECT Max(atDate) cumDate FROM ccass.quotes WHERE noclose=False And atDate <'" & MSdate(xd) & "' AND issueID =" & issueID, con)
            d = DBdate(rs("cumDate"))
            rs.Close()
        End If
        con.Close()
        con = Nothing
        Return d
    End Function
    Function StripDbl(ByVal s As String) As Double
        Dim c As Integer, t As String
        'Find the first digit
        For c = 1 To Len(s)
            If IsNumeric(Mid(s, c, 1)) Then Exit For
        Next
        s = Mid(s, c)
        For c = 1 To Len(s)
            t = Mid(s, c, 1)
            If Not (t = "," Or t = "." Or IsNumeric(t)) Then Exit For
        Next
        t = Left(s, c - 1)
        If IsNumeric(t) Then
            StripDbl = CDbl(t)
            s = Mid(s, c)
            If (Left(s, 4) = " per" And Left(s, 8) <> " per HDR" And Left(s, 7) <> " per sh") Or Left(s, 5) = " for " Then
                c = StripInt(s)
                If c > 0 Then StripDbl /= c
            End If
        Else
            Return Nothing
        End If
    End Function
    Function StripInt(ByVal s As String) As Integer
        'strip an integer out of a string, coping with commas in the format which Val function cannot do
        Dim c As Integer, t As String
        'Find the first digit
        For c = 1 To Len(s)
            If IsNumeric(Mid(s, c, 1)) Then Exit For
        Next
        'Find the end of the number by looking for a non-number or non-comma
        s = Mid(s, c)
        For c = 1 To Len(s)
            t = Mid(s, c, 1)
            If Not (t = "," Or IsNumeric(t)) Then Exit For
        Next
        s = Left(s, c - 1)
        If IsNumeric(s) Then Return CInt(s) Else Return Nothing
    End Function


    Function GetEnts(ByRef token As String, ByVal sc As String) As String
        'Get the entitlements of one stock
        'the widget is on a different server, www1.
        Dim r As String, x As Byte
        r = ""
        For x = 1 To 20
            'possible throttling, but after 20 tries we may have the wrong stock code
            r = GetWeb(GetLog("StockEnt") & token & "&sym=" & CInt(sc))
            If InStr(r, "data") <> 0 Then Exit For
            Call WaitNSec(10)
            token = HKEXtoken()
        Next
        Return r
    End Function
    Sub SetAdjAll()
        'recalculate all adjustment factors
        'firs update value of warrants attached to rights issues and distributions in specie
        Call SetLinkedPrice()
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset, x As Integer
        Call OpenEnigma(con)
        rs.Open("SELECT DISTINCT issueID FROM events", con)
        Do Until rs.EOF
            Call SetAdj(CInt(rs("IssueID").Value)) 'this also calls setCumAdj
            x += 1
            Console.WriteLine(x & vbTab & rs("IssueID").Value.ToString)
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
        con = Nothing
    End Sub
    Sub SetAdj(IssueID As Integer)
        'adjust for distributions in one issue. First make sure cumDates are correct
        Call SetCumDateIssue(IssueID)
        Dim con As New ADODB.Connection, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset,
            sql As String,
            eventID, SEHKcurr, currID As Integer,
            cumDate As Date,
            adjust, Price, cumPrice, newS, oldS As Double
        Call OpenEnigma(con)
        'do splits & consols
        sql = "SELECT * FROM events WHERE eventType=4 AND NOT isNull(new) AND NOT isNull(old) AND issueID=" & IssueID & " ORDER BY exDate"
        rs1.Open(sql, con)
        Do Until rs1.EOF
            newS = CDbl(rs1("new").Value)
            oldS = CDbl(rs1("old").Value)
            adjust = oldS / newS
            con.Execute("UPDATE enigma.events SET adjust=" & adjust & " WHERE eventID=" & rs1("eventID").Value.ToString)
            Console.WriteLine("Split/consol" & vbTab & rs1("ExDate").Value.ToString & vbTab & "new:" & newS & vbTab & "old:" & oldS & vbTab & adjust)
            rs1.MoveNext()
        Loop
        rs1.Close()
        'do bonus issues and scrip-only dividends
        rs1.Open("SELECT * FROM events WHERE eventType IN(5,15) AND Not isNull(New) AND Not isNull(Old) AND issueID=" & IssueID & " ORDER BY exDate", con)
        Do Until rs1.EOF
            newS = CDbl(rs1("new").Value)
            oldS = CDbl(rs1("old").Value)
            adjust = oldS / (newS + oldS)
            con.Execute("UPDATE enigma.events SET adjust=" & adjust & " WHERE eventID=" & rs1("eventID").Value.ToString)
            Console.WriteLine("Bonus/scrip" & vbTab & rs1("ExDate").Value.ToString & vbTab & "new:" & newS & vbTab & "old:" & oldS & vbTab & adjust)
            rs1.MoveNext()
        Loop
        rs1.Close()
        'do distributions
        rs1.Open("SELECT SEHKcurr FROM issue WHERE ID1=" & IssueID, con)
        SEHKcurr = DBint(rs1("SEHKcurr")) 'Nothing=0=HKD
        rs1.Close()
        rs1.Open("SELECT * FROM events JOIN capchangetypes ON events.eventType=capchangetypes.CapChangeType" &
            " WHERE (NOT isNull(cumDate)) AND price<>0 AND isNull(cancelDate) AND dist=True AND issueID=" & IssueID &
            " ORDER BY cumDate,exDate,eventID", con)
        Do Until rs1.EOF
            eventID = CInt(rs1("eventID").Value)
            cumDate = CDate(rs1("cumDate").Value)
            currID = CInt(rs1("currID").Value)
            If currID = SEHKcurr Then
                Price = CDbl(rs1("Price").Value)
            Else
                'different distribution currency to quote currency
                'priceHKD is the distribution value in quoted currency
                Price = DBdbl(rs1("PriceHKD"))
            End If
            If Price <> 0 Then
                cumPrice = Math.Round(CDbl(con.Execute("SELECT closing FROM ccass.quotes WHERE issueID=" & IssueID &
                                                       " AND atDate='" & MSdate(cumDate) & "'").Fields(0).Value), 3) *
                                                       SuspFactor(cumDate, CDate(rs1("ExDate").Value), IssueID, eventID, True)
                adjust = 1 - Price / cumPrice
                Console.WriteLine(cumDate & vbTab & "event:" & eventID & vbTab & "cum:" & cumPrice & vbTab & "price:" & Price & vbTab & adjust)
                sql = "UPDATE enigma.events SET cumPrice=" & cumPrice & ",adjust="
                If adjust > 0 Then sql &= adjust Else sql &= "NULL"
                sql = sql & " WHERE eventID=" & eventID
                con.Execute(sql)
            End If
            rs1.MoveNext()
        Loop
        rs1.Close()
        'do rights issues and open offers
        sql = "SELECT events.eventID,events.cumDate,events.exDate,events.cumPrice,events.Price,events.priceHKD,events.new,events.old,events_1.adjust " &
            "FROM events LEFT JOIN events AS events_1 ON events.afterEvent = events_1.eventID " &
            "WHERE (Not isNull(events.cumDate)) AND events.price<>0 AND events.eventType IN(2,8) " &
            "AND events.issueID=" & IssueID & " ORDER BY cumDate,events.exDate,eventID"
        rs1.Open(sql, con)
        Do Until rs1.EOF
            eventID = CInt(rs1("eventID").Value)
            cumDate = CDate(rs1("cumDate").Value)
            newS = CDbl(rs1("New").Value)
            oldS = CDbl(rs1("Old").Value)
            Price = CDbl(rs1("Price").Value)
            If Not IsDBNull(rs1("PriceHKD").Value) Then Price = CDbl(rs1("PriceHKD").Value)
            cumPrice = Math.Round(CDbl(con.Execute("SELECT closing FROM ccass.quotes WHERE issueID=" & IssueID &
                                                   " AND atDate='" & MSdate(cumDate) & "'").Fields(0).Value), 3) *
                                                   SuspFactor(cumDate, CDate(rs1("ExDate").Value), IssueID, eventID, False)
            'now adjust subscription price for value of any bonus warrants or distributions attached to the rights
            sql = "SELECT sum(new*price/old) as rebate FROM events WHERE afterEvent=" & eventID & " AND eventType IN (45,46,51)" &
                " AND isNull(cancelDate) GROUP BY afterEvent"
            rs2.Open(sql, con)
            If Not rs2.EOF Then
                If Not IsDBNull(rs2("rebate").Value) Then Price -= CDbl(rs2("rebate").Value)
            End If
            rs2.Close()
            If IsDBNull(rs1("adjust").Value) Then adjust = 1 Else adjust = CDbl(rs1("adjust").Value)
            adjust = (newS * Price / cumPrice / adjust + oldS) / (newS + oldS)
            If adjust > 1 Then adjust = 1 'no take-up if strike is above market
            sql = "UPDATE enigma.events SET cumPrice=" & cumPrice & ",adjust="
            If adjust > 0 Then sql &= adjust Else sql &= "NULL"
            sql = sql & " WHERE eventID=" & eventID
            con.Execute(sql)
            Console.WriteLine("rights/OO" & vbTab & cumDate & vbTab & "cum:" & cumPrice & vbTab & "price:" & Price & vbTab & "new:" &
                              newS & vbTab & "old:" & oldS & vbTab & adjust)
            rs1.MoveNext()
        Loop
        rs1.Close()
        con.Close()
        con = Nothing
        Call SetCumAdj(IssueID)
    End Sub
    Sub SetLinkedPrice()
        'set the value of distributions in specie and warrants attached to rights issues/open offers, per warrant only if not set before
        'view pre/post results in findLinkedWarrPrice query
        'first value attachments
        Dim con As New ADODB.Connection, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset, Price As Double, currID As Integer
        Call OpenEnigma(con)
        rs1.Open("SELECT * FROM events WHERE eventType IN(45,46,51) AND NOT isNull(issue2) AND NOT isNull(cumDate) " &
                 "AND isNull(cancelDate) AND isNull(price)", con)
        Do Until rs1.EOF
            rs2.Open("SELECT closing from ccass.quotes where issueID=" & CInt(rs1("issue2").Value) & " AND closing<>0 AND noclose=false AND atDate>='" &
                      MSdate(CDate(rs1("cumDate").Value)) & "' ORDER BY atDate LIMIT 1", con)
            If Not rs2.EOF Then
                Price = CDbl(rs2("closing").Value)
                rs2.Close()
                rs2.Open("SELECT SEHKcurr FROM issue WHERE ID1=" & CInt(rs1("issue2").Value), con)
                currID = DBint(rs2("SEHKcurr"))
                con.Execute("UPDATE events SET price=" & Price & ",currID=" & currID & " WHERE eventID=" & CInt(rs1("eventID").Value))
            End If
            rs2.Close()
            rs1.MoveNext()
        Loop
        rs1.Close()
        'set value of distributions in specie (including warrants), on a per-share basis for the distributor
        'view pre/post results in findSpecie query
        rs1.Open("SELECT * FROM events WHERE eventType IN(18,25,50) and not isnull(issue2) AND not isnull(cumDate) " &
                  "AND isnull(cancelDate) AND isnull(price)", con)
        Do Until rs1.EOF
            rs2.Open("SELECT closing from ccass.quotes where issueID=" & CInt(rs1("issue2").Value) & " AND closing<>0 AND noclose=false AND atDate>='" &
                      MSdate(CDate(rs1("cumDate").Value)) & "' ORDER BY atDate LIMIT 1", con)
            If Not rs2.EOF Then
                Price = Math.Round(CDbl(rs2("closing").Value) * CDbl(rs1("New").Value) / CDbl(rs1("Old").Value), 8)
                rs2.Close()
                rs2.Open("SELECT SEHKcurr FROM issue WHERE ID1=" & CInt(rs1("issue2").Value), con)
                currID = DBint(rs2("SEHKcurr"))
                con.Execute("UPDATE events SET price=" & Price & ",currID=" & currID & " WHERE eventID=" & CInt(rs1("eventID").Value))
            End If
            rs2.Close()
            rs1.MoveNext()
        Loop
        rs1.Close()
        con.Close()
        con = Nothing
    End Sub
    Function SuspFactor(ByVal cumDate As Date, ByVal ExDate As Date, ByVal IssueID As Integer, ByVal eventID As Integer, ByVal dist As Boolean) As Double
        'calculate adjustment factor to cumPrice for events prior to the target event but after the cumDate
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset, sql As String
        Call OpenEnigma(con)
        'get combined adjustment for prior events (any type) with same cumDate and EARLIER exDate
        'this could happen if the stock is suspended between two or more successive exDates
        sql = "SELECT EXP(SUM(LOG(adjust))) factor FROM events" &
            " WHERE (NOT isNull(adjust)) AND isNull(cancelDate) AND issueID=" & IssueID &
            " AND cumDate='" & MSdate(cumDate) & "' AND exDate<'" & MSdate(ExDate) & "' GROUP BY issueID,cumDate"
        rs.Open(sql, con)
        If Not rs.EOF Then SuspFactor = CDbl(rs("factor").Value) Else SuspFactor = 1
        rs.Close()
        'now adjust for distributions (but not other events) with same cumDate and SAME exDate
        'if dist=False then this is a rights or open offer, so include all distributions, otherwise just those with lower eventID
        'we calculate adjustments in order of eventID, so this works in sequence
        'bonus shares, splits, rights issues are assumed not to rank for distributions with same ex-date
        'because the bonus shares or rights shares have not yet been issued
        'so we must manually adjust the distribution amount if they do
        sql = "SELECT EXP(SUM(LOG(adjust))) factor FROM events JOIN capchangetypes ON events.eventType=capchangetypes.CapChangeType" &
            " WHERE (Not isNull(adjust)) AND isNull(cancelDate) And dist=True And issueID=" & IssueID &
            " AND cumDate='" & MSdate(cumDate) & "' AND exDate='" & MSdate(ExDate) & "'"
        If dist Then sql = sql & " AND eventID<" & eventID
        sql &= " GROUP BY issueID,cumDate"
        rs.Open(sql, con)
        If Not rs.EOF Then SuspFactor *= CDbl(rs("factor").Value)
        rs.Close()
        con.Close()
        con = Nothing
    End Function
    Sub SetCumAdj(IssueID As Integer)
        'calculate cumulative adjustment factors
        'call this after adding, removing or editing events of an issue
        Console.WriteLine("Setting cumulative adjustments for issueID: " & IssueID)
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset
        Call OpenEnigma(con)
        Dim x As Integer, cumAdjust As Double, ExDate As String
        x = 0
        cumAdjust = 1
        con.Execute("DELETE FROM adjustments WHERE issueID=" & IssueID)
        rs.Open("SELECT exDate,EXP(SUM(LOG(adjust))) AS product FROM events WHERE issueID=" & IssueID &
                 " AND isNull(cancelDate) AND exDate<=CURDATE() AND Not isNull(adjust)" &
                 " GROUP BY exDate ORDER BY exDate", con)
        Do Until rs.EOF
            x += 1
            cumAdjust *= CDbl(rs("product").Value)
            ExDate = MSdate(CDate(rs("exDate").Value))
            Console.WriteLine(IssueID & vbTab & ExDate & vbTab & cumAdjust)
            con.Execute("INSERT INTO adjustments(issueID,exDate,cumAdjust) VALUES (" & IssueID & ",'" & ExDate & "'," & cumAdjust & ")")
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
        con = Nothing
        Console.WriteLine("Cumulative adjustments done for issueID: " & IssueID)
    End Sub
    Sub ParseQuotes(resp As String, atDate As Date, StockExID As Integer)
        On Error GoTo repErr
        Dim stockCode, prevClo, closing, ask, bid, high, low, vol, turn, oneline, shortName, sql As String,
            x As Integer,
            noCode, oldFormat, newsusp, susp, noclose As Boolean,
            con As New ADODB.Connection With {
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient 'for find method
            }, rs As New ADODB.Recordset
        Call OpenCCASS(con)
        If InStr(resp, "NO TRADING DUE TO") + InStr(resp, "No trading was recorded") > 0 Then 'typhoon
            con.Execute("REPLACE INTO ccass.specialdays (specialDate,noAM,noPM,noSettle) " & Valsql({atDate, True, True, True}))
            con.Close()
            con = Nothing
            Exit Sub
        End If
        Call PrepDQS(resp, atDate, noCode, oldFormat, StockExID)
        'now start on the rows
        x = 0
        rs.Open("SELECT IssueID,shortName,parallel FROM shortnames WHERE (fromDate<=" & Sqv(atDate) & " OR fromDate is Null) AND " &
                 "(toDate>" & Sqv(atDate) & " OR toDate is Null) AND shortname NOT LIKE 'EFN%' ORDER BY shortName", con)
        Do
            oneline = GetLine(resp)
            If oldFormat Then oneline = Mid(oneline, 2)
            If Left(oneline, 3) = "---" Then Exit Do 'end of quotations list
            If noCode Then
                stockCode = ""
            Else
                stockCode = Trim(Mid(oneline, 2, 5))
                'trim off the stock code
                oneline = Mid(oneline, 7)
            End If
            shortName = Trim(Mid(oneline, 2, 17))
            If InStr(shortName, "#") <> 0 Or InStr(shortName, "@") <> 0 Or InStr(shortName, "*") <> 0 Then
                'CBBC, cash-settled warrants, physical-settled warrant names (#,@,*)
                If InStr(oneline, "SUSPENDED") = 0 Then resp = MoveLines(resp, 1)
            Else
                Console.WriteLine(stockCode & " " & shortName)
                rs.MoveFirst()
                rs.Find("shortName='" & Apos(shortName) & "'")
                If rs.EOF Then
                    'stock not in shortnames
                    Console.WriteLine("Not found")
                    If InStr(oneline, "SUSPENDED") = 0 Then resp = MoveLines(resp, 1)
                Else
                    If Left(oneline, 1) = "#" Then
                        newsusp = True
                        susp = True
                        Console.Write("TODAY ")
                    Else
                        newsusp = False
                        susp = False
                    End If
                    If InStr(oneline, "SUSPENDED") <> 0 Or InStr(oneline, "TRADING HALTED") <> 0 Then
                        susp = True
                        prevClo = "0"
                        ask = "0"
                        high = "0"
                        vol = "0"
                        closing = "0"
                        noclose = True
                        bid = "0"
                        low = "0"
                        turn = "0"
                        Console.WriteLine(" SUSPENDED or HALTED")
                    Else
                        'there was a trading period today, even if it is now suspended
                        prevClo = GetNum(oneline, 23, 8)
                        ask = GetNum(oneline, 32, 8)
                        high = GetNum(oneline, 41, 8)
                        vol = GetNum(oneline, 50, 19)
                        oneline = GetLine(resp)
                        If oldFormat Then oneline = Right(oneline, Len(oneline) - 1)
                        If Not noCode Then oneline = Right(oneline, Len(oneline) - 6)
                        closing = GetNum(oneline, 23, 8)
                        noclose = (CDbl(closing) = 0) 'very unlikely
                        bid = GetNum(oneline, 32, 8)
                        low = GetNum(oneline, 41, 8)
                        turn = GetNum(oneline, 50, 19)
                        Console.WriteLine(prevClo & vbTab & closing & vbTab & ask & vbTab & bid & vbTab & high & vbTab & low & vbTab & vol & vbTab & turn)
                    End If
                    sql = "REPLACE INTO "
                    If IsDBNull(rs("issueID").Value) Then
                        sql &= "unquotes (stockCode,atDate,prevClose,closing,ask,bid,high,low,vol,turn,susp,newsusp,noclose)" &
                            Valsql({stockCode, atDate, prevClo, closing, ask, bid, high, low, vol, turn, susp, newsusp, noclose})
                    Else
                        If CBool(rs("parallel").Value) Then
                            sql &= "pquotes"
                        Else
                            sql &= "quotes"
                        End If
                        sql &= " (issueID,atDate,prevClose,closing,ask,bid,high,low,vol,turn,susp,newsusp,noclose)" &
                            Valsql({CInt(rs("IssueID").Value), atDate, prevClo, closing, ask, bid, high, low, vol, turn, susp, newsusp, noclose})
                    End If
                    con.Execute(sql)
                    x += 1
                End If
            End If
        Loop
        rs.Close()
        con.Close()
        con = Nothing
        Console.WriteLine(MSdate(atDate) & vbTab & x & " stock quotes updated")
        Exit Sub
repErr:
        Call ErrMail("ParseQuotes failed", Err, oneline)
    End Sub
    Function GetNum(ByVal str As String, ByVal start As Integer, ByVal length As Integer) As String
        'return a mysql-friendly number string
        str = Trim(Mid(str, start, length))
        str = Replace(str, ",", "")
        If str = "-" Or str = "N/A" Then str = "0"
        Return str
    End Function
    Sub ParseShortNames(ByVal resp As String, atDate As Date, StockExID As Integer)
        On Error GoTo repErr
        Dim dateStr, oneline, shortName, stockCode, suffix, prefix As String,
            issueID, x As Integer,
            oldFormat, noCode, parallel As Boolean,
            con As New ADODB.Connection, rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
        Call OpenCCASS(con)
        dateStr = MSdate(atDate)
        If InStr(resp, "NO TRADING DUE TO") + InStr(resp, "No trading was recorded") > 0 Then 'typhoon
            con.Execute("REPLACE INTO specialdays (specialDate,noAM,noPM,noSettle) " & Valsql({atDate, True, True, True}))
            con.Close()
            con = Nothing
            Exit Sub
        End If
        Call PrepDQS(resp, atDate, noCode, oldFormat, StockExID)
        Console.WriteLine("DQS prepared")
        'now start on the rows
        x = 0
        Do
            oneline = GetLine(resp)
            If oldFormat Then oneline = Right(oneline, Len(oneline) - 1)
            If Left(oneline, 3) = "---" Then Exit Do 'end of quotations list
            parallel = False
            If noCode Then
                shortName = Trim(Mid(oneline, 2, 20))
                stockCode = ""
            Else
                stockCode = Trim(Mid(oneline, 2, 5))
                shortName = Trim(Mid(oneline, 8, 17))
            End If
            Console.WriteLine(stockCode & vbTab & shortName)
            If InStr(shortName, "#") + InStr(shortName, "@") + InStr(shortName, "*") = 0 Then
                'skip CBBC, cash-settled warrants, physical-settled warrant names (#,@,*)
                'look for name in current names
                rs.Open("SELECT  * FROM shortnames USE INDEX (nameindex) WHERE isNull(toDate) AND shortName='" & Apos(shortName) & "' AND stockExID=" & StockExID, con)
                If rs.EOF Then
                    'this is a new current name on this board
                    If Right(shortName, 4) = "-OLD" Then parallel = True
                    issueID = GetIssueID(stockCode, shortName, dateStr, StockExID)
                    If stockCode = "" Then
                        con.Execute("INSERT INTO shortnames (shortName,fromDate,useDate,stockExID,parallel) " &
                            Valsql({shortName, atDate, atDate, StockExID, parallel}))
                    Else
                        If issueID = 0 Then
                            con.Execute("INSERT INTO shortnames (stockCode,shortName,fromDate,useDate,stockExID,parallel) " &
                                Valsql({stockCode, shortName, atDate, atDate, StockExID, parallel}))
                        Else
                            con.Execute("INSERT INTO shortnames (issueID,stockCode,shortName,fromDate,useDate,stockExID,parallel) " &
                                Valsql({issueID, stockCode, shortName, atDate, atDate, StockExID, parallel}))
                        End If
                    End If
                    Console.WriteLine("New name" & vbTab & shortName & vbTab & stockCode & vbTab & issueID)
                Else
                    'this is an existing name
                    If noCode Then
                        con.Execute("UPDATE shortnames" & Setsql("useDate", {atDate}) & "ID=" & rs("ID").Value.ToString)
                    Else
                        If IsDBNull(rs("stockCode").Value) Then
                            'but now we know the stock code (must be 27-Nov-06 or later) and set useDate while we are at it
                            con.Execute("UPDATE shortnames" & Setsql("stockCode,useDate", {stockCode, atDate}) & "ID=" & rs("ID").Value.ToString)
                            'and can check the issueID
                            issueID = GetIssueID(stockCode, shortName, dateStr, StockExID)
                            If issueID <> 0 Then con.Execute("UPDATE shortnames" & Setsql("issueID", {issueID}) & "ID=" & rs("ID").Value.ToString)
                        ElseIf CInt(stockCode) = CInt(rs("stockCode").Value) Then
                            'same stock code (not null) and name
                            con.Execute("UPDATE shortnames" & Setsql("useDate", {atDate}) & "ID=" & rs("ID").Value.ToString)
                        Else
                            'old Name with a new stock code - e.g. a temporary counter or a move to main board
                            issueID = CInt(rs("issueID").Value)
                            con.Execute("INSERT INTO shortnames (stockCode,shortName,issueID,fromDate,useDate,stockExID) " &
                                Valsql({stockCode, shortName, issueID, atDate, atDate, StockExID}))
                            Console.WriteLine("New code" & vbTab & shortName & vbTab & stockCode & vbTab & issueID)
                        End If
                    End If
                End If
                rs.Close()
            End If
            If InStr(oneline, "SUSPENDED") = 0 And InStr(oneline, "TRADING HALTED") = 0 Then resp = MoveLines(resp, 1) 'skip next line
            x += 1
        Loop
        'for any rights issues which have disappeared, set delisting date in stocklistings if not already there
        'cannot do this for normal issues as they would appear delisted whenever they go parallel!
        rs.Open("SELECT * FROM shortnames WHERE Right(shortname,4)=' RTS' AND StockExID=" & StockExID & " AND isNull(toDate) And useDate<'" & dateStr & "'", con)
        Do Until rs.EOF
            If Not IsDBNull(rs("issueID").Value) Then con.Execute("UPDATE enigma.stocklistings SET DelistDate='" & dateStr & "',FinalTradeDate='" &
                                                                  MSdate(CDate(rs("useDate").Value)) & "' WHERE issueID=" & rs("issueID").Value.ToString &
                                                                  " AND isNull(DelistDate) AND StockExID=" & StockExID)
            rs.MoveNext()
        Loop
        rs.Close()
        'having run through the day, find new parallel counters based on the prefix-suffix combo, taking the issueID from the normal stock code
        rs.Open("SELECT * FROM shortNames WHERE isNull(issueID) AND (shortName NOT LIKE 'EFN %') AND (shortName LIKE '%-%') AND fromDate='" & dateStr & "'",
                con, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        Do Until rs.EOF
            shortName = rs("shortName").Value.ToString
            x = InStrRev(shortName, "-")
            If x > 0 Then
                suffix = Mid(shortName, x + 1)
                If suffix Like "*#K" Or IsNumeric(suffix) Or suffix = "OLD" Then
                    prefix = Left(shortName, x)
                    rs2.Open("SELECT * FROM shortNames WHERE Not isNull(issueID) AND parallel=False AND fromDate='" & dateStr & "' AND shortName LIKE '" & Apos(prefix) & "%'", con)
                    If Not rs2.EOF Then
                        rs("parallel").Value = True
                        rs("issueID").Value = rs2("issueID").Value
                        rs.Update()
                    End If
                    rs2.Close()
                End If
            End If
            rs.MoveNext()
        Loop
        rs.Close()
        'any name not found today is not in use today. So we must run this procedure in sequential dates
        con.Execute("UPDATE shortNames SET toDate='" & dateStr & "' WHERE isNull(toDate) AND useDate<'" & dateStr & "' AND stockExID=" & StockExID)
        con.Close()
        con = Nothing
        Exit Sub
repErr:
        ErrMail("ParseShortNames failed", Err, oneline)
    End Sub
    Function GetLine(ByRef str As String) As String
        Dim endline As Integer
        endline = InStr(str, Chr(13))
        GetLine = Left(str, endline - 1)
        str = Right(str, Len(str) - endline - 1)
    End Function
    Function MoveLines(ByVal str As String, ByVal lines As Integer) As String
        For lines = 1 To lines
            str = Right(str, Len(str) - InStr(str, Chr(13)) - 1)
        Next
        Return str
    End Function
    Sub PrepDQS(ByRef resp As String, ByVal atDate As Date, ByRef noCode As Boolean, ByRef oldFormat As Boolean, ByVal StockExID As Integer)
        'prepare a raw DQS contents for analysis - called by parseQuotes and parseShortNames
        'returning noCode, oldFormat, resp
        Dim endline As Integer
        noCode = (atDate < #11/27/2006#) 'no stock codes before that
        If Not noCode Then
            oldFormat = False
            endline = InStr(resp, "<a name = ""quotations"">") 'look for keywords 5 lines above start
            resp = Right(resp, Len(resp) - endline) 'cut the top off
            endline = InStr(resp, "---")
            resp = Left(resp, endline + 100) 'cut the bottom off, to speed things up. Allow for a carriage return which we need.
            'older files were 75 columns, now they are 74
            resp = MoveLines(resp, 2) 'move to column headings
            If InStr(resp, "Name") = 9 Then oldFormat = True
            'strip out <pre> tags which confuse parsing. Old format has leading space
            If oldFormat Then
                resp = Replace(resp, " <pre><font size=""1"">", "")
            Else
                resp = Replace(resp, "<pre><font size=""1"">", "")
            End If
            If atDate < #4/4/2011# Then
                resp = Replace(resp, "</pre>", "")
            Else
                'on 4-Apr-2011 a new form of garbage appeared:
                resp = Replace(resp, "</font></pre><pre><font size='1'>", "")
                'and the HTML started using &amp; instead of &, shifting all the columns
                resp = Replace(resp, "&amp;", "&")
                'and sometimes (notably in front of Goldbond, 0172) a double return was inserted, so replace it with a single one
                resp = Replace(resp, Chr(13) & Chr(13), Chr(13))
            End If
        Else
            'no stock codes, line width is 69/68 except for suspensions
            oldFormat = True
            If StockExID = 20 Then oldFormat = False
            endline = InStr(resp, "PRV.CLO") 'look for keywords 3 lines above start
            resp = Right(resp, Len(resp) - endline) 'cut the top off
            endline = InStr(resp, "---")
            resp = Left(resp, endline + 100)
        End If
        resp = MoveLines(resp, 3)
    End Sub
    Function GetIssueID(ByVal stockCode As String, ByVal shortName As String, ByVal dStr As String, ByVal StockExID As Integer) As Integer
        'try to find issueID corresponding to a stock code or short name
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset, root, suffix As String,
            issuer, SEHKcurr, i As Integer,
            acceptDate As Date
        Call OpenCCASS(con)
        suffix = Right(shortName, 4)
        i = 0
        If stockCode <> "" Then
            Console.WriteLine(stockCode)
            rs.Open("SELECT enigma.getIssueID(" & stockCode & ",'" & dStr & "') issueID", con)
            If Not rs.EOF And Not IsDBNull(rs("issueID").Value) Then
                'A change of shortname for existing stock
                i = CInt(rs("issueID").Value)
            Else
                'could be a parallel counter (not in stocklistings) so try searching yesterday's valid codes
                rs.Close()
                rs.Open("SELECT issueID FROM shortnames WHERE isNull(toDate) and Not isNull(issueID) and stockCode=" & stockCode, con)
                If Not rs.EOF Then
                    i = CInt(rs("issueID").Value)
                Else
                    'either no stockcode or new stockcode. Now try the name
                    If suffix = "-OLD" Or suffix = "-NEW" Or suffix = " RTS" Then
                        root = Left(shortName, Len(shortName) - 4)
                    Else
                        root = shortName
                    End If
                    rs.Close()
                    rs.Open("SELECT issueID FROM shortnames WHERE isNull(toDate) AND shortName='" & Apos(root) & "'", con)
                    If Not rs.EOF Then
                        i = CInt(rs("issueID").Value)
                        If suffix = " RTS" Then
                            'create a rights issue for a known issuer and return its issueID
                            rs.Close()
                            rs.Open("SELECT * FROM enigma.issue WHERE ID1=" & i, con)
                            issuer = CInt(rs("issuer").Value)
                            SEHKcurr = DBint(rs("SEHKcurr"))
                            rs.Close()
                            'find the rights issue event
                            rs.Open("SELECT * FROM enigma.events WHERE eventType=2 AND issueID=" & i &
                                    " AND acceptDate>'" & dStr & "' LIMIT 1", con)
                            If rs.EOF Then
                                acceptDate = Nothing
                            Else
                                acceptDate = DBdate(rs("acceptDate"))
                            End If
                            If acceptDate = Nothing Then
                                con.Execute("INSERT INTO enigma.issue (issuer,typeID,SEHKcurr) " & Valsql({issuer, 2, SEHKcurr}))
                            Else
                                con.Execute("INSERT INTO enigma.issue (issuer,typeID,SEHKcurr,expmat)" & Valsql({issuer, 2, SEHKcurr, acceptDate}))
                            End If
                            i = LastID(con)
                            If stockCode <> "" Then
                                con.Execute("INSERT INTO enigma.stocklistings (IssueID,StockExID,StockCode,FirstTradeDate)" &
                                            Valsql({i, StockExID, stockCode, dStr}))
                            End If
                            Console.WriteLine("New rights issue" & vbTab & stockCode & vbTab & i & vbTab & acceptDate)
                        End If
                    End If
                End If
            End If
            rs.Close()
        End If
        con.Close()
        con = Nothing
        Return i
    End Function
    Sub GetShorts()
        'fetch the latest short position files, if any
        Dim con As New ADODB.Connection,
            lastShort As Date,
            archive, e, r, dest, lastFile, row, td, URL As String,
            x, rPos As Integer
        archive = GetLog("ShortsFolder")
        Call OpenEnigma(con)
        lastShort = CDate(con.Execute("SELECT max(atDate) FROM sfcshort").Fields(0).Value)
        lastFile = Format(lastShort, "yyyyMMdd") & ".pdf"
        Console.WriteLine(lastFile)
        r = GetWeb("https://www.sfc.hk/en/Regulatory-functions/Market/Short-position-reporting/Aggregated-reportable-short-positions-of-specified-shares")
        'extract the table
        Call TagContID(1, r, "spr_table", r)
        'extract the tbody
        Call TagCont(1, r, "tbody", r)
        x = 1
        'discard the header
        row = ""
        Call TagCont(x, r, "tr", row)
        Do
            row = ""
            td = ""
            e = ""
            Call TagCont(x, r, "tr", row)
            rPos = InStr(row, "<td") + 3 'skip date cell
            'fetch PDF
            Call TagCont(rPos, row, "td", td)
            URL = GetAttrib(td, "href")
            URL = Left(URL, InStr(URL, "?") - 1) 'added 2022-10-26 when SFC appended revision querystring
            Console.WriteLine(URL)
            If Right(URL, 12) = lastFile Then
                Console.WriteLine("No more files")
                Exit Do
            End If
            'fetch csv
            Call TagCont(rPos, row, "td", td)
            URL = GetAttrib(td, "href")
            URL = Left(URL, InStr(URL, "?") - 1)
            dest = archive & Right(URL, 12)
            Call Download(URL, dest, e)
            If e = "" Then Call ProcShorts(dest)
        Loop
        con.Close()
        con = Nothing
    End Sub
    Sub ProcShorts(f As String)
        'process a single CSV file of shorts at location f
        Dim d, s, shares, val, sc, sName, shorts(), c() As String,
            y As Integer,
            con As New ADODB.Connection
        Call OpenEnigma(con)
        s = My.Computer.FileSystem.ReadAllText(f)
        'strip off the last newline before split
        If Right(s, 1) = Chr(10) Then s = Left(s, Len(s) - 1)
        shorts = Split(s, Chr(10)) 'lines are only terminated by newline so cannot use "Line Input.."
        'file has a header in row 0
        For y = 1 To UBound(shorts)
            c = Split(shorts(y), ",")
            d = MSdateDMY(c(0))
            sc = c(1)
            sName = c(2)
            shares = c(3)
            val = c(4)
            If val = "n.a." Then val = "0"
            Console.WriteLine(d & vbTab & sc & vbTab & shares & vbTab & val)
            con.Execute("INSERT IGNORE INTO sfcshort (issueID,atDate,stockCode,stockName,shares,value) VALUES(getIssueID(" &
                sc & ",'" & d & "'),'" & d & "'," & sc & ",'" & Apos(sName) & "'," & shares & "," & val & ")")
        Next
        con.Close()
        con = Nothing
    End Sub
End Module
