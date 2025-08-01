Option Compare Text
Option Explicit On
Imports JSONkit
Imports ScraperKit
Imports persons
Imports System.IO 'for streamreader in MapCRBR

Module CR
    Sub Main()
        'NB WHEN TESTING CHOOSE CORRECT ProcFileL or ProcFileF!
        'Call ProcFileF("\\webbserver2\c\home\David\db\CR\F20231225.csv")
        Call GetChangesLF("L")
        Call GetChangesLF("F")
        Call GetMissingDomF()
        Call GetMissingTypeL()
    End Sub
    Sub TestRow()
        Dim r() As String, x As Integer
        r = ReadCSVrow(",GDS WanGuo Holdings Limited,,F0028528,,11/11/2020")
        For x = 0 To UBound(r)
            Console.WriteLine(x & vbTab & r(x))
        Next
        Console.ReadKey()
    End Sub
    Sub CRBR()
        On Error GoTo repErr
        '1-off conversion work following 2023-12-27 launch of BR number at CR
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset,
           r(), s, crn As String
        Call OpenEnigma(con)
        Console.WriteLine("Reading file")
        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(GetLog("CRfolder") & "MapCRBRnonHK.csv")
        s = reader.ReadLine 'skip the header row
        Do
            s = reader.ReadLine
            If s = "" Then Exit Do
            r = ReadCSVrow(s)
            crn = r(0)
            rs.Open("SELECT fregID FROM oldcrf WHERE crn='" & crn & "'", con)
        If rs.EOF Then
                Console.WriteLine("Not found CRno: " & crn)
                'con.Execute("INSERT INTO missingcr (CR,BRN,eName,cName) VALUES('" & crn & "','" & r(1) & "'," & Sqv(r(2)) & "," & Sqv(r(3)) & ")")
                con.Execute("INSERT INTO missingcr (CR,BRN) VALUES('" & r(0) & "','" & r(1) & "')")
            Else
                con.Execute("UPDATE oldcrf SET brn='" & r(1) & "' WHERE fregID=" & DBint(rs("fregID")))
                Console.WriteLine(crn & vbTab & r(1) & vbTab & r(2))
            End If
            rs.Close()
        Loop
        reader.Dispose()
        con.Close()
        con = Nothing
        Exit Sub
repErr:
        Call ErrMail("CRBR failed at or after CRno:" & crn, Err)
        Console.WriteLine(Err)
        Console.ReadKey()
    End Sub

    Sub SetNewID()
        'this was an attempt on 2023-12-14 to give freg a new ID, but the slave would not replicate it because table was out of sync
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset, x As Integer = 0, hostDom As Integer, regID As String
        Call OpenEnigma(con)
        rs.Open("SELECT * FROM freg ORDER BY hostDom,regID", con)
        Do Until rs.EOF
            hostDom = DBint(rs("hostDom"))
            regID = rs("regID").Value.ToString
            x += 1
            con.Execute("UPDATE freg SET ID2=" & x & " WHERE hostDom=" & hostDom & " AND regID='" & Apos(regID) & "'")
            Console.WriteLine(x & vbTab & vbTab & hostDom & vbTab & regID)
            Console.ReadKey()
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
    End Sub
    Sub GetMissingDomF()
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset
        Call OpenEnigma(con)
        rs.Open("SELECT regID FROM organisations o JOIN freg f ON personID=orgID WHERE isNUll(domicile) AND regDate>'1994-11-07' AND hostdom=1 order by regDate", con)
        Do Until rs.EOF
            Call GetDomF(rs("regID").Value.ToString)
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
        con = Nothing
    End Sub
    Sub GetDomF(regID As String, Optional ByRef found As Boolean = False)
        'get the domicile of a foreign company using API
        Dim o, incPlace As String,
            domicile, orgID As Integer,
            con As New ADODB.Connection
        found = False
        Call OpenEnigma(con)
        orgID = DBint(con.Execute("SELECT orgID FROM freg WHERE regID='" & regID & "'").Fields(0))
        o = GetWeb("https://data.cr.gov.hk/cr/api/api/v1/api_builder/json/foreign/search?query[0][key1]=Brn&query[0][key2]=equal&query[0][key3]=" & regID)
        If o = "[]" Or o = "" Then
            Console.WriteLine(regID & vbTab & "No response")
        Else
            'problem with 2 different variables in registry. One is ctryEnName, the other is fullCtryEnName. In Turks & Caicos, one uses " AND " rather than " & "
            incPlace = Replace(GetVal(o, "Place_of_Incorporation"), " AND ", " & ")
            domicile = CInt(con.Execute("SELECT IFNULL((SELECT ID FROM domiciles WHERE HKCR='" & incPlace & "'),0)").Fields(0).Value)
            If domicile = 0 Then
                Console.WriteLine(regID & vbTab & "Domicile not found: " & incPlace)
                Call SendMail("HKCR domicile not found:" & incPlace)
            Else
                found = True
                con.Execute("UPDATE organisations" & Setsql("domicile", {domicile}) & "personID=" & orgID)
                Console.WriteLine(regID & vbTab & "Domicile: " & domicile & vbTab & incPlace)
            End If
        End If
    End Sub
    Sub GetMissingTypeL()
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset
        Call OpenEnigma(con)
        rs.Open("SELECT incID FROM organisations WHERE domicile=1 AND isNUll(orgType) AND NOT isNull(incID) ORDER BY incDate DESC", con)
        Do Until rs.EOF
            Call GetTypeL(rs("incID").Value.ToString)
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
        con = Nothing
    End Sub
    Sub GetTypeL(incID As String, Optional ByRef found As Boolean = False)
        'get the HK company type from API
        Dim o, typeStr As String,
            orgID, orgType As Integer,
            con As New ADODB.Connection
        found = False
        Call OpenEnigma(con)
        orgID = DBint(con.Execute("SELECT personID FROM organisations WHERE domicile=1 AND incID='" & incID & "'").Fields(0))
        o = GetWeb("https://data.cr.gov.hk/cr/api/api/v1/api_builder/json/local/search?query[0][key1]=Brn&query[0][key2]=equal&query[0][key3]=" & incID)
        If o = "[]" Or o = "" Then
            Console.WriteLine(incID & vbTab & "No response")
        Else
            typeStr = GetVal(o, "Company_Type")
            orgType = CInt(con.Execute("SELECT IFNULL((SELECT orgType FROM orgtypes WHERE HKCRname='" & typeStr & "'),0)").Fields(0).Value)
            If orgType = 0 Then
                Console.WriteLine(incID & vbTab & "OrgType not found: " & typeStr)
                Call SendMail("HKCR orgType not found:" & incID & " " & typeStr)
            Else
                found = True
                con.Execute("UPDATE organisations" & Setsql("orgType", {orgType}) & "personID=" & orgID)
                Console.WriteLine(incID & vbTab & "orgType: " & orgType & vbTab & typeStr)
            End If
        End If
        con.Close()
        con = Nothing
    End Sub
    Sub FixData()
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset, found As Boolean
        Call OpenEnigma(con)
        'check missing orgtypes
        rs.Open("SELECT incID FROM organisations WHERE domicile=1 AND isnull(orgtype) AND incID RLIKE '^[0-9]' order by inciD", con)
        Do Until rs.EOF
            Call CRAPI(rs("incID").Value.ToString, found)
            rs.MoveNext()
        Loop
        rs.Close()
        Exit Sub
        'check missing incDates
        rs.Open("SELECT incID FROM organisations WHERE domicile=1 AND isNull(incDate) AND incID RLIKE '^[0-9]' order by inciD", con)
        Do Until rs.EOF
            Call CRAPI(rs("incID").Value.ToString, found)
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
        con = Nothing
    End Sub
    Sub CRAPI(incID As String, ByRef found As Boolean)
        Dim URLbase, o, crType, en, cn As String,
            incDate As Date,
            orgType, p As Integer,
            con As New ADODB.Connection, rs As New ADODB.Recordset
        Call OpenEnigma(con)
        incID = Right("000000" & incID, 7)
        URLbase = "https://data.cr.gov.hk/cr/api/api/v1/api_builder/json/local/search?query[0][key1]=Cr_no&query[0][key2]=equal&query[0][key3]="
        o = GetWeb(URLbase & incID)
        If o <> "[]" Then
            incDate = ReadDMY(GetVal(o, "Date_of_Incorporation"))
            'sometimes API returns first live company which at time of coding was HK & Shanghai Hotels incorporated in 1866
            found = (Year(incDate) > 1866)
            If found Then
                crType = GetVal(o, "Company_Type")
                orgType = CInt(con.Execute("SELECT IFNULL((SELECT orgType From orgtypes WHERE HKCRname='" & crType & "'),0)").Fields(0).Value)
                If orgType = 0 Then
                    Console.WriteLine(incID & rs("name1").Value.ToString)
                    Console.WriteLine("Type not found: " & crType)
                    Console.ReadKey()
                End If
                rs.Open("SELECT * FROM organisations WHERE domicile=1 AND incID='" & incID & "'", con)
                If rs.EOF Then
                    en = GetVal(o, "English_Company_Name")
                    cn = GetVal(o, "Chinese_Company_Name")
                    If en = "NULL" Then en = "HKCR:" & incID & " (no English name)" Else en = ConfName(en)
                    If cn = "NULL" Then cn = ""
                    Call NameResOrg(1, en, CDate(incDate), Nothing, 1, incID)
                    con.Execute("INSERT INTO persons() VALUES()")
                    p = LastID(con)
                    con.Execute("INSERT INTO organisations(personID,orgType,domicile,incID,incDate,name1,cName,incupd)" & Valsql({p, orgType, 1, incID, incDate, en, cn, MSdateTime(Now)}))
                Else
                    p = CInt(rs("personID").Value)
                    en = rs("name1").Value.ToString
                    cn = rs("cName").Value.ToString
                    If IsDBNull(rs("incDate").Value) Then con.Execute("UPDATE organisations" & Setsql("incDate", {incDate}) & "personID=" & p)
                    con.Execute("UPDATE organisations SET orgType=" & orgType & " WHERE personID=" & p)
                End If
                rs.Close()
                Console.WriteLine(incID & vbTab & incDate & vbTab & orgType & vbTab & en & " " & cn)
            End If
        Else
            found = False
            Console.WriteLine(incID & vbTab & "No response")
        End If
        con.Close()
        con = Nothing
    End Sub
    Sub CRAPIF(regID As String, ByRef found As Boolean)
        Dim URLbase, o, crType, en, cn, name, incPlace, dStr As String,
        mostLatin, score As Double,
        orgType, p, domicile As Integer,
        con As New ADODB.Connection, rs As New ADODB.Recordset
        Call OpenEnigma(con)
        URLbase = "https://data.cr.gov.hk/cr/api/api/v1/api_builder/json/foreign/search?query[0][key1]=Cr_no&query[0][key2]=equal&query[0][key3]="
        o = GetWeb(URLbase & regID)
        If o <> "[]" Then
            found = True
            dStr = GetVal(o, "Date_of_Registration")
            If dStr = "null" Then dStr = "" Else dStr = MSdateDMY(dStr)
            crType = GetVal(o, "Company_Type")
            incPlace = GetVal(o, "Place_of_Incorporation")
            orgType = CInt(con.Execute("SELECT IFNULL((SELECT orgType FROM orgtypes WHERE HKCRname='" & crType & "'),0)").Fields(0).Value)
            If orgType = 0 Then
                Console.WriteLine(regID & vbTab & "Type not found: " & crType)
                Call SendMail("OrgType in HKCR non-HK not found for regID:" & regID & " and orgType:" & crType)
                found = False
                con.Close()
                con = Nothing
                Exit Sub
            End If
            domicile = CInt(con.Execute("SELECT IFNULL((SELECT ID FROM domiciles WHERE HKCR='" & incPlace & "'),0)").Fields(0).Value)
            If domicile = 0 Then
                Console.WriteLine(regID & vbTab & "Domicile not found: " & incPlace)
                Call SendMail("HKCR domicile not found:" & incPlace)
                found = False
                con.Close()
                Exit Sub
            End If
            rs.Open("SELECT * FROM organisations JOIN freg ON personID=orgID WHERE hostDom=1 AND regID='" & regID & "'", con)
            If rs.EOF Then
                en = ""
                cn = ""
                'foreign companies can have multiple corporate names
                'English and Chinese names can appear in any order, so we have to score them to figure out which is which!
                name = GetVal(o, "Corporate_Name")
                mostLatin = LatinScore(name)
                'provisionally determine first name is either English or Chinese
                If mostLatin > 0.5 Then en = name Else cn = name
                name = GetVal(o, "Other_Corporate_Name_s")
                If name <> "null" Then
                    score = LatinScore(name)
                    If score > mostLatin And score > 0.5 Then
                        en = name
                    ElseIf score < mostLatin And score <= 0.5 Then
                        cn = name
                    End If
                End If
                If en = "" Then en = "HKCR:" & regID & " (no English name)" Else en = ConfName(en)
                Call NameResOrg(1, en, Nothing, Nothing, domicile, regID)
                'we sent the regID not the incID, so adjust the naming if this was used
                en = Replace(en, "(incorp. ID:", "(HKCR:")
                con.Execute("INSERT INTO persons() VALUES ()")
                p = LastID(con)
                con.Execute("INSERT INTO organisations (orgType,personID,domicile,Name1,cName) " & Valsql({orgType, p, domicile, en, cn}))
                con.Execute("INSERT INTO freg(hostDom,orgID,regID,regDate,regUpd)" & Valsql({1, p, regID, dStr, MSdateTime(Now)}))
            Else
                p = CInt(rs("personID").Value)
                en = rs("name1").Value.ToString
                cn = rs("cName").Value.ToString
                If IsDBNull(rs("regDate").Value) Then con.Execute("UPDATE freg" & Setsql("regDate", {dStr}) & "orgID=" & p)
                If IsDBNull(rs("domicile").Value) Then con.Execute("UPDATE organisations SET domicile=" & domicile & " WHERE personID=" & p)
            End If
            rs.Close()
            Console.WriteLine(regID & " " & dStr & " " & domicile & vbTab & en & " " & cn)
        Else
            found = False
            Console.WriteLine(regID & vbTab & "No response")
        End If
        con.Close()
        con = Nothing
    End Sub
    Sub GetOrgTypes()
        'for new local companies without an orgType
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset,
            incID As String, found As Boolean
        Call OpenEnigma(con)
        rs.Open("SELECT * FROM organisations WHERE domicile=1 AND incID rlike '^[0-9]' AND isNull(orgType) ORDER BY incDate,incID", con)
        Do Until rs.EOF
            incID = rs("incID").Value.ToString
            Call CRAPI(incID, found)
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
        con = Nothing
    End Sub
    Sub GetDomiciles()
        'for newly-registered non-HK companies without a domicile
        Dim con As New ADODB.Connection, rs As New ADODB.Recordset,
            regID As String, found As Boolean
        Call OpenEnigma(con)
        rs.Open("SELECT * FROM organisations JOIN freg ON personID=orgID WHERE isNull(domicile) AND hostDom=1 AND regDate>'2020-11-22' ORDER BY regDate,regID", con)
        Do Until rs.EOF
            regID = rs("regID").Value.ToString
            Call CRAPIF(regID, found)
            rs.MoveNext()
        Loop
        rs.Close()
        con.Close()
        con = Nothing
    End Sub

    Sub GetNewLocos()
        'use the CR API to get new companies - this only serves numeric company numbers, not LF etc
        Dim lastNum As Integer, found As Boolean
        lastNum = CInt(GetLog("lastHKCRloco"))
        found = True
        Do Until Not found
            lastNum += 1
            'Console.WriteLine("Seeking company number " & lastNum)
            Call CRAPI(lastNum.ToString, found)
            If found Then PutLog("lastHKCRloco", lastNum.ToString)
        Loop
    End Sub
    Sub GetNewFocos()
        'use the CR API to get new non-HK companies F1234567
        Dim lastNum As Integer, found = True
        lastNum = CInt(GetLog("lastHKCRfoco"))
        Do Until Not found
            lastNum += 1
            'Console.WriteLine("Seeking company number " & lastNum)
            Call CRAPIF("F" & Right("00" & lastNum.ToString, 7), found)
            If found Then PutLog("lastHKCRfoco", lastNum.ToString)
        Loop
    End Sub
    Sub GetCRfiles()
        'One-time download all files, no overwrite
        Dim dest, e, CRfolder, URL As String, d As Date
        d = #2019-09-16#
        CRfolder = GetLog("CRfolder")
        Do Until d > Today
            e = ""
            URL = "http://www.cr.gov.hk/docs/wrpt/RNC063/wk_new&changednamecoys_" & d.ToString("yyyyMMdd") & ".csv"
            dest = CRfolder & Right(URL, 12)
            Call Download(URL, dest, e, False)
            d = d.AddDays(1)
        Loop
    End Sub
    Sub GetChanges()
        'Old version for files before 2021-12-06
        On Error GoTo RepErr
        Dim CRfolder, URL, dest, e As String,
            d As Date
        CRfolder = GetLog("CRfolder")
        d = CDate(GetLog("LastCRchanges"))
        Do Until d = Today
            e = ""
            d = d.AddDays(1)
            URL = "http://www.cr.gov.hk/docs/wrpt/RNC063/wk_new&changednamecoys_" & d.ToString("yyyyMMdd") & ".csv"
            dest = CRfolder & Right(URL, 12)
            Call Download(URL, dest, e, False, False)
            If e = "" Then
                Console.WriteLine("Found it")
                Call ProcFile(dest)
                Call PutLog("LastCRchanges", MSdate(d))
            End If
            Console.WriteLine(URL)
        Loop
        Exit Sub
RepErr:
        Call ErrMail("CR GetChanges failed", Err, "File date: " & MSdate(d))
    End Sub
    Sub GetChangesLF(LF As String)
        'New version from 2021-12-06 LF is either "L" for Local or "F" for Foreign
        On Error GoTo RepErr
        Dim CRfolder, URL, dest, e As String,
            d As Date
        CRfolder = GetLog("CRfolder")
        d = CDate(GetLog("LastCRchanges" & LF))
        Do Until d = Today
            e = ""
            d = d.AddDays(1)
            URL = "http://www.cr.gov.hk/docs/wrpt/RNC063/RNC063" & LF & "_" & d.ToString("yyyyMMdd") & ".csv"
            dest = CRfolder & LF & Right(URL, 12)
            Call Download(URL, dest, e, False, False)
            If e = "" Then
                Console.WriteLine("Found it")
                If LF = "F" Then Call ProcFileF(dest) Else Call ProcFileL(dest)
                Call PutLog("LastCRchanges" & LF, MSdate(d))
            End If
            Console.WriteLine(URL)
        Loop
        Exit Sub
RepErr:
        Call ErrMail("CR GetChanges failed", Err, "File: " & LF & MSdate(d))
    End Sub


    Sub ProcDateFile(d As Date)
        'can use this to process a file out of sequence - the file for 2021-03-29 was not downloaded properly
        Dim dest, CRfolder As String
        CRfolder = GetLog("CRfolder")
        dest = CRfolder & d.ToString("yyyyMMdd") & ".csv"
        Call ProcFile(dest)
    End Sub
    Sub ProcFileL(dest As String)
        'New from files dated 2021-12-06
        'process the CR CSV file for Local companies at dest
        On Error GoTo RepErr
        Dim s, c(), r(), name, ensql, oldName, cName, cnsql, oldCname, incID, incdStr, chgdStr, orgType As String,
            x, p, tempID As Integer,
            incDate As Date,
            con As New ADODB.Connection, rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
        Call OpenEnigma(con)
        s = My.Computer.FileSystem.ReadAllText(dest)
        s = RemCSVbreaks(s) 'some files have erroneous carriage returns inside names, e.g. 20230626 local companies
        c = SplitCSVrows(s)
        'ignore c(0) header row
        For x = 1 To UBound(c)
            r = ReadCSVrow(c(x))
            name = ConfName(CleanStr(r(1)))
            cName = CleanStr(r(2))
            incID = r(3)
            incdStr = MSdateDMY(r(4))
            chgdStr = MSdateDMY(r(5))
            If name = "" Then name = "HKCR:" & incID & " (no English name)"
            If r(4) > "" Then 'newCo
                'Only do this if we haven't already added it
                If CBool(con.Execute("SELECT EXISTS (SELECT * From organisations WHERE domicile=1 AND incID='" & incID & "')").Fields(0).Value) Then
                    Console.WriteLine("Already added: " & incID & " " & incdStr & " " & name)
                Else
                    Call NameResOrg(1, name, CDate(incdStr), Nothing, 1, incID)
                    If Not IsNumeric(incID) And Left(incID, 1) <> "C" Then
                        orgType = con.Execute("SELECT orgType FROM hkcrtypes WHERE prefix='" & Left(incID, 2) & "'").Fields(0).Value.ToString
                    Else
                        orgType = "" 'the file doesn't specifically tell us the orgType
                    End If
                    con.Execute("INSERT INTO persons() VALUES ()")
                    p = LastID(con)
                    con.Execute("INSERT INTO organisations (domicile,personID,Name1,cName,incID,orgType,incDate,incUpd) " & Valsql({1, p, name, cName, incID, orgType, incdStr, MSdateTime(Now)}))
                    Console.WriteLine("Added " & incID & " " & incdStr & " " & name)
                End If
            Else
                'changed a name, possibly
                rs.Open("SELECT * FROM organisations WHERE domicile=1 AND incID='" & incID & "'", con)
                If rs.EOF Then
                    'org not found in DB, could be an LP/LF/OF/RT/SO
                    'or renamed shortly after incorporation - the CR files then exclude the incorporation
                    'so just add the current name and incID and flag the missing data in the missingorg table
                    If Left(name, 5) <> "HKCR:" Then Call NameResOrg(1, name, Nothing, Nothing, 1, incID)
                    If Not IsNumeric(incID) And Left(incID, 1) <> "C" Then
                        orgType = con.Execute("SELECT orgType FROM hkcrtypes WHERE prefix='" & Left(incID, 2) & "'").Fields(0).Value.ToString
                    Else
                        orgType = "" 'the file doesn't specifically tell us the orgType
                    End If
                    con.Execute("INSERT INTO persons() VALUES ()")
                    p = LastID(con)
                    con.Execute("INSERT INTO organisations (domicile,personID,Name1,cName,incID,incDate,orgType,incUpd)" & Valsql({1, p, name, cName, incID, incdStr, orgType, MSdateTime(Now)}))
                    Console.WriteLine("Added " & incID & " " & name)
                    Console.WriteLine("Org was missing, find old names & incDate for:" & incID & " " & incdStr & " " & name)
                    con.Execute("INSERT IGNORE INTO missingorg (orgID) VALUES (" & p & ")")
                Else
                    p = CInt(rs("personID").Value)
                    oldCname = rs("cName").Value.ToString
                    oldName = rs("Name1").Value.ToString
                    If Left(oldName, 5) <> "HKCR:" Then oldName = TrimName(oldName) 'remove extensions for comparison
                    If (StrComp(name, oldName, vbBinaryCompare) <> 0 Or StrComp(cName, oldCname, vbBinaryCompare) <> 0) And Not CBool(rs("incName").Value) Then
                        'one or both names has changed
                        'don't disturb it if incName is True (meaning, ignore HKCR Names, use place of incorporation) 
                        'but we may be processing an old file and the name has changed since then, so try to find the name change
                        If Left(name, 5) = "HKCR:" Then ensql = "isNull(oldName)" Else ensql = "oldName='" & Apos(name) & "'"
                        If cName = "" Then cnsql = "isNull(oldcName)" Else cnsql = "oldcName='" & cName & "'"
                        rs2.Open("SELECT * FROM namechanges WHERE personID=" & p & " AND " & ensql & " AND " & cnsql & " AND dateChanged >='" & chgdStr & "'", con)
                        If Not rs2.EOF Then
                            Console.WriteLine("New name has subsequently changed on " & MSdate(CDate(rs2("dateChanged").Value)) & " for orgID:" & p & " " & name)
                        Else
                            If Left(name, 5) <> "HKCR:" Then
                                tempID = p
                                incDate = DBdate(rs("incDate"))
                                Call NameResOrg(tempID, name, incDate, Nothing, 1, incID)
                            End If
                            con.Execute("UPDATE organisations" & Setsql("name1,cName", {name, cName}) & "personID=" & p)
                            If Left(oldName, 5) = "HKCR:" Then oldName = ""
                            con.Execute("INSERT INTO namechanges (personID,oldName,oldcName,dateChanged)" & Valsql({p, oldName, oldCname, chgdStr}))
                            Console.WriteLine("New name orgID:" & p & " " & chgdStr & " " & name)
                        End If
                        rs2.Close()
                    Else
                        Console.WriteLine("Already renamed orgID:" & p & " " & chgdStr & " " & name)
                    End If
                End If
                rs.Close()
            End If
        Next
        con.Close()
        con = Nothing
        Exit Sub
RepErr:
        Call ErrMail("CR ProcFileL failed", Err, "File: " & dest & vbCrLf & "Last name: " & name)
    End Sub
    Sub ProcFileF(dest As String)
        'New from files dated 2021-12-06
        'process the CR CSV file for Foreign companies at dest
        On Error GoTo RepErr
        Dim s, c(), name, ensql, oldName, cName, cnsql, oldCname, incID, dStr, r1(), r2() As String,
            x, y, p, tempID, domicile As Integer,
            mostLatin, leastLatin, score As Double,
            newCo As Boolean,
            incDate As Date,
            con As New ADODB.Connection, rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
        Call OpenEnigma(con)
        s = My.Computer.FileSystem.ReadAllText(dest)
        s = RemCSVbreaks(s) 'some files have erroneous carriage returns inside names, e.g. 20230626 local companies
        c = SplitCSVrows(s)
        'ignore c(0) header row
        x = 1
        r1 = ReadCSVrow(c(x))
        Do Until x > UBound(c)
            domicile = Nothing
            name = ""
            cName = ""
            incID = r1(3)
            dStr = MSdateDMY(r1(4) & r1(5))
            newCo = r1(5) = ""
            'foreign companies can have multiple corporate names sharing same sequence number in csv file
            'English and Chinese names can appear in any row, so we have to score them to figure out which is which!
            mostLatin = LatinScore(r1(1))
            leastLatin = mostLatin
            'provisionally determine first name is either English or Chinese
            If mostLatin > 0.5 Then name = r1(1) Else cName = r1(1)
            y = x + 1
            'look ahead for same company
            Do Until y > UBound(c)
                r2 = ReadCSVrow(c(y))
                If r2(0) = r1(0) Then
                    'same company & event with another corporate name
                    score = LatinScore(r2(1))
                    If score > mostLatin And score > 0.5 Then
                        name = CleanStr(r2(1))
                        mostLatin = score
                    ElseIf score < leastLatin And score <= 0.5 Then
                        cName = CleanStr(r2(1))
                        leastLatin = score
                    End If
                    y += 1
                Else
                    'reached different co.
                    r1 = r2
                    Exit Do
                End If
            Loop
            x = y
            If name = "" Then
                name = "HKCR:" & incID & " (no English name)"
            Else
                name = ConfName(name)
            End If
            'now process the changes
            If newCo Then
                'Non-HK company registered in HK
                'Only do this if we haven't already added it
                If CBool(con.Execute("SELECT EXISTS(SELECT * FROM organisations o JOIN freg f ON o.personID=f.orgID WHERE hostdom=1 AND regID='" & incID & "')").Fields(0).Value) Then
                    Console.WriteLine("Already added: " & incID & " " & dStr & " " & name)
                Else
                    Call NameResOrg(1, name, Nothing, Nothing, Nothing, incID)
                    name = Replace(name, "(incorp. ID:", "(HKCR:")
                    con.Execute("INSERT INTO persons() VALUES ()")
                    p = LastID(con)
                    con.Execute("INSERT INTO organisations (orgType,personID,Name1,cName)" & Valsql({27, p, name, cName}))
                    con.Execute("INSERT INTO freg(hostDom,orgID,regID,regDate,regUpd)" & Valsql({1, p, incID, dStr, MSdateTime(Now)}))
                    Console.WriteLine("Added " & incID & " " & dStr & " " & name & vbTab & cName)
                End If
            Else
                'changed a name, possibly
                rs.Open("SELECT * FROM organisations o JOIN freg f ON o.personID=f.orgID WHERE hostdom=1 AND regID='" & incID & "'", con)
                If Not rs.EOF Then
                    domicile = DBint(rs("domicile"))
                    incID = rs("incID").Value.ToString
                End If
                If rs.EOF Then
                    'org not found in DB, could be renamed shortly after registration - the CR files then exclude the registration
                    'so just add the current name and incID and flag the missing data in the missingorg table
                    If Left(name, 5) <> "HKCR:" Then Call NameResOrg(1, name, Nothing, Nothing, Nothing, incID)
                    name = Replace(name, "(incorp. ID:", "(HKCR:")
                    con.Execute("INSERT INTO persons() VALUES ()")
                    p = LastID(con)
                    con.Execute("INSERT INTO organisations (orgType,personID,Name1,cName)" & Valsql({27, p, name, cName}))
                    con.Execute("INSERT INTO freg(hostDom,orgID,regID,regUpd)" & Valsql({1, p, incID, MSdateTime(Now)}))
                    Console.WriteLine("Added " & incID & " " & name)
                    Console.WriteLine("Org was missing, find old names & incDate for:" & incID & " " & dStr & " " & name)
                    con.Execute("INSERT IGNORE INTO missingorg (orgID) VALUES (" & p & ")")
                Else
                    p = CInt(rs("personID").Value)
                    oldCname = rs("cName").Value.ToString
                    oldName = rs("Name1").Value.ToString
                    If Left(oldName, 5) <> "HKCR:" Then oldName = TrimName(oldName) 'remove extensions for comparison
                    If (StrComp(name, oldName, vbBinaryCompare) <> 0 Or StrComp(cName, oldCname, vbBinaryCompare) <> 0) And
                                    Left(name, 9) <> "KABUSHIKI" And Left(name, 6) <> "THANAK" And Not CBool(rs("incName").Value) Then
                        'one or both names has changed
                        'don't disturb it if incName is True (meaning, ignore HKCR Names, use place of incorporation) 
                        'but we may be processing an old file and the name has changed since then, so try to find the name change
                        If Left(name, 5) = "HKCR:" Then ensql = "isNull(oldName)" Else ensql = "oldName='" & Apos(name) & "'"
                        If cName = "" Then cnsql = "isNull(oldcName)" Else cnsql = "oldcName='" & Apos(cName) & "'"
                        rs2.Open("SELECT * FROM namechanges WHERE personID=" & p & " AND " & ensql & " AND " & cnsql & " AND dateChanged >='" & dStr & "'", con)
                        If Not rs2.EOF Then
                            Console.WriteLine("New name has subsequently changed on " & MSdate(CDate(rs2("dateChanged").Value)) & " for orgID:" & p & " " & name)
                        Else
                            If Left(name, 5) <> "HKCR:" Then
                                tempID = p
                                incDate = DBdate(rs("incDate"))
                                Call NameResOrg(tempID, name, incDate, Nothing, domicile, incID)
                            End If
                            con.Execute("UPDATE organisations" & Setsql("name1,cName", {name, cName}) & "personID=" & p)
                            If Left(oldName, 5) = "HKCR:" Then oldName = ""
                            con.Execute("INSERT INTO namechanges (personID,oldName,oldcName,dateChanged)" & Valsql({p, oldName, oldCname, dStr}))
                            Console.WriteLine("New name orgID:" & p & " " & dStr & " " & name & vbTab & cName)
                        End If
                        rs2.Close()
                    Else
                        Console.WriteLine("Already renamed orgID:" & p & " " & dStr & " " & name)
                    End If
                End If
                rs.Close()
            End If
        Loop
        con.Close()
        con = Nothing
        Exit Sub
RepErr:
        Call ErrMail("CR ProcFileF failed", Err, "File: " & dest & vbCrLf & "Last name: " & name)
    End Sub

    Sub ProcFile(dest As String)
        'Old version for files before 2021-12-06 which combined Local and Foreign companies
        'process the CR CSV file at dest
        On Error GoTo RepErr
        Dim s, c(), r(), name, ensql, oldName, cName, cnsql, oldCname, incID, dStr, orgType As String,
            x, p, tempID, domicile As Integer,
            newCo As Boolean,
            rowDate, incDate As Date,
            con As New ADODB.Connection, rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
        Call OpenEnigma(con)
        s = My.Computer.FileSystem.ReadAllText(dest)
        'strip off the last newline before split
        If Right(s, 1) = Chr(10) Then s = Left(s, Len(s) - 1)
        c = Split(s, Chr(10))
        For x = 2 To UBound(c)
            r = ReadCSVrow(c(x))
            name = ConfName(r(1))
            cName = r(2)
            dStr = r(4) & r(5)
            newCo = r(5) = ""
            If Left(r(3), 11) = "carrying on" Then
                'spread over two rows with HKCR ID on second row and date on a random row
                x += 1
                r = ReadCSVrow(c(x))
                If dStr = "" Then
                    'date is on second row
                    dStr = r(4) & r(5)
                    newCo = r(5) = ""
                End If
                Console.WriteLine("Carrying on business as:" & r(1) & r(2))
            End If
            incID = r(3)
            If Left(incID, 1) = "F" And IsNumeric(Mid(incID, 2, 1)) Then
                domicile = Nothing
            Else
                domicile = 1
                If IsNumeric(Left(incID, 1)) Then incID = Right("000000" & incID, 7)
            End If
            If name = "" Then name = "HKCR:" & incID & " (no English name)"
            rowDate = ReadDMY(dStr)
            dStr = "'" & MSdate(rowDate) & "'"
            If newCo Then
                If cName = "" Then cName = "NULL" Else cName = "'" & Apos(cName) & "'"
                If domicile = 1 Then
                    'Hong Kong company
                    'Only do this if we haven't already added it
                    If CBool(con.Execute("SELECT EXISTS (SELECT * From organisations WHERE domicile=1 AND incID='" & incID & "')").Fields(0).Value) Then
                        Console.WriteLine("Already added: " & incID & " " & Mid(dStr, 2, 10) & " " & name)
                    Else
                        Call NameResOrg(1, name, rowDate, Nothing, 1, incID)
                        If Not IsNumeric(incID) Then
                            orgType = con.Execute("SELECT orgType FROM hkcrtypes WHERE prefix='" & Left(incID, 2) & "'").Fields(0).Value.ToString
                        Else
                            orgType = "NULL" 'the file doesn't specifically tell us the orgType
                        End If
                        con.Execute("INSERT INTO persons() VALUES ()")
                        p = LastID(con)
                        con.Execute("INSERT INTO organisations (domicile,personID,Name1,cName,incID,orgType,incDate,incUpd) " &
                                            "VALUES (1," & p & ",'" & Apos(name) & "'," & cName & ",'" & incID & "'," & orgType & "," & dStr & ",NOW())")
                        Console.WriteLine("Added " & incID & " " & Mid(dStr, 2, 10) & " " & name)
                    End If
                Else
                    'Non-HK company registered in HK
                    'Only do this if we haven't already added it
                    If CBool(con.Execute("SELECT EXISTS(SELECT * FROM organisations o JOIN freg f ON o.personID=f.orgID WHERE hostdom=1 AND regID='" & incID & "')").Fields(0).Value) Then
                        Console.WriteLine("Already added: " & incID & " " & Mid(dStr, 2, 10) & " " & name)
                    Else
                        Call NameResOrg(1, name, Nothing, Nothing, Nothing, incID)
                        name = Replace(name, "(incorp. ID:", "(HKCR:")
                        con.Execute("INSERT INTO persons() VALUES ()")
                        p = LastID(con)
                        con.Execute("INSERT INTO organisations (orgType,personID,Name1,cName) " &
                                            "VALUES (27," & p & ",'" & Apos(name) & "'," & cName & ")")
                        con.Execute("INSERT INTO freg(hostDom,orgID,regID,regDate,regUpd) VALUES(1," & p & ",'" & incID & "'," & dStr & ",NOW())")
                        Console.WriteLine("Added " & incID & " " & Mid(dStr, 2, 10) & " " & name)
                    End If
                End If
            Else
                'changed a name, possibly
                If domicile = 1 Then
                    rs.Open("SELECT * FROM organisations WHERE domicile=1 AND incID='" & incID & "'", con)
                Else
                    rs.Open("SELECT * FROM organisations o JOIN freg f ON o.personID=f.orgID WHERE hostdom=1 AND regID='" & incID & "'", con)
                    If Not rs.EOF Then
                        domicile = DBint(rs("domicile"))
                        incID = rs("incID").Value.ToString
                    End If
                End If
                If rs.EOF Then
                    'org not found in DB, could be an LP/LF/OF/RT/SO
                    'or renamed shortly after incorporation - the CR files then exclude the incorporation
                    'so just add the current name and incID and flag the missing data in the missingorg table
                    If cName = "" Then cName = "NULL" Else cName = "'" & Apos(cName) & "'"
                    If domicile = 1 Then
                        If Left(name, 5) <> "HKCR:" Then Call NameResOrg(1, name, Nothing, Nothing, 1, incID)
                        If Not IsNumeric(incID) Then
                            orgType = con.Execute("SELECT orgType FROM hkcrtypes WHERE prefix='" & Left(incID, 2) & "'").Fields(0).Value.ToString
                        Else
                            orgType = "NULL" 'the file doesn't specifically tell us the orgType
                        End If
                        con.Execute("INSERT INTO persons() VALUES ()")
                        p = LastID(con)
                        con.Execute("INSERT INTO organisations (domicile,personID,Name1,cName,incID,orgType,incUpd) " &
                                            "VALUES (1," & p & ",'" & Apos(name) & "'," & cName & ",'" & incID & "'," & orgType & ",NOW())")
                        Console.WriteLine("Added " & incID & " " & name)
                    Else
                        If Left(name, 5) <> "HKCR:" Then Call NameResOrg(1, name, Nothing, Nothing, Nothing, incID)
                        name = Replace(name, "(incorp. ID:", "(HKCR:")
                        con.Execute("INSERT INTO persons() VALUES ()")
                        p = LastID(con)
                        con.Execute("INSERT INTO organisations (orgType,personID,Name1,cName) " &
                                            "VALUES (27," & p & ",'" & Apos(name) & "'," & cName & ")")
                        con.Execute("INSERT INTO freg(hostDom,orgID,regID,regUpd) VALUES(1," & p & ",'" & incID & "'," & ",NOW())")
                        Console.WriteLine("Added " & incID & " " & name)
                    End If
                    Console.WriteLine("Org was missing, find old names & incDate for:" & incID & " " & Mid(dStr, 2, 10) & " " & name)
                    con.Execute("INSERT IGNORE INTO missingorg (orgID) VALUES (" & p & ")")
                Else
                    p = CInt(rs("personID").Value)
                    oldCname = rs("cName").Value.ToString
                    oldName = rs("Name1").Value.ToString
                    If Left(oldName, 5) <> "HKCR:" Then
                        'remove extensions for comparison
                        oldName = TrimName(oldName)
                    End If
                    If (StrComp(name, oldName, vbBinaryCompare) <> 0 Or StrComp(cName, oldCname, vbBinaryCompare) <> 0) And
                                    Left(name, 9) <> "KABUSHIKI" And Left(name, 6) <> "THANAK" And Not CBool(rs("incName").Value) Then
                        'one or both names has changed
                        'don't disturb it if incName is True (meaning, ignore HKCR Names, use place of incorporation) 
                        'but we may be processing an old file and the name has changed since then, so try to find the name change
                        If Left(name, 5) = "HKCR:" Then
                            ensql = "isNull(oldName)"
                        Else
                            ensql = "oldName='" & Apos(name) & "'"
                        End If
                        If cName = "" Then
                            cName = "NULL"
                            cnsql = "isNull(oldcName)"
                        Else
                            cName = "'" & Apos(cName) & "'"
                            cnsql = "oldcName=" & cName
                        End If
                        rs2.Open("SELECT * FROM namechanges WHERE personID=" & p & " AND " & ensql & " AND " & cnsql & " AND dateChanged >=" & dStr, con)
                        If Not rs2.EOF Then
                            Console.WriteLine("New name has subsequently changed on " & MSdate(CDate(rs2("dateChanged").Value)) & " for orgID:" & p & " " & name)
                        Else
                            If Left(name, 5) <> "HKCR:" Then
                                tempID = p
                                incDate = DBdate(rs("incDate"))
                                Call NameResOrg(tempID, name, incDate, Nothing, domicile, incID)
                            End If
                            con.Execute("UPDATE organisations SET name1='" & Apos(name) & "', cName=" & cName & " WHERE personID=" & p)
                            If Left(oldName, 5) = "HKCR:" Then
                                oldName = "NULL"
                            Else
                                oldName = "'" & Apos(oldName) & "'"
                            End If
                            If oldCname = "" Then
                                oldCname = "NULL"
                            Else
                                oldCname = "'" & Apos(oldCname) & "'"
                            End If
                            con.Execute("INSERT INTO namechanges (personID,oldName,oldcName,dateChanged) VALUES (" &
                                            p & "," & oldName & "," & oldCname & "," & dStr & ")")
                            Console.WriteLine("New name orgID:" & p & " " & Mid(dStr, 2, 10) & " " & name)
                        End If
                        rs2.Close()
                    Else
                        Console.WriteLine("Already renamed orgID:" & p & " " & Mid(dStr, 2, 10) & " " & name)
                    End If
                End If
                rs.Close()
            End If
        Next
        con.Close()
        con = Nothing
        Exit Sub
RepErr:
        Call ErrMail("CR ProcFile failed", Err, "File: " & dest & vbCrLf & "Last name: " & name)
    End Sub
    Function ConfName(n As String) As String
        'conform a name which contains the definite article. Preserve case
        Dim s As String
        If Right(n, 5) = "-THE-" Then
            s = Left(n, Len(n) - 5) & "(" & Left(Right(n, 4), 3) & ")"
        ElseIf Right(n, 6) = "-THE-." Then
            s = Left(n, Len(n) - 6) & "(" & Left(Right(n, 5), 3) & ")"
        ElseIf Right(n, 7) = "- THE -" Then
            s = Left(n, Len(n) - 7) & "(" & Left(Right(n, 5), 3) & ")"
        ElseIf Left(n, 4) = "The " Then
            s = Mid(n, 5) & " (" & Left(n, 3) & ")"
        Else
            s = n
        End If
        'remove spaces inside parentheses
        s = Replace(s, "( ", "(")
        s = Replace(s, " )", ")")
        s = StripSpace(s)
        'the largest name we have found in ICRIS is 150 characters (ICRISID:74516). We need an extra 15 for date tags
        'we have used varchar(255) in the table
        s = Left(s, 240)
        s = Trim(s)
        Return s
    End Function

End Module
