Option Explicit

'==============================================================
' OFV KJORETOYOPPSLAG - kjoretoyinfo + alle registreringer
'
' Leser en liste med Regnr (kolonne C) og/eller VIN (kolonne D)
' fra rad 5 og nedover pa "Oppslag"-arket (eller det aktive
' arket, hvis "Oppslag" ikke finnes). For hvert kjoretoy i
' listen slas det opp mot Statens vegvesen (kun forstegangs-
' registrering) og OFV Transactions API (alle eierskifter/
' registreringer), og resultatet stables etter hverandre i
' et eget ark: "Kjoretoyrapport" - en blokk per bil.
'
' Formal: raskt se om hver bil star registrert pa et selskap
' eller en privatperson, og se hele eierhistorikken.
'==============================================================


'==============================================================
' KONFIGURASJON
'==============================================================

Private Const INPUT_SHEET As String = "Oppslag"
Private Const REPORT_SHEET As String = "Kjoretoyrapport"

Private Const FIRST_DATA_ROW As Long = 5
Private Const COL_REGNR As Long = 3   ' kolonne C
Private Const COL_VIN As Long = 4     ' kolonne D

Private Const REPORT_HEADER_ROW As Long = 5

Private Const CELL_OFV_KEY As String = "B1"
Private Const CELL_SVV_KEY As String = "B2"

Private Const OFV_URL As String = _
    "https://api.ofv.no/transactions/v1/"

Private Const SVV_URL As String = _
    "https://akfell-datautlevering.atlas.vegvesen.no/" & _
    "enkeltoppslag/kjoretoydata?"

' Vidt datointervall for a fange opp HELE historikken til bilen.
Private Const OFV_DATE_FROM As String = "1950-01-01"

Private Const MAX_RETRIES As Long = 4
Private Const RETRY_WAIT_MS As Long = 3000
Private Const API_PAUSE_MS As Long = 150


'==============================================================
' VENTEFUNKSJON UTEN WIN32-KALL
'
' Windows Defender sin ASR-regel "Blokker Win32 API-kall fra
' Office-makro" blokkerer enhver Declare ... Lib "kernel32".
' SafePause gir en ikke-blokkerende pause i millisekunder med
' rent VBA (Timer + DoEvents), uten ekstern bibliotekdeklarasjon.
'==============================================================

Private Sub SafePause(ByVal milliseconds As Long)

    Dim startTime As Double
    Dim currentTime As Double
    Dim elapsedMilliseconds As Double

    If milliseconds <= 0 Then Exit Sub

    startTime = Timer

    Do
        DoEvents
        currentTime = Timer

        If currentTime >= startTime Then
            elapsedMilliseconds = (currentTime - startTime) * 1000#
        Else
            elapsedMilliseconds = _
                ((86400# - startTime) + currentTime) * 1000#
        End If

    Loop While elapsedMilliseconds < milliseconds

End Sub


'==============================================================
' HOVEDMAKRO
'==============================================================

Public Sub OFV_SlaOppKjoretoy()

    Dim wsInput As Worksheet
    Dim wsReport As Worksheet

    Dim ofvKey As String
    Dim svvKey As String
    Dim dateToISO As String

    Dim queue As Object
    Dim key As Variant
    Dim vehicleData As Variant
    Dim queueKey As String

    Dim regNo As String
    Dim vin As String
    Dim identifier As String
    Dim useVin As Boolean

    Dim vehicleInfo As Object
    Dim transactions As Collection

    Dim lastRegRow As Long
    Dim lastVinRow As Long
    Dim lastInputRow As Long
    Dim inputRow As Long

    Dim currentRow As Long
    Dim vehicleIndex As Long
    Dim totalVehicles As Long

    Dim stage As String

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldCursor As Variant
    Dim applicationChanged As Boolean

    On Error GoTo FatalError

    stage = "finner input-arket"

    On Error Resume Next
    Set wsInput = ThisWorkbook.Worksheets(INPUT_SHEET)
    On Error GoTo 0

    If wsInput Is Nothing Then Set wsInput = ActiveSheet

    stage = "leser regnr/VIN-listen fra " & wsInput.Name & _
        "!C" & FIRST_DATA_ROW & " og nedover"

    lastRegRow = _
        wsInput.Cells(wsInput.Rows.Count, COL_REGNR).End(xlUp).Row

    lastVinRow = _
        wsInput.Cells(wsInput.Rows.Count, COL_VIN).End(xlUp).Row

    lastInputRow = Application.Max(lastRegRow, lastVinRow)

    Set queue = CreateObject("Scripting.Dictionary")
    queue.CompareMode = vbTextCompare

    For inputRow = FIRST_DATA_ROW To lastInputRow

        regNo = NormalizeIdentifier(wsInput.Cells(inputRow, COL_REGNR).Value)
        vin = NormalizeIdentifier(wsInput.Cells(inputRow, COL_VIN).Value)

        queueKey = BuildVehicleKey(vin, regNo)

        If Len(queueKey) > 0 Then
            If Not queue.Exists(queueKey) Then
                queue.Add queueKey, Array(regNo, vin)
            End If
        End If

    Next inputRow

    If queue.Count = 0 Then
        MsgBox _
            "Fant ingen registreringsnummer eller VIN i " & _
            wsInput.Name & "!C" & FIRST_DATA_ROW & " og nedover.", _
            vbInformation, "Kjoretoyoppslag"
        Exit Sub
    End If

    stage = "leser API-nokler"

    ofvKey = Trim$(CStr(wsInput.Range(CELL_OFV_KEY).Value))
    svvKey = Trim$(CStr(wsInput.Range(CELL_SVV_KEY).Value))

    If Len(ofvKey) = 0 Then
        MsgBox _
            "Fant ingen OFV-nokkel i " & wsInput.Name & "!" & _
            CELL_OFV_KEY & ". Skriv inn OFV API-nokkelen der.", _
            vbExclamation, "Kjoretoyoppslag"
        Exit Sub
    End If

    If Len(svvKey) = 0 Then
        MsgBox _
            "Fant ingen Vegvesen-nokkel i " & wsInput.Name & "!" & _
            CELL_SVV_KEY & ". Skriv inn Vegvesenet-nokkelen der.", _
            vbExclamation, "Kjoretoyoppslag"
        Exit Sub
    End If

    dateToISO = Format$(Date, "yyyy-mm-dd")

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation
    oldCursor = Application.Cursor

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait
    applicationChanged = True

    totalVehicles = queue.Count

    stage = "forbereder rapportarket"

    Set wsReport = GetOrCreateReportSheet(REPORT_SHEET)
    WriteGlobalHeader wsReport, totalVehicles

    currentRow = REPORT_HEADER_ROW + 1
    vehicleIndex = 0

    For Each key In queue.Keys

        vehicleIndex = vehicleIndex + 1
        vehicleData = queue(key)

        regNo = CStr(vehicleData(0))
        vin = CStr(vehicleData(1))
        useVin = (Len(vin) > 0)

        If useVin Then
            identifier = vin
        Else
            identifier = regNo
        End If

        '------------------------------------------------------
        ' STATENS VEGVESEN - forstegangsregistrering
        '------------------------------------------------------

        stage = "henter forstegangsregistrering (" & identifier & ")"

        Application.StatusBar = "Bil " & vehicleIndex & " av " & _
            totalVehicles & " | Statens vegvesen | " & identifier

        Set vehicleInfo = FetchVehicleInfoFromSVV(svvKey, regNo, vin)

        '------------------------------------------------------
        ' OFV - alle registreringer/eierskifter
        '------------------------------------------------------

        stage = "henter registreringer fra OFV (" & identifier & ")"

        Application.StatusBar = "Bil " & vehicleIndex & " av " & _
            totalVehicles & " | OFV | " & identifier

        Set transactions = FetchOFVTransactions( _
            ofvKey, identifier, useVin, regNo, vin, _
            OFV_DATE_FROM, dateToISO)

        '------------------------------------------------------
        ' SKRIV BLOKK FOR DENNE BILEN
        '------------------------------------------------------

        currentRow = WriteVehicleBlock( _
            wsReport, currentRow, vehicleIndex, _
            regNo, vin, vehicleInfo, transactions)

        SafePause API_PAUSE_MS

    Next key

    stage = "formaterer rapportarket"

    ApplyReportColumnWidths wsReport

    Application.Calculation = oldCalculation
    Application.StatusBar = False

    RestoreApplicationState _
        oldScreenUpdating, oldEnableEvents, oldCalculation, oldCursor

    applicationChanged = False

    wsReport.Activate
    wsReport.Range("B2").Select

    MsgBox _
        totalVehicles & " kjoretoy er slatt opp og lagt i arket """ & _
        REPORT_SHEET & """.", _
        vbInformation, "Kjoretoyoppslag"

    Exit Sub

FatalError:

    Dim errorNumber As Long
    Dim errorDescription As String

    errorNumber = Err.Number
    errorDescription = Err.Description

    Application.StatusBar = False

    If applicationChanged Then
        RestoreApplicationState _
            oldScreenUpdating, oldEnableEvents, oldCalculation, oldCursor
    End If

    MsgBox _
        "Oppslaget ble avbrutt." & vbCrLf & vbCrLf & _
        "Trinn: " & stage & vbCrLf & _
        "Feil " & errorNumber & ": " & errorDescription, _
        vbCritical, "Kjoretoyoppslag"

End Sub


Private Sub RestoreApplicationState( _
    ByVal screenUpdatingValue As Boolean, _
    ByVal enableEventsValue As Boolean, _
    ByVal calculationValue As XlCalculation, _
    ByVal cursorValue As Variant)

    Application.StatusBar = False
    Application.Cursor = cursorValue
    Application.ScreenUpdating = screenUpdatingValue
    Application.EnableEvents = enableEventsValue
    Application.Calculation = calculationValue

End Sub


'==============================================================
' STATENS VEGVESEN
'==============================================================

Private Function FetchVehicleInfoFromSVV( _
    ByVal apiKey As String, _
    ByVal regNo As String, _
    ByVal vin As String) As Object

    Dim result As Object
    Dim responseText As String
    Dim statusText As String
    Dim isoDate As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("Status") = "Statens vegvesen - ingen dato"

    If Len(regNo) > 0 Then
        responseText = GetSVVResponse( _
            apiKey, "kjennemerke", regNo, statusText)
    End If

    If Len(responseText) = 0 And Len(vin) > 0 Then
        responseText = GetSVVResponse( _
            apiKey, "understellsnummer", vin, statusText)
    End If

    If Len(responseText) = 0 Then
        If Len(statusText) > 0 Then result("Status") = statusText
        Set FetchVehicleInfoFromSVV = result
        Exit Function
    End If

    isoDate = VariantToString(JSON_ExtractValue( _
        responseText, "registrertForstegangNorgeDato"))

    If Len(isoDate) < 10 Then
        isoDate = VariantToString(JSON_ExtractValue( _
            responseText, "registrertForstegangDato"))
    End If

    If Len(isoDate) >= 10 Then
        result("FirstRegistrationDate") = DateFromISO(isoDate)
        result("Status") = "Statens vegvesen"
    End If

    Set FetchVehicleInfoFromSVV = result

End Function


Private Function GetSVVResponse( _
    ByVal apiKey As String, _
    ByVal filterName As String, _
    ByVal identifier As String, _
    ByRef statusText As String) As String

    Dim http As Object
    Dim url As String
    Dim attempt As Long
    Dim statusCode As Long
    Dim responseText As String
    Dim lastError As String

    url = SVV_URL & filterName & "=" & identifier
    statusText = vbNullString

    For attempt = 1 To MAX_RETRIES

        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

        statusCode = 0
        responseText = vbNullString

        On Error Resume Next

        http.Open "GET", url, False
        http.setRequestHeader "SVV-Authorization", "Apikey " & apiKey
        http.setRequestHeader "Accept", "application/json"
        http.setTimeouts 10000, 10000, 30000, 30000
        http.Send

        statusCode = http.Status
        responseText = http.responseText

        If Err.Number <> 0 Then
            lastError = Err.Description
            statusCode = 0
            Err.Clear
        End If

        On Error GoTo 0

        Select Case statusCode

            Case 200
                statusText = "Statens vegvesen"
                GetSVVResponse = responseText
                Exit Function

            Case 401
                statusText = _
                    "Feil: Vegvesenet 401 - kontroller API-nokkel"
                Exit Function

            Case 403
                statusText = _
                    "Feil: Vegvesenet 403 - tilgang eller kvote"
                Exit Function

            Case 404
                statusText = "Statens vegvesen - ingen treff"
                Exit Function

            Case 429, 500, 502, 503, 504
                lastError = CStr(statusCode) & ": " & responseText
                SafePause RETRY_WAIT_MS * attempt

            Case Else

                If statusCode <> 0 Then
                    statusText = "Feil: Vegvesenet HTTP " & statusCode
                    Exit Function
                Else
                    SafePause RETRY_WAIT_MS * attempt
                End If

        End Select

    Next attempt

    statusText = "Feil: Vegvesenet - " & lastError

End Function


'==============================================================
' OFV-TRANSAKSJONER (alle registreringer for kjoretoyet)
'==============================================================

Private Function FetchOFVTransactions( _
    ByVal apiKey As String, _
    ByVal identifier As String, _
    ByVal useVin As Boolean, _
    ByVal originalRegNo As String, _
    ByVal originalVin As String, _
    ByVal dateFromISO As String, _
    ByVal dateToISO As String) As Collection

    Dim rows As New Collection
    Dim items As Collection

    Dim filterKey As String
    Dim cursor As String
    Dim body As String
    Dim statusText As String
    Dim responseText As String
    Dim transactionsJSON As String
    Dim paginationJSON As String

    Dim item As Variant
    Dim hasNextValue As Variant
    Dim cursorValue As Variant
    Dim hasNext As Boolean

    If useVin Then
        filterKey = "chassisNumber"
    Else
        filterKey = "regNo"
    End If

    cursor = vbNullString

    Do

        body = "{""filters"":{"
        body = body & """" & filterKey & """:"""
        body = body & JsonEscape(identifier) & ""","
        body = body & """transactionDateFrom"":"""
        body = body & dateFromISO & ""","
        body = body & """transactionDateTo"":"""
        body = body & dateToISO & """},"
        body = body & """pagination"":{""first"":1000"

        If Len(cursor) > 0 Then
            body = body & ",""cursor"":"""
            body = body & JsonEscape(cursor) & """"
        End If

        body = body & "},"
        body = body & """sorting"":{"
        body = body & """orderBy"":""transactionDate"","
        body = body & """orderDirection"":""ASC""}}"

        responseText = PostOFVWithRetries(apiKey, body, statusText)

        If statusText <> "OK" Then
            rows.Add BuildEmptyTransactionRow( _
                originalRegNo, originalVin, statusText)
            Set FetchOFVTransactions = rows
            Exit Function
        End If

        transactionsJSON = JSON_ExtractObject(responseText, "transactions")
        Set items = JSON_ArrayAllElements(transactionsJSON)

        For Each item In items
            rows.Add BuildTransactionRow( _
                CStr(item), originalRegNo, originalVin)
        Next item

        paginationJSON = JSON_ExtractObject(responseText, "pagination")
        hasNextValue = JSON_ExtractValue(paginationJSON, "hasNextPage")
        cursorValue = JSON_ExtractValue(paginationJSON, "endCursor")

        hasNext = False

        Select Case VarType(hasNextValue)
            Case vbBoolean
                hasNext = CBool(hasNextValue)
            Case vbString
                hasNext = (LCase$(CStr(hasNextValue)) = "true")
            Case vbByte, vbInteger, vbLong
                hasNext = (hasNextValue <> 0)
            Case vbSingle, vbDouble, vbCurrency
                hasNext = (hasNextValue <> 0)
        End Select

        If hasNext Then

            If IsNull(cursorValue) Then Exit Do
            If IsEmpty(cursorValue) Then Exit Do

            cursor = CStr(cursorValue)
            If Len(cursor) = 0 Then Exit Do

            SafePause API_PAUSE_MS

        Else
            Exit Do
        End If

    Loop

    If rows.Count = 0 Then
        rows.Add BuildEmptyTransactionRow( _
            originalRegNo, originalVin, "Ingen registreringer funnet")
    End If

    Set FetchOFVTransactions = rows

End Function


Private Function BuildEmptyTransactionRow( _
    ByVal originalRegNo As String, _
    ByVal originalVin As String, _
    ByVal statusText As String) As Object

    Dim result As Object

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    result("RegNo") = originalRegNo
    result("ChassisNumber") = originalVin
    result("Status") = statusText

    Set BuildEmptyTransactionRow = result

End Function


Private Function BuildTransactionRow( _
    ByVal transactionJSON As String, _
    ByVal originalRegNo As String, _
    ByVal originalVin As String) As Object

    Dim result As Object
    Dim fromJSON As String
    Dim toJSON As String
    Dim fromOwnerJSON As String
    Dim toOwnerJSON As String
    Dim companyJSON As String
    Dim extractedValue As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    extractedValue = JSON_ExtractValue(transactionJSON, "regNo")

    If Len(VariantToString(extractedValue)) > 0 Then
        result("RegNo") = extractedValue
    Else
        result("RegNo") = originalRegNo
    End If

    extractedValue = JSON_ExtractValue(transactionJSON, "chassisNumber")

    If Len(VariantToString(extractedValue)) > 0 Then
        result("ChassisNumber") = extractedValue
    Else
        result("ChassisNumber") = originalVin
    End If

    result("MakeName") = JSON_ExtractValue(transactionJSON, "makeName")
    result("ModelName") = JSON_ExtractValue(transactionJSON, "modelName")

    result("RegistrationType") = _
        JSON_ExtractValue(transactionJSON, "registrationType")

    result("FuelGroup") = JSON_ExtractValue(transactionJSON, "fuelGroup")
    result("IsLeased") = JSON_ExtractValue(transactionJSON, "isLeased")

    result("IsUsedImported") = _
        JSON_ExtractValue(transactionJSON, "isUsedImported")

    result("ChassisName") = JSON_ExtractValue(transactionJSON, "chassisName")

    result("Transmission") = _
        JSON_ExtractValue(transactionJSON, "transmission")

    result("VehicleGroupName") = _
        JSON_ExtractValue(transactionJSON, "vehicleGroupName")

    result("LastApprovedInspectionDate") = DateFromISO(VariantToString( _
        JSON_ExtractValue(transactionJSON, "lastApprovedInspectionDate")))

    result("NextInspectionDate") = DateFromISO(VariantToString( _
        JSON_ExtractValue(transactionJSON, "nextInspectionDate")))

    result("TransactionNumber") = _
        JSON_ExtractValue(transactionJSON, "transactionNumber")

    result("TransactionDate") = DateFromISO(VariantToString( _
        JSON_ExtractValue(transactionJSON, "transactionDate")))

    fromJSON = JSON_ExtractObject(transactionJSON, "from")
    toJSON = JSON_ExtractObject(transactionJSON, "to")

    fromOwnerJSON = JSON_ExtractObject(fromJSON, "owner")
    toOwnerJSON = JSON_ExtractObject(toJSON, "owner")

    result("FromOwnerType") = JSON_ExtractValue(fromOwnerJSON, "type")

    companyJSON = JSON_ExtractObject(fromOwnerJSON, "companyInfo")
    result("FromOwnerCompanyName") = JSON_ExtractValue(companyJSON, "name")
    result("FromOwnerCounty") = JSON_ExtractValue(fromOwnerJSON, "countyName")

    result("ToOwnerType") = JSON_ExtractValue(toOwnerJSON, "type")

    companyJSON = JSON_ExtractObject(toOwnerJSON, "companyInfo")
    result("ToOwnerCompanyName") = JSON_ExtractValue(companyJSON, "name")
    result("ToOwnerCounty") = JSON_ExtractValue(toOwnerJSON, "countyName")

    result("Status") = "OK"

    Set BuildTransactionRow = result

End Function


Private Function PostOFVWithRetries( _
    ByVal apiKey As String, _
    ByVal body As String, _
    ByRef statusText As String) As String

    Dim attempt As Long
    Dim statusCode As Long
    Dim lastError As String
    Dim responseText As String
    Dim http As Object

    statusText = "OK"

    For attempt = 1 To MAX_RETRIES

        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

        statusCode = 0
        responseText = vbNullString

        On Error Resume Next

        http.Open "POST", OFV_URL, False
        http.setRequestHeader "Ocp-Apim-Subscription-Key", apiKey
        http.setRequestHeader "Content-Type", "application/json"
        http.setTimeouts 10000, 10000, 30000, 30000
        http.Send body

        statusCode = http.Status
        responseText = http.responseText

        If Err.Number <> 0 Then
            lastError = Err.Description
            statusCode = 0
            Err.Clear
        End If

        On Error GoTo 0

        Select Case statusCode

            Case 200
                PostOFVWithRetries = responseText
                Exit Function

            Case 401
                statusText = "Feil: OFV 401 - kontroller API-nokkelen"
                Exit Function

            Case 403
                statusText = "Feil: OFV 403 - tilgang eller kvote"
                Exit Function

            Case 429, 500, 502, 503, 504
                lastError = CStr(statusCode) & ": " & responseText
                SafePause RETRY_WAIT_MS * attempt

            Case Else

                If statusCode <> 0 Then
                    statusText = "Feil: OFV HTTP " & statusCode
                    Exit Function
                Else
                    SafePause RETRY_WAIT_MS * attempt
                End If

        End Select

    Next attempt

    statusText = "Feil: OFV - " & lastError

End Function


'==============================================================
' RAPPORTARK
'==============================================================

Private Function GetOrCreateReportSheet( _
    ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
            ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set GetOrCreateReportSheet = ws

End Function


Private Sub WriteGlobalHeader( _
    ByVal ws As Worksheet, ByVal vehicleCount As Long)

    ws.Range("B2").Value = "Kjoretoyrapport"

    With ws.Range("B2")
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 120)
    End With

    ws.Range("B3").Value = _
        "Generert: " & Format$(Now, "dd.mm.yyyy hh:nn") & _
        "   |   Antall kjoretoy: " & vehicleCount

    ws.Range("B3").Font.Italic = True
    ws.Range("B3").Font.Color = RGB(100, 100, 100)

    WriteTableHeaderRow ws, REPORT_HEADER_ROW

End Sub


Private Sub WriteTableHeaderRow(ByVal ws As Worksheet, ByVal r As Long)

    ws.Range("B" & r).Value = "Regnr"
    ws.Range("C" & r).Value = "VIN"
    ws.Range("D" & r).Value = "Merke"
    ws.Range("E" & r).Value = "Modell"
    ws.Range("F" & r).Value = "Drivstoff"
    ws.Range("G" & r).Value = "Forstegangsreg."
    ws.Range("H" & r).Value = "Naaverende registrering"
    ws.Range("I" & r).Value = "Dato"
    ws.Range("J" & r).Value = "Type registrering"
    ws.Range("K" & r).Value = "Solgt av"
    ws.Range("L" & r).Value = "Solgt til"
    ws.Range("M" & r).Value = "Fylke (kjoper)"
    ws.Range("N" & r).Value = "Status"

    With ws.Range("B" & r & ":N" & r)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30
    End With

End Sub


'==============================================================
' EN BLOKK PER BIL - kjoretoyinfo vises 1 gang (flettet celler),
' alle registreringer listes radvis rett under.
'==============================================================

Private Function WriteVehicleBlock( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal vehicleIndex As Long, _
    ByVal regNoInput As String, _
    ByVal vinInput As String, _
    ByVal vehicleInfo As Object, _
    ByVal transactions As Collection) As Long

    Dim okTxns As Collection
    Dim allTxns() As Object
    Dim t As Object
    Dim latest As Object
    Dim hasLatest As Boolean

    Dim currentOwnerLabel As String
    Dim ownerColor As Long
    Dim ownerFontColor As Long
    Dim blockFillColor As Long

    Dim firstRegText As String
    Dim r As Long
    Dim i As Long
    Dim n As Long
    Dim rowCount As Long
    Dim lastRow As Long

    '----------------------------------------------------------
    ' Finn "siste" (nyeste) OK-registrering
    '----------------------------------------------------------

    Set okTxns = New Collection

    For Each t In transactions
        If VariantToString(t("Status")) = "OK" Then
            okTxns.Add t
        End If
    Next t

    hasLatest = (okTxns.Count > 0)

    If hasLatest Then
        Set latest = okTxns(okTxns.Count)
    End If

    n = transactions.Count
    rowCount = Application.Max(1, n)
    lastRow = startRow + rowCount - 1

    ' Tving tekstformat FOR verdiene skrives inn. Data fra OFV/SVV
    ' (f.eks. firmanavn) er ekstern, ukontrollert tekst - uten dette
    ' vil Excel tolke en verdi som starter med f.eks. "=", "+", "-"
    ' eller "@" som en formel i stedet for ren tekst.
    ws.Range("B" & startRow & ":N" & lastRow).NumberFormat = "@"

    '----------------------------------------------------------
    ' Naaverende registrering (selskap eller privatperson)
    '----------------------------------------------------------

    If hasLatest Then

        currentOwnerLabel = FriendlyOwnerLabel( _
            VariantToString(latest("ToOwnerType")), _
            VariantToString(latest("ToOwnerCompanyName")))

        If VariantToString(latest("ToOwnerType")) = "Privat" Then
            ownerColor = RGB(198, 239, 206)
            ownerFontColor = RGB(0, 97, 0)
        ElseIf Len(VariantToString( _
            latest("ToOwnerCompanyName"))) > 0 Then
            ownerColor = RGB(255, 235, 156)
            ownerFontColor = RGB(156, 101, 0)
        Else
            ownerColor = RGB(221, 235, 247)
            ownerFontColor = RGB(31, 78, 120)
        End If

    Else
        currentOwnerLabel = "Ukjent (ingen registreringer funnet)"
        ownerColor = RGB(242, 242, 242)
        ownerFontColor = RGB(89, 89, 89)
    End If

    '----------------------------------------------------------
    ' Forstegangsregistrert (kun fra Statens vegvesen)
    '----------------------------------------------------------

    If vehicleInfo.Exists("FirstRegistrationDate") Then
        If IsDate(vehicleInfo("FirstRegistrationDate")) Then
            firstRegText = Format$( _
                vehicleInfo("FirstRegistrationDate"), "dd.mm.yyyy")
        End If
    End If

    If Len(firstRegText) = 0 Then firstRegText = "Ukjent"

    If vehicleIndex Mod 2 = 0 Then
        blockFillColor = RGB(246, 249, 252)
    Else
        blockFillColor = RGB(255, 255, 255)
    End If

    '----------------------------------------------------------
    ' Kjoretoyinfo - flettet, vises EN gang for hele bilen
    '----------------------------------------------------------

    ws.Range("B" & startRow & ":B" & lastRow).Merge
    ws.Range("B" & startRow).Value = _
        FirstNonEmpty(FieldOrEmpty(latest, "RegNo"), regNoInput)

    ws.Range("C" & startRow & ":C" & lastRow).Merge
    ws.Range("C" & startRow).Value = _
        FirstNonEmpty(FieldOrEmpty(latest, "ChassisNumber"), vinInput)

    ws.Range("D" & startRow & ":D" & lastRow).Merge
    ws.Range("D" & startRow).Value = FieldOrEmpty(latest, "MakeName")

    ws.Range("E" & startRow & ":E" & lastRow).Merge
    ws.Range("E" & startRow).Value = FieldOrEmpty(latest, "ModelName")

    ws.Range("F" & startRow & ":F" & lastRow).Merge
    ws.Range("F" & startRow).Value = FieldOrEmpty(latest, "FuelGroup")

    ws.Range("G" & startRow & ":G" & lastRow).Merge
    ws.Range("G" & startRow).Value = firstRegText

    ws.Range("H" & startRow & ":H" & lastRow).Merge
    ws.Range("H" & startRow).Value = currentOwnerLabel

    With ws.Range("B" & startRow & ":G" & lastRow)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    '----------------------------------------------------------
    ' Registreringsrader (I:N) - en rad per registrering,
    ' eldste oeverst (leses som historikk nedover).
    '----------------------------------------------------------

    r = startRow

    If n = 0 Then

        ws.Range("N" & r).Value = "Ingen registreringer funnet i OFV"

    Else

        ReDim allTxns(1 To n)

        i = 0
        For Each t In transactions
            i = i + 1
            Set allTxns(i) = t
        Next t

        For i = 1 To n

            Set t = allTxns(i)

            If VariantToString(t("Status")) = "OK" Then

                If IsDate(t("TransactionDate")) Then
                    ws.Range("I" & r).Value = _
                        Format$(t("TransactionDate"), "dd.mm.yyyy")
                End If

                ws.Range("J" & r).Value = _
                    FieldOrEmpty(t, "RegistrationType")

                ws.Range("K" & r).Value = FriendlyOwnerLabel( _
                    VariantToString(t("FromOwnerType")), _
                    VariantToString(t("FromOwnerCompanyName")))

                ws.Range("L" & r).Value = FriendlyOwnerLabel( _
                    VariantToString(t("ToOwnerType")), _
                    VariantToString(t("ToOwnerCompanyName")))

                ws.Range("M" & r).Value = FieldOrEmpty(t, "ToOwnerCounty")
                ws.Range("N" & r).Value = "OK"

            Else
                ws.Range("N" & r).Value = VariantToString(t("Status"))
            End If

            r = r + 1

        Next i

    End If

    '----------------------------------------------------------
    ' Bakgrunn, uthevet "naaverende registrering" og kantlinjer
    '----------------------------------------------------------

    With ws.Range("B" & startRow & ":N" & lastRow)
        .Interior.Color = blockFillColor
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("H" & startRow)
        .Font.Bold = True
        .Font.Color = ownerFontColor
        .Interior.Color = ownerColor
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("B" & startRow & ":N" & lastRow).Borders
        .LineStyle = xlContinuous
        .Color = RGB(221, 226, 243)
        .Weight = xlThin
    End With

    With ws.Range("B" & startRow & ":N" & lastRow) _
        .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(150, 150, 150)
        .Weight = xlMedium
    End With

    WriteVehicleBlock = lastRow + 1

End Function


Private Sub ApplyReportColumnWidths(ByVal ws As Worksheet)

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 15
    ws.Columns("H").ColumnWidth = 24
    ws.Columns("I").ColumnWidth = 13
    ws.Columns("J").ColumnWidth = 20
    ws.Columns("K").ColumnWidth = 22
    ws.Columns("L").ColumnWidth = 22
    ws.Columns("M").ColumnWidth = 16
    ws.Columns("N").ColumnWidth = 26

End Sub


Private Function FieldOrEmpty( _
    ByVal dict As Object, ByVal key As String) As String

    If dict Is Nothing Then
        FieldOrEmpty = vbNullString
        Exit Function
    End If

    If dict.Exists(key) Then
        FieldOrEmpty = VariantToString(dict(key))
    Else
        FieldOrEmpty = vbNullString
    End If

End Function


Private Function FirstNonEmpty( _
    ByVal a As String, ByVal b As String) As String

    If Len(Trim$(a)) > 0 Then
        FirstNonEmpty = a
    Else
        FirstNonEmpty = b
    End If

End Function


Private Function FriendlyOwnerLabel( _
    ByVal ownerType As String, ByVal companyName As String) As String

    If Len(Trim$(ownerType)) = 0 And Len(Trim$(companyName)) = 0 Then
        FriendlyOwnerLabel = ""
    ElseIf ownerType = "Privat" Then
        FriendlyOwnerLabel = "Privatperson"
    ElseIf Len(Trim$(companyName)) > 0 Then
        FriendlyOwnerLabel = companyName
    Else
        FriendlyOwnerLabel = ownerType
    End If

End Function


'==============================================================
' TEKST OG IDENTIFIKATORER
'==============================================================

Private Function NormalizeIdentifier(ByVal value As Variant) As String

    NormalizeIdentifier = UCase$(Trim$(Replace( _
        CStr(value & vbNullString), " ", vbNullString)))

End Function


Private Function BuildVehicleKey( _
    ByVal vin As String, ByVal regNo As String) As String

    If Len(vin) > 0 Then
        BuildVehicleKey = "VIN|" & vin
    ElseIf Len(regNo) > 0 Then
        BuildVehicleKey = "REG|" & regNo
    Else
        BuildVehicleKey = vbNullString
    End If

End Function


Private Function VariantToString(ByVal value As Variant) As String

    If IsNull(value) Or IsEmpty(value) Then
        VariantToString = vbNullString
    Else
        VariantToString = CStr(value)
    End If

End Function


Private Function JsonEscape(ByVal value As String) As String

    Dim result As String

    result = Replace(value, "\", "\\")
    result = Replace(result, Chr$(34), "\" & Chr$(34))

    JsonEscape = result

End Function


Private Function DateFromISO(ByVal isoDate As String) As Variant

    If Len(isoDate) < 10 Then
        DateFromISO = Empty
        Exit Function
    End If

    On Error GoTo InvalidDate

    DateFromISO = DateSerial( _
        CInt(Mid$(isoDate, 1, 4)), _
        CInt(Mid$(isoDate, 6, 2)), _
        CInt(Mid$(isoDate, 9, 2)))

    Exit Function

InvalidDate:
    DateFromISO = Empty

End Function


'==============================================================
' JSON (enkel tekstbasert parser - ingen eksterne referanser)
'==============================================================

Private Function JSON_FindMatchingBrace( _
    ByVal json As String, ByVal openPosition As Long) As Long

    Dim openCharacter As String
    Dim closeCharacter As String
    Dim currentCharacter As String
    Dim depth As Long
    Dim i As Long
    Dim insideString As Boolean

    openCharacter = Mid$(json, openPosition, 1)

    If openCharacter = "{" Then
        closeCharacter = "}"
    ElseIf openCharacter = "[" Then
        closeCharacter = "]"
    Else
        Exit Function
    End If

    For i = openPosition To Len(json)

        currentCharacter = Mid$(json, i, 1)

        If insideString Then

            If currentCharacter = "\" Then
                i = i + 1
            ElseIf currentCharacter = Chr$(34) Then
                insideString = False
            End If

        Else

            If currentCharacter = Chr$(34) Then
                insideString = True
            ElseIf currentCharacter = openCharacter Then
                depth = depth + 1
            ElseIf currentCharacter = closeCharacter Then

                depth = depth - 1

                If depth = 0 Then
                    JSON_FindMatchingBrace = i
                    Exit Function
                End If

            End If

        End If

    Next i

End Function


Private Function JSON_SkipWhitespace( _
    ByVal json As String, ByVal position As Long) As Long

    Dim p As Long
    Dim character As String

    p = position

    Do While p <= Len(json)

        character = Mid$(json, p, 1)

        If character = " " Or character = vbLf Or _
           character = vbCr Or character = vbTab Then
            p = p + 1
        Else
            Exit Do
        End If

    Loop

    JSON_SkipWhitespace = p

End Function


Private Function JSON_ExtractObject( _
    ByVal json As String, ByVal key As String) As String

    Dim keyPosition As Long
    Dim valuePosition As Long
    Dim endPosition As Long
    Dim firstCharacter As String

    If Len(json) = 0 Then Exit Function

    keyPosition = InStr( _
        1, json, Chr$(34) & key & Chr$(34) & ":", vbBinaryCompare)

    If keyPosition = 0 Then Exit Function

    valuePosition = keyPosition + Len(key) + 3
    valuePosition = JSON_SkipWhitespace(json, valuePosition)

    firstCharacter = Mid$(json, valuePosition, 1)

    If firstCharacter <> "{" And firstCharacter <> "[" Then Exit Function

    endPosition = JSON_FindMatchingBrace(json, valuePosition)

    If endPosition = 0 Then Exit Function

    JSON_ExtractObject = Mid$( _
        json, valuePosition, endPosition - valuePosition + 1)

End Function


Private Function JSON_ExtractValue( _
    ByVal json As String, ByVal key As String) As Variant

    Dim keyPosition As Long
    Dim valuePosition As Long
    Dim endPosition As Long
    Dim i As Long
    Dim j As Long

    Dim firstCharacter As String
    Dim result As String
    Dim rawValue As String
    Dim character As String
    Dim nextCharacter As String

    JSON_ExtractValue = Null

    If Len(json) = 0 Then Exit Function

    keyPosition = InStr( _
        1, json, Chr$(34) & key & Chr$(34) & ":", vbBinaryCompare)

    If keyPosition = 0 Then Exit Function

    valuePosition = keyPosition + Len(key) + 3
    valuePosition = JSON_SkipWhitespace(json, valuePosition)

    firstCharacter = Mid$(json, valuePosition, 1)

    If firstCharacter = Chr$(34) Then

        i = valuePosition + 1

        Do While i <= Len(json)

            character = Mid$(json, i, 1)

            If character = "\" Then

                nextCharacter = Mid$(json, i + 1, 1)

                Select Case nextCharacter
                    Case "n"
                        result = result & vbLf
                    Case "r"
                        result = result & vbCr
                    Case "t"
                        result = result & vbTab
                    Case Chr$(34)
                        result = result & Chr$(34)
                    Case "\"
                        result = result & "\"
                    Case Else
                        result = result & nextCharacter
                End Select

                i = i + 2

            ElseIf character = Chr$(34) Then
                Exit Do
            Else
                result = result & character
                i = i + 1
            End If

        Loop

        JSON_ExtractValue = result

    ElseIf firstCharacter = "{" Or firstCharacter = "[" Then

        endPosition = JSON_FindMatchingBrace(json, valuePosition)

        If endPosition > 0 Then
            JSON_ExtractValue = Mid$( _
                json, valuePosition, endPosition - valuePosition + 1)
        End If

    Else

        j = valuePosition

        Do While j <= Len(json)

            character = Mid$(json, j, 1)

            If character = "," Or character = "}" Or _
               character = "]" Then Exit Do

            rawValue = rawValue & character
            j = j + 1

        Loop

        rawValue = Trim$(rawValue)

        Select Case LCase$(rawValue)

            Case "true"
                JSON_ExtractValue = True

            Case "false"
                JSON_ExtractValue = False

            Case "null"
                JSON_ExtractValue = Null

            Case Else

                If IsNumeric(rawValue) Then
                    JSON_ExtractValue = CDbl(rawValue)
                Else
                    JSON_ExtractValue = rawValue
                End If

        End Select

    End If

End Function


Private Function JSON_ArrayAllElements( _
    ByVal jsonArray As String) As Collection

    Dim result As New Collection
    Dim position As Long
    Dim endPosition As Long
    Dim i As Long

    Dim firstCharacter As String
    Dim element As String
    Dim rawValue As String
    Dim character As String

    If Len(jsonArray) < 2 Then
        Set JSON_ArrayAllElements = result
        Exit Function
    End If

    If Left$(jsonArray, 1) <> "[" Then
        Set JSON_ArrayAllElements = result
        Exit Function
    End If

    position = JSON_SkipWhitespace(jsonArray, 2)

    Do While position <= Len(jsonArray)

        If Mid$(jsonArray, position, 1) = "]" Then Exit Do

        firstCharacter = Mid$(jsonArray, position, 1)
        element = vbNullString

        If firstCharacter = "{" Or firstCharacter = "[" Then

            endPosition = JSON_FindMatchingBrace(jsonArray, position)

            If endPosition = 0 Then Exit Do

            element = Mid$( _
                jsonArray, position, endPosition - position + 1)

            position = endPosition + 1

        Else

            i = position
            rawValue = vbNullString

            Do While i <= Len(jsonArray)

                character = Mid$(jsonArray, i, 1)

                If character = "," Or character = "]" Then Exit Do

                rawValue = rawValue & character
                i = i + 1

            Loop

            element = Trim$(rawValue)
            position = i

        End If

        result.Add element

        position = JSON_SkipWhitespace(jsonArray, position)

        If position <= Len(jsonArray) Then
            If Mid$(jsonArray, position, 1) = "," Then
                position = JSON_SkipWhitespace(jsonArray, position + 1)
            End If
        End If

    Loop

    Set JSON_ArrayAllElements = result

End Function
