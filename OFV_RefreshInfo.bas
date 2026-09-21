Option Explicit

' Produksjonsmakro for prosjektet "Python program for bilregistrering":
' Input-arket (KjoretoyInput), Resultat-arket (Transaksjoner-tabellen),
' Oversikt (KPI-er + pivot + regnr/forstegangsreg-liste) og
' Kontroll solgte biler oppdateres alle fra samme OFV-datahenting.
'
' Kort om oppsettet:
'   - Bruker kun OFV Transactions API. Statens Vegvesen (SVV/"VVS") er
'     ikke i bruk noe sted i denne filen.
'   - HTTP-kallene bruker WinHttp.WinHttpRequest.5.1 (samme klient som
'     testscriptet som lostte lagringsproblemene tidligere).
'   - OFV_URL = https://api.ofv.no/transactions/v1/, bekreftet via
'     "Try it"-konsollen i Azure APIM-portalen (se kommentar ved
'     konstanten). Kallet henter ALLE transaksjoner for hvert regnr/VIN
'     - uten datofilter - sa firstRegistrationDate garantert finnes sa
'     lenge OFV har minst en transaksjon noen gang for kjoretoyet.
'     Fra dato/Til dato (Input!B2:B3) brukes kun til a avgrense hvilken
'     transaksjon som regnes som "naermeste eierskifte" i Kontroll
'     solgte biler - Resultat/Transaksjoner-tabellen viser hele
'     historikken, og kan filtreres pa TransaksjonsDato med det
'     innebygde Excel-autofilteret.
'   - KjoretoyInput-tabellen pa Input (bygges automatisk om den
'     mangler) har na 4 kolonner: Regnr, VIN, Bokfort (dato du fyller
'     inn manuelt) og Forstegangsregistrering (skrives tilbake av
'     makroen etter hver kjoring).

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
#End If

'==============================================================
' KONFIGURASJON
'==============================================================

Private Const INPUT_SHEET As String = "Input"
Private Const RESULT_SHEET As String = "Resultat"
Private Const OVERVIEW_SHEET As String = "Oversikt"
Private Const CONTROL_SHEET As String = "Kontroll solgte biler"

Private Const INPUT_TABLE As String = "KjoretoyInput"
Private Const RESULT_TABLE As String = "Transaksjoner"
Private Const PIVOT_NAME As String = "TransaksjonsPivot"

Private Const FIRST_ROW As Long = 5
Private Const COL_REGNR As Long = 2
Private Const COL_VIN As Long = 3
Private Const COL_BOKFORT As Long = 4
Private Const COL_FORSTEREG As Long = 5

' Bekreftet via "Try it"-konsollen i Azure APIM-portalen
' (https://data.ofv.no/api-details#api=transactions-api-v1&operation=query-transactions):
' POST https://api.ofv.no/transactions/v1/ - dette er den faktiske
' verten APIet ruter pa, uavhengig av at portalens eget domene er
' data.ofv.no.
Private Const OFV_URL As String = _
    "https://api.ofv.no/transactions/v1/"

Private Const MAX_RETRIES As Long = 4
Private Const RETRY_WAIT_MS As Long = 3000
Private Const API_PAUSE_MS As Long = 150

' Fargekoder for "Dager avvik" i Kontroll solgte biler
' (samme paletter som Excels innebygde Bra/Noytral/Darlig-stiler).
Private Const COLOR_GREEN_FILL As Long = 13561798   ' RGB(198,239,206)
Private Const COLOR_GREEN_FONT As Long = 24832      ' RGB(0,97,0)
Private Const COLOR_YELLOW_FILL As Long = 10284031  ' RGB(255,235,156)
Private Const COLOR_YELLOW_FONT As Long = 26012     ' RGB(156,101,0)
Private Const COLOR_RED_FILL As Long = 13551615     ' RGB(255,199,206)
Private Const COLOR_RED_FONT As Long = 393372       ' RGB(156,0,6)


'==============================================================
' HOVEDMAKRO
'==============================================================

Public Sub OFV_RefreshInfo()

    Dim wsInput As Worksheet
    Dim wsResult As Worksheet
    Dim wsOverview As Worksheet
    Dim wsControl As Worksheet
    Dim loResult As ListObject

    Dim queue As Object
    Dim hitVehicles As Object
    Dim vehicleRowsByKey As Object
    Dim firstRegByKey As Object
    Dim resultRow As Object
    Dim kontrollRow As Object

    Dim allRows As Collection
    Dim vehicleRows As Collection
    Dim kontrollRows As Collection
    Dim vehicleTxRows As Collection

    Dim fieldMap As Variant
    Dim vehicleData As Variant
    Dim bokfortValue As Variant
    Dim output() As Variant

    Dim ofvKey As String
    Dim dateFrom As Variant
    Dim dateTo As Variant
    Dim stage As String

    Dim lastRegRow As Long
    Dim lastVinRow As Long
    Dim lastInputRow As Long
    Dim oldLastRow As Long
    Dim newLastRow As Long
    Dim fieldCount As Long
    Dim outputRows As Long

    Dim transactionCount As Long
    Dim noTransactionCount As Long
    Dim ofvErrorCount As Long

    Dim r As Long
    Dim c As Long
    Dim currentVehicle As Long
    Dim totalVehicles As Long

    Dim regNo As String
    Dim vin As String
    Dim queueKey As String
    Dim identifier As String
    Dim dictionaryKey As String
    Dim statusText As String

    Dim useVin As Boolean
    Dim key As Variant
    Dim value As Variant

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldCursor As Variant
    Dim applicationChanged As Boolean

    On Error GoTo FatalError

    stage = "finner arkene"
    API_ShowStatus "Forbereder", stage

    Set wsInput = GetRequiredSheet(ThisWorkbook, INPUT_SHEET)
    Set wsResult = GetRequiredSheet(ThisWorkbook, RESULT_SHEET)
    Set wsOverview = GetRequiredSheet(ThisWorkbook, OVERVIEW_SHEET)
    Set wsControl = GetRequiredSheet(ThisWorkbook, CONTROL_SHEET)

    If wsResult.ProtectContents Then
        Err.Raise vbObjectError + 1000, , _
            "Resultat-arket er beskyttet."
    End If

    If wsControl.ProtectContents Then
        Err.Raise vbObjectError + 1001, , _
            "Kontroll solgte biler er beskyttet."
    End If

    stage = "oppretter/kontrollerer KjoretoyInput-tabellen"
    API_ShowStatus "Forbereder", stage

    EnsureInputTable wsInput

    stage = "leser API-nokkel"
    API_ShowStatus "Forbereder", stage

    ' Bruker det navngitte omradet OFV_API hvis det finnes i
    ' arbeidsboken, ellers leses nokkelen direkte fra Input!A1.
    ofvKey = Trim$(CStr( _
        ReadConfigValue( _
            ThisWorkbook, "OFV_API", wsInput.Range("A1"))))

    If Len(ofvKey) = 0 Then
        MsgBox _
            "Fant ingen OFV-nokkel. Legg den enten i det " & _
            "navngitte omradet OFV_API, eller direkte i " & _
            "celle A1 pa arket " & INPUT_SHEET & ".", _
            vbExclamation, "API-oppdatering"
        GoTo SafeExit
    End If

    stage = "leser datoperioden"
    API_ShowStatus "Forbereder", stage

    ' Samme prinsipp for datoperioden: navngitt omrade hvis det
    ' finnes, ellers Input!B2 (fra-dato) og Input!B3 (til-dato).
    dateFrom = ReadConfigValue( _
        ThisWorkbook, "OFV_DateFrom", wsInput.Range("B2"))

    dateTo = ReadConfigValue( _
        ThisWorkbook, "OFV_DateTo", wsInput.Range("B3"))

    If Not IsDate(dateFrom) Or Not IsDate(dateTo) Then
        MsgBox "Fyll inn gyldige datoer i Input!B2:B3.", _
            vbExclamation, "API-oppdatering"
        GoTo SafeExit
    End If

    If CDate(dateFrom) > CDate(dateTo) Then
        MsgBox "Fra-dato kan ikke vaere senere enn til-dato.", _
            vbExclamation, "API-oppdatering"
        GoTo SafeExit
    End If

    stage = "leser kjoretoylisten"
    API_ShowStatus "Forbereder", stage

    lastRegRow = wsInput.Cells( _
        wsInput.rows.Count, COL_REGNR).End(xlUp).Row

    lastVinRow = wsInput.Cells( _
        wsInput.rows.Count, COL_VIN).End(xlUp).Row

    lastInputRow = Application.Max(lastRegRow, lastVinRow)

    Set queue = CreateObject("Scripting.Dictionary")
    queue.CompareMode = vbTextCompare

    For r = FIRST_ROW To lastInputRow

        regNo = NormalizeIdentifier( _
            wsInput.Cells(r, COL_REGNR).value)

        vin = NormalizeIdentifier( _
            wsInput.Cells(r, COL_VIN).value)

        bokfortValue = wsInput.Cells(r, COL_BOKFORT).value

        queueKey = BuildVehicleKey(vin, regNo)

        If Len(queueKey) > 0 Then
            If Not queue.Exists(queueKey) Then
                queue.Add queueKey, Array(regNo, vin, bokfortValue)
            End If
        End If

    Next r

    If queue.Count = 0 Then
        MsgBox "Fant ingen registreringsnummer eller VIN.", _
            vbInformation, "API-oppdatering"
        GoTo SafeExit
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation
    oldCursor = Application.cursor

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.cursor = xlWait
    applicationChanged = True

    Set allRows = New Collection

    Set hitVehicles = _
        CreateObject("Scripting.Dictionary")

    hitVehicles.CompareMode = vbTextCompare

    Set vehicleRowsByKey = _
        CreateObject("Scripting.Dictionary")

    vehicleRowsByKey.CompareMode = vbTextCompare

    Set kontrollRows = New Collection

    Set firstRegByKey = CreateObject("Scripting.Dictionary")
    firstRegByKey.CompareMode = vbTextCompare

    fieldMap = GetFieldMap()
    fieldCount = UBound(fieldMap) + 1
    totalVehicles = queue.Count

    '==========================================================
    ' HENT DATA (kun OFV)
    '==========================================================

    For Each key In queue.Keys

        currentVehicle = currentVehicle + 1
        vehicleData = queue(key)

        regNo = CStr(vehicleData(0))
        vin = CStr(vehicleData(1))

        useVin = (Len(vin) > 0)

        If useVin Then
            identifier = vin
        Else
            identifier = regNo
        End If

        stage = "henter eierskifter fra OFV"

        API_ShowStatus _
            "OFV", _
            "Eierskifter", _
            identifier, _
            currentVehicle, _
            totalVehicles

        Set vehicleRows = FetchOFVTransactions( _
            ofvKey, _
            identifier, _
            useVin, _
            regNo, _
            vin)

        If Not vehicleRowsByKey.Exists(CStr(key)) Then
            vehicleRowsByKey.Add CStr(key), New Collection
        End If

        For Each resultRow In vehicleRows

            statusText = VariantToString( _
                resultRow("Status"))

            If statusText = "OK" Then

                transactionCount = transactionCount + 1

                If Not hitVehicles.Exists(CStr(key)) Then
                    hitVehicles.Add CStr(key), True
                End If

            ElseIf Left$(statusText, 5) = "Feil:" Then

                ofvErrorCount = ofvErrorCount + 1

            Else

                noTransactionCount = _
                    noTransactionCount + 1

            End If

            allRows.Add resultRow
            vehicleRowsByKey(CStr(key)).Add resultRow

        Next resultRow

        API_ShowStatus _
            "Fullfort", _
            "OFV behandlet", _
            identifier, _
            currentVehicle, _
            totalVehicles

        Sleep API_PAUSE_MS

    Next key

    '==========================================================
    ' KONTROLL SOLGTE BILER - en rad per kjoretoy
    '==========================================================

    For Each key In queue.Keys

        vehicleData = queue(key)

        Set vehicleTxRows = Nothing

        If vehicleRowsByKey.Exists(CStr(key)) Then
            Set vehicleTxRows = vehicleRowsByKey(CStr(key))
        End If

        Set kontrollRow = BuildKontrollRow( _
            CStr(vehicleData(0)), _
            CStr(vehicleData(1)), _
            vehicleData(2), _
            vehicleTxRows, _
            CDate(dateFrom), _
            CDate(dateTo))

        kontrollRows.Add kontrollRow

        firstRegByKey(CStr(key)) = kontrollRow("Forstegangsregistrert")

    Next key

    stage = "skriver forstegangsregistrering til Input"
    API_ShowStatus "Excel", stage

    WriteFirstRegistrationToInput wsInput, firstRegByKey

    '==========================================================
    ' RESULTAT
    '==========================================================

    stage = "oppdaterer Resultat"
    API_ShowStatus "Excel", "Oppdaterer Resultat"

    outputRows = allRows.Count

    Set loResult = GetOrCreateResultTable( _
        wsResult, fieldMap)

    oldLastRow = _
        loResult.Range.Row + _
        loResult.Range.rows.Count - 1

    If outputRows > 0 Then
        newLastRow = outputRows + 1
    Else
        newLastRow = 2
    End If

    Set loResult = ResizeResultTable( _
        wsResult, loResult, newLastRow, fieldCount)

    If Not loResult.DataBodyRange Is Nothing Then
        loResult.DataBodyRange.ClearContents
    End If

    If oldLastRow > newLastRow Then

        With wsResult.Range( _
            wsResult.Cells(newLastRow + 1, 1), _
            wsResult.Cells(oldLastRow, fieldCount))

            .ClearContents
            .ClearFormats

        End With

    End If

    For c = LBound(fieldMap) To UBound(fieldMap)

        loResult.HeaderRowRange.Cells( _
            1, c + 1).value = fieldMap(c)(1)

    Next c

    If outputRows > 0 Then

        ReDim output( _
            1 To outputRows, _
            1 To fieldCount)

        For r = 1 To outputRows

            Set resultRow = allRows(r)

            For c = LBound(fieldMap) To UBound(fieldMap)

                dictionaryKey = CStr(fieldMap(c)(0))

                If dictionaryKey = _
                    "CalculatedSellerType" Or _
                   dictionaryKey = _
                    "CalculatedBuyerType" Then

                    output(r, c + 1) = Empty

                ElseIf resultRow.Exists( _
                    dictionaryKey) Then

                    value = resultRow(dictionaryKey)

                    If IsNull(value) Or IsEmpty(value) Then
                        output(r, c + 1) = Empty
                    Else
                        output(r, c + 1) = value
                    End If

                Else
                    output(r, c + 1) = Empty
                End If

            Next c

        Next r

        loResult.DataBodyRange.value = output
        ApplyCalculatedColumns loResult

    End If

    FormatResultTable wsResult, loResult
    FormatResultTableGrouping wsResult, loResult, allRows

    '==========================================================
    ' OVERSIKT
    '==========================================================

    stage = "oppdaterer Oversikt"
    API_ShowStatus "Excel", "Oppdaterer Oversikt"

    UpdateOverviewKPIs wsOverview
    RebuildOverviewPivot wsOverview, loResult
    WriteFirstRegistrationOverviewList wsOverview, kontrollRows

    '==========================================================
    ' KONTROLLARK
    '==========================================================

    stage = "oppdaterer Kontroll solgte biler"

    API_ShowStatus _
        "Excel", _
        "Oppdaterer Kontroll solgte biler"

    UpdateControlSheet _
        wsControl, kontrollRows, totalVehicles, _
        CDate(dateFrom), CDate(dateTo)

    Application.Calculation = oldCalculation

    If oldCalculation = xlCalculationManual Then

        wsResult.Calculate
        wsOverview.Calculate
        wsControl.Calculate

    Else
        Application.CalculateFull
    End If

    API_ShowStatus _
        "Ferdig", _
        "Alle data og ark er oppdatert"

    RestoreApplicationState _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldCalculation, _
        oldCursor

    applicationChanged = False

    MsgBox _
        "Oppdateringen er ferdig." & vbCrLf & vbCrLf & _
        totalVehicles & " kjoretoy lest." & vbCrLf & _
        transactionCount & _
        " OFV-eierskifter funnet." & vbCrLf & _
        noTransactionCount & _
        " uten OFV-eierskifter i perioden." & vbCrLf & _
        ofvErrorCount & " OFV-feil.", _
        vbInformation, "API-oppdatering"

    Exit Sub

SafeExit:

    Application.StatusBar = False
    Exit Sub

FatalError:

    Dim errorNumber As Long
    Dim errorDescription As String

    errorNumber = Err.Number
    errorDescription = Err.Description

    Application.StatusBar = False

    If applicationChanged Then

        RestoreApplicationState _
            oldScreenUpdating, _
            oldEnableEvents, _
            oldCalculation, _
            oldCursor

    End If

    MsgBox _
        "Oppdateringen ble avbrutt." & vbCrLf & vbCrLf & _
        "Trinn: " & stage & vbCrLf & _
        "Feil " & errorNumber & ": " & _
        errorDescription, _
        vbCritical, "API-oppdatering"

End Sub


'==============================================================
' STATUSLINJE
'==============================================================

Private Sub API_ShowStatus( _
    ByVal sourceName As String, _
    ByVal activity As String, _
    Optional ByVal identifier As String = "", _
    Optional ByVal currentVehicle As Long = 0, _
    Optional ByVal totalVehicles As Long = 0)

    Dim message As String
    Dim percentage As Double

    message = sourceName & " | " & activity

    If currentVehicle > 0 And totalVehicles > 0 Then

        percentage = currentVehicle / totalVehicles

        message = message & _
            " | " & currentVehicle & _
            " av " & totalVehicles

        message = message & _
            " | " & Format$(percentage, "0%")

    End If

    If Len(identifier) > 0 Then
        message = message & " | " & identifier
    End If

    Application.StatusBar = message
    DoEvents

End Sub


Private Sub RestoreApplicationState( _
    ByVal screenUpdatingValue As Boolean, _
    ByVal enableEventsValue As Boolean, _
    ByVal calculationValue As XlCalculation, _
    ByVal cursorValue As Variant)

    Application.StatusBar = False
    Application.cursor = cursorValue
    Application.ScreenUpdating = screenUpdatingValue
    Application.EnableEvents = enableEventsValue
    Application.Calculation = calculationValue

End Sub


'==============================================================
' OFV-TRANSAKSJONER
'==============================================================

' Henter ALLE transaksjoner for kjoretoyet - ikke bare de innenfor
' Fra dato/Til dato. Dette gir ett kall per kjoretoy (ikke to), og
' garanterer at firstRegistrationDate blir funnet sa lenge OFV har
' minst en transaksjon noen gang for kjoretoyet, uavhengig av om den
' ligger innenfor perioden. Periodeavgrensningen for "naermeste
' eierskifte" gjores i BuildKontrollRow, og for Transaksjoner-tabellen
' kan brukeren selv filtrere pa TransaksjonsDato med det innebygde
' Excel-autofilteret.
Private Function FetchOFVTransactions( _
    ByVal apiKey As String, _
    ByVal identifier As String, _
    ByVal useVin As Boolean, _
    ByVal originalRegNo As String, _
    ByVal originalVin As String) As Collection

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
        body = body & JsonEscape(identifier) & """},"
        body = body & """pagination"":{""first"":1000"

        If Len(cursor) > 0 Then
            body = body & ",""cursor"":"""
            body = body & JsonEscape(cursor) & """"
        End If

        body = body & "},"
        body = body & """sorting"":{"
        body = body & """orderBy"":""transactionDate"","
        body = body & """orderDirection"":""DESC""}}"

        responseText = PostOFVWithRetries( _
            apiKey, body, statusText)

        If statusText <> "OK" Then

            rows.Add BuildEmptyTransactionRow( _
                identifier, useVin, originalRegNo, _
                originalVin, statusText)

            Set FetchOFVTransactions = rows
            Exit Function

        End If

        transactionsJSON = JSON_ExtractObject( _
            responseText, "transactions")

        Set items = JSON_ArrayAllElements( _
            transactionsJSON)

        For Each item In items

            rows.Add BuildTransactionRow( _
                identifier, CStr(item), useVin, _
                originalRegNo, originalVin)

        Next item

        paginationJSON = JSON_ExtractObject( _
            responseText, "pagination")

        hasNextValue = JSON_ExtractValue( _
            paginationJSON, "hasNextPage")

        cursorValue = JSON_ExtractValue( _
            paginationJSON, "endCursor")

        hasNext = False

        Select Case VarType(hasNextValue)

            Case vbBoolean
                hasNext = CBool(hasNextValue)

            Case vbString
                hasNext = _
                    (LCase$(CStr(hasNextValue)) = "true")

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

            Sleep API_PAUSE_MS

        Else
            Exit Do
        End If

    Loop

    If rows.Count = 0 Then

        rows.Add BuildEmptyTransactionRow( _
            identifier, useVin, originalRegNo, _
            originalVin, _
            "Ingen eierskifter i perioden")

    End If

    Set FetchOFVTransactions = rows

End Function


Private Function BuildEmptyTransactionRow( _
    ByVal identifier As String, _
    ByVal useVin As Boolean, _
    ByVal originalRegNo As String, _
    ByVal originalVin As String, _
    ByVal statusText As String) As Object

    Dim result As Object

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    result("Input") = identifier

    If useVin Then
        result("Kilde") = "VIN"
    Else
        result("Kilde") = "Regnr"
    End If

    result("RegNo") = originalRegNo
    result("ChassisNumber") = originalVin
    result("Status") = statusText

    Set BuildEmptyTransactionRow = result

End Function


Private Function BuildTransactionRow( _
    ByVal identifier As String, _
    ByVal transactionJSON As String, _
    ByVal useVin As Boolean, _
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

    result("Input") = identifier

    If useVin Then
        result("Kilde") = "VIN"
    Else
        result("Kilde") = "Regnr"
    End If

    extractedValue = JSON_ExtractValue( _
        transactionJSON, "regNo")

    If Len(VariantToString(extractedValue)) > 0 Then
        result("RegNo") = extractedValue
    Else
        result("RegNo") = originalRegNo
    End If

    extractedValue = JSON_ExtractValue( _
        transactionJSON, "chassisNumber")

    If Len(VariantToString(extractedValue)) > 0 Then
        result("ChassisNumber") = extractedValue
    Else
        result("ChassisNumber") = originalVin
    End If

    result("MakeName") = _
        JSON_ExtractValue(transactionJSON, "makeName")

    result("ModelName") = _
        JSON_ExtractValue(transactionJSON, "modelName")

    result("RegistrationType") = _
        JSON_ExtractValue(transactionJSON, "registrationType")

    result("FuelGroup") = _
        JSON_ExtractValue(transactionJSON, "fuelGroup")

    result("IsLeased") = _
        JSON_ExtractValue(transactionJSON, "isLeased")

    result("IsUsedImported") = _
        JSON_ExtractValue(transactionJSON, "isUsedImported")

    ' Forstegangsregistrering kommer direkte fra OFV sin egen
    ' transaksjon - ikke fra en egen Vegvesen-oppslag lenger.
    result("FirstRegistrationDate") = _
        DateFromISO(VariantToString( _
            JSON_ExtractValue( _
                transactionJSON, "firstRegistrationDate")))

    result("TransactionNumber") = _
        JSON_ExtractValue(transactionJSON, "transactionNumber")

    result("TransactionDate") = _
        DateFromISO(VariantToString( _
            JSON_ExtractValue( _
                transactionJSON, "transactionDate")))

    fromJSON = JSON_ExtractObject(transactionJSON, "from")
    toJSON = JSON_ExtractObject(transactionJSON, "to")

    fromOwnerJSON = JSON_ExtractObject(fromJSON, "owner")
    toOwnerJSON = JSON_ExtractObject(toJSON, "owner")

    result("FromOwnerType") = _
        JSON_ExtractValue(fromOwnerJSON, "type")

    companyJSON = _
        JSON_ExtractObject(fromOwnerJSON, "companyInfo")

    result("FromOwnerCompanyName") = _
        JSON_ExtractValue(companyJSON, "name")

    result("FromOwnerCounty") = _
        JSON_ExtractValue(fromOwnerJSON, "countyName")

    result("ToOwnerType") = _
        JSON_ExtractValue(toOwnerJSON, "type")

    companyJSON = _
        JSON_ExtractObject(toOwnerJSON, "companyInfo")

    result("ToOwnerCompanyName") = _
        JSON_ExtractValue(companyJSON, "name")

    result("ToOwnerCounty") = _
        JSON_ExtractValue(toOwnerJSON, "countyName")

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

        Set http = CreateObject( _
            "WinHttp.WinHttpRequest.5.1")

        statusCode = 0
        responseText = vbNullString

        On Error Resume Next

        http.SetTimeouts 10000, 10000, 30000, 30000

        http.Open "POST", OFV_URL, False

        http.SetRequestHeader _
            "Content-Type", "application/json"

        http.SetRequestHeader _
            "Cache-Control", "no-cache"

        http.SetRequestHeader _
            "Ocp-Apim-Subscription-Key", apiKey

        http.Send body

        statusCode = http.Status
        responseText = http.ResponseText

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
                statusText = _
                    "Feil: OFV 401 - kontroller API-nokkelen"
                Exit Function

            Case 403
                statusText = _
                    "Feil: OFV 403 - tilgang eller kvote"
                Exit Function

            Case 404
                statusText = _
                    "Feil: OFV-endepunktet ble ikke funnet"
                Exit Function

            Case 429, 500, 502, 503, 504

                lastError = _
                    CStr(statusCode) & ": " & responseText

                Sleep RETRY_WAIT_MS * attempt

            Case Else

                If statusCode <> 0 Then

                    statusText = _
                        "Feil: OFV HTTP " & statusCode

                    Exit Function

                Else
                    Sleep RETRY_WAIT_MS * attempt
                End If

        End Select

    Next attempt

    statusText = "Feil: OFV - " & lastError

End Function


'==============================================================
' FELTMAPPING - 22 KOLONNER A:V
'==============================================================

Private Function GetFieldMap() As Variant

    Dim fields(0 To 21) As Variant

    fields(0) = Array("Input", "Input", False)
    fields(1) = Array("Kilde", "Kilde", False)
    fields(2) = Array("RegNo", "RegNo", False)
    fields(3) = Array( _
        "ChassisNumber", "Chassisnummer", False)
    fields(4) = Array("MakeName", "Merke", False)
    fields(5) = Array("ModelName", "Modell", False)
    fields(6) = Array( _
        "RegistrationType", "RegistreringsType", False)
    fields(7) = Array( _
        "FuelGroup", "Drivstoffgruppe", False)
    fields(8) = Array("IsLeased", "Leaset", False)
    fields(9) = Array( _
        "IsUsedImported", "Bruktimportert", False)
    fields(10) = Array( _
        "FirstRegistrationDate", _
        "ForstegangsRegistrering", True)
    fields(11) = Array( _
        "TransactionNumber", _
        "TransaksjonsNummer", False)
    fields(12) = Array( _
        "TransactionDate", _
        "Eierskiftedato", True)
    fields(13) = Array( _
        "CalculatedSellerType", _
        "SelgerType", False)
    fields(14) = Array( _
        "CalculatedBuyerType", _
        BuyerTypeHeader(), False)
    fields(15) = Array( _
        "FromOwnerType", _
        "SelgerEierType", False)
    fields(16) = Array( _
        "FromOwnerCompanyName", _
        "SelgerEierFirma", False)
    fields(17) = Array( _
        "FromOwnerCounty", _
        "SelgerEierFylke", False)
    fields(18) = Array( _
        "ToOwnerType", _
        "KjoperEierType", False)
    fields(19) = Array( _
        "ToOwnerCompanyName", _
        "KjoperEierFirma", False)
    fields(20) = Array( _
        "ToOwnerCounty", BuyerCountyHeader(), False)
    fields(21) = Array("Status", "Status", False)

    GetFieldMap = fields

End Function


'==============================================================
' RESULTATTABELL
'==============================================================

Private Function GetOrCreateResultTable( _
    ByVal ws As Worksheet, _
    ByVal fieldMap As Variant) As ListObject

    Dim lo As ListObject
    Dim target As Range
    Dim fieldCount As Long
    Dim c As Long

    fieldCount = UBound(fieldMap) + 1

    On Error Resume Next
    Set lo = ws.ListObjects(RESULT_TABLE)
    On Error GoTo 0

    If lo Is Nothing Then

        For c = LBound(fieldMap) To UBound(fieldMap)
            ws.Cells(1, c + 1).value = fieldMap(c)(1)
        Next c

        Set target = ws.Range( _
            ws.Cells(1, 1), _
            ws.Cells(2, fieldCount))

        Set lo = ws.ListObjects.Add( _
            xlSrcRange, target, , xlYes)

        lo.Name = RESULT_TABLE
        lo.DataBodyRange.ClearContents

    End If

    Set GetOrCreateResultTable = lo

End Function


Private Function ResizeResultTable( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject, _
    ByVal lastRow As Long, _
    ByVal fieldCount As Long) As ListObject

    Dim target As Range
    Dim styleName As String

    Set target = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(lastRow, fieldCount))

    styleName = lo.TableStyle

    On Error Resume Next

    If ws.FilterMode Then ws.ShowAllData

    lo.AutoFilter.ShowAllData
    lo.ShowAutoFilter = False

    Err.Clear
    lo.Resize target

    If Err.Number <> 0 Then

        Err.Clear
        lo.Unlist

        Set lo = ws.ListObjects.Add( _
            xlSrcRange, target, , xlYes)

        lo.Name = RESULT_TABLE

    End If

    On Error GoTo 0

    lo.ShowAutoFilter = True

    If Len(styleName) > 0 Then
        lo.TableStyle = styleName
    Else
        lo.TableStyle = "TableStyleMedium2"
    End If

    Set ResizeResultTable = lo

End Function


Private Sub ApplyCalculatedColumns(ByVal lo As ListObject)

    If lo.DataBodyRange Is Nothing Then Exit Sub

    With lo.ListColumns("SelgerType").DataBodyRange

        .Formula = _
            "=IF(AND([@SelgerEierType]=""""," & _
            "[@SelgerEierFirma]=""""),""""," & _
            "IF([@SelgerEierType]=""Privat"",""Privat""," & _
            "IF([@SelgerEierFirma]<>""""," & _
            "[@SelgerEierFirma],[@SelgerEierType])))"

    End With

    With lo.ListColumns(BuyerTypeHeader()).DataBodyRange

        .Formula = _
            "=IF(AND([@KjoperEierType]=""""," & _
            "[@KjoperEierFirma]=""""),""""," & _
            "IF([@KjoperEierType]=""Privat"",""Privat""," & _
            "IF([@KjoperEierFirma]<>""""," & _
            "[@KjoperEierFirma],[@KjoperEierType])))"

    End With

End Sub


Private Sub FormatResultTable( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject)

    Dim lastRow As Long

    lastRow = lo.Range.Row + lo.Range.rows.Count - 1

    lo.TableStyle = "TableStyleMedium2"
    lo.ShowTableStyleRowStripes = False
    lo.ShowAutoFilter = True

    With lo.HeaderRowRange
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 34
    End With

    If lastRow >= 2 Then
        ws.Range("K2:K" & lastRow).NumberFormat = "dd.mm.yyyy"
        ws.Range("L2:L" & lastRow).NumberFormat = "0"
        ws.Range("M2:M" & lastRow).NumberFormat = "dd.mm.yyyy"
    End If

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 9
    ws.Columns("C").ColumnWidth = 11
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 11
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 24
    ws.Columns("H").ColumnWidth = 15
    ws.Columns("I:J").ColumnWidth = 12
    ws.Columns("K:M").ColumnWidth = 16
    ws.Columns("N:O").ColumnWidth = 25
    ws.Columns("P").ColumnWidth = 13
    ws.Columns("Q").ColumnWidth = 25
    ws.Columns("R:S").ColumnWidth = 13
    ws.Columns("T").ColumnWidth = 25
    ws.Columns("U").ColumnWidth = 15
    ws.Columns("V").ColumnWidth = 28

End Sub


' Gjor det tydelig hvor en bils transaksjoner slutter og neste
' begynner (topplinje ved skifte av "Input"-verdi), og uthever raden
' som Kontroll solgte biler faktisk brukte til kontrollen (merket
' ErKontrollMatch=True pa det samme delte objektet som ligger i
' allRows - se BuildKontrollRow).
Private Sub FormatResultTableGrouping( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject, _
    ByVal allRows As Collection)

    Dim firstDataRow As Long
    Dim r As Long
    Dim wsRow As Long
    Dim resultRow As Object
    Dim forrigeInput As String
    Dim denneInput As String

    If lo.DataBodyRange Is Nothing Then Exit Sub

    firstDataRow = lo.HeaderRowRange.Row + 1
    forrigeInput = vbNullString

    For r = 1 To allRows.Count

        Set resultRow = allRows(r)
        wsRow = firstDataRow + r - 1

        denneInput = VariantToString(resultRow("Input"))

        If r > 1 And denneInput <> forrigeInput Then

            With ws.Range( _
                ws.Cells(wsRow, 1), _
                ws.Cells(wsRow, lo.ListColumns.Count)).Borders(xlEdgeTop)

                .LineStyle = xlContinuous
                .Color = RGB(31, 78, 120)
                .Weight = xlMedium

            End With

        End If

        If resultRow.Exists("ErKontrollMatch") Then

            If resultRow("ErKontrollMatch") = True Then

                With ws.Range( _
                    ws.Cells(wsRow, 1), ws.Cells(wsRow, lo.ListColumns.Count))

                    .Font.Bold = True
                    .Interior.Color = RGB(238, 244, 251)

                End With

            End If

        End If

        forrigeInput = denneInput

    Next r

End Sub


'==============================================================
' KONTROLL SOLGTE BILER
'==============================================================

' Bygger hele arket pa nytt hver kjoring - tittel, forklaring av
' kontrollregelen, KPI-bokser, fargekodelegende, og en hovedrad per
' kjoretoy (sortert etter storst dagers avvik forst) etterfulgt av en
' kompakt detaljrad per ovrig eierskifte. Ingen levende Excel-formler
' her lenger: alt regnes ut i VBA og skrives som faste verdier,
' akkurat som Resultat og Oversikt for ovrig.
Private Sub UpdateControlSheet( _
    ByVal ws As Worksheet, _
    ByVal kontrollRows As Collection, _
    ByVal totalVehicles As Long, _
    ByVal dateFrom As Date, _
    ByVal dateTo As Date)

    Const HEADER_ROW As Long = 12
    Const FIRST_DATA_ROW As Long = 13

    Dim row As Object
    Dim txRow As Variant
    Dim alleTx As Collection
    Dim sortertRader As Collection
    Dim r As Long
    Dim lastRow As Long
    Dim groupStartRow As Long

    Dim bucket0 As Long
    Dim bucket1til15 As Long
    Dim bucketOver15 As Long

    Dim kontrollertText As String
    Dim dagerAvvik As Variant
    Dim bucketColor As Long
    Dim bucketFontColor As Long
    Dim matchetTx As Object

    ws.Cells.Clear

    '----------------------------------------------------------
    ' Tittel og forklaring av kontrollregelen
    '----------------------------------------------------------

    ws.Range("A1:M1").Merge
    ws.Range("A1").value = CONTROL_SHEET

    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ws.rows(1).RowHeight = 26

    ws.Range("A2:M2").Merge
    ws.Range("A2").value = _
        "Kontrollregel: For hver bil ser vi forst etter det SISTE " & _
        "eierskiftet OFV har registrert innenfor +/- 1 maned rundt " & _
        "bokfort dato. Finnes ett eller flere slike, brukes det " & _
        "nyeste av dem (" & Chr$(171) & "Eierskifte i perioden" & _
        Chr$(187) & ")."

    ws.Range("A3:M3").Merge
    ws.Range("A3").value = _
        "Finnes ingen eierskifte i det vinduet, brukes i stedet " & _
        "bilens SISTE registrerte eierskifte totalt, uansett dato " & _
        "(" & Chr$(171) & "Siste registrerte eierskifte" & Chr$(187) & _
        "), og dager avvik regnes fra den datoen i stedet."

    ws.Range("A4:M4").Merge
    ws.Range("A4").value = _
        "Forstegangsregistrering pavirker aldri dette valget - den " & _
        "vises kun som egen opplysning. Hver bil har en uthevet " & _
        "hovedrad med full kontrollinfo (den matchede transaksjonen), " & _
        "etterfulgt av alle bilens ovrige eierskifter med kun dato " & _
        "og transaksjonsinfo. Tabellen er sortert med storst dagers " & _
        "avvik forst."

    ws.Range("A2:A4").Font.Italic = True
    ws.rows("2:4").RowHeight = 15

    '----------------------------------------------------------
    ' Sorter hovedradene etter dagers avvik, storst forst
    '----------------------------------------------------------

    Set sortertRader = SorterKontrollRadPaAvvik(kontrollRows)

    '----------------------------------------------------------
    ' Tell opp KPI-er
    '----------------------------------------------------------

    For Each row In sortertRader

        kontrollertText = VariantToString(row("Kontrollert"))

        If kontrollertText = "Ja" Then

            dagerAvvik = row("DagerAvvik")

            If IsNumeric(dagerAvvik) Then

                If CLng(dagerAvvik) = 0 Then
                    bucket0 = bucket0 + 1
                ElseIf CLng(dagerAvvik) <= 15 Then
                    bucket1til15 = bucket1til15 + 1
                Else
                    bucketOver15 = bucketOver15 + 1
                End If

            End If

        End If

    Next row

    '----------------------------------------------------------
    ' KPI-bokser (rad 6-8)
    '----------------------------------------------------------

    ws.Range("A6:B6").Merge : ws.Range("A6").value = "Inputbiler"
    ws.Range("G6").value = "Fra dato"
    ws.Range("H6").value = "Til dato"
    ws.Range("I6:M6").Merge
    ws.Range("I6").value = "FARGEKODER - DAGER AVVIK"

    With ws.Range("A6:M6")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("A7:B7").Merge
    ws.Range("A7").value = totalVehicles

    ws.Range("G7").value = dateFrom
    ws.Range("H7").value = dateTo
    ws.Range("G7:H7").NumberFormat = "dd.mm.yyyy"

    With ws.Range("A7:H7")
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("I7").value = "0 dager"
    ws.Range("J7:K7").Merge : ws.Range("J7").value = "1-15 dager"
    ws.Range("L7:M7").Merge : ws.Range("L7").value = "Over 15 dager"

    ws.Range("I8").value = bucket0
    ws.Range("J8:K8").Merge : ws.Range("J8").value = bucket1til15
    ws.Range("L8:M8").Merge : ws.Range("L8").value = bucketOver15

    With ws.Range("I7:I8")
        .Interior.Color = COLOR_GREEN_FILL
        .Font.Color = COLOR_GREEN_FONT
    End With

    With ws.Range("J7:K8")
        .Interior.Color = COLOR_YELLOW_FILL
        .Font.Color = COLOR_YELLOW_FONT
    End With

    With ws.Range("L7:M8")
        .Interior.Color = COLOR_RED_FILL
        .Font.Color = COLOR_RED_FONT
    End With

    With ws.Range("I7:M8")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.rows("6:8").RowHeight = 20

    '----------------------------------------------------------
    ' Kolonneoverskrifter
    '----------------------------------------------------------

    ws.Range("A" & HEADER_ROW).value = "API-treff"
    ws.Range("B" & HEADER_ROW).value = "Regnr / input"
    ws.Range("C" & HEADER_ROW).value = "Chassisnummer"
    ws.Range("D" & HEADER_ROW).value = "Modell"
    ws.Range("E" & HEADER_ROW).value = "Bokfort dato"
    ws.Range("F" & HEADER_ROW).value = "Forstegangsregistrert"
    ws.Range("G" & HEADER_ROW).value = "Transaksjonsdato"
    ws.Range("H" & HEADER_ROW).value = "Kontrollgrunnlag / type"
    ws.Range("I" & HEADER_ROW).value = "Dager avvik"
    ws.Range("J" & HEADER_ROW).value = "Kontrollert"
    ws.Range("K" & HEADER_ROW).value = "Selger"
    ws.Range("L" & HEADER_ROW).value = "Kjoper"
    ws.Range("M" & HEADER_ROW).value = "Status eierskifte"

    With ws.Range("A" & HEADER_ROW & ":M" & HEADER_ROW)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 34
    End With

    '----------------------------------------------------------
    ' Hovedrad per kjoretoy + detaljrad per ovrig eierskifte
    '----------------------------------------------------------

    r = FIRST_DATA_ROW

    For Each row In sortertRader

        groupStartRow = r

        ws.Cells(r, 1).value = VariantToString(row("ApiTreff"))
        ws.Cells(r, 2).value = VariantToString(row("RegnrInput"))
        ws.Cells(r, 3).value = VariantToString(row("Chassisnummer"))
        ws.Cells(r, 4).value = VariantToString(row("Modell"))
        ws.Cells(r, 5).value = row("BokfortDato")
        ws.Cells(r, 6).value = row("Forstegangsregistrert")
        ws.Cells(r, 7).value = row("KontrollTransaksjonDato")
        ws.Cells(r, 8).value = VariantToString(row("Kontrollgrunnlag"))
        ws.Cells(r, 9).value = row("DagerAvvik")
        ws.Cells(r, 10).value = VariantToString(row("Kontrollert"))
        ws.Cells(r, 11).value = VariantToString(row("Selger"))
        ws.Cells(r, 12).value = VariantToString(row("Kjoper"))
        ws.Cells(r, 13).value = VariantToString(row("StatusEierskifte"))

        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 13))
            .Font.Bold = True
            .Interior.Color = RGB(238, 244, 251)
        End With

        dagerAvvik = row("DagerAvvik")

        If IsNumeric(dagerAvvik) Then

            If CLng(dagerAvvik) = 0 Then
                bucketColor = COLOR_GREEN_FILL
                bucketFontColor = COLOR_GREEN_FONT
            ElseIf CLng(dagerAvvik) <= 15 Then
                bucketColor = COLOR_YELLOW_FILL
                bucketFontColor = COLOR_YELLOW_FONT
            Else
                bucketColor = COLOR_RED_FILL
                bucketFontColor = COLOR_RED_FONT
            End If

            With ws.Range(ws.Cells(r, 7), ws.Cells(r, 7))
                .Interior.Color = bucketColor
                .Font.Color = bucketFontColor
            End With

            With ws.Range(ws.Cells(r, 9), ws.Cells(r, 9))
                .Interior.Color = bucketColor
                .Font.Color = bucketFontColor
                .Font.Bold = True
            End With

        End If

        r = r + 1

        ' Detaljrader: alle ovrige eierskifter for samme bil, kun
        ' transaksjonsdato og transaksjonsinfo (den matchede
        ' transaksjonen er allerede vist i full i hovedraden over).
        Set alleTx = row("AlleTransaksjoner")

        If Not alleTx Is Nothing Then

            Set matchetTx = row("MatchetTransaksjon")

            For Each txRow In alleTx

                If matchetTx Is Nothing Or _
                   Not txRow Is matchetTx Then

                    ws.Cells(r, 7).value = txRow("TransactionDate")
                    ws.Cells(r, 8).value = _
                        VariantToString(txRow("RegistrationType"))
                    ws.Cells(r, 11).value = ComputeOwnerLabel( _
                        VariantToString(txRow("FromOwnerType")), _
                        VariantToString(txRow("FromOwnerCompanyName")))
                    ws.Cells(r, 12).value = ComputeOwnerLabel( _
                        VariantToString(txRow("ToOwnerType")), _
                        VariantToString(txRow("ToOwnerCompanyName")))
                    ws.Cells(r, 13).value = _
                        VariantToString(txRow("Status"))

                    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 13))
                        .Font.Italic = True
                        .Font.Color = RGB(90, 90, 90)
                    End With

                    r = r + 1

                End If

            Next txRow

        End If

        ' Tykk topplinje over hver ny bil, sa gruppene er lette a se.
        With ws.Range( _
            ws.Cells(groupStartRow, 1), ws.Cells(groupStartRow, 13)).Borders(xlEdgeTop)

            .LineStyle = xlContinuous
            .Color = RGB(31, 78, 120)
            .Weight = xlMedium

        End With

    Next row

    lastRow = r - 1
    If lastRow < FIRST_DATA_ROW Then lastRow = FIRST_DATA_ROW

    With ws.Range("A" & FIRST_DATA_ROW & ":M" & lastRow)
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .rows.RowHeight = 18
    End With

    ws.Range("E" & FIRST_DATA_ROW & ":G" & lastRow).NumberFormat = _
        "dd.mm.yyyy"

    ws.Range("I" & FIRST_DATA_ROW & ":I" & lastRow).NumberFormat = "0"

    With ws.Range("A" & HEADER_ROW & ":M" & lastRow).Borders
        .LineStyle = xlContinuous
        .Color = RGB(217, 226, 243)
        .Weight = xlThin
    End With

    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E:G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 24
    ws.Columns("I:J").ColumnWidth = 12
    ws.Columns("K:L").ColumnWidth = 25
    ws.Columns("M").ColumnWidth = 28

End Sub


' Sorterer kontrollradene etter DagerAvvik, storst forst. Rader uten
' et tallavvik (ikke kontrollert / ingen treff) regnes som -1 og
' havner dermed sist.
Private Function SorterKontrollRadPaAvvik( _
    ByVal kontrollRows As Collection) As Collection

    Dim sortert As New Collection
    Dim row As Variant
    Dim i As Long
    Dim inserted As Boolean
    Dim thisVal As Double
    Dim otherVal As Double

    For Each row In kontrollRows

        inserted = False
        thisVal = AvvikSorteringsverdi(row("DagerAvvik"))

        For i = 1 To sortert.Count

            otherVal = AvvikSorteringsverdi(sortert(i)("DagerAvvik"))

            If thisVal > otherVal Then
                sortert.Add row, Before:=i
                inserted = True
                Exit For
            End If

        Next i

        If Not inserted Then sortert.Add row

    Next row

    Set SorterKontrollRadPaAvvik = sortert

End Function


Private Function AvvikSorteringsverdi(ByVal dagerAvvik As Variant) As Double

    If IsNumeric(dagerAvvik) Then
        AvvikSorteringsverdi = CDbl(dagerAvvik)
    Else
        AvvikSorteringsverdi = -1
    End If

End Function


' Bygger kontroll-raden for ett kjoretoy.
'
' Kontrollregel (i denne rekkefolgen):
'   1. Se etter det SISTE eierskiftet innenfor +/- 1 maned rundt
'      bokfort dato. Finnes ett eller flere, brukes det nyeste av dem.
'   2. Finnes ingen eierskifte i det vinduet, brukes i stedet bilens
'      SISTE registrerte eierskifte totalt (uansett dato), og avviket
'      regnes fra den datoen - markert tydelig som "utenfor perioden"
'      i ApiTreff/Kontrollgrunnlag, slik at det ikke forveksles med et
'      treff innenfor vinduet.
'   3. Forstegangsregistrering pavirker ALDRI dette valget - den vises
'      kun som egen, uavhengig kolonne (forste gang bilen ble
'      registrert, hentet fra hvilken som helst av bilens
'      transaksjoner).
'
' Alle bilens eierskifter samles ogsa i "AlleTransaksjoner" (sortert
' pa dato), slik at UpdateControlSheet kan vise dem som egne rader
' under kjoretoyets hovedrad. Den valgte transaksjonen merkes med
' ErKontrollMatch=True direkte pa det delte JSON-objektet, slik at
' Resultat-arket kan kjenne igjen og utheve akkurat den samme raden.
Private Function BuildKontrollRow( _
    ByVal regNo As String, _
    ByVal vin As String, _
    ByVal bokfortRaw As Variant, _
    ByVal vehicleTxRows As Collection, _
    ByVal periodeFra As Date, _
    ByVal periodeTil As Date) As Object

    Dim result As Object
    Dim txRow As Variant
    Dim bokfortDate As Variant

    Dim firstRegDate As Variant
    Dim modelName As String
    Dim chassisNo As String
    Dim errorStatus As String
    Dim hasAnyOkRow As Boolean

    Dim vinduFra As Date
    Dim vinduTil As Date

    Dim matchITxDato As Variant
    Dim matchITxRow As Object
    Dim sisteTxDato As Variant
    Dim sisteTxRow As Object

    Dim alleTransaksjoner As Collection
    Dim sortertListe As Collection

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    bokfortDate = Empty
    If IsDate(bokfortRaw) Then bokfortDate = CDate(bokfortRaw)

    If Not IsEmpty(bokfortDate) Then
        vinduFra = DateAdd("m", -1, CDate(bokfortDate))
        vinduTil = DateAdd("m", 1, CDate(bokfortDate))
    End If

    firstRegDate = Empty
    modelName = vbNullString
    chassisNo = vin
    errorStatus = vbNullString
    hasAnyOkRow = False

    matchITxDato = Empty
    Set matchITxRow = Nothing
    sisteTxDato = Empty
    Set sisteTxRow = Nothing

    Set alleTransaksjoner = New Collection

    If Not vehicleTxRows Is Nothing Then

        For Each txRow In vehicleTxRows

            If VariantToString(txRow("Status")) = "OK" Then

                hasAnyOkRow = True
                alleTransaksjoner.Add txRow

                If IsEmpty(firstRegDate) Then
                    If IsDate(txRow("FirstRegistrationDate")) Then
                        firstRegDate = txRow("FirstRegistrationDate")
                    End If
                End If

                If Len(modelName) = 0 Then
                    modelName = VariantToString(txRow("ModelName"))
                End If

                If Len(VariantToString( _
                    txRow("ChassisNumber"))) > 0 Then

                    chassisNo = VariantToString( _
                        txRow("ChassisNumber"))

                End If

                If IsDate(txRow("TransactionDate")) Then

                    ' Siste registrerte eierskifte totalt (reserve).
                    If IsEmpty(sisteTxDato) Then

                        sisteTxDato = txRow("TransactionDate")
                        Set sisteTxRow = txRow

                    ElseIf CDate(txRow("TransactionDate")) > _
                        CDate(sisteTxDato) Then

                        sisteTxDato = txRow("TransactionDate")
                        Set sisteTxRow = txRow

                    End If

                    ' Siste eierskifte innenfor +/- 1 maned-vinduet.
                    If Not IsEmpty(bokfortDate) And _
                       CDate(txRow("TransactionDate")) >= vinduFra And _
                       CDate(txRow("TransactionDate")) <= vinduTil Then

                        If IsEmpty(matchITxDato) Then

                            matchITxDato = txRow("TransactionDate")
                            Set matchITxRow = txRow

                        ElseIf CDate(txRow("TransactionDate")) > _
                            CDate(matchITxDato) Then

                            matchITxDato = txRow("TransactionDate")
                            Set matchITxRow = txRow

                        End If

                    End If

                End If

            ElseIf Left$(VariantToString(txRow("Status")), 5) = _
                "Feil:" Then

                If Len(errorStatus) = 0 Then
                    errorStatus = VariantToString(txRow("Status"))
                End If

            End If

        Next txRow

    End If

    result("RegnrInput") = regNo
    result("Chassisnummer") = chassisNo
    result("Modell") = modelName
    result("BokfortDato") = bokfortDate
    result("Forstegangsregistrert") = firstRegDate
    result("KontrollTransaksjonDato") = Empty
    result("Kontrollgrunnlag") = vbNullString
    result("DagerAvvik") = Empty
    result("Selger") = vbNullString
    result("Kjoper") = vbNullString
    result("StatusEierskifte") = vbNullString
    Set result("MatchetTransaksjon") = Nothing

    If IsEmpty(bokfortDate) Then

        If hasAnyOkRow Then
            result("ApiTreff") = "Mangler bokfort dato"
        ElseIf Len(errorStatus) > 0 Then
            result("ApiTreff") = errorStatus
        Else
            result("ApiTreff") = "Ingen treff"
        End If

        result("Kontrollert") = "Nei"

    ElseIf Not matchITxRow Is Nothing Then

        result("KontrollTransaksjonDato") = matchITxDato
        result("Kontrollgrunnlag") = "Eierskifte i perioden"
        result("DagerAvvik") = Abs(CLng( _
            CDate(matchITxDato) - CDate(bokfortDate)))
        result("ApiTreff") = "Treff OFV eierskifte i perioden"
        result("Kontrollert") = "Ja"
        Set result("MatchetTransaksjon") = matchITxRow

        result("Selger") = ComputeOwnerLabel( _
            VariantToString(matchITxRow("FromOwnerType")), _
            VariantToString(matchITxRow("FromOwnerCompanyName")))

        result("Kjoper") = ComputeOwnerLabel( _
            VariantToString(matchITxRow("ToOwnerType")), _
            VariantToString(matchITxRow("ToOwnerCompanyName")))

        result("StatusEierskifte") = _
            VariantToString(matchITxRow("Status"))

        matchITxRow("ErKontrollMatch") = True

    ElseIf Not sisteTxRow Is Nothing Then

        result("KontrollTransaksjonDato") = sisteTxDato
        result("Kontrollgrunnlag") = "Siste registrerte eierskifte"
        result("DagerAvvik") = Abs(CLng( _
            CDate(sisteTxDato) - CDate(bokfortDate)))
        result("ApiTreff") = "Siste eierskifte (utenfor perioden)"
        result("Kontrollert") = "Ja"
        Set result("MatchetTransaksjon") = sisteTxRow

        result("Selger") = ComputeOwnerLabel( _
            VariantToString(sisteTxRow("FromOwnerType")), _
            VariantToString(sisteTxRow("FromOwnerCompanyName")))

        result("Kjoper") = ComputeOwnerLabel( _
            VariantToString(sisteTxRow("ToOwnerType")), _
            VariantToString(sisteTxRow("ToOwnerCompanyName")))

        result("StatusEierskifte") = _
            VariantToString(sisteTxRow("Status"))

        sisteTxRow("ErKontrollMatch") = True

    Else

        If Len(errorStatus) > 0 Then
            result("ApiTreff") = errorStatus
        Else
            result("ApiTreff") = "Ingen treff"
        End If

        result("Kontrollert") = "Nei"

    End If

    ' Alle bilens eierskifter, sortert kronologisk (eldst forst),
    ' til bruk for detaljradene i Kontroll solgte biler.
    Set sortertListe = SorterTransaksjonerPaDato(alleTransaksjoner)
    Set result("AlleTransaksjoner") = sortertListe

    Set BuildKontrollRow = result

End Function


' Enkel innsettingssortering (fa elementer per bil, ytelse er ikke
' et tema) - stigende pa TransactionDate.
Private Function SorterTransaksjonerPaDato( _
    ByVal txRows As Collection) As Collection

    Dim sortert As New Collection
    Dim txRow As Variant
    Dim i As Long
    Dim inserted As Boolean

    For Each txRow In txRows

        inserted = False

        For i = 1 To sortert.Count

            If CDate(txRow("TransactionDate")) < _
                CDate(sortert(i)("TransactionDate")) Then

                sortert.Add txRow, Before:=i
                inserted = True
                Exit For

            End If

        Next i

        If Not inserted Then sortert.Add txRow

    Next txRow

    Set SorterTransaksjonerPaDato = sortert

End Function


' Samme utledning som formlene i Resultat-tabellen
' (SelgerType/KjoperType): Privat vinner over firmanavn, ellers
' firmanavn hvis det finnes, ellers den ra eiertype-teksten.
Private Function ComputeOwnerLabel( _
    ByVal ownerType As String, _
    ByVal companyName As String) As String

    If Len(ownerType) = 0 And Len(companyName) = 0 Then
        ComputeOwnerLabel = vbNullString
    ElseIf ownerType = "Privat" Then
        ComputeOwnerLabel = "Privat"
    ElseIf Len(companyName) > 0 Then
        ComputeOwnerLabel = companyName
    Else
        ComputeOwnerLabel = ownerType
    End If

End Function


'==============================================================
' OVERSIKT
'==============================================================

' Enkel liste - Regnr/input og forstegangsregistrering for hvert
' kjoretoy som ble kjort denne runden. Star til hoyre for KPI-boksene
' (kolonne P) sa den ikke kolliderer med pivottabellen i A11.
Private Sub WriteFirstRegistrationOverviewList( _
    ByVal ws As Worksheet, _
    ByVal kontrollRows As Collection)

    Const START_COL As String = "P"
    Const START_ROW As Long = 4

    Dim row As Object
    Dim r As Long

    ws.Range("P" & (START_ROW - 1) & ":Q" & _
        (ws.rows.Count)).ClearContents

    ws.Range(START_COL & (START_ROW - 1)).value = "Regnr / input"
    ws.Range("Q" & (START_ROW - 1)).value = "Forstegangsregistrering"

    With ws.Range(START_COL & (START_ROW - 1) & ":Q" & (START_ROW - 1))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
    End With

    r = START_ROW

    For Each row In kontrollRows

        ws.Cells(r, "P").value = VariantToString(row("RegnrInput"))
        ws.Cells(r, "Q").value = row("Forstegangsregistrert")

        r = r + 1

    Next row

    ws.Range("Q" & START_ROW & ":Q" & (r - 1)).NumberFormat = "dd.mm.yyyy"

    ws.Columns("P").ColumnWidth = 16
    ws.Columns("Q").ColumnWidth = 22

End Sub


Private Sub UpdateOverviewKPIs(ByVal ws As Worksheet)

    ws.Range("A4").value = "Antall kjoretoy"
    ws.Range("C4").value = "OFV-treff (OK)"
    ws.Range("F4").value = "Uten treff / feil"

    ws.Range("A4:A4,C4:C4,F4:F4").Font.Bold = True

    ws.Range("A5").Formula = _
        "=SUMPRODUCT(--(((KjoretoyInput[Regnr]<>"""")+" & _
        "(KjoretoyInput[VIN]<>""""))>0))"

    ws.Range("C5").Formula = _
        "=COUNTIF(Transaksjoner[Status],""OK"")"

    ws.Range("F5").Formula = _
        "=COUNTIF(Transaksjoner[Status],""<>OK"")"

End Sub


Private Sub RebuildOverviewPivot( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject)

    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim fields As Variant
    Dim i As Long

    Do While ws.PivotTables.Count > 0
        ws.PivotTables(1).TableRange2.Clear
    Loop

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=lo.Name)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A11"), _
        TableName:=PIVOT_NAME)

    fields = Array( _
        "RegNo", _
        "Merke", _
        "Modell", _
        "Chassisnummer", _
        "Drivstoffgruppe", _
        "ForstegangsRegistrering", _
        "Leaset", _
        "TransaksjonsNummer", _
        "Eierskiftedato", _
        "SelgerType", _
        BuyerTypeHeader(), _
        "Status")

    pt.ManualUpdate = True

    For i = LBound(fields) To UBound(fields)

        Set pf = pt.PivotFields(CStr(fields(i)))

        pf.Orientation = xlRowField
        pf.position = i + 1

        On Error Resume Next

        pf.Subtotals = Array( _
            False, False, False, False, _
            False, False, False, False, _
            False, False, False, False)

        On Error GoTo 0

    Next i

    pt.RowAxisLayout xlTabularRow
    pt.RowGrand = False
    pt.ColumnGrand = False
    pt.TableStyle2 = "PivotStyleMedium2"
    pt.ManualUpdate = False
    pt.RefreshTable

    ws.Range("A2").value = _
        "Alle data (inkludert forstegangsregistrering) " & _
        "hentes fra OFV."

End Sub


'==============================================================
' TEKST OG IDENTIFIKATORER
'==============================================================

' Henter en arkfane ved navn og gir en presis feilmelding
' (istedenfor "Subscript out of range") hvis fanen ikke finnes -
' f.eks. ved skrivefeil, ekstra mellomrom eller feil store/sma bokstaver.
Private Function GetRequiredSheet( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then

        Err.Raise vbObjectError + 1010, "GetRequiredSheet", _
            "Fant ikke arkfanen """ & sheetName & """. " & _
            "Sjekk at en arkfane med akkurat dette navnet " & _
            "finnes i arbeidsboken (store/sma bokstaver og " & _
            "mellomrom ma stemme noyaktig)."

    End If

    Set GetRequiredSheet = ws

End Function


' Oppretter Excel-tabellen KjoretoyInput pa Input-arket hvis den
' ikke finnes fra for. Resten av koden (og formlene i Oversikt og
' Kontroll solgte biler) refererer til KjoretoyInput[Regnr] og
' KjoretoyInput[VIN] som strukturerte referanser - det krever et
' ekte tabellobjekt, ikke bare rader med tekst. Overskriftsraden
' forventes rett over FIRST_ROW (dvs. rad 4 nar FIRST_ROW er 5),
' med Regnr i kolonne B og VIN i kolonne C.
Private Sub EnsureInputTable(ByVal ws As Worksheet)

    Dim lo As ListObject
    Dim headerRow As Long
    Dim lastDataRow As Long
    Dim target As Range

    On Error Resume Next
    Set lo = ws.ListObjects(INPUT_TABLE)
    On Error GoTo 0

    If Not lo Is Nothing Then Exit Sub

    headerRow = FIRST_ROW - 1

    lastDataRow = Application.Max( _
        ws.Cells(ws.rows.Count, COL_REGNR).End(xlUp).Row, _
        ws.Cells(ws.rows.Count, COL_VIN).End(xlUp).Row, _
        ws.Cells(ws.rows.Count, COL_BOKFORT).End(xlUp).Row, _
        ws.Cells(ws.rows.Count, COL_FORSTEREG).End(xlUp).Row)

    If lastDataRow < FIRST_ROW Then
        lastDataRow = FIRST_ROW
    End If

    ' Bokfort- og Forstegangsregistrering-kolonnene har kanskje ingen
    ' overskrift enna (de er nye).
    If Len(Trim$(CStr( _
        ws.Cells(headerRow, COL_BOKFORT).value & vbNullString))) = 0 Then

        ws.Cells(headerRow, COL_BOKFORT).value = "Bokfort"

    End If

    If Len(Trim$(CStr( _
        ws.Cells(headerRow, COL_FORSTEREG).value & vbNullString))) = 0 Then

        ws.Cells(headerRow, COL_FORSTEREG).value = "Forstegangsregistrering"

    End If

    Set target = ws.Range( _
        ws.Cells(headerRow, COL_REGNR), _
        ws.Cells(lastDataRow, COL_FORSTEREG))

    Set lo = ws.ListObjects.Add(xlSrcRange, target, , xlYes)
    lo.Name = INPUT_TABLE

    ' Tving eksakte kolonnenavn uansett hva som sto i overskriftscellene,
    ' slik at KjoretoyInput[Regnr]/[VIN]/[Bokfort] alltid treffer.
    lo.ListColumns(1).Name = "Regnr"
    lo.ListColumns(2).Name = "VIN"
    lo.ListColumns(3).Name = "Bokfort"
    lo.ListColumns(4).Name = "Forstegangsregistrering"

    ws.Range(ws.Cells(FIRST_ROW, COL_BOKFORT), _
        ws.Cells(lastDataRow, COL_FORSTEREG)).NumberFormat = "dd.mm.yyyy"

End Sub


' Skriver forstegangsregistreringsdato tilbake til Input-arket for
' hvert regnr/VIN som ble kjort, slik at den er synlig direkte pa
' Input og ikke bare i Resultat/Oversikt/Kontroll solgte biler.
' firstRegByKey er nokkelet pa samme BuildVehicleKey-format som queue.
Private Sub WriteFirstRegistrationToInput( _
    ByVal ws As Worksheet, _
    ByVal firstRegByKey As Object)

    Dim lastRow As Long
    Dim r As Long
    Dim rowRegNo As String
    Dim rowVin As String
    Dim rowKey As String
    Dim regDate As Variant

    lastRow = Application.Max( _
        ws.Cells(ws.rows.Count, COL_REGNR).End(xlUp).Row, _
        ws.Cells(ws.rows.Count, COL_VIN).End(xlUp).Row)

    For r = FIRST_ROW To lastRow

        rowRegNo = NormalizeIdentifier(ws.Cells(r, COL_REGNR).value)
        rowVin = NormalizeIdentifier(ws.Cells(r, COL_VIN).value)
        rowKey = BuildVehicleKey(rowVin, rowRegNo)

        If Len(rowKey) > 0 Then

            If firstRegByKey.Exists(rowKey) Then

                regDate = firstRegByKey(rowKey)

                If IsDate(regDate) Then
                    ws.Cells(r, COL_FORSTEREG).value = CDate(regDate)
                Else
                    ws.Cells(r, COL_FORSTEREG).ClearContents
                End If

            End If

        End If

    Next r

    ws.Range(ws.Cells(FIRST_ROW, COL_FORSTEREG), _
        ws.Cells(lastRow, COL_FORSTEREG)).NumberFormat = "dd.mm.yyyy"

End Sub


' Leser verdien fra et navngitt omrade hvis det finnes i
' arbeidsboken, ellers fra en gitt reserveCelle. Brukes til
' oppsettsverdier (API-nokkel, datoperiode) som enten kan ligge i
' et navngitt omrade eller direkte i en fast celle - istedenfor at
' et manglende navngitt omrade gir "Application-defined or
' object-defined error" (feil 1004).
Private Function ReadConfigValue( _
    ByVal wb As Workbook, _
    ByVal namedRangeName As String, _
    ByVal fallbackCell As Range) As Variant

    Dim result As Variant

    result = Empty

    On Error Resume Next
    result = wb.Names(namedRangeName).RefersToRange.value
    On Error GoTo 0

    If IsEmpty(result) Then

        result = fallbackCell.value

    ElseIf VarType(result) = vbString Then

        If Len(Trim$(CStr(result))) = 0 Then
            result = fallbackCell.value
        End If

    End If

    ReadConfigValue = result

End Function


Private Function BuyerTypeHeader() As String
    BuyerTypeHeader = "Kj" & ChrW(248) & "perType"
End Function


Private Function BuyerCountyHeader() As String

    BuyerCountyHeader = _
        "Kj" & ChrW(248) & "perEierFylke"

End Function


Private Function NormalizeIdentifier( _
    ByVal value As Variant) As String

    NormalizeIdentifier = _
        UCase$(Trim$(Replace( _
            CStr(value & vbNullString), _
            " ", vbNullString)))

End Function


Private Function BuildVehicleKey( _
    ByVal vin As String, _
    ByVal regNo As String) As String

    If Len(vin) > 0 Then
        BuildVehicleKey = "VIN|" & vin
    ElseIf Len(regNo) > 0 Then
        BuildVehicleKey = "REG|" & regNo
    Else
        BuildVehicleKey = vbNullString
    End If

End Function


Private Function VariantToString( _
    ByVal value As Variant) As String

    If IsNull(value) Or IsEmpty(value) Then
        VariantToString = vbNullString
    Else
        VariantToString = CStr(value)
    End If

End Function


Private Function JsonEscape( _
    ByVal value As String) As String

    Dim result As String

    result = Replace(value, "\", "\\")
    result = Replace( _
        result, Chr$(34), "\" & Chr$(34))

    JsonEscape = result

End Function


Private Function DateFromISO( _
    ByVal isoDate As String) As Variant

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
' JSON
'==============================================================

Private Function JSON_FindMatchingBrace( _
    ByVal json As String, _
    ByVal openPosition As Long) As Long

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
    ByVal json As String, _
    ByVal position As Long) As Long

    Dim p As Long
    Dim character As String

    p = position

    Do While p <= Len(json)

        character = Mid$(json, p, 1)

        If character = " " Or _
           character = vbLf Or _
           character = vbCr Or _
           character = vbTab Then

            p = p + 1

        Else
            Exit Do
        End If

    Loop

    JSON_SkipWhitespace = p

End Function


Private Function JSON_ExtractObject( _
    ByVal json As String, _
    ByVal key As String) As String

    Dim keyPosition As Long
    Dim valuePosition As Long
    Dim endPosition As Long
    Dim firstCharacter As String

    If Len(json) = 0 Then Exit Function

    keyPosition = InStr( _
        1, json, _
        Chr$(34) & key & Chr$(34) & ":", _
        vbBinaryCompare)

    If keyPosition = 0 Then Exit Function

    valuePosition = keyPosition + Len(key) + 3
    valuePosition = JSON_SkipWhitespace( _
        json, valuePosition)

    firstCharacter = Mid$( _
        json, valuePosition, 1)

    If firstCharacter <> "{" And _
       firstCharacter <> "[" Then Exit Function

    endPosition = JSON_FindMatchingBrace( _
        json, valuePosition)

    If endPosition = 0 Then Exit Function

    JSON_ExtractObject = Mid$( _
        json, valuePosition, _
        endPosition - valuePosition + 1)

End Function


Private Function JSON_ExtractValue( _
    ByVal json As String, _
    ByVal key As String) As Variant

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
        1, json, _
        Chr$(34) & key & Chr$(34) & ":", _
        vbBinaryCompare)

    If keyPosition = 0 Then Exit Function

    valuePosition = keyPosition + Len(key) + 3
    valuePosition = JSON_SkipWhitespace( _
        json, valuePosition)

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

    ElseIf firstCharacter = "{" Or _
           firstCharacter = "[" Then

        endPosition = JSON_FindMatchingBrace( _
            json, valuePosition)

        If endPosition > 0 Then

            JSON_ExtractValue = Mid$( _
                json, valuePosition, _
                endPosition - valuePosition + 1)

        End If

    Else

        j = valuePosition

        Do While j <= Len(json)

            character = Mid$(json, j, 1)

            If character = "," Or _
               character = "}" Or _
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

        If Mid$(jsonArray, position, 1) = "]" Then
            Exit Do
        End If

        firstCharacter = Mid$(jsonArray, position, 1)
        element = vbNullString

        If firstCharacter = "{" Or _
           firstCharacter = "[" Then

            endPosition = JSON_FindMatchingBrace( _
                jsonArray, position)

            If endPosition = 0 Then Exit Do

            element = Mid$( _
                jsonArray, position, _
                endPosition - position + 1)

            position = endPosition + 1

        Else

            i = position
            rawValue = vbNullString

            Do While i <= Len(jsonArray)

                character = Mid$(jsonArray, i, 1)

                If character = "," Or _
                   character = "]" Then Exit Do

                rawValue = rawValue & character
                i = i + 1

            Loop

            element = Trim$(rawValue)
            position = i

        End If

        result.Add element

        position = JSON_SkipWhitespace( _
            jsonArray, position)

        If position <= Len(jsonArray) Then

            If Mid$(jsonArray, position, 1) = "," Then
                position = JSON_SkipWhitespace( _
                    jsonArray, position + 1)
            End If

        End If

    Loop

    Set JSON_ArrayAllElements = result

End Function
