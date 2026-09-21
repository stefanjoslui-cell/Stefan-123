Option Explicit

' Produksjonsmakro for prosjektet "Python program for bilregistrering":
' Input-arket (KjoretoyInput), Resultat-arket (Transaksjoner-tabellen),
' Oversikt (KPI-er + pivot) og Kontroll solgte biler beholdes uendret.
'
' Endret i forhold til forrige versjon:
'   - Statens Vegvesen (SVV/"VVS") er fjernet fullstendig. Ingen kall,
'     ingen nokkel, ingen status- eller tellevariabler for den kilden.
'     ForstegangsRegistrering hentes na direkte fra OFV sitt eget felt
'     "firstRegistrationDate" pa hver transaksjon - den var allerede der,
'     sa vi mister ingen funksjonalitet ved a droppe SVV-kallet.
'   - HTTP-kallene mot OFV bruker WinHttp.WinHttpRequest.5.1 (samme som
'     i testscriptet som lostte lagringsproblemene dine), ikke
'     MSXML2.ServerXMLHTTP.6.0 som den forrige versjonen brukte.
'   - OFV_URL peker na pa det bekreftet fungerende endepunktet
'     (https://data.ofv.no/transactions/v1/query). Det som stod i
'     forrige versjon (https://api.ofv.no/transactions/v1/, hentet fra
'     swagger-dokumentets "host") gir 404 i praksis.

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

Private Const OFV_URL As String = _
    "https://data.ofv.no/transactions/v1/query"

Private Const MAX_RETRIES As Long = 4
Private Const RETRY_WAIT_MS As Long = 3000
Private Const API_PAUSE_MS As Long = 150


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
    Dim resultRow As Object

    Dim allRows As Collection
    Dim vehicleRows As Collection

    Dim fieldMap As Variant
    Dim vehicleData As Variant
    Dim output() As Variant

    Dim ofvKey As String
    Dim dateFrom As Variant
    Dim dateTo As Variant
    Dim DateFromISO As String
    Dim dateToISO As String
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

    stage = "leser API-nokkel"
    API_ShowStatus "Forbereder", stage

    ofvKey = Trim$(CStr( _
        ThisWorkbook.Names("OFV_API").RefersToRange.value))

    If Len(ofvKey) = 0 Then
        MsgBox "Fant ingen OFV-nokkel i OFV_API.", _
            vbExclamation, "API-oppdatering"
        GoTo SafeExit
    End If

    stage = "leser datoperioden"
    API_ShowStatus "Forbereder", stage

    dateFrom = ThisWorkbook.Names( _
        "OFV_DateFrom").RefersToRange.value

    dateTo = ThisWorkbook.Names( _
        "OFV_DateTo").RefersToRange.value

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

    DateFromISO = Format$(CDate(dateFrom), "yyyy-mm-dd")
    dateToISO = Format$(CDate(dateTo), "yyyy-mm-dd")

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

        queueKey = BuildVehicleKey(vin, regNo)

        If Len(queueKey) > 0 Then
            If Not queue.Exists(queueKey) Then
                queue.Add queueKey, Array(regNo, vin)
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
            vin, _
            DateFromISO, _
            dateToISO)

        For Each resultRow In vehicleRows

            statusText = VariantToString( _
                resultRow("Status"))

            If statusText = "OK" Then

                transactionCount = transactionCount + 1

                If Not hitVehicles.Exists(queueKey) Then
                    hitVehicles.Add queueKey, True
                End If

            ElseIf Left$(statusText, 5) = "Feil:" Then

                ofvErrorCount = ofvErrorCount + 1

            Else

                noTransactionCount = _
                    noTransactionCount + 1

            End If

            allRows.Add resultRow

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

    '==========================================================
    ' OVERSIKT
    '==========================================================

    stage = "oppdaterer Oversikt"
    API_ShowStatus "Excel", "Oppdaterer Oversikt"

    UpdateOverviewKPIs wsOverview
    RebuildOverviewPivot wsOverview, loResult

    '==========================================================
    ' KONTROLLARK
    '==========================================================

    stage = "oppdaterer Kontroll solgte biler"

    API_ShowStatus _
        "Excel", _
        "Oppdaterer Kontroll solgte biler"

    UpdateControlSheet wsControl

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

Private Function FetchOFVTransactions( _
    ByVal apiKey As String, _
    ByVal identifier As String, _
    ByVal useVin As Boolean, _
    ByVal originalRegNo As String, _
    ByVal originalVin As String, _
    ByVal DateFromISO As String, _
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
        body = body & DateFromISO & ""","
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
            "Accept", "application/json"

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


'==============================================================
' KONTROLL SOLGTE BILER
'==============================================================

Private Sub UpdateControlSheet(ByVal ws As Worksheet)

    Dim loInput As ListObject
    Dim inputData As Variant
    Dim inputCount As Long
    Dim lastRow As Long
    Dim r As Long

    Set loInput = ThisWorkbook.Worksheets( _
        INPUT_SHEET).ListObjects(INPUT_TABLE)

    If Not loInput.DataBodyRange Is Nothing Then

        inputData = loInput.DataBodyRange.Value2

        For r = 1 To UBound(inputData, 1)

            If Len(Trim$(CStr( _
                inputData(r, 1) & vbNullString))) > 0 Or _
               Len(Trim$(CStr( _
                inputData(r, 2) & vbNullString))) > 0 Then

                inputCount = inputCount + 1

            End If

        Next r

    End If

    lastRow = 9 + Application.Max(1, inputCount)

    'Den permanente spillformelen ligger i A10.
    'Den skal aldri slettes eller skrives på nytt.
    If Not ws.Range("A10").HasFormula Then

        Err.Raise vbObjectError + 1100, _
            "UpdateControlSheet", _
            "Den dynamiske formelen mangler i " & _
            CONTROL_SHEET & "!A10."

    End If

    ws.Range("A9").value = _
        "Datakilde naermeste hendelse"

    ws.Range("B9").value = "Regnr / input"
    ws.Range("C9").value = "Chassisnummer"
    ws.Range("D9").value = "Modell"
    ws.Range("E9").value = "Bokfort dato"
    ws.Range("F9").value = "Forstegangsregistrert"
    ws.Range("G9").value = "Naermeste eierskiftedato"
    ws.Range("H9").value = "Naermeste hendelse"
    ws.Range("I9").value = "Dager avvik"
    ws.Range("J9").value = "Kontrollert"
    ws.Range("K9").value = "Selger ved eierskifte"
    ws.Range("L9").value = "Kjoper ved eierskifte"
    ws.Range("M9").value = "Status eierskifte"

    ws.Range("A5").Formula2 = _
        "=SUMPRODUCT(--(((KjoretoyInput[Regnr]<>"""")+" & _
        "(KjoretoyInput[VIN]<>""""))>0))"

    ws.Range("C5").Formula2 = _
        "=COUNTIF(INDEX(A10#,0,10),""Ja"")"

    ws.Range("E5").Formula2 = _
        "=IFERROR(ROWS(FILTER(INDEX(A10#,0,2)," & _
        "(INDEX(A10#,0,9)=0)*" & _
        "(INDEX(A10#,0,10)=""Ja""))),0)"

    ws.Range("I6").Formula2 = _
        "=COUNTIFS(INDEX(A10#,0,9),0," & _
        "INDEX(A10#,0,10),""Ja"")"

    ws.Range("J6").Formula2 = _
        "=COUNTIFS(INDEX(A10#,0,9),"">=1""," & _
        "INDEX(A10#,0,9),""<=15""," & _
        "INDEX(A10#,0,10),""Ja"")"

    ws.Range("K6").Formula2 = _
        "=COUNTIFS(INDEX(A10#,0,9),"">15""," & _
        "INDEX(A10#,0,10),""Ja"")"

    ws.Range("A10").Calculate
    ws.Calculate

    FormatControlSheet ws, lastRow

End Sub


Private Sub FormatControlSheet( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long)

    Dim dataRange As Range
    Dim fullRange As Range

    Set dataRange = ws.Range("A10:M" & lastRow)
    Set fullRange = ws.Range("A9:M" & lastRow)

    With ws.Range("A9:M9")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 34
    End With

    With dataRange
        .Font.Size = 10
        .Font.Color = RGB(31, 31, 31)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .RowHeight = 18
    End With

    With fullRange.Borders
        .LineStyle = xlContinuous
        .Color = RGB(217, 226, 243)
        .Weight = xlThin
    End With

    ws.Range("E10:G" & lastRow).NumberFormat = "dd.mm.yyyy"
    ws.Range("I10:I" & lastRow).NumberFormat = "0"

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


'==============================================================
' OVERSIKT
'==============================================================

Private Sub UpdateOverviewKPIs(ByVal ws As Worksheet)

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
