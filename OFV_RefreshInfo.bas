Option Explicit

' Produksjonsmakro for prosjektet "Python program for bilregistrering":
' tre uavhengige kontroller, hver med egen input-tabell pa Input-arket,
' egen Resultat-fane (full transaksjonshistorikk) og egen kontrollfane.
' Input-arket LESES fra, men aldri skrevet til eller ellers endret av
' denne makroen.
'
' Kort om oppsettet (alle koordinater matcher Input-arkets faktiske
' layout - tre seksjoner ved siden av hverandre):
'   - B1 = OFV API-nokkel (eller navngitt omrade OFV_API)
'   - B2 = Statens Vegvesen API-nokkel (eller navngitt omrade SVV_API)
'
'   Kontroll 1 - "Kontroll solgte biler" (kolonne A-D):
'     B8         = Juridisk enhet (Selger) - organisasjonsnummer
'     A12 og ned = Regnr, B12 og ned = VIN (kun en av de to fylles ut
'                  per rad - koden kjenner dem automatisk fra
'                  hverandre pa lengde, se ErVIN)
'     D12 og ned = Bokfort dato
'
'   Kontroll 2 - "Varekjop bruktbil" (kolonne F-H):
'     G8         = Juridisk enhet (Org Nr) - kjoper, organisasjonsnummer
'     G9 / G10   = Dato fra / Dato til (perioden det sjekkes OFV-kjop i)
'     F12 og ned = Regnr, G12 og ned = VIN
'     H12 og ned = Bokfort dato
'
'   Kontroll 3 - "Kontroll Demobil" (kolonne J-L):
'     K8         = Juridisk enhet - organisasjonsnummer
'     K9 / K10   = Dato fra / Dato til
'     J12 og ned = Regnr, K12 og ned = VIN
'     L12 og ned = Bokfort inn dato
'
' - OFV Transactions API er eneste datakilde for eierskiftehistorikk.
'   Kontroll 1 og 3 henter ETT kall per regnr/VIN som gir hele
'   transaksjonshistorikken (ingen datofilter), sortert nyeste forst.
'   Kontroll 2 henter i tillegg EN liste basert pa organisasjonsnummer
'   (toOrganizationNumber) + datointervall, uavhengig av regnr.
' - Statens Vegvesen (SVV) brukes KUN som reserve i kontroll 1: hvis
'   et kjoretoy ikke har noen transaksjoner i det hele tatt fra OFV,
'   hentes forstegangsregistreringsdatoen fra SVV i stedet. Mangler
'   SVV-nokkelen, hoppes SVV-oppslaget bare over.
' - HTTP-kallene (bade OFV og SVV) bruker WinHttp.WinHttpRequest.5.1.
' - OFV_URL = https://api.ofv.no/transactions/v1/, bekreftet via
'   "Try it"-konsollen i Azure APIM-portalen.
' - Fremdriftsvindu: hvis en UserForm ved navn "frmFremdrift" finnes
'   i prosjektet (med Label-kontroller "lblOFV" og "lblSVV"), vises
'   den som en ikke-blokkerende popup mens makroen kjorer.
' - Knappen "Oppdater" er koblet til Sub OFV_Oppdater, som forst spor
'   (via en enkel InputBox) hvilken av de tre kontrollene som skal
'   kjores (eller alle tre), og deretter kaller riktig delmakro(er).

'==============================================================
' KONFIGURASJON
'==============================================================

Private Const INPUT_SHEET As String = "Input"

' --- Kontroll 1: Kontroll solgte biler ---
Private Const RESULT_SHEET_1 As String = "Resultat"
Private Const CONTROL_SHEET_1 As String = "Kontroll solgte biler"
Private Const RESULT_TABLE_1 As String = "Transaksjoner"

Private Const ORG_CELL_1 As String = "B8"
Private Const FIRST_ROW_1 As Long = 12
Private Const COL_REGNR_1 As Long = 1   ' A
Private Const COL_VIN_1 As Long = 2     ' B
Private Const COL_BOKFORT_1 As Long = 4 ' D

' --- Kontroll 2: Varekjop bruktbil ---
Private Const RESULT_SHEET_2 As String = "Resultat Varekjop"
Private Const CONTROL_SHEET_2 As String = "Kontroll Varekjop Bruktbil"
Private Const RESULT_TABLE_2 As String = "TransaksjonerVarekjop"

Private Const ORG_CELL_2 As String = "G8"
Private Const DATOFRA_CELL_2 As String = "G9"
Private Const DATOTIL_CELL_2 As String = "G10"
Private Const FIRST_ROW_2 As Long = 12
Private Const COL_REGNR_2 As Long = 6   ' F
Private Const COL_VIN_2 As Long = 7     ' G
Private Const COL_BOKFORT_2 As Long = 8 ' H

' --- Kontroll 3: Kontroll Demobil ---
Private Const RESULT_SHEET_3 As String = "Resultat Demobil"
Private Const CONTROL_SHEET_3 As String = "Kontroll Demobil"
Private Const RESULT_TABLE_3 As String = "TransaksjonerDemobil"

Private Const ORG_CELL_3 As String = "K8"
Private Const DATOFRA_CELL_3 As String = "K9"
Private Const DATOTIL_CELL_3 As String = "K10"
Private Const FIRST_ROW_3 As Long = 12
Private Const COL_REGNR_3 As Long = 10   ' J
Private Const COL_VIN_3 As Long = 11     ' K
Private Const COL_BOKFORT_3 As Long = 12 ' L (Bokfort inn dato)

' Bekreftet via "Try it"-konsollen i Azure APIM-portalen
' (https://data.ofv.no/api-details#api=transactions-api-v1&operation=query-transactions):
' POST https://api.ofv.no/transactions/v1/ - dette er den faktiske
' verten APIet ruter pa, uavhengig av at portalens eget domene er
' data.ofv.no.
Private Const OFV_URL As String = _
    "https://api.ofv.no/transactions/v1/"

Private Const SVV_URL As String = _
    "https://akfell-datautlevering.atlas.vegvesen.no/" & _
    "enkeltoppslag/kjoretoydata?"

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

' Grenser for fargekoding av "Dager avvik" i Kontroll solgte biler:
' 0-2 dager = gronn, 3-14 dager = gul/oransje, 15+ dager = rod.
Private Const AVVIK_GRONN_MAX As Long = 2
Private Const AVVIK_GUL_MAX As Long = 14

' RegistreringsType-verdi som (etter avtale) betyr at kjoretoyet
' forlot den registrerte eierens aktive bestand i denne transaksjonen -
' brukes til "avregistrert"-sjekken i Varekjop bruktbil.
Private Const REGTYPE_IKKE_BESTAND As String = _
    "Juridisk eierskifte (ikke i bestand)"

' Instans av UserForm-en "frmFremdrift" (fremdriftsvindu-popup), satt
' av VisFremdriftVindu. Nothing hvis skjemaet ikke finnes/ikke ble
' opprettet - resten av koden sjekker alltid for dette, sa fravaer av
' skjemaet aldri stopper selve API-oppdateringen (kun popup-vinduet
' uteblir).
Private gFremdriftForm As Object

' Satt av knappene i UserForm-en "frmVelgKontroll" (velg-kontroll-
' vinduet), lest av OFV_RefreshInfo etter at vinduet lukkes. Ma vaere
' Public siden frmVelgKontroll sin kode (i et eget skjema-modul)
' skriver til den. "" = avbrutt/lukket uten valg.
Public gValgKontroll As String


' Ren VBA-pause (ingen Win32/kernel32-kall) - Timer er innebygd i
' Excel/VBA. Samme navn og signatur som den gamle Sleep Lib
' "kernel32"-erklaeringen, sa alle eksisterende Sleep-kall i filen
' virker uendret.
Private Sub Sleep(ByVal milliseconds As Long)

    Dim startTime As Double
    Dim elapsedMs As Double

    startTime = Timer

    Do
        DoEvents

        elapsedMs = (Timer - startTime) * 1000

        If elapsedMs < 0 Then
            elapsedMs = elapsedMs + 86400000  ' midnatt-rullering
        End If

    Loop Until elapsedMs >= milliseconds

End Sub


'==============================================================
' HOVEDMAKRO
'==============================================================

' Knappen "Oppdater" er koblet til denne. Apner et lite valgvindu
' (UserForm-en "frmVelgKontroll", hvis den er bygget - se
' fremgangsmate i chatten) der du klikker hvilken kontroll som skal
' kjores. Er ikke skjemaet bygget enna, faller den tilbake til en
' enkel InputBox med samme valg, sa knappen alltid virker.
Public Sub OFV_RefreshInfo()

    Dim valg As String
    Dim velgForm As Object

    gValgKontroll = vbNullString

    On Error Resume Next
    Set velgForm = VBA.UserForms.Add("frmVelgKontroll")
    On Error GoTo 0

    If Not velgForm Is Nothing Then

        velgForm.Show vbModal

        valg = gValgKontroll

        On Error Resume Next
        Unload velgForm
        On Error GoTo 0

    Else

        valg = Trim$(InputBox( _
            "Hvilken kontroll vil du kjore?" & vbCrLf & vbCrLf & _
            "1 = Kontroll solgte biler" & vbCrLf & _
            "2 = Varekjop bruktbil" & vbCrLf & _
            "3 = Kontroll Demobil" & vbCrLf & _
            "4 = Alle tre", _
            "Velg kontroll", "4"))

    End If

    Select Case valg

        Case ""
            ' Avbrutt av bruker - gjor ingenting.

        Case "1"
            KjorKontrollSolgteBiler

        Case "2"
            KjorVarekjopBruktbil

        Case "3"
            KjorKontrollDemobil

        Case "4"
            KjorKontrollSolgteBiler
            KjorVarekjopBruktbil
            KjorKontrollDemobil

        Case Else
            MsgBox "Ugyldig valg: """ & valg & """." & vbCrLf & _
                "Skriv 1, 2, 3 eller 4.", _
                vbExclamation, "Velg kontroll"

    End Select

End Sub


' Tre egne makroer til bruk hvis du heller vil ha en knapp per
' kontroll (i stedet for/i tillegg til InputBox-valget i
' OFV_RefreshInfo over). Koble hver av disse til sin egen knapp i
' Excel - se fremgangsmate i chatten.
Public Sub Knapp_KontrollSolgteBiler()
    KjorKontrollSolgteBiler
End Sub

Public Sub Knapp_VarekjopBruktbil()
    KjorVarekjopBruktbil
End Sub

Public Sub Knapp_KontrollDemobil()
    KjorKontrollDemobil
End Sub


'==============================================================
' KONTROLL 1: KONTROLL SOLGTE BILER
'==============================================================

Private Sub KjorKontrollSolgteBiler()

    Dim wsInput As Worksheet
    Dim wsResult As Worksheet
    Dim wsControl As Worksheet
    Dim loResult As ListObject

    Dim queue As Object
    Dim hitVehicles As Object
    Dim vehicleRowsByKey As Object
    Dim svvInfoByKey As Object
    Dim svvInfo As Object
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
    Dim svvKey As String
    Dim selgerOrgNo As String
    Dim stage As String

    Dim lastInputRow As Long
    Dim oldLastRow As Long
    Dim newLastRow As Long
    Dim fieldCount As Long
    Dim outputRows As Long

    Dim transactionCount As Long
    Dim noTransactionCount As Long
    Dim ofvErrorCount As Long
    Dim svvDateCount As Long
    Dim svvErrorCount As Long
    Dim vehicleHasOkRow As Boolean

    Dim r As Long
    Dim c As Long
    Dim currentVehicle As Long
    Dim totalVehicles As Long

    Dim regNo As String
    Dim vin As String
    Dim inputIdent As String
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
    Set wsResult = GetRequiredSheet(ThisWorkbook, RESULT_SHEET_1)
    Set wsControl = GetRequiredSheet(ThisWorkbook, CONTROL_SHEET_1)

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

    ' Bruker det navngitte omradet OFV_API hvis det finnes i
    ' arbeidsboken, ellers leses nokkelen direkte fra Input!B1.
    ofvKey = Trim$(CStr( _
        ReadConfigValue( _
            ThisWorkbook, "OFV_API", wsInput.Range("B1"))))

    If Len(ofvKey) = 0 Then
        MsgBox _
            "Fant ingen OFV-nokkel. Legg den enten i det " & _
            "navngitte omradet OFV_API, eller direkte i " & _
            "celle B1 pa arket " & INPUT_SHEET & ".", _
            vbExclamation, "API-oppdatering"
        GoTo SafeExit
    End If

    ' SVV er kun en reserve for forstegangsregistrering nar OFV ikke
    ' har noen transaksjoner - mangler nokkelen, hoppes SVV bare over.
    svvKey = Trim$(CStr( _
        ReadConfigValue( _
            ThisWorkbook, "SVV_API", wsInput.Range("B2"))))

    ' Selger-orgnr for denne kontrollen - valgfri (selvhandel-sjekken
    ' hoppes over hvis cellen er tom, se BuildKontrollRow).
    selgerOrgNo = Trim$(CStr( _
        wsInput.Range(ORG_CELL_1).value & vbNullString))

    stage = "leser kjoretoylisten"
    API_ShowStatus "Forbereder", stage

    lastInputRow = LastRowInEitherColumn( _
        wsInput, FIRST_ROW_1, COL_REGNR_1, COL_VIN_1)

    Set queue = CreateObject("Scripting.Dictionary")
    queue.CompareMode = vbTextCompare

    For r = FIRST_ROW_1 To lastInputRow

        inputIdent = ReadInputIdentifier( _
            wsInput, r, COL_REGNR_1, COL_VIN_1)

        regNo = vbNullString
        vin = vbNullString

        If Len(inputIdent) > 0 Then

            If ErVIN(inputIdent) Then
                vin = inputIdent
            Else
                regNo = inputIdent
            End If

        End If

        bokfortValue = wsInput.Cells(r, COL_BOKFORT_1).value

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

    VisFremdriftVindu

    Set allRows = New Collection

    Set hitVehicles = _
        CreateObject("Scripting.Dictionary")

    hitVehicles.CompareMode = vbTextCompare

    Set vehicleRowsByKey = _
        CreateObject("Scripting.Dictionary")

    vehicleRowsByKey.CompareMode = vbTextCompare

    Set kontrollRows = New Collection

    Set svvInfoByKey = CreateObject("Scripting.Dictionary")
    svvInfoByKey.CompareMode = vbTextCompare

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

        vehicleHasOkRow = False

        For Each resultRow In vehicleRows

            statusText = VariantToString( _
                resultRow("Status"))

            If statusText = "OK" Then

                transactionCount = transactionCount + 1
                vehicleHasOkRow = True

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

        ' SVV er kun en reserve: bare nar OFV ikke ga noen brukbar
        ' transaksjon for kjoretoyet, og bare hvis SVV-nokkelen finnes.
        If Not vehicleHasOkRow And Len(svvKey) > 0 Then

            stage = "henter forstegangsregistrering fra Statens vegvesen"

            API_ShowStatus _
                "SVV", _
                "Forstegangsregistrering", _
                identifier, _
                currentVehicle, _
                totalVehicles

            Set svvInfo = FetchVehicleInfoFromSVV(svvKey, regNo, vin)
            Set svvInfoByKey(CStr(key)) = svvInfo

            If svvInfo.Exists("FirstRegistrationDate") Then
                If IsDate(svvInfo("FirstRegistrationDate")) Then
                    svvDateCount = svvDateCount + 1
                End If
            End If

            If Left$(VariantToString(svvInfo("Status")), 5) = _
                "Feil:" Then

                svvErrorCount = svvErrorCount + 1

            End If

            Sleep API_PAUSE_MS

        End If

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

        Set svvInfo = Nothing

        If svvInfoByKey.Exists(CStr(key)) Then
            Set svvInfo = svvInfoByKey(CStr(key))
        End If

        Set kontrollRow = BuildKontrollRow( _
            CStr(vehicleData(0)), _
            CStr(vehicleData(1)), _
            vehicleData(2), _
            vehicleTxRows, _
            svvInfo, _
            selgerOrgNo)

        kontrollRows.Add kontrollRow

    Next key

    '==========================================================
    ' RESULTAT
    '==========================================================

    stage = "oppdaterer Resultat"
    API_ShowStatus "Excel", "Oppdaterer Resultat"

    outputRows = allRows.Count

    Set loResult = GetOrCreateResultTable( _
        wsResult, fieldMap, RESULT_TABLE_1)

    oldLastRow = _
        loResult.Range.Row + _
        loResult.Range.rows.Count - 1

    If outputRows > 0 Then
        newLastRow = outputRows + 1
    Else
        newLastRow = 2
    End If

    Set loResult = ResizeResultTable( _
        wsResult, loResult, newLastRow, fieldCount, RESULT_TABLE_1)

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

        Dim forrigeInputR1 As String
        Dim denneInputR1 As String
        Dim erForsteIGruppeR1 As Boolean

        forrigeInputR1 = vbNullString

        For r = 1 To outputRows

            Set resultRow = allRows(r)

            denneInputR1 = VariantToString(resultRow("Input"))
            erForsteIGruppeR1 = _
                (r = 1 Or denneInputR1 <> forrigeInputR1)

            For c = LBound(fieldMap) To UBound(fieldMap)

                dictionaryKey = CStr(fieldMap(c)(0))

                If dictionaryKey = _
                    "CalculatedSellerType" Or _
                   dictionaryKey = _
                    "CalculatedBuyerType" Then

                    output(r, c + 1) = Empty

                ElseIf Not erForsteIGruppeR1 And _
                    IsCarLevelField(dictionaryKey) Then

                    ' Kun nyeste transaksjon (forste rad i gruppen)
                    ' viser bilinfo - eldre transaksjoner for samme
                    ' bil viser bare transaksjonsspesifikk info.
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

            forrigeInputR1 = denneInputR1

        Next r

        loResult.DataBodyRange.value = output
        ApplyCalculatedColumns loResult

    End If

    FormatResultTable wsResult, loResult
    FormatResultTableGrouping wsResult, loResult, allRows

    '==========================================================
    ' KONTROLLARK
    '==========================================================

    stage = "oppdaterer Kontroll solgte biler"

    API_ShowStatus _
        "Excel", _
        "Oppdaterer Kontroll solgte biler"

    UpdateControlSheet wsControl, kontrollRows, totalVehicles

    Application.Calculation = oldCalculation

    If oldCalculation = xlCalculationManual Then

        wsResult.Calculate
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

    SkjulFremdriftVindu

    MsgBox _
        "Oppdateringen er ferdig." & vbCrLf & vbCrLf & _
        totalVehicles & " kjoretoy lest." & vbCrLf & _
        transactionCount & _
        " OFV-eierskifter funnet." & vbCrLf & _
        noTransactionCount & _
        " uten OFV-eierskifter." & vbCrLf & _
        ofvErrorCount & " OFV-feil." & vbCrLf & _
        svvDateCount & _
        " forstegangsregistreringer hentet fra SVV (reserve)." & vbCrLf & _
        svvErrorCount & " SVV-feil.", _
        vbInformation, "API-oppdatering"

    Exit Sub

SafeExit:

    Application.StatusBar = False
    SkjulFremdriftVindu
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

    SkjulFremdriftVindu

    MsgBox _
        "Oppdateringen ble avbrutt." & vbCrLf & vbCrLf & _
        "Trinn: " & stage & vbCrLf & _
        "Feil " & errorNumber & ": " & _
        errorDescription, _
        vbCritical, "API-oppdatering"

End Sub


'==============================================================
' KONTROLL 2: VAREKJOP BRUKTBIL
'==============================================================

' Henter FORST alle biler OFV sier er kjopt (toOrganizationNumber) av
' juridisk enhet G8 i perioden G9-G10 - ett samlekall, uavhengig av
' regnr. Deretter hentes hele egen transaksjonshistorikk for hver bil
' i input-listen (F:H), for a sjekke om OFV-listen stemmer med
' bokforingen og om bilen senere er avregistrert (solgt ut av
' bestand) av samme juridisk enhet.
Private Sub KjorVarekjopBruktbil()

    Dim wsInput As Worksheet
    Dim wsResult As Worksheet
    Dim wsControl As Worksheet
    Dim loResult As ListObject

    Dim queue As Object
    Dim vehicleRowsByKey As Object
    Dim ofvListeRegNr As Object
    Dim ofvListeVin As Object
    Dim resultRow As Object
    Dim kontrollRow As Object

    Dim allRows As Collection
    Dim vehicleRows As Collection
    Dim kontrollRows As Collection
    Dim vehicleTxRows As Collection
    Dim ofvListeRader As Collection
    Dim manglendeIBokforing As Collection

    Dim fieldMap As Variant
    Dim vehicleData As Variant
    Dim bokfortValue As Variant
    Dim output() As Variant

    Dim ofvKey As String
    Dim buyerOrgNo As String
    Dim buyerOrgName As String
    Dim dateFraRaw As Variant
    Dim dateTilRaw As Variant
    Dim stage As String

    Dim lastInputRow As Long
    Dim oldLastRow As Long
    Dim newLastRow As Long
    Dim fieldCount As Long
    Dim outputRows As Long

    Dim r As Long
    Dim c As Long
    Dim currentVehicle As Long
    Dim totalVehicles As Long

    Dim regNo As String
    Dim vin As String
    Dim inputIdent As String
    Dim queueKey As String
    Dim identifier As String
    Dim dictionaryKey As String

    Dim useVin As Boolean
    Dim erIOFVListe As Boolean
    Dim finnesIInput As Boolean
    Dim ofvRegNorm As String
    Dim ofvVinNorm As String

    Dim key As Variant
    Dim value As Variant
    Dim item As Variant

    Dim forrigeInputR2 As String
    Dim denneInputR2 As String
    Dim erForsteIGruppeR2 As Boolean

    Dim errorNumber2 As Long
    Dim errorDescription2 As String

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldCursor As Variant
    Dim applicationChanged As Boolean

    On Error GoTo FatalError2

    stage = "finner arkene (Varekjop bruktbil)"
    API_ShowStatus "Forbereder", stage

    Set wsInput = GetRequiredSheet(ThisWorkbook, INPUT_SHEET)
    Set wsResult = GetRequiredSheet(ThisWorkbook, RESULT_SHEET_2)
    Set wsControl = GetRequiredSheet(ThisWorkbook, CONTROL_SHEET_2)

    If wsResult.ProtectContents Then
        Err.Raise vbObjectError + 1100, , _
            RESULT_SHEET_2 & "-arket er beskyttet."
    End If

    If wsControl.ProtectContents Then
        Err.Raise vbObjectError + 1101, , _
            CONTROL_SHEET_2 & " er beskyttet."
    End If

    stage = "leser API-nokkel"

    ofvKey = Trim$(CStr( _
        ReadConfigValue(ThisWorkbook, "OFV_API", wsInput.Range("B1"))))

    If Len(ofvKey) = 0 Then
        MsgBox "Fant ingen OFV-nokkel. Legg den i celle B1 pa " & _
            INPUT_SHEET & ".", vbExclamation, "Varekjop bruktbil"
        GoTo SafeExit2
    End If

    buyerOrgNo = Trim$(CStr(wsInput.Range(ORG_CELL_2).value & vbNullString))

    If Len(buyerOrgNo) = 0 Then
        MsgBox "Fyll ut Juridisk enhet (Org Nr) i celle " & _
            ORG_CELL_2 & " for Varekjop bruktbil.", _
            vbExclamation, "Varekjop bruktbil"
        GoTo SafeExit2
    End If

    dateFraRaw = TolkBokfortDato(wsInput.Range(DATOFRA_CELL_2).value)
    dateTilRaw = TolkBokfortDato(wsInput.Range(DATOTIL_CELL_2).value)

    If IsEmpty(dateFraRaw) Or IsEmpty(dateTilRaw) Then
        MsgBox "Fyll ut gyldig Dato fra / Dato til (" & _
            DATOFRA_CELL_2 & "/" & DATOTIL_CELL_2 & _
            ") for Varekjop bruktbil.", vbExclamation, "Varekjop bruktbil"
        GoTo SafeExit2
    End If

    lastInputRow = LastRowInEitherColumn( _
        wsInput, FIRST_ROW_2, COL_REGNR_2, COL_VIN_2)

    Set queue = CreateObject("Scripting.Dictionary")
    queue.CompareMode = vbTextCompare

    For r = FIRST_ROW_2 To lastInputRow

        inputIdent = ReadInputIdentifier( _
            wsInput, r, COL_REGNR_2, COL_VIN_2)

        regNo = vbNullString
        vin = vbNullString

        If Len(inputIdent) > 0 Then
            If ErVIN(inputIdent) Then
                vin = inputIdent
            Else
                regNo = inputIdent
            End If
        End If

        bokfortValue = wsInput.Cells(r, COL_BOKFORT_2).value

        queueKey = BuildVehicleKey(vin, regNo)

        If Len(queueKey) > 0 Then
            If Not queue.Exists(queueKey) Then
                queue.Add queueKey, Array(regNo, vin, bokfortValue)
            End If
        End If

    Next r

    If queue.Count = 0 Then
        MsgBox "Fant ingen registreringsnummer eller VIN i " & _
            "Varekjop bruktbil-listen.", vbInformation, "Varekjop bruktbil"
        GoTo SafeExit2
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

    VisFremdriftVindu

    Set allRows = New Collection
    Set vehicleRowsByKey = CreateObject("Scripting.Dictionary")
    vehicleRowsByKey.CompareMode = vbTextCompare
    Set kontrollRows = New Collection
    Set manglendeIBokforing = New Collection

    Set ofvListeRegNr = CreateObject("Scripting.Dictionary")
    ofvListeRegNr.CompareMode = vbTextCompare
    Set ofvListeVin = CreateObject("Scripting.Dictionary")
    ofvListeVin.CompareMode = vbTextCompare

    fieldMap = GetFieldMap()
    fieldCount = UBound(fieldMap) + 1
    totalVehicles = queue.Count

    '======================================================
    ' Steg 1: hent listen over ALLE biler OFV sier er kjopt av
    ' dette orgnr i perioden - KUN orgnr + dato, uten regnr.
    '======================================================

    stage = "henter kjopsliste fra OFV (orgnr + periode)"
    API_ShowStatus "OFV", "Henter kjopsliste for org " & buyerOrgNo

    Set ofvListeRader = FetchOFVTransactionsByBuyerOrg( _
        ofvKey, buyerOrgNo, CDate(dateFraRaw), CDate(dateTilRaw))

    buyerOrgName = vbNullString

    For Each item In ofvListeRader

        If Len(VariantToString(item("RegNo"))) > 0 Then

            If Not ofvListeRegNr.Exists( _
                NormalizeIdentifier(item("RegNo"))) Then

                ofvListeRegNr.Add NormalizeIdentifier(item("RegNo")), True

            End If

        End If

        If Len(VariantToString(item("ChassisNumber"))) > 0 Then

            If Not ofvListeVin.Exists( _
                NormalizeIdentifier(item("ChassisNumber"))) Then

                ofvListeVin.Add _
                    NormalizeIdentifier(item("ChassisNumber")), True

            End If

        End If

        If Len(buyerOrgName) = 0 Then

            If NormalizeIdentifier(item("ToOwnerOrgNo")) = _
                NormalizeIdentifier(buyerOrgNo) Then

                buyerOrgName = VariantToString(item("ToOwnerCompanyName"))

            End If

        End If

    Next item

    '======================================================
    ' Steg 2: full transaksjonshistorikk per bil i input-listen
    '======================================================

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

        API_ShowStatus "OFV", "Eierskifter", identifier, _
            currentVehicle, totalVehicles

        Set vehicleRows = FetchOFVTransactions( _
            ofvKey, identifier, useVin, regNo, vin)

        If Not vehicleRowsByKey.Exists(CStr(key)) Then
            vehicleRowsByKey.Add CStr(key), New Collection
        End If

        For Each resultRow In vehicleRows
            allRows.Add resultRow
            vehicleRowsByKey(CStr(key)).Add resultRow
        Next resultRow

        Sleep API_PAUSE_MS

    Next key

    For Each key In queue.Keys

        vehicleData = queue(key)

        Set vehicleTxRows = Nothing
        If vehicleRowsByKey.Exists(CStr(key)) Then
            Set vehicleTxRows = vehicleRowsByKey(CStr(key))
        End If

        erIOFVListe = False

        If Len(CStr(vehicleData(0))) > 0 Then
            If ofvListeRegNr.Exists( _
                NormalizeIdentifier(CStr(vehicleData(0)))) Then
                erIOFVListe = True
            End If
        End If

        If Len(CStr(vehicleData(1))) > 0 Then
            If ofvListeVin.Exists( _
                NormalizeIdentifier(CStr(vehicleData(1)))) Then
                erIOFVListe = True
            End If
        End If

        Set kontrollRow = BuildVarekjopRow( _
            CStr(vehicleData(0)), CStr(vehicleData(1)), _
            vehicleData(2), vehicleTxRows, erIOFVListe, buyerOrgNo)

        kontrollRows.Add kontrollRow

    Next key

    ' Ekstra: biler OFV sier er kjopt av selskapet i perioden, men som
    ' IKKE finnes i det hele tatt i bokforingslisten (input).
    For Each item In ofvListeRader

        ofvRegNorm = NormalizeIdentifier(item("RegNo"))
        ofvVinNorm = NormalizeIdentifier(item("ChassisNumber"))
        finnesIInput = False

        For Each key In queue.Keys

            vehicleData = queue(key)

            If (Len(ofvRegNorm) > 0 And _
                NormalizeIdentifier(CStr(vehicleData(0))) = ofvRegNorm) Or _
               (Len(ofvVinNorm) > 0 And _
                NormalizeIdentifier(CStr(vehicleData(1))) = ofvVinNorm) Then

                finnesIInput = True
                Exit For

            End If

        Next key

        If Not finnesIInput Then manglendeIBokforing.Add item

    Next item

    '======================================================
    ' RESULTAT
    '======================================================

    stage = "oppdaterer " & RESULT_SHEET_2
    API_ShowStatus "Excel", stage

    outputRows = allRows.Count

    Set loResult = GetOrCreateResultTable(wsResult, fieldMap, RESULT_TABLE_2)

    oldLastRow = loResult.Range.Row + loResult.Range.rows.Count - 1

    If outputRows > 0 Then
        newLastRow = outputRows + 1
    Else
        newLastRow = 2
    End If

    Set loResult = ResizeResultTable( _
        wsResult, loResult, newLastRow, fieldCount, RESULT_TABLE_2)

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
        loResult.HeaderRowRange.Cells(1, c + 1).value = fieldMap(c)(1)
    Next c

    If outputRows > 0 Then

        ReDim output(1 To outputRows, 1 To fieldCount)

        forrigeInputR2 = vbNullString

        For r = 1 To outputRows

            Set resultRow = allRows(r)

            denneInputR2 = VariantToString(resultRow("Input"))
            erForsteIGruppeR2 = (r = 1 Or denneInputR2 <> forrigeInputR2)

            For c = LBound(fieldMap) To UBound(fieldMap)

                dictionaryKey = CStr(fieldMap(c)(0))

                If dictionaryKey = "CalculatedSellerType" Or _
                   dictionaryKey = "CalculatedBuyerType" Then

                    output(r, c + 1) = Empty

                ElseIf Not erForsteIGruppeR2 And _
                    IsCarLevelField(dictionaryKey) Then

                    output(r, c + 1) = Empty

                ElseIf resultRow.Exists(dictionaryKey) Then

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

            forrigeInputR2 = denneInputR2

        Next r

        loResult.DataBodyRange.value = output
        ApplyCalculatedColumns loResult

    End If

    FormatResultTable wsResult, loResult
    FormatResultTableGrouping wsResult, loResult, allRows

    '======================================================
    ' KONTROLLARK
    '======================================================

    stage = "oppdaterer " & CONTROL_SHEET_2
    API_ShowStatus "Excel", stage

    If Len(buyerOrgName) = 0 Then buyerOrgName = buyerOrgNo

    UpdateVarekjopControlSheet wsControl, kontrollRows, _
        manglendeIBokforing, totalVehicles, buyerOrgNo, buyerOrgName, _
        CDate(dateFraRaw), CDate(dateTilRaw)

    Application.Calculation = oldCalculation

    If oldCalculation = xlCalculationManual Then
        wsResult.Calculate
        wsControl.Calculate
    Else
        Application.CalculateFull
    End If

    API_ShowStatus "Ferdig", "Varekjop bruktbil er oppdatert"

    RestoreApplicationState oldScreenUpdating, oldEnableEvents, _
        oldCalculation, oldCursor

    applicationChanged = False

    SkjulFremdriftVindu

    MsgBox "Varekjop bruktbil er oppdatert." & vbCrLf & vbCrLf & _
        totalVehicles & " biler lest fra input." & vbCrLf & _
        ofvListeRader.Count & " OFV-transaksjoner funnet for org " & _
        buyerOrgNo & " i perioden." & vbCrLf & _
        manglendeIBokforing.Count & _
        " biler OFV viser kjopt, men som mangler i bokforingslisten.", _
        vbInformation, "Varekjop bruktbil"

    Exit Sub

SafeExit2:
    Application.StatusBar = False
    SkjulFremdriftVindu
    Exit Sub

FatalError2:

    errorNumber2 = Err.Number
    errorDescription2 = Err.Description

    Application.StatusBar = False

    If applicationChanged Then

        RestoreApplicationState oldScreenUpdating, oldEnableEvents, _
            oldCalculation, oldCursor

    End If

    SkjulFremdriftVindu

    MsgBox "Varekjop bruktbil ble avbrutt." & vbCrLf & vbCrLf & _
        "Trinn: " & stage & vbCrLf & _
        "Feil " & errorNumber2 & ": " & errorDescription2, _
        vbCritical, "Varekjop bruktbil"

End Sub


' Henter ALLE OFV-transaksjoner der gitt orgnr star som KJOPER
' (toOrganizationNumber), innenfor et datointervall - ett samlekall
' uavhengig av regnr/VIN, med cursor-paginering akkurat som
' FetchOFVTransactions. Feil her stopper hele Varekjop bruktbil-
' kjoringen (se FatalError2), i motsetning til per-kjoretoy-kallene
' som heller skriver en feilrad og fortsetter.
Private Function FetchOFVTransactionsByBuyerOrg( _
    ByVal apiKey As String, _
    ByVal orgNo As String, _
    ByVal dateFra As Date, _
    ByVal dateTil As Date) As Collection

    Dim rows As New Collection
    Dim items As Collection

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

    cursor = vbNullString

    Do

        body = "{""filters"":{"
        body = body & """toOrganizationNumber"":"""
        body = body & JsonEscape(orgNo) & ""","
        body = body & """transactionDateFrom"":"""
        body = body & Format$(dateFra, "yyyy-mm-dd") & ""","
        body = body & """transactionDateTo"":"""
        body = body & Format$(dateTil, "yyyy-mm-dd") & """},"
        body = body & """pagination"":{""first"":1000"

        If Len(cursor) > 0 Then
            body = body & ",""cursor"":"""
            body = body & JsonEscape(cursor) & """"
        End If

        body = body & "},"
        body = body & """sorting"":{"
        body = body & """orderBy"":""transactionDate"","
        body = body & """orderDirection"":""DESC""}}"

        responseText = PostOFVWithRetries(apiKey, body, statusText)

        If statusText <> "OK" Then

            Err.Raise vbObjectError + 1120, _
                "FetchOFVTransactionsByBuyerOrg", statusText

        End If

        transactionsJSON = JSON_ExtractObject(responseText, "transactions")
        Set items = JSON_ArrayAllElements(transactionsJSON)

        For Each item In items

            rows.Add BuildTransactionRow( _
                vbNullString, CStr(item), False, _
                vbNullString, vbNullString)

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

            Sleep API_PAUSE_MS

        Else
            Exit Do
        End If

    Loop

    Set FetchOFVTransactionsByBuyerOrg = rows

End Function


' Bygger kontrollraden for en bil i Varekjop bruktbil-listen.
' erIOFVListe: True hvis regnr/VIN ble funnet i OFV sin liste over
' biler kjopt av buyerOrgNo i perioden (fra det orgnr-baserte
' samlekallet - IKKE fra bilens egen transaksjonshistorikk).
Private Function BuildVarekjopRow( _
    ByVal regNo As String, _
    ByVal vin As String, _
    ByVal bokfortRaw As Variant, _
    ByVal vehicleTxRows As Collection, _
    ByVal erIOFVListe As Boolean, _
    ByVal buyerOrgNo As String) As Object

    Dim result As Object
    Dim txRow As Variant
    Dim bokfortDate As Variant
    Dim modelName As String
    Dim chassisNo As String
    Dim regNoResolved As String
    Dim buyerOrgNormalisert As String

    Dim kjoptTxRow As Object
    Dim avregistrertTxRow As Object

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    bokfortDate = TolkBokfortDato(bokfortRaw)
    modelName = vbNullString
    chassisNo = vin
    regNoResolved = regNo
    buyerOrgNormalisert = NormalizeIdentifier(buyerOrgNo)

    Set kjoptTxRow = Nothing
    Set avregistrertTxRow = Nothing

    If Not vehicleTxRows Is Nothing Then

        For Each txRow In vehicleTxRows

            If VariantToString(txRow("Status")) = "OK" Then

                If Len(modelName) = 0 Then
                    modelName = VariantToString(txRow("ModelName"))
                End If

                If Len(VariantToString(txRow("ChassisNumber"))) > 0 Then
                    chassisNo = VariantToString(txRow("ChassisNumber"))
                End If

                If Len(VariantToString(txRow("RegNo"))) > 0 Then
                    regNoResolved = VariantToString(txRow("RegNo"))
                End If

                ' Kjopt av juridisk enhet: en transaksjon der
                ' selskapet star som kjoper (til-siden). Bruker den
                ' seneste hvis flere.
                If NormalizeIdentifier(txRow("ToOwnerOrgNo")) = _
                    buyerOrgNormalisert Then

                    If kjoptTxRow Is Nothing Then

                        Set kjoptTxRow = txRow

                    ElseIf IsDate(txRow("TransactionDate")) And _
                        IsDate(kjoptTxRow("TransactionDate")) Then

                        If CDate(txRow("TransactionDate")) > _
                            CDate(kjoptTxRow("TransactionDate")) Then

                            Set kjoptTxRow = txRow

                        End If

                    End If

                End If

                ' Avregistrert av juridisk enhet: en transaksjon der
                ' selskapet star som selger (fra-siden) OG
                ' registreringstypen viser at bilen forlot bestanden.
                ' Bruker den tidligste slike (forste avregistrering
                ' etter kjop).
                If NormalizeIdentifier(txRow("FromOwnerOrgNo")) = _
                    buyerOrgNormalisert And _
                    VariantToString(txRow("RegistrationType")) = _
                    REGTYPE_IKKE_BESTAND Then

                    If avregistrertTxRow Is Nothing Then

                        Set avregistrertTxRow = txRow

                    ElseIf IsDate(txRow("TransactionDate")) And _
                        IsDate(avregistrertTxRow("TransactionDate")) Then

                        If CDate(txRow("TransactionDate")) < _
                            CDate(avregistrertTxRow("TransactionDate")) Then

                            Set avregistrertTxRow = txRow

                        End If

                    End If

                End If

            End If

        Next txRow

    End If

    result("RegnrInput") = regNoResolved
    result("Chassisnummer") = chassisNo
    result("Modell") = modelName
    result("BokfortDato") = bokfortDate
    result("IOFVListe") = IIf(erIOFVListe, "Ja", "Nei")
    result("IBokfort") = IIf(IsEmpty(bokfortDate), "Nei", "Ja")

    If Not kjoptTxRow Is Nothing Then
        result("KjoptDato") = kjoptTxRow("TransactionDate")
        result("KjoptAvEnhet") = "Ja"
    Else
        result("KjoptDato") = Empty
        result("KjoptAvEnhet") = "Nei"
    End If

    If Not avregistrertTxRow Is Nothing Then

        result("AvregistrertDato") = avregistrertTxRow("TransactionDate")
        result("Avregistrert") = "Ja"

    Else

        result("AvregistrertDato") = Empty

        If kjoptTxRow Is Nothing Then
            result("Avregistrert") = vbNullString
        Else
            result("Avregistrert") = "Nei"
        End If

    End If

    If erIOFVListe And IsEmpty(bokfortDate) Then

        result("Status") = "Avvik: OFV viser kjop, mangler i bokforing"

    ElseIf Not erIOFVListe And Not IsEmpty(bokfortDate) Then

        result("Status") = _
            "Avvik: Bokfort, ikke bekreftet kjopt av OFV i perioden"

    ElseIf erIOFVListe And Not IsEmpty(bokfortDate) And _
        result("Avregistrert") = "Ja" Then

        result("Status") = "OK - kjopt og avregistrert"

    ElseIf erIOFVListe And Not IsEmpty(bokfortDate) Then

        result("Status") = "OK - kjopt, ikke avregistrert enna"

    Else
        result("Status") = "Ingen treff"
    End If

    Set BuildVarekjopRow = result

End Function


' Bygger hele arket Kontroll Varekjop Bruktbil pa nytt hver kjoring.
Private Sub UpdateVarekjopControlSheet( _
    ByVal ws As Worksheet, _
    ByVal kontrollRows As Collection, _
    ByVal manglendeIBokforing As Collection, _
    ByVal totalVehicles As Long, _
    ByVal buyerOrgNo As String, _
    ByVal buyerOrgName As String, _
    ByVal dateFra As Date, _
    ByVal dateTil As Date)

    Const HEADER_ROW As Long = 9
    Const FIRST_DATA_ROW As Long = 10

    Dim row As Object
    Dim item As Variant
    Dim r As Long
    Dim lastRow As Long
    Dim antallOk As Long
    Dim antallAvvik As Long
    Dim statusText As String

    ws.Cells.Clear

    ws.Range("A1:J1").Merge
    ws.Range("A1").value = ws.Name

    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ws.rows(1).RowHeight = 26

    ws.Range("A2:J2").Merge
    ws.Range("A2").value = _
        "Kontrollregel: Alle biler OFV sier er kjopt (" & _
        buyerOrgName & ", orgnr " & buyerOrgNo & ") i perioden " & _
        Format$(dateFra, "dd.mm.yyyy") & " - " & _
        Format$(dateTil, "dd.mm.yyyy") & _
        " sjekkes mot bokforingslisten (Regnr/VIN + Bokfort dato)."

    ws.Range("A3:J3").Merge
    ws.Range("A3").value = _
        "En bil regnes som bekreftet varekjop nar den bade er " & _
        "bokfort OG senere avregistrert (solgt ut av bestand) av " & _
        "samme juridiske enhet - se kolonnene KjoptDato/" & _
        "AvregistrertDato/Avregistrert og Status."

    ws.Range("A2:A3").Font.Italic = True
    ws.rows("2:3").RowHeight = 15

    For Each row In kontrollRows

        statusText = VariantToString(row("Status"))

        If Left$(statusText, 2) = "OK" Then
            antallOk = antallOk + 1
        ElseIf Left$(statusText, 5) = "Avvik" Then
            antallAvvik = antallAvvik + 1
        End If

    Next row

    ws.Range("A5:B5").Merge : ws.Range("A5").value = "Inputbiler"
    ws.Range("C5:D5").Merge : ws.Range("C5").value = "OK"
    ws.Range("E5:F5").Merge : ws.Range("E5").value = "Avvik"
    ws.Range("G5:J5").Merge
    ws.Range("G5").value = "OFV-kjop uten bokforing"

    With ws.Range("A5:J5")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("A6:B6").Merge : ws.Range("A6").value = totalVehicles
    ws.Range("C6:D6").Merge : ws.Range("C6").value = antallOk
    ws.Range("E6:F6").Merge : ws.Range("E6").value = antallAvvik
    ws.Range("G6:J6").Merge
    ws.Range("G6").value = manglendeIBokforing.Count

    With ws.Range("A6:B6")
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("C6:D6")
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = COLOR_GREEN_FILL
        .Font.Color = COLOR_GREEN_FONT
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("E6:F6")
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = COLOR_RED_FILL
        .Font.Color = COLOR_RED_FONT
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("G6:J6")
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = COLOR_YELLOW_FILL
        .Font.Color = COLOR_YELLOW_FONT
        .HorizontalAlignment = xlCenter
    End With

    ws.rows("5:6").RowHeight = 20

    ws.Range("A" & HEADER_ROW).value = "Regnr"
    ws.Range("B" & HEADER_ROW).value = "Chassisnummer"
    ws.Range("C" & HEADER_ROW).value = "Modell"
    ws.Range("D" & HEADER_ROW).value = "Bokfort dato"
    ws.Range("E" & HEADER_ROW).value = "I OFV-liste"
    ws.Range("F" & HEADER_ROW).value = "Kjopt dato"
    ws.Range("G" & HEADER_ROW).value = "Avregistrert dato"
    ws.Range("H" & HEADER_ROW).value = "Avregistrert"
    ws.Range("I" & HEADER_ROW).value = "Status"

    With ws.Range("A" & HEADER_ROW & ":I" & HEADER_ROW)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30
    End With

    r = FIRST_DATA_ROW

    For Each row In kontrollRows

        ws.Cells(r, 1).value = VariantToString(row("RegnrInput"))
        ws.Cells(r, 2).value = VariantToString(row("Chassisnummer"))
        ws.Cells(r, 3).value = VariantToString(row("Modell"))
        ws.Cells(r, 4).value = row("BokfortDato")
        ws.Cells(r, 5).value = VariantToString(row("IOFVListe"))
        ws.Cells(r, 6).value = row("KjoptDato")
        ws.Cells(r, 7).value = row("AvregistrertDato")
        ws.Cells(r, 8).value = VariantToString(row("Avregistrert"))
        ws.Cells(r, 9).value = VariantToString(row("Status"))

        statusText = VariantToString(row("Status"))

        If Left$(statusText, 2) = "OK" Then

            With ws.Range(ws.Cells(r, 9), ws.Cells(r, 9))
                .Interior.Color = COLOR_GREEN_FILL
                .Font.Color = COLOR_GREEN_FONT
            End With

        ElseIf Left$(statusText, 5) = "Avvik" Then

            With ws.Range(ws.Cells(r, 9), ws.Cells(r, 9))
                .Interior.Color = COLOR_RED_FILL
                .Font.Color = COLOR_RED_FONT
            End With

        End If

        r = r + 1

    Next row

    lastRow = r - 1
    If lastRow < FIRST_DATA_ROW Then lastRow = FIRST_DATA_ROW

    ws.Range("D" & FIRST_DATA_ROW & ":D" & lastRow).NumberFormat = _
        "dd.mm.yyyy"
    ws.Range("F" & FIRST_DATA_ROW & ":G" & lastRow).NumberFormat = _
        "dd.mm.yyyy"

    ' Ekstra blokk: biler OFV sier er kjopt av selskapet i perioden,
    ' men som ikke finnes i det hele tatt i bokforingslisten.
    r = lastRow + 3

    ws.Range("A" & r & ":I" & r).Merge
    ws.Range("A" & r).value = _
        "Biler OFV viser kjopt av " & buyerOrgName & _
        ", men som mangler i bokforingslisten"

    With ws.Range("A" & r)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With

    r = r + 1

    ws.Range("A" & r).value = "Regnr"
    ws.Range("B" & r).value = "Chassisnummer"
    ws.Range("C" & r).value = "Kjopt dato"
    ws.Range("D" & r).value = "Selger"

    With ws.Range("A" & r & ":D" & r)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
    End With

    r = r + 1

    For Each item In manglendeIBokforing

        ws.Cells(r, 1).value = VariantToString(item("RegNo"))
        ws.Cells(r, 2).value = VariantToString(item("ChassisNumber"))
        ws.Cells(r, 3).value = item("TransactionDate")

        ws.Cells(r, 4).value = ComputeOwnerLabel( _
            VariantToString(item("FromOwnerType")), _
            VariantToString(item("FromOwnerCompanyName")))

        r = r + 1

    Next item

    If r > FIRST_DATA_ROW Then
        ws.Range("C" & (r - manglendeIBokforing.Count) & ":C" & _
            (r - 1)).NumberFormat = "dd.mm.yyyy"
    End If

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D:G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 34

End Sub


'==============================================================
' KONTROLL 3: KONTROLL DEMOBIL
'==============================================================

' For hver bil i input-listen (J:L) hentes hele transaksjonshistorikken
' (samme per-kjoretoy-kall som Kontroll solgte biler). Kontrollen sjekker
' om SISTE registrerte eierskifte i historikken har juridisk enhet (K8)
' som kjoper - altså at bilen fremdeles star registrert pa selskapet som
' demobil, uten noe salg etterpa.
Private Sub KjorKontrollDemobil()

    Dim wsInput As Worksheet
    Dim wsResult As Worksheet
    Dim wsControl As Worksheet
    Dim loResult As ListObject

    Dim queue As Object
    Dim vehicleRowsByKey As Object
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
    Dim companyOrgNo As String
    Dim dateFraRaw As Variant
    Dim dateTilRaw As Variant
    Dim stage As String

    Dim lastInputRow As Long
    Dim oldLastRow As Long
    Dim newLastRow As Long
    Dim fieldCount As Long
    Dim outputRows As Long

    Dim r As Long
    Dim c As Long
    Dim currentVehicle As Long
    Dim totalVehicles As Long

    Dim regNo As String
    Dim vin As String
    Dim inputIdent As String
    Dim queueKey As String
    Dim identifier As String
    Dim dictionaryKey As String

    Dim useVin As Boolean
    Dim key As Variant
    Dim value As Variant

    Dim forrigeInputR3 As String
    Dim denneInputR3 As String
    Dim erForsteIGruppeR3 As Boolean

    Dim errorNumber3 As Long
    Dim errorDescription3 As String

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldCursor As Variant
    Dim applicationChanged As Boolean

    On Error GoTo FatalError3

    stage = "finner arkene (Kontroll Demobil)"
    API_ShowStatus "Forbereder", stage

    Set wsInput = GetRequiredSheet(ThisWorkbook, INPUT_SHEET)
    Set wsResult = GetRequiredSheet(ThisWorkbook, RESULT_SHEET_3)
    Set wsControl = GetRequiredSheet(ThisWorkbook, CONTROL_SHEET_3)

    If wsResult.ProtectContents Then
        Err.Raise vbObjectError + 1200, , _
            RESULT_SHEET_3 & "-arket er beskyttet."
    End If

    If wsControl.ProtectContents Then
        Err.Raise vbObjectError + 1201, , _
            CONTROL_SHEET_3 & " er beskyttet."
    End If

    stage = "leser API-nokkel"

    ofvKey = Trim$(CStr( _
        ReadConfigValue(ThisWorkbook, "OFV_API", wsInput.Range("B1"))))

    If Len(ofvKey) = 0 Then
        MsgBox "Fant ingen OFV-nokkel. Legg den i celle B1 pa " & _
            INPUT_SHEET & ".", vbExclamation, "Kontroll Demobil"
        GoTo SafeExit3
    End If

    companyOrgNo = Trim$(CStr(wsInput.Range(ORG_CELL_3).value & vbNullString))

    If Len(companyOrgNo) = 0 Then
        MsgBox "Fyll ut Juridisk enhet i celle " & ORG_CELL_3 & _
            " for Kontroll Demobil.", vbExclamation, "Kontroll Demobil"
        GoTo SafeExit3
    End If

    ' Dato fra/til brukes kun til visning i kontrollarket - selve
    ' sjekken bruker alltid bilens NYESTE transaksjon, uansett dato,
    ' siden hvert kjoretoy hentes med ett ufiltrert kall.
    dateFraRaw = TolkBokfortDato(wsInput.Range(DATOFRA_CELL_3).value)
    dateTilRaw = TolkBokfortDato(wsInput.Range(DATOTIL_CELL_3).value)

    lastInputRow = LastRowInEitherColumn( _
        wsInput, FIRST_ROW_3, COL_REGNR_3, COL_VIN_3)

    Set queue = CreateObject("Scripting.Dictionary")
    queue.CompareMode = vbTextCompare

    For r = FIRST_ROW_3 To lastInputRow

        inputIdent = ReadInputIdentifier( _
            wsInput, r, COL_REGNR_3, COL_VIN_3)

        regNo = vbNullString
        vin = vbNullString

        If Len(inputIdent) > 0 Then
            If ErVIN(inputIdent) Then
                vin = inputIdent
            Else
                regNo = inputIdent
            End If
        End If

        bokfortValue = wsInput.Cells(r, COL_BOKFORT_3).value

        queueKey = BuildVehicleKey(vin, regNo)

        If Len(queueKey) > 0 Then
            If Not queue.Exists(queueKey) Then
                queue.Add queueKey, Array(regNo, vin, bokfortValue)
            End If
        End If

    Next r

    If queue.Count = 0 Then
        MsgBox "Fant ingen registreringsnummer eller VIN i " & _
            "Kontroll Demobil-listen.", vbInformation, "Kontroll Demobil"
        GoTo SafeExit3
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

    VisFremdriftVindu

    Set allRows = New Collection
    Set vehicleRowsByKey = CreateObject("Scripting.Dictionary")
    vehicleRowsByKey.CompareMode = vbTextCompare
    Set kontrollRows = New Collection

    fieldMap = GetFieldMap()
    fieldCount = UBound(fieldMap) + 1
    totalVehicles = queue.Count

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

        API_ShowStatus "OFV", "Eierskifter", identifier, _
            currentVehicle, totalVehicles

        Set vehicleRows = FetchOFVTransactions( _
            ofvKey, identifier, useVin, regNo, vin)

        If Not vehicleRowsByKey.Exists(CStr(key)) Then
            vehicleRowsByKey.Add CStr(key), New Collection
        End If

        For Each resultRow In vehicleRows
            allRows.Add resultRow
            vehicleRowsByKey(CStr(key)).Add resultRow
        Next resultRow

        Sleep API_PAUSE_MS

    Next key

    For Each key In queue.Keys

        vehicleData = queue(key)

        Set vehicleTxRows = Nothing
        If vehicleRowsByKey.Exists(CStr(key)) Then
            Set vehicleTxRows = vehicleRowsByKey(CStr(key))
        End If

        Set kontrollRow = BuildDemobilRow( _
            CStr(vehicleData(0)), CStr(vehicleData(1)), _
            vehicleData(2), vehicleTxRows, companyOrgNo)

        kontrollRows.Add kontrollRow

    Next key

    '======================================================
    ' RESULTAT
    '======================================================

    stage = "oppdaterer " & RESULT_SHEET_3
    API_ShowStatus "Excel", stage

    outputRows = allRows.Count

    Set loResult = GetOrCreateResultTable(wsResult, fieldMap, RESULT_TABLE_3)

    oldLastRow = loResult.Range.Row + loResult.Range.rows.Count - 1

    If outputRows > 0 Then
        newLastRow = outputRows + 1
    Else
        newLastRow = 2
    End If

    Set loResult = ResizeResultTable( _
        wsResult, loResult, newLastRow, fieldCount, RESULT_TABLE_3)

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
        loResult.HeaderRowRange.Cells(1, c + 1).value = fieldMap(c)(1)
    Next c

    If outputRows > 0 Then

        ReDim output(1 To outputRows, 1 To fieldCount)

        forrigeInputR3 = vbNullString

        For r = 1 To outputRows

            Set resultRow = allRows(r)

            denneInputR3 = VariantToString(resultRow("Input"))
            erForsteIGruppeR3 = (r = 1 Or denneInputR3 <> forrigeInputR3)

            For c = LBound(fieldMap) To UBound(fieldMap)

                dictionaryKey = CStr(fieldMap(c)(0))

                If dictionaryKey = "CalculatedSellerType" Or _
                   dictionaryKey = "CalculatedBuyerType" Then

                    output(r, c + 1) = Empty

                ElseIf Not erForsteIGruppeR3 And _
                    IsCarLevelField(dictionaryKey) Then

                    output(r, c + 1) = Empty

                ElseIf resultRow.Exists(dictionaryKey) Then

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

            forrigeInputR3 = denneInputR3

        Next r

        loResult.DataBodyRange.value = output
        ApplyCalculatedColumns loResult

    End If

    FormatResultTable wsResult, loResult
    FormatResultTableGrouping wsResult, loResult, allRows

    '======================================================
    ' KONTROLLARK
    '======================================================

    stage = "oppdaterer " & CONTROL_SHEET_3
    API_ShowStatus "Excel", stage

    UpdateDemobilControlSheet wsControl, kontrollRows, totalVehicles, _
        companyOrgNo, dateFraRaw, dateTilRaw

    Application.Calculation = oldCalculation

    If oldCalculation = xlCalculationManual Then
        wsResult.Calculate
        wsControl.Calculate
    Else
        Application.CalculateFull
    End If

    API_ShowStatus "Ferdig", "Kontroll Demobil er oppdatert"

    RestoreApplicationState oldScreenUpdating, oldEnableEvents, _
        oldCalculation, oldCursor

    applicationChanged = False

    SkjulFremdriftVindu

    MsgBox "Kontroll Demobil er oppdatert." & vbCrLf & vbCrLf & _
        totalVehicles & " biler lest fra input.", _
        vbInformation, "Kontroll Demobil"

    Exit Sub

SafeExit3:
    Application.StatusBar = False
    SkjulFremdriftVindu
    Exit Sub

FatalError3:

    errorNumber3 = Err.Number
    errorDescription3 = Err.Description

    Application.StatusBar = False

    If applicationChanged Then

        RestoreApplicationState oldScreenUpdating, oldEnableEvents, _
            oldCalculation, oldCursor

    End If

    SkjulFremdriftVindu

    MsgBox "Kontroll Demobil ble avbrutt." & vbCrLf & vbCrLf & _
        "Trinn: " & stage & vbCrLf & _
        "Feil " & errorNumber3 & ": " & errorDescription3, _
        vbCritical, "Kontroll Demobil"

End Sub


' Bygger kontrollraden for en bil i Kontroll Demobil-listen. Sjekker om
' bilens NYESTE registrerte transaksjon (uansett dato) har juridisk
' enhet (companyOrgNo) som kjoper - altsa at bilen fremdeles star
' registrert pa selskapet, uten noe senere salg.
Private Function BuildDemobilRow( _
    ByVal regNo As String, _
    ByVal vin As String, _
    ByVal bokfortInnRaw As Variant, _
    ByVal vehicleTxRows As Collection, _
    ByVal companyOrgNo As String) As Object

    Dim result As Object
    Dim txRow As Variant
    Dim bokfortDate As Variant
    Dim modelName As String
    Dim chassisNo As String
    Dim regNoResolved As String
    Dim companyOrgNormalisert As String

    Dim nyesteTxRow As Object
    Dim companyNavn As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    bokfortDate = TolkBokfortDato(bokfortInnRaw)
    modelName = vbNullString
    chassisNo = vin
    regNoResolved = regNo
    companyOrgNormalisert = NormalizeIdentifier(companyOrgNo)
    companyNavn = vbNullString

    Set nyesteTxRow = Nothing

    If Not vehicleTxRows Is Nothing Then

        For Each txRow In vehicleTxRows

            If VariantToString(txRow("Status")) = "OK" Then

                If Len(modelName) = 0 Then
                    modelName = VariantToString(txRow("ModelName"))
                End If

                If Len(VariantToString(txRow("ChassisNumber"))) > 0 Then
                    chassisNo = VariantToString(txRow("ChassisNumber"))
                End If

                If Len(VariantToString(txRow("RegNo"))) > 0 Then
                    regNoResolved = VariantToString(txRow("RegNo"))
                End If

                If Len(companyNavn) = 0 Then

                    If NormalizeIdentifier(txRow("ToOwnerOrgNo")) = _
                        companyOrgNormalisert Then

                        companyNavn = _
                            VariantToString(txRow("ToOwnerCompanyName"))

                    End If

                End If

                If IsDate(txRow("TransactionDate")) Then

                    If nyesteTxRow Is Nothing Then

                        Set nyesteTxRow = txRow

                    ElseIf CDate(txRow("TransactionDate")) > _
                        CDate(nyesteTxRow("TransactionDate")) Then

                        Set nyesteTxRow = txRow

                    End If

                End If

            End If

        Next txRow

    End If

    result("RegnrInput") = regNoResolved
    result("Chassisnummer") = chassisNo
    result("Modell") = modelName
    result("BokfortInnDato") = bokfortDate
    result("SisteTransaksjonsDato") = Empty
    result("SisteKjoper") = vbNullString
    result("EidAvEnhet") = vbNullString

    If nyesteTxRow Is Nothing Then

        result("Status") = "Ingen treff - ingen OFV-transaksjoner funnet"

    Else

        result("SisteTransaksjonsDato") = nyesteTxRow("TransactionDate")

        result("SisteKjoper") = ComputeOwnerLabel( _
            VariantToString(nyesteTxRow("ToOwnerType")), _
            VariantToString(nyesteTxRow("ToOwnerCompanyName")))

        If NormalizeIdentifier(nyesteTxRow("ToOwnerOrgNo")) = _
            companyOrgNormalisert Then

            result("EidAvEnhet") = "Ja"

            If Len(companyNavn) = 0 Then
                companyNavn = _
                    VariantToString(nyesteTxRow("ToOwnerCompanyName"))
            End If

            result("Status") = "OK - fortsatt eid av " & _
                IIf(Len(companyNavn) > 0, companyNavn, companyOrgNo)

        Else

            result("EidAvEnhet") = "Nei"
            result("Status") = "Avvik - siste eierskifte er til " & _
                result("SisteKjoper")

        End If

    End If

    Set BuildDemobilRow = result

End Function


' Bygger hele arket Kontroll Demobil pa nytt hver kjoring.
Private Sub UpdateDemobilControlSheet( _
    ByVal ws As Worksheet, _
    ByVal kontrollRows As Collection, _
    ByVal totalVehicles As Long, _
    ByVal companyOrgNo As String, _
    ByVal dateFraRaw As Variant, _
    ByVal dateTilRaw As Variant)

    Const HEADER_ROW As Long = 9
    Const FIRST_DATA_ROW As Long = 10

    Dim row As Object
    Dim r As Long
    Dim lastRow As Long
    Dim antallOk As Long
    Dim antallAvvik As Long
    Dim statusText As String
    Dim periodeTekst As String

    ws.Cells.Clear

    ws.Range("A1:G1").Merge
    ws.Range("A1").value = ws.Name

    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ws.rows(1).RowHeight = 26

    If IsDate(dateFraRaw) And IsDate(dateTilRaw) Then

        periodeTekst = " Periode (kun til visning): " & _
            Format$(CDate(dateFraRaw), "dd.mm.yyyy") & " - " & _
            Format$(CDate(dateTilRaw), "dd.mm.yyyy") & "."

    End If

    ws.Range("A2:G2").Merge
    ws.Range("A2").value = _
        "Kontrollregel: For hver bil sjekkes bilens NYESTE " & _
        "registrerte OFV-transaksjon (uansett dato) - kjoperen der " & _
        "skal vaere juridisk enhet (orgnr " & companyOrgNo & _
        "). Er kjoperen et annet selskap eller en privatperson, " & _
        "er bilen sannsynligvis solgt videre og flagges som avvik." & _
        periodeTekst

    ws.Range("A2").Font.Italic = True
    ws.rows(2).RowHeight = 30

    For Each row In kontrollRows

        statusText = VariantToString(row("Status"))

        If Left$(statusText, 2) = "OK" Then
            antallOk = antallOk + 1
        ElseIf Left$(statusText, 5) = "Avvik" Then
            antallAvvik = antallAvvik + 1
        End If

    Next row

    ws.Range("A4:B4").Merge : ws.Range("A4").value = "Inputbiler"
    ws.Range("C4:D4").Merge : ws.Range("C4").value = "Fortsatt hos enhet"
    ws.Range("E4:F4").Merge : ws.Range("E4").value = "Avvik (videresolgt)"

    With ws.Range("A4:F4")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("A5:B5").Merge : ws.Range("A5").value = totalVehicles
    ws.Range("C5:D5").Merge : ws.Range("C5").value = antallOk
    ws.Range("E5:F5").Merge : ws.Range("E5").value = antallAvvik

    With ws.Range("A5:B5")
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("C5:D5")
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = COLOR_GREEN_FILL
        .Font.Color = COLOR_GREEN_FONT
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("E5:F5")
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = COLOR_RED_FILL
        .Font.Color = COLOR_RED_FONT
        .HorizontalAlignment = xlCenter
    End With

    ws.rows("4:5").RowHeight = 20

    ws.Range("A" & HEADER_ROW).value = "Regnr"
    ws.Range("B" & HEADER_ROW).value = "Chassisnummer"
    ws.Range("C" & HEADER_ROW).value = "Modell"
    ws.Range("D" & HEADER_ROW).value = "Bokfort inn dato"
    ws.Range("E" & HEADER_ROW).value = "Siste transaksjonsdato"
    ws.Range("F" & HEADER_ROW).value = "Siste kjoper"
    ws.Range("G" & HEADER_ROW).value = "Status"

    With ws.Range("A" & HEADER_ROW & ":G" & HEADER_ROW)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30
    End With

    r = FIRST_DATA_ROW

    For Each row In kontrollRows

        ws.Cells(r, 1).value = VariantToString(row("RegnrInput"))
        ws.Cells(r, 2).value = VariantToString(row("Chassisnummer"))
        ws.Cells(r, 3).value = VariantToString(row("Modell"))
        ws.Cells(r, 4).value = row("BokfortInnDato")
        ws.Cells(r, 5).value = row("SisteTransaksjonsDato")
        ws.Cells(r, 6).value = VariantToString(row("SisteKjoper"))
        ws.Cells(r, 7).value = VariantToString(row("Status"))

        statusText = VariantToString(row("Status"))

        If Left$(statusText, 2) = "OK" Then

            With ws.Range(ws.Cells(r, 7), ws.Cells(r, 7))
                .Interior.Color = COLOR_GREEN_FILL
                .Font.Color = COLOR_GREEN_FONT
            End With

        ElseIf Left$(statusText, 5) = "Avvik" Then

            With ws.Range(ws.Cells(r, 7), ws.Cells(r, 7))
                .Interior.Color = COLOR_RED_FILL
                .Font.Color = COLOR_RED_FONT
            End With

        End If

        r = r + 1

    Next row

    lastRow = r - 1
    If lastRow < FIRST_DATA_ROW Then lastRow = FIRST_DATA_ROW

    ws.Range("D" & FIRST_DATA_ROW & ":E" & lastRow).NumberFormat = _
        "dd.mm.yyyy"

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D:E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 25
    ws.Columns("G").ColumnWidth = 38

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
    Dim popupMessage As String
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

    ' Penere, kompakt popup-tekst med en tekstbasert fremdriftslinje -
    ' egen formatering fra den mer detaljerte statuslinje-teksten over.
    popupMessage = activity

    If currentVehicle > 0 And totalVehicles > 0 Then

        popupMessage = popupMessage & "  " & _
            LagFremdriftsbar(percentage) & "  " & _
            Format$(percentage, "0%") & _
            "  (" & currentVehicle & "/" & totalVehicles & ")"

    End If

    If Len(identifier) > 0 Then
        popupMessage = popupMessage & "  " & identifier
    End If

    Select Case sourceName
        Case "OFV"
            OppdaterFremdriftLinje "lblOFV", "OFV: " & popupMessage
        Case "SVV"
            OppdaterFremdriftLinje "lblSVV", "SVV: " & popupMessage
    End Select

    DoEvents

End Sub


' Tekstbasert fremdriftslinje av Unicode-blokktegn (fylt/tom), til
' bruk i popup-vinduet. Ingen ekstra kontroller trengs i UserForm-en -
' hele "loading bar"-effekten er bare formatert tekst i den samme
' Label-en som resten av statuslinjen.
Private Function LagFremdriftsbar( _
    ByVal andel As Double, _
    Optional ByVal bredde As Long = 16) As String

    Dim fylte As Long
    Dim i As Long
    Dim bar As String

    If andel < 0 Then andel = 0
    If andel > 1 Then andel = 1

    fylte = CLng(andel * bredde)

    For i = 1 To bredde

        If i <= fylte Then
            bar = bar & ChrW(9608)  ' full blokk
        Else
            bar = bar & ChrW(9617)  ' lys skyggelegging
        End If

    Next i

    LagFremdriftsbar = bar

End Function


'==============================================================
' FREMDRIFTSVINDU (popup under kjoring)
'==============================================================

' Viser popup-vinduet "frmFremdrift" hvis det finnes i prosjektet
' (Insert > UserForm i VBA-editoren, navngitt eksakt "frmFremdrift",
' med to Label-kontroller navngitt "lblOFV" og "lblSVV"). Sent-bundet
' (VBA.UserForms.Add med et navn som streng) slik at hele filen
' fortsatt kompilerer selv om skjemaet ikke er opprettet enna.
Private Sub VisFremdriftVindu()

    On Error Resume Next

    Set gFremdriftForm = Nothing
    Set gFremdriftForm = VBA.UserForms.Add("frmFremdrift")

    If Not gFremdriftForm Is Nothing Then

        gFremdriftForm.Caption = "Oppdaterer API-data ..."

        gFremdriftForm.Controls("lblOFV").Font.Bold = True
        gFremdriftForm.Controls("lblOFV").Font.Name = "Consolas"

        gFremdriftForm.Controls("lblSVV").Font.Bold = True
        gFremdriftForm.Controls("lblSVV").Font.Name = "Consolas"

        OppdaterFremdriftLinje "lblOFV", "OFV: Venter ..."
        OppdaterFremdriftLinje "lblSVV", "SVV: Venter ..."

        gFremdriftForm.Show vbModeless

    End If

    On Error GoTo 0

End Sub


' Setter Caption pa en navngitt kontroll i fremdriftsvinduet, hvis
' vinduet er apent og kontrollen finnes. Feil her (f.eks. feil
' kontrollnavn) svelges bevisst - popup-teksten er kun kosmetisk.
Private Sub OppdaterFremdriftLinje( _
    ByVal kontrollNavn As String, _
    ByVal tekst As String)

    If gFremdriftForm Is Nothing Then Exit Sub

    On Error Resume Next
    gFremdriftForm.Controls(kontrollNavn).Caption = tekst
    On Error GoTo 0

    DoEvents

End Sub


' Lukker popup-vinduet. Trygg a kalle selv om det aldri ble apnet.
Private Sub SkjulFremdriftVindu()

    On Error Resume Next

    If Not gFremdriftForm Is Nothing Then
        Unload gFremdriftForm
        Set gFremdriftForm = Nothing
    End If

    On Error GoTo 0

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
' STATENS VEGVESEN (reserve for forstegangsregistrering)
'==============================================================

' Kalles kun nar OFV ikke har noen transaksjoner for kjoretoyet.
' Prover regnr forst, deretter VIN hvis regnr ikke gir treff.
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

        If Len(statusText) > 0 Then
            result("Status") = statusText
        End If

        Set FetchVehicleInfoFromSVV = result
        Exit Function

    End If

    isoDate = VariantToString( _
        JSON_ExtractValue( _
            responseText, "registrertForstegangNorgeDato"))

    If Len(isoDate) < 10 Then

        isoDate = VariantToString( _
            JSON_ExtractValue( _
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

        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

        statusCode = 0
        responseText = vbNullString

        On Error Resume Next

        http.SetTimeouts 10000, 10000, 30000, 30000

        http.Open "GET", url, False

        http.SetRequestHeader _
            "SVV-Authorization", "Apikey " & apiKey

        http.SetRequestHeader "Accept", "application/json"

        http.Send

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

                lastError = _
                    CStr(statusCode) & ": " & responseText

                Sleep RETRY_WAIT_MS * attempt

            Case Else

                If statusCode <> 0 Then

                    statusText = _
                        "Feil: Vegvesenet HTTP " & statusCode

                    Exit Function

                Else
                    Sleep RETRY_WAIT_MS * attempt
                End If

        End Select

    Next attempt

    statusText = "Feil: Vegvesenet - " & lastError

End Function


'==============================================================
' OFV-TRANSAKSJONER
'==============================================================

' Henter ALLE transaksjoner for kjoretoyet - ett kall per kjoretoy,
' uten datofilter, sortert nyeste forst. Garanterer at
' firstRegistrationDate blir funnet sa lenge OFV har minst en
' transaksjon noen gang for kjoretoyet. For Transaksjoner-tabellen
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

    ' Antatt feltnavn for organisasjonsnummer (samme monster som
    ' filternavnet "fromOrganizationNumber"/"toOrganizationNumber") -
    ' sjekk mot en faktisk OFV-respons og juster her hvis feltet
    ' faktisk heter noe annet i companyInfo.
    result("FromOwnerOrgNo") = _
        VariantToString( _
            JSON_ExtractValue(companyJSON, "organizationNumber"))

    result("FromOwnerCounty") = _
        JSON_ExtractValue(fromOwnerJSON, "countyName")

    result("ToOwnerType") = _
        JSON_ExtractValue(toOwnerJSON, "type")

    companyJSON = _
        JSON_ExtractObject(toOwnerJSON, "companyInfo")

    result("ToOwnerCompanyName") = _
        JSON_ExtractValue(companyJSON, "name")

    result("ToOwnerOrgNo") = _
        VariantToString( _
            JSON_ExtractValue(companyJSON, "organizationNumber"))

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

    Dim fields(0 To 23) As Variant

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
        "FromOwnerOrgNo", "SelgerOrgNr", False)
    fields(18) = Array( _
        "FromOwnerCounty", _
        "SelgerEierFylke", False)
    fields(19) = Array( _
        "ToOwnerType", _
        "KjoperEierType", False)
    fields(20) = Array( _
        "ToOwnerCompanyName", _
        "KjoperEierFirma", False)
    fields(21) = Array( _
        "ToOwnerOrgNo", "KjoperOrgNr", False)
    fields(22) = Array( _
        "ToOwnerCounty", BuyerCountyHeader(), False)
    fields(23) = Array("Status", "Status", False)

    GetFieldMap = fields

End Function


'==============================================================
' RESULTATTABELL
'==============================================================

Private Function GetOrCreateResultTable( _
    ByVal ws As Worksheet, _
    ByVal fieldMap As Variant, _
    ByVal tableName As String) As ListObject

    Dim lo As ListObject
    Dim target As Range
    Dim fieldCount As Long
    Dim c As Long

    fieldCount = UBound(fieldMap) + 1

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
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

        lo.Name = tableName
        lo.DataBodyRange.ClearContents

    End If

    Set GetOrCreateResultTable = lo

End Function


Private Function ResizeResultTable( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject, _
    ByVal lastRow As Long, _
    ByVal fieldCount As Long, _
    ByVal tableName As String) As ListObject

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

        lo.Name = tableName

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
    ws.Columns("R").ColumnWidth = 13
    ws.Columns("S").ColumnWidth = 13
    ws.Columns("T").ColumnWidth = 13
    ws.Columns("U").ColumnWidth = 25
    ws.Columns("V").ColumnWidth = 13
    ws.Columns("W").ColumnWidth = 15
    ws.Columns("X").ColumnWidth = 28

End Sub


' Gjor det tydelig hvor en bils transaksjoner slutter og neste
' begynner: topplinje ved skifte av "Input"-verdi, og kun den
' FORSTE raden for hver bil er fet/uthevet - resten av bilens
' transaksjoner star i vanlig skrift.
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

        If r = 1 Or denneInput <> forrigeInput Then

            With ws.Range( _
                ws.Cells(wsRow, 1), ws.Cells(wsRow, lo.ListColumns.Count))

                .Font.Bold = True
                .Interior.Color = RGB(238, 244, 251)

            End With

        Else

            With ws.Range( _
                ws.Cells(wsRow, 1), ws.Cells(wsRow, lo.ListColumns.Count))

                .Font.Bold = False
                .Interior.ColorIndex = xlColorIndexNone

            End With

        End If

        forrigeInput = denneInput

    Next r

End Sub


'==============================================================
' KONTROLL SOLGTE BILER
'==============================================================

' Bygger hele arket pa nytt hver kjoring - tittel, forklaring av
' kontrollregelen, KPI-bokser, fargekodelegende, og EN rad per kjoretoy
' (sortert etter storst dagers avvik forst) med kun den matchede
' transaksjonen - full historikk ligger i Resultat-arket. Ingen levende
' Excel-formler her: alt regnes ut i VBA og skrives som faste verdier.
Private Sub UpdateControlSheet( _
    ByVal ws As Worksheet, _
    ByVal kontrollRows As Collection, _
    ByVal totalVehicles As Long)

    Const HEADER_ROW As Long = 12
    Const FIRST_DATA_ROW As Long = 13

    Dim row As Object
    Dim sortertRader As Collection
    Dim r As Long
    Dim lastRow As Long

    Dim bucket0 As Long
    Dim bucket1til15 As Long
    Dim bucketOver15 As Long

    Dim kontrollertText As String
    Dim dagerAvvik As Variant
    Dim bucketColor As Long
    Dim bucketFontColor As Long

    ws.Cells.Clear

    '----------------------------------------------------------
    ' Tittel og forklaring av kontrollregelen
    '----------------------------------------------------------

    ws.Range("A1:M1").Merge
    ws.Range("A1").value = ws.Name

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
        "Kontrollregel: Bokfort dato sjekkes mot den OFV-transaksjonen " & _
        "som ligger NAERMEST bokfort dato i tid (uansett om den er for " & _
        "eller etter). Dager avvik er antall dager mellom denne datoen " & _
        "og bokfort dato."

    ws.Range("A3:M3").Merge
    ws.Range("A3").value = _
        "Har OFV ingen transaksjoner i det hele tatt for kjoretoyet, " & _
        "brukes forstegangsregistreringsdato fra Statens vegvesen " & _
        "(SVV) i stedet - bade som kontrollgrunnlag og i kolonnen " & _
        "Forstegangsregistrert. Kolonnen Kilde helt til venstre " & _
        "viser om treffet kommer fra OFV eller SVV. Er Juridisk " & _
        "enhet (Selger) fylt ut i B8, sjekkes det i tillegg om samme " & _
        "selskap star oppfort som bade selger og kjoper i den matchede " & _
        "transaksjonen (Selvhandel) - da flagges raden rod uansett " & _
        "dagers avvik."

    ws.Range("A4:M4").Merge
    ws.Range("A4").value = _
        "Hver bil har kun EN rad her, med den matchede " & _
        "transaksjonen/registreringen. Full transaksjonshistorikk for " & _
        "hver bil ligger i arket Resultat. Tabellen er sortert med " & _
        "storst dagers avvik forst."

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

                If VariantToString(row("Selvhandel")) = "Ja" Then
                    bucketOver15 = bucketOver15 + 1
                ElseIf CLng(dagerAvvik) <= AVVIK_GRONN_MAX Then
                    bucket0 = bucket0 + 1
                ElseIf CLng(dagerAvvik) <= AVVIK_GUL_MAX Then
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
    ws.Range("K6:M6").Merge
    ws.Range("K6").value = "FARGEKODER - DAGER AVVIK"

    With ws.Range("A6:M6")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("A7:B7").Merge
    ws.Range("A7").value = totalVehicles

    With ws.Range("A7:B7")
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("K7").value = "0-2 dager"
    ws.Range("L7").value = "3-14 dager"
    ws.Range("M7").value = "15+ dager"

    ws.Range("K8").value = bucket0
    ws.Range("L8").value = bucket1til15
    ws.Range("M8").value = bucketOver15

    With ws.Range("K7:K8")
        .Interior.Color = COLOR_GREEN_FILL
        .Font.Color = COLOR_GREEN_FONT
    End With

    With ws.Range("L7:L8")
        .Interior.Color = COLOR_YELLOW_FILL
        .Font.Color = COLOR_YELLOW_FONT
    End With

    With ws.Range("M7:M8")
        .Interior.Color = COLOR_RED_FILL
        .Font.Color = COLOR_RED_FONT
    End With

    With ws.Range("K7:M8")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.rows("6:8").RowHeight = 20

    '----------------------------------------------------------
    ' Kolonneoverskrifter
    '----------------------------------------------------------

    ws.Range("A" & HEADER_ROW).value = "Kilde"
    ws.Range("B" & HEADER_ROW).value = "API-treff"
    ws.Range("C" & HEADER_ROW).value = "Regnr"
    ws.Range("D" & HEADER_ROW).value = "Chassisnummer"
    ws.Range("E" & HEADER_ROW).value = "Modell"
    ws.Range("F" & HEADER_ROW).value = "Forstegangsregistrert"
    ws.Range("G" & HEADER_ROW).value = "Bokfort dato"
    ws.Range("H" & HEADER_ROW).value = "Transaksjonsdato"
    ws.Range("I" & HEADER_ROW).value = "RegistreringsType"
    ws.Range("J" & HEADER_ROW).value = "Dager avvik"
    ws.Range("K" & HEADER_ROW).value = "Selger"
    ws.Range("L" & HEADER_ROW).value = "Kjoper"
    ws.Range("M" & HEADER_ROW).value = "Selvhandel"

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
    ' En rad per kjoretoy - kun den matchede transaksjonen. Full
    ' historikk ligger i Resultat-arket.
    '----------------------------------------------------------

    r = FIRST_DATA_ROW

    For Each row In sortertRader

        ws.Cells(r, 1).value = VariantToString(row("Kilde"))
        ws.Cells(r, 2).value = VariantToString(row("ApiTreff"))
        ws.Cells(r, 3).value = VariantToString(row("RegnrInput"))
        ws.Cells(r, 4).value = VariantToString(row("Chassisnummer"))
        ws.Cells(r, 5).value = VariantToString(row("Modell"))
        ws.Cells(r, 6).value = row("Forstegangsregistrert")
        ws.Cells(r, 7).value = row("BokfortDato")
        ws.Cells(r, 8).value = row("KontrollTransaksjonDato")
        ws.Cells(r, 9).value = VariantToString(row("RegistreringsType"))
        ws.Cells(r, 10).value = row("DagerAvvik")
        ws.Cells(r, 11).value = VariantToString(row("Selger"))
        ws.Cells(r, 12).value = VariantToString(row("Kjoper"))
        ws.Cells(r, 13).value = VariantToString(row("Selvhandel"))

        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 13))
            .Font.Bold = True
            .Interior.Color = RGB(238, 244, 251)
        End With

        dagerAvvik = row("DagerAvvik")

        If IsNumeric(dagerAvvik) Then

            If CLng(dagerAvvik) <= AVVIK_GRONN_MAX Then
                bucketColor = COLOR_GREEN_FILL
                bucketFontColor = COLOR_GREEN_FONT
            ElseIf CLng(dagerAvvik) <= AVVIK_GUL_MAX Then
                bucketColor = COLOR_YELLOW_FILL
                bucketFontColor = COLOR_YELLOW_FONT
            Else
                bucketColor = COLOR_RED_FILL
                bucketFontColor = COLOR_RED_FONT
            End If

            ' Selvhandel (samme selskap som kjoper og selger)
            ' overstyrer alltid til rod, uansett dagers avvik.
            If VariantToString(row("Selvhandel")) = "Ja" Then
                bucketColor = COLOR_RED_FILL
                bucketFontColor = COLOR_RED_FONT
            End If

            With ws.Range(ws.Cells(r, 8), ws.Cells(r, 8))
                .Interior.Color = bucketColor
                .Font.Color = bucketFontColor
            End With

            With ws.Range(ws.Cells(r, 10), ws.Cells(r, 10))
                .Interior.Color = bucketColor
                .Font.Color = bucketFontColor
                .Font.Bold = True
            End With

            If VariantToString(row("Selvhandel")) = "Ja" Then

                With ws.Range(ws.Cells(r, 13), ws.Cells(r, 13))
                    .Interior.Color = COLOR_RED_FILL
                    .Font.Color = COLOR_RED_FONT
                    .Font.Bold = True
                End With

            End If

        End If

        r = r + 1

    Next row

    lastRow = r - 1
    If lastRow < FIRST_DATA_ROW Then lastRow = FIRST_DATA_ROW

    With ws.Range("A" & FIRST_DATA_ROW & ":M" & lastRow)
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .rows.RowHeight = 18
    End With

    ws.Range("F" & FIRST_DATA_ROW & ":H" & lastRow).NumberFormat = _
        "dd.mm.yyyy"

    ws.Range("J" & FIRST_DATA_ROW & ":J" & lastRow).NumberFormat = "0"

    With ws.Range("A" & HEADER_ROW & ":M" & lastRow).Borders
        .LineStyle = xlContinuous
        .Color = RGB(217, 226, 243)
        .Weight = xlThin
    End With

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F:H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 24
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K:L").ColumnWidth = 25
    ws.Columns("M").ColumnWidth = 14

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


' Bygger kontroll-raden for ett kjoretoy - kun EN rad per bil, med den
' transaksjonen som ligger NAERMEST bokfort dato i hele bilens OFV-
' transaksjonshistorikk (nyeste og eldste, ikke bare siste registrerte).
' Full historikk for bilen vises ikke her, men i Resultat-arket. Har
' OFV ingen transaksjoner i det hele tatt, brukes
' forstegangsregistreringsdato fra SVV i stedet - bade som
' kontrollgrunnlag og i kolonnen Forstegangsregistrert. Kolonnen
' "Kilde" viser om treffet endte opp som OFV, SVV eller Ingen.
'
' Valgfri tilleggskontroll (kun nar selgerOrgNo er fylt ut): den
' matchede transaksjonen skal vaere et salg FRA selgerOrgNo. Er OGSA
' kjoperen selgerOrgNo (samme selskap pa begge sider), flagges raden
' rod uansett dagers avvik - se result("Selvhandel").
Private Function BuildKontrollRow( _
    ByVal regNo As String, _
    ByVal vin As String, _
    ByVal bokfortRaw As Variant, _
    ByVal vehicleTxRows As Collection, _
    ByVal svvInfo As Object, _
    ByVal selgerOrgNo As String) As Object

    Dim result As Object
    Dim txRow As Variant
    Dim bokfortDate As Variant

    Dim firstRegDate As Variant
    Dim modelName As String
    Dim chassisNo As String
    Dim regNoResolved As String
    Dim errorStatus As String
    Dim hasAnyOkRow As Boolean
    Dim harBokfortDato As Boolean

    Dim naermesteTxRow As Object
    Dim naermesteDiff As Double
    Dim diffDager As Double

    Dim selgerOrgNormalisert As String
    Dim fraOrgNr As String
    Dim tilOrgNr As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    bokfortDate = TolkBokfortDato(bokfortRaw)
    harBokfortDato = Not IsEmpty(bokfortDate)

    firstRegDate = Empty
    modelName = vbNullString
    chassisNo = vin
    regNoResolved = regNo
    errorStatus = vbNullString
    hasAnyOkRow = False

    naermesteDiff = -1
    Set naermesteTxRow = Nothing

    selgerOrgNormalisert = NormalizeIdentifier(selgerOrgNo)

    If Not vehicleTxRows Is Nothing Then

        For Each txRow In vehicleTxRows

            If VariantToString(txRow("Status")) = "OK" Then

                hasAnyOkRow = True

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

                If Len(VariantToString(txRow("RegNo"))) > 0 Then
                    regNoResolved = VariantToString(txRow("RegNo"))
                End If

                If IsDate(txRow("TransactionDate")) And harBokfortDato Then

                    ' Den transaksjonen som ligger naermest bokfort
                    ' dato (i antall dager), uansett om den er for
                    ' eller etter.
                    diffDager = Abs(CDbl( _
                        CDate(txRow("TransactionDate")) - _
                        CDate(bokfortDate)))

                    If naermesteDiff < 0 Or _
                        diffDager < naermesteDiff Then

                        naermesteDiff = diffDager
                        Set naermesteTxRow = txRow

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

    result("RegnrInput") = regNoResolved
    result("Chassisnummer") = chassisNo
    result("Modell") = modelName
    result("BokfortDato") = bokfortDate
    result("Forstegangsregistrert") = firstRegDate
    result("KontrollTransaksjonDato") = Empty
    result("RegistreringsType") = vbNullString
    result("DagerAvvik") = Empty
    result("Selger") = vbNullString
    result("Kjoper") = vbNullString
    result("Kilde") = "Ingen"
    result("Selvhandel") = vbNullString

    ' Reserve: OFV har ingen transaksjon i det hele tatt for
    ' kjoretoyet - bruk forstegangsregistreringsdato fra SVV, bade
    ' som visningsverdi og som grunnlag for kontrollen.
    If Not hasAnyOkRow And Not svvInfo Is Nothing Then

        If svvInfo.Exists("FirstRegistrationDate") Then

            If IsDate(svvInfo("FirstRegistrationDate")) Then

                firstRegDate = svvInfo("FirstRegistrationDate")
                result("Forstegangsregistrert") = firstRegDate

            End If

        End If

    End If

    If IsEmpty(bokfortDate) Then

        If hasAnyOkRow Or Not IsEmpty(firstRegDate) Then
            result("ApiTreff") = "Mangler bokfort dato"
        ElseIf Len(errorStatus) > 0 Then
            result("ApiTreff") = errorStatus
        Else
            result("ApiTreff") = "Ingen treff"
        End If

        result("Kontrollert") = "Nei"

    ElseIf Not naermesteTxRow Is Nothing Then

        result("KontrollTransaksjonDato") = _
            naermesteTxRow("TransactionDate")
        result("RegistreringsType") = _
            VariantToString(naermesteTxRow("RegistrationType"))
        result("DagerAvvik") = naermesteDiff
        result("ApiTreff") = "Treff OFV eierskifte"
        result("Kontrollert") = "Ja"
        result("Kilde") = "OFV"

        result("Selger") = ComputeOwnerLabel( _
            VariantToString(naermesteTxRow("FromOwnerType")), _
            VariantToString(naermesteTxRow("FromOwnerCompanyName")))

        result("Kjoper") = ComputeOwnerLabel( _
            VariantToString(naermesteTxRow("ToOwnerType")), _
            VariantToString(naermesteTxRow("ToOwnerCompanyName")))

        ' Valgfri selvhandel-sjekk - kun nar B8 er fylt ut.
        If Len(selgerOrgNormalisert) > 0 Then

            fraOrgNr = NormalizeIdentifier( _
                naermesteTxRow("FromOwnerOrgNo"))
            tilOrgNr = NormalizeIdentifier( _
                naermesteTxRow("ToOwnerOrgNo"))

            If fraOrgNr = selgerOrgNormalisert And _
               tilOrgNr = selgerOrgNormalisert Then

                result("Selvhandel") = "Ja"

            Else
                result("Selvhandel") = "Nei"
            End If

        End If

    ElseIf Not IsEmpty(firstRegDate) Then

        result("KontrollTransaksjonDato") = firstRegDate
        result("RegistreringsType") = "Forstegangsregistrering (SVV)"
        result("DagerAvvik") = Abs(CLng( _
            CDate(firstRegDate) - CDate(bokfortDate)))
        result("ApiTreff") = "Treff SVV forstegangsregistrering"
        result("Kontrollert") = "Ja"
        result("Kilde") = "SVV"

    Else

        If svvInfo Is Nothing Then

            If Len(errorStatus) > 0 Then
                result("ApiTreff") = errorStatus
            Else
                result("ApiTreff") = "Ingen treff"
            End If

        Else

            result("ApiTreff") = VariantToString(svvInfo("Status"))

        End If

        result("Kontrollert") = "Nei"

    End If

    Set BuildKontrollRow = result

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


' Leser ett kjoretoy-input fra en av to nabokolonner (Regnr/VIN) - kun
' en av de to fylles ut per rad. Brukes av alle tre kontrollene, som
' hver har sitt eget kolonnepar for dette pa Input-arket.
Private Function ReadInputIdentifier( _
    ByVal ws As Worksheet, _
    ByVal r As Long, _
    ByVal colRegnr As Long, _
    ByVal colVin As Long) As String

    Dim verdi As String

    verdi = NormalizeIdentifier(ws.Cells(r, colRegnr).value)

    If Len(verdi) = 0 Then
        verdi = NormalizeIdentifier(ws.Cells(r, colVin).value)
    End If

    ReadInputIdentifier = verdi

End Function


' Siste rad med data i enten Regnr- eller VIN-kolonnen for en av
' kontrollenes input-seksjon (radene kan ha data i bare en av de to).
Private Function LastRowInEitherColumn( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal colRegnr As Long, _
    ByVal colVin As Long) As Long

    Dim lastRegnr As Long
    Dim lastVin As Long

    lastRegnr = ws.Cells(ws.rows.Count, colRegnr).End(xlUp).Row
    lastVin = ws.Cells(ws.rows.Count, colVin).End(xlUp).Row

    LastRowInEitherColumn = WorksheetFunction.Max( _
        lastRegnr, lastVin, firstRow - 1)

End Function


' Felt som beskriver selve kjoretoyet (ikke transaksjonen) - blankes
' ut pa alle rader unntatt den forste/nyeste for hver bil i
' Resultat-tabellene, slik at bilinfo kun vises en gang per kjoretoy.
Private Function IsCarLevelField(ByVal fieldKey As String) As Boolean

    Select Case fieldKey

        Case "RegNo", "ChassisNumber", "MakeName", "ModelName", _
             "FuelGroup", "IsLeased", "IsUsedImported", _
             "FirstRegistrationDate"

            IsCarLevelField = True

        Case Else
            IsCarLevelField = False

    End Select

End Function


Private Function NormalizeIdentifier( _
    ByVal value As Variant) As String

    NormalizeIdentifier = _
        UCase$(Trim$(Replace( _
            CStr(value & vbNullString), _
            " ", vbNullString)))

End Function


' Skiller VIN fra regnr i den kombinerte input-kolonnen. VIN
' (chassisnummer) er alltid noyaktig 17 tegn per ISO 3779 - en fast,
' internasjonal standard. Norske regnr kan derimot variere i lengde
' (personlige skilt, eldre formater osv.), sa vi sjekker IKKE et fast
' bokstav/tall-monster for regnr - alt som ikke er 17 tegn regnes som
' regnr.
Private Function ErVIN(ByVal tekst As String) As Boolean
    ErVIN = (Len(tekst) = 17)
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


' Tolker Bokfort dato fra Input-arket. Dette er den eneste datoen i
' hele arket som en bruker skriver inn for hand (alle andre datoer
' kommer fra OFV/SVV sitt eget ISO-format og tolkes av DateFromISO,
' som er helt uavhengig av regionsinnstillinger).
'
' - Er cellen en ekte Excel-dato (uansett hvilket tallformat den
'   VISES i - dd.mm.aaaa, mm.dd.aaaa osv. spiller ingen rolle, Excel
'   lagrer den som et tall), brukes den direkte og uten tvetydighet.
' - Er cellen tekst, tolkes den EKSPLISITT som dag.maned.ar (norsk
'   standard), uavhengig av hvilke regionsinnstillinger som star pa
'   maskinen som kjorer makroen - IKKE via CDate, som ville tolket
'   teksten ulikt fra pc til pc.
Private Function TolkBokfortDato( _
    ByVal raw As Variant) As Variant

    Dim tekst As String
    Dim deler() As String
    Dim dag As Long
    Dim maned As Long
    Dim ar As Long

    If VarType(raw) = vbDate Then
        TolkBokfortDato = CDate(raw)
        Exit Function
    End If

    tekst = Trim$(CStr(raw & vbNullString))

    If Len(tekst) = 0 Then
        TolkBokfortDato = Empty
        Exit Function
    End If

    tekst = Replace(tekst, "/", ".")
    tekst = Replace(tekst, "-", ".")

    deler = Split(tekst, ".")

    If UBound(deler) = 2 Then

        If IsNumeric(deler(0)) And IsNumeric(deler(1)) And _
           IsNumeric(deler(2)) Then

            dag = CLng(deler(0))
            maned = CLng(deler(1))
            ar = CLng(deler(2))

            If ar < 100 Then ar = ar + 2000

            If maned >= 1 And maned <= 12 And _
               dag >= 1 And dag <= 31 And _
               ar >= 1900 And ar <= 2100 Then

                On Error GoTo UgyldigDato
                TolkBokfortDato = DateSerial(ar, maned, dag)
                Exit Function

            End If

        End If

    End If

    TolkBokfortDato = Empty
    Exit Function

UgyldigDato:
    TolkBokfortDato = Empty

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
