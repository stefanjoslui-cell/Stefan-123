Attribute VB_Name = "modOFV"
Option Explicit

' =====================================================================
' OFV Transactions API - Excel VBA-versjon
' updated 2026-09-19
'
' Portert fra ofv_transactions_lookup.py (samme funksjonalitet: for
' hvert kjoretoy hentes ALLE registreringer/eierskifter som faller inn
' i et datointervall, og hver registrering blir sin egen rad i
' resultatet - ikke bare den siste).
'
' Forutsetter to ark i arbeidsboken:
'   "Input"    - API-nokkel, fra-/til-dato, og Regnr/VIN-listen
'   "Resultat" - tomt ark som fylles ut av makroen for hver kjoring
'
' Navngitte omrader (Formler > Navnebehandling), pa arket "Input":
'   OFV_API       - cellen med API-nokkelen
'   OFV_DateFrom  - cellen med fra-dato (formatert som dato)
'   OFV_DateTo    - cellen med til-dato (formatert som dato)
'
' Regnr star i kolonne B fra rad FIRST_ROW og nedover, VIN i kolonne C
' (samme rader). En rad kan ha enten Regnr eller VIN - har den begge,
' brukes VIN. Utdata skrives til arket "Resultat", en rad per treff.
' =====================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ---------------------------------------------------------------------
' Konfigurasjon
' ---------------------------------------------------------------------
Private Const INPUT_SHEET As String = "Input"
Private Const RESULT_SHEET As String = "Resultat"
Private Const FIRST_ROW As Long = 5           ' forste datarad pa Input-arket
Private Const COL_REGNR As Long = 2           ' B - input
Private Const COL_VIN As Long = 3             ' C - input

Private Const OFV_BASE_URL As String = "https://api.ofv.no/transactions/v1/"
Private Const OFV_MAX_RETRIES As Long = 4
Private Const OFV_RETRY_WAIT_MS As Long = 3000
Private Const OFV_PAUSE_MS As Long = 150
Private Const OFV_SORT_DIRECTION As String = "ASC"  ' ASC = eldste forst, DESC = nyeste forst


' =====================================================================
' HOVEDMAKRO - kjor denne (Alt+F8, eller koble til en knapp)
' =====================================================================
Public Sub OFV_RefreshInfo()
    Dim wsInput As Worksheet, wsResult As Worksheet
    On Error Resume Next
    Set wsInput = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set wsResult = ThisWorkbook.Worksheets(RESULT_SHEET)
    On Error GoTo 0
    If wsInput Is Nothing Or wsResult Is Nothing Then
        MsgBox "Fant ikke arkene '" & INPUT_SHEET & "' og/eller '" & RESULT_SHEET & "'.", vbExclamation, "OFV"
        Exit Sub
    End If

    Dim strApiKey As String
    On Error Resume Next
    strApiKey = Trim$(ThisWorkbook.Names("OFV_API").RefersToRange.Value)
    On Error GoTo 0
    If Len(strApiKey) = 0 Then
        MsgBox "Fant ingen API-nokkel i navngitt omrade 'OFV_API'. Sjekk Formler > Navnebehandling.", _
               vbExclamation, "OFV"
        Exit Sub
    End If

    Dim datFrom As Variant, datTo As Variant
    On Error Resume Next
    datFrom = ThisWorkbook.Names("OFV_DateFrom").RefersToRange.Value
    datTo = ThisWorkbook.Names("OFV_DateTo").RefersToRange.Value
    On Error GoTo 0
    If Not IsDate(datFrom) Or Not IsDate(datTo) Then
        MsgBox "Fyll inn gyldige datoer i cellene for fra-dato og til-dato pa arket '" & INPUT_SHEET & "'.", _
               vbExclamation, "OFV"
        Exit Sub
    End If

    Dim strDateFromIso As String, strDateToIso As String
    strDateFromIso = Format$(datFrom, "yyyy-mm-dd")
    strDateToIso = Format$(datTo, "yyyy-mm-dd")

    ' finn siste datarad pa Input (lengste av Regnr- og VIN-kolonnen)
    Dim lngLastRowB As Long, lngLastRowC As Long, lngLastRow As Long
    lngLastRowB = wsInput.Cells(wsInput.Rows.Count, COL_REGNR).End(xlUp).Row
    lngLastRowC = wsInput.Cells(wsInput.Rows.Count, COL_VIN).End(xlUp).Row
    lngLastRow = lngLastRowB
    If lngLastRowC > lngLastRow Then lngLastRow = lngLastRowC
    If lngLastRow < FIRST_ROW Then
        MsgBox "Fant ingen Regnr eller VIN fra rad " & FIRST_ROW & " og nedover pa arket '" & INPUT_SHEET & "'.", _
               vbInformation, "OFV"
        Exit Sub
    End If

    ' bygg liste over unike kjoretoy (Regnr eller VIN, VIN har forrang)
    Dim objQueue As Object
    Set objQueue = CreateObject("Scripting.Dictionary")
    objQueue.CompareMode = vbTextCompare

    Dim r As Long, strVin As String, strReg As String, strKey As String
    For r = FIRST_ROW To lngLastRow
        strVin = Trim$(Replace(CStr(wsInput.Cells(r, COL_VIN).Value & vbNullString), " ", vbNullString))
        strReg = Trim$(Replace(CStr(wsInput.Cells(r, COL_REGNR).Value & vbNullString), " ", vbNullString))
        strKey = OFV_BuildKey(strVin, strReg)
        If Len(strKey) > 0 Then
            If Not objQueue.Exists(strKey) Then objQueue.Add strKey, strKey
        End If
    Next r

    If objQueue.Count = 0 Then
        MsgBox "Fant ingen Regnr eller VIN fra rad " & FIRST_ROW & " og nedover.", vbInformation, "OFV"
        Exit Sub
    End If

    Application.Cursor = xlWait
    Application.ScreenUpdating = False

    Dim objFieldMap As Variant
    objFieldMap = OFV_GetFieldMap()
    Dim m As Long

    ' tom resultatarket (behold ev. tidligere innhold under overskriftene,
    ' men fjern det siden radantallet varierer fra kjoring til kjoring)
    Dim lngResultLastRow As Long
    lngResultLastRow = wsResult.Cells(wsResult.Rows.Count, 1).End(xlUp).Row
    If lngResultLastRow > 1 Then
        wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(lngResultLastRow, UBound(objFieldMap) + 1)).ClearContents
    End If

    For m = LBound(objFieldMap) To UBound(objFieldMap)
        wsResult.Cells(1, m + 1).Value = objFieldMap(m)(1)
    Next m
    wsResult.Rows(1).Font.Bold = True

    Dim varKey As Variant, i As Long, t As Long, lngErr As Long, lngOutRow As Long, lngTotalRows As Long
    t = objQueue.Count
    lngOutRow = 2

    For Each varKey In objQueue.Keys
        i = i + 1
        Application.StatusBar = "OFV: Henter data (" & i & " av " & t & ") ..."

        Dim strParts() As String, blnIsVin As Boolean, strIdentifier As String
        strParts = Split(CStr(varKey), "|")
        blnIsVin = (strParts(0) = "VIN")
        strIdentifier = strParts(1)

        Dim colRows As Collection
        Set colRows = OFV_FetchAllTransactionRows(strApiKey, strIdentifier, blnIsVin, strDateFromIso, strDateToIso)

        Dim objRow As Object
        For Each objRow In colRows
            If objRow("Status") <> "OK" Then lngErr = lngErr + 1
            For m = LBound(objFieldMap) To UBound(objFieldMap)
                Dim strDictKey As String
                strDictKey = objFieldMap(m)(0)
                If objRow.Exists(strDictKey) Then
                    wsResult.Cells(lngOutRow, m + 1).Value = objRow(strDictKey)
                End If
            Next m
            lngOutRow = lngOutRow + 1
            lngTotalRows = lngTotalRows + 1
        Next objRow

        Sleep OFV_PAUSE_MS
    Next varKey

    ' formater dato-kolonnene (3. element i objFieldMap) som ekte datoer
    If lngOutRow > 2 Then
        For m = LBound(objFieldMap) To UBound(objFieldMap)
            If objFieldMap(m)(2) = True Then
                wsResult.Range(wsResult.Cells(2, m + 1), wsResult.Cells(lngOutRow - 1, m + 1)).NumberFormat = "dd.mm.yyyy"
            End If
        Next m
    End If

    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True

    MsgBox "OFV: Ferdig. " & lngTotalRows & " rad(er) for " & t & " kjoretoy (" & lngErr & " feil/uten treff).", _
           vbInformation, "OFV"
End Sub


Private Function OFV_BuildKey(strVin As String, strReg As String) As String
    ' VIN har forrang over Regnr hvis begge er fylt ut.
    If Len(strVin) > 0 Then
        OFV_BuildKey = "VIN|" & UCase$(strVin)
    ElseIf Len(strReg) > 0 Then
        OFV_BuildKey = "REG|" & UCase$(strReg)
    Else
        OFV_BuildKey = vbNullString
    End If
End Function


Private Function OFV_GetFieldMap() As Variant
    ' {dictionary-nokkel i objRow, kolonneoverskrift i Resultat-arket, er dato}
    OFV_GetFieldMap = Array( _
        Array("Input", "Input", False), _
        Array("Kilde", "Kilde", False), _
        Array("RegNo", "RegNo", False), _
        Array("ChassisNumber", "Chassisnummer", False), _
        Array("MakeName", "Merke", False), _
        Array("ModelName", "Modell", False), _
        Array("RegistrationType", "RegistreringsType", False), _
        Array("FuelGroup", "Drivstoffgruppe", False), _
        Array("IsLeased", "Leaset", False), _
        Array("IsUsedImported", "Bruktimportert", False), _
        Array("FirstRegistrationDate", "ForstegangsRegistrering", True), _
        Array("TransactionNumber", "TransaksjonsNummer", False), _
        Array("TransactionDate", "Eierskiftedato", True), _
        Array("FromOwnerType", "SelgerEierType", False), _
        Array("FromOwnerCompanyName", "SelgerEierFirma", False), _
        Array("FromOwnerCounty", "SelgerEierFylke", False), _
        Array("FromOwnerMunicipality", "SelgerEierKommune", False), _
        Array("FromUserCounty", "SelgerBrukerFylke", False), _
        Array("ToOwnerType", "KjoperEierType", False), _
        Array("ToOwnerCompanyName", "KjoperEierFirma", False), _
        Array("Status", "Status", False) _
    )
End Function


' =====================================================================
' Henter ALLE transaksjoner for ett kjoretoy innenfor datointervallet
' (med paginering), og returnerer en Collection av Dictionary-objekter
' - en per registrering. Ingen treff eller feil gir en enkelt rad med
' status satt tilsvarende (matcher ofv_transactions_lookup.py).
' =====================================================================
Private Function OFV_FetchAllTransactionRows(strApiKey As String, strIdentifier As String, blnIsVin As Boolean, _
                                              strDateFromIso As String, strDateToIso As String) As Collection
    Dim colRows As New Collection
    Dim strFilterKey As String
    strFilterKey = IIf(blnIsVin, "chassisNumber", "regNo")

    Dim strCursor As String
    strCursor = vbNullString

    Do
        Dim strBody As String
        strBody = "{""filters"":{""" & strFilterKey & """:""" & OFV_JsonEscape(strIdentifier) & """," & _
                  """transactionDateFrom"":""" & strDateFromIso & """," & _
                  """transactionDateTo"":""" & strDateToIso & """}," & _
                  """pagination"":{""first"":1000"
        If Len(strCursor) > 0 Then
            strBody = strBody & ",""cursor"":""" & OFV_JsonEscape(strCursor) & """"
        End If
        strBody = strBody & "}," & _
                  """sorting"":{""orderBy"":""transactionDate"",""orderDirection"":""" & OFV_SORT_DIRECTION & """}}"

        Dim strStatus As String
        Dim strResponse As String
        strResponse = OFV_PostWithRetries(strApiKey, strBody, strStatus)

        If strStatus <> "OK" Then
            colRows.Add OFV_BuildEmptyFields(strIdentifier, blnIsVin, strStatus)
            Set OFV_FetchAllTransactionRows = colRows
            Exit Function
        End If

        Dim strTransactionsArray As String
        strTransactionsArray = JSON_ExtractObject(strResponse, "transactions")

        Dim colItems As Collection
        Set colItems = JSON_ArrayAllElements(strTransactionsArray)

        Dim varItem As Variant
        For Each varItem In colItems
            colRows.Add OFV_BuildFieldsFromTransaction(strIdentifier, CStr(varItem), blnIsVin)
        Next varItem

        Dim strPaginationObj As String
        strPaginationObj = JSON_ExtractObject(strResponse, "pagination")

        Dim varHasNext As Variant, varCursor As Variant
        varHasNext = JSON_ExtractValue(strPaginationObj, "hasNextPage")
        varCursor = JSON_ExtractValue(strPaginationObj, "endCursor")

        Dim blnHasNext As Boolean
        blnHasNext = False
        If Not IsNull(varHasNext) Then blnHasNext = CBool(varHasNext)

        If blnHasNext And Len(CStr(varCursor & vbNullString)) > 0 Then
            strCursor = CStr(varCursor)
            Sleep OFV_PAUSE_MS
        Else
            Exit Do
        End If
    Loop

    If colRows.Count = 0 Then
        colRows.Add OFV_BuildEmptyFields(strIdentifier, blnIsVin, "Ingen registreringer i perioden")
    End If

    Set OFV_FetchAllTransactionRows = colRows
End Function


Private Function OFV_BuildEmptyFields(strIdentifier As String, blnIsVin As Boolean, strStatus As String) As Object
    Dim objFields As Object
    Set objFields = CreateObject("Scripting.Dictionary")
    objFields.CompareMode = vbTextCompare
    objFields("Input") = strIdentifier
    objFields("Kilde") = IIf(blnIsVin, "VIN", "Regnr")
    objFields("Status") = strStatus
    Set OFV_BuildEmptyFields = objFields
End Function


Private Function OFV_BuildFieldsFromTransaction(strIdentifier As String, strTxnJson As String, blnIsVin As Boolean) As Object
    Dim objFields As Object
    Set objFields = CreateObject("Scripting.Dictionary")
    objFields.CompareMode = vbTextCompare

    objFields("Input") = strIdentifier
    objFields("Kilde") = IIf(blnIsVin, "VIN", "Regnr")
    objFields("RegNo") = JSON_ExtractValue(strTxnJson, "regNo")
    objFields("ChassisNumber") = JSON_ExtractValue(strTxnJson, "chassisNumber")
    objFields("MakeName") = JSON_ExtractValue(strTxnJson, "makeName")
    objFields("ModelName") = JSON_ExtractValue(strTxnJson, "modelName")
    objFields("RegistrationType") = JSON_ExtractValue(strTxnJson, "registrationType")
    objFields("FuelGroup") = JSON_ExtractValue(strTxnJson, "fuelGroup")
    objFields("IsLeased") = JSON_ExtractValue(strTxnJson, "isLeased")
    objFields("IsUsedImported") = JSON_ExtractValue(strTxnJson, "isUsedImported")
    objFields("FirstRegistrationDate") = OFV_DateFromISO(CStr(JSON_ExtractValue(strTxnJson, "firstRegistrationDate") & vbNullString))
    objFields("TransactionNumber") = JSON_ExtractValue(strTxnJson, "transactionNumber")
    objFields("TransactionDate") = OFV_DateFromISO(CStr(JSON_ExtractValue(strTxnJson, "transactionDate") & vbNullString))

    Dim strFromObj As String, strToObj As String
    Dim strFromOwner As String, strFromUser As String, strToOwner As String, strCompanyObj As String
    strFromObj = JSON_ExtractObject(strTxnJson, "from")
    strToObj = JSON_ExtractObject(strTxnJson, "to")
    strFromOwner = JSON_ExtractObject(strFromObj, "owner")
    strFromUser = JSON_ExtractObject(strFromObj, "user")
    strToOwner = JSON_ExtractObject(strToObj, "owner")

    objFields("FromOwnerType") = JSON_ExtractValue(strFromOwner, "type")
    strCompanyObj = JSON_ExtractObject(strFromOwner, "companyInfo")
    objFields("FromOwnerCompanyName") = JSON_ExtractValue(strCompanyObj, "name")
    objFields("FromOwnerCounty") = JSON_ExtractValue(strFromOwner, "countyName")
    objFields("FromOwnerMunicipality") = JSON_ExtractValue(strFromOwner, "municipalityName")
    objFields("FromUserCounty") = JSON_ExtractValue(strFromUser, "countyName")

    objFields("ToOwnerType") = JSON_ExtractValue(strToOwner, "type")
    strCompanyObj = JSON_ExtractObject(strToOwner, "companyInfo")
    objFields("ToOwnerCompanyName") = JSON_ExtractValue(strCompanyObj, "name")

    objFields("Status") = "OK"

    Set OFV_BuildFieldsFromTransaction = objFields
End Function


' =====================================================================
' HTTP - POST mot OFV Transactions-API, med retry pa 429/5xx
' =====================================================================
Private Function OFV_PostWithRetries(strApiKey As String, strBody As String, ByRef strStatus As String) As String
    Dim lngAttempt As Long, strLastError As String
    strStatus = "OK"

    For lngAttempt = 1 To OFV_MAX_RETRIES
        Dim objHTTP As Object
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

        Dim lngStatusCode As Long, strResponseText As String
        On Error Resume Next
        objHTTP.Open "POST", OFV_BASE_URL, False
        objHTTP.setRequestHeader "Ocp-Apim-Subscription-Key", strApiKey
        objHTTP.setRequestHeader "Content-Type", "application/json"
        objHTTP.Send strBody
        lngStatusCode = objHTTP.Status
        strResponseText = objHTTP.ResponseText
        If Err.Number <> 0 Then
            strLastError = "VBA-feil: " & Err.Description
            lngStatusCode = 0
        End If
        On Error GoTo 0

        Select Case lngStatusCode
            Case 200
                OFV_PostWithRetries = strResponseText
                Exit Function
            Case 401
                strStatus = "Feil: 401 Unauthorized - sjekk API-nokkelen (OFV_API)"
                Exit Function
            Case 403
                strStatus = "Feil: 403 Forbidden - kvote overskredet"
                Exit Function
            Case 429, 500, 502, 503, 504
                strLastError = lngStatusCode & ": " & strResponseText
                Sleep OFV_RETRY_WAIT_MS * lngAttempt
            Case Else
                If lngStatusCode <> 0 Then
                    strStatus = "Feil: " & lngStatusCode & " " & strResponseText
                    Exit Function
                Else
                    Sleep OFV_RETRY_WAIT_MS * lngAttempt
                End If
        End Select
    Next lngAttempt

    strStatus = "Feil: Ga opp etter " & OFV_MAX_RETRIES & " forsok. Siste feil: " & strLastError
End Function


Private Function OFV_JsonEscape(strValue As String) As String
    Dim strResult As String
    strResult = Replace(strValue, "\", "\\")
    strResult = Replace(strResult, Chr(34), "\" & Chr(34))
    OFV_JsonEscape = strResult
End Function


Private Function OFV_DateFromISO(strISO As String) As Variant
    If Len(strISO) < 10 Then
        OFV_DateFromISO = Null
        Exit Function
    End If
    On Error GoTo Feil
    OFV_DateFromISO = DateSerial(CInt(Mid$(strISO, 1, 4)), CInt(Mid$(strISO, 6, 2)), CInt(Mid$(strISO, 9, 2)))
    Exit Function
Feil:
    OFV_DateFromISO = Null
End Function


' =====================================================================
' Generiske JSON-hjelpefunksjoner (haandterer nested objekter/arrays -
' OFV-svaret har flere nivaer: transactions[].from.owner.companyInfo osv.)
' =====================================================================
Private Function JSON_FindMatchingBrace(strJSON As String, lngOpenPos As Long) As Long
    Dim strOpen As String, strClose As String
    strOpen = Mid$(strJSON, lngOpenPos, 1)
    If strOpen = "{" Then
        strClose = "}"
    ElseIf strOpen = "[" Then
        strClose = "]"
    Else
        JSON_FindMatchingBrace = 0
        Exit Function
    End If

    Dim lngDepth As Long, i As Long, blnInString As Boolean, strChar As String
    lngDepth = 0
    blnInString = False
    For i = lngOpenPos To Len(strJSON)
        strChar = Mid$(strJSON, i, 1)
        If blnInString Then
            If strChar = "\" Then
                i = i + 1
            ElseIf strChar = Chr(34) Then
                blnInString = False
            End If
        Else
            If strChar = Chr(34) Then
                blnInString = True
            ElseIf strChar = strOpen Then
                lngDepth = lngDepth + 1
            ElseIf strChar = strClose Then
                lngDepth = lngDepth - 1
                If lngDepth = 0 Then
                    JSON_FindMatchingBrace = i
                    Exit Function
                End If
            End If
        End If
    Next i
    JSON_FindMatchingBrace = 0
End Function


Private Function JSON_SkipWhitespace(strJSON As String, lngPos As Long) As Long
    Dim p As Long, strChar As String
    p = lngPos
    Do While p <= Len(strJSON)
        strChar = Mid$(strJSON, p, 1)
        If strChar = " " Or strChar = vbLf Or strChar = vbCr Or strChar = vbTab Then
            p = p + 1
        Else
            Exit Do
        End If
    Loop
    JSON_SkipWhitespace = p
End Function


Private Function JSON_ExtractObject(strJSON As String, strKey As String) As String
    ' Finner "strKey":{...} eller "strKey":[...] og returnerer hele
    ' verdien inkl. ytre klammer/parenteser. Tom streng hvis nokkelen
    ' ikke finnes, er null, eller ikke er et objekt/array.
    If Len(strJSON) = 0 Then Exit Function

    Dim lngKeyPos As Long, lngPos As Long
    lngKeyPos = InStr(1, strJSON, Chr(34) & strKey & Chr(34) & ":", vbBinaryCompare)
    If lngKeyPos = 0 Then Exit Function

    lngPos = lngKeyPos + Len(strKey) + 3
    lngPos = JSON_SkipWhitespace(strJSON, lngPos)

    Dim strFirstChar As String
    strFirstChar = Mid$(strJSON, lngPos, 1)
    If strFirstChar <> "{" And strFirstChar <> "[" Then Exit Function

    Dim lngEnd As Long
    lngEnd = JSON_FindMatchingBrace(strJSON, lngPos)
    If lngEnd = 0 Then Exit Function

    JSON_ExtractObject = Mid$(strJSON, lngPos, lngEnd - lngPos + 1)
End Function


Private Function JSON_ExtractValue(strJSON As String, strKey As String) As Variant
    ' Henter en skalarverdi (streng/tall/bool/null) for strKey.
    JSON_ExtractValue = Null
    If Len(strJSON) = 0 Then Exit Function

    Dim lngKeyPos As Long, lngPos As Long
    lngKeyPos = InStr(1, strJSON, Chr(34) & strKey & Chr(34) & ":", vbBinaryCompare)
    If lngKeyPos = 0 Then Exit Function

    lngPos = lngKeyPos + Len(strKey) + 3
    lngPos = JSON_SkipWhitespace(strJSON, lngPos)

    Dim strFirstChar As String
    strFirstChar = Mid$(strJSON, lngPos, 1)

    If strFirstChar = Chr(34) Then
        Dim i As Long, strResult As String, c As String, nc As String
        i = lngPos + 1
        Do While i <= Len(strJSON)
            c = Mid$(strJSON, i, 1)
            If c = "\" Then
                nc = Mid$(strJSON, i + 1, 1)
                Select Case nc
                    Case "n": strResult = strResult & vbLf
                    Case "r": strResult = strResult & vbCr
                    Case "t": strResult = strResult & vbTab
                    Case Chr(34): strResult = strResult & Chr(34)
                    Case "\": strResult = strResult & "\"
                    Case Else: strResult = strResult & nc
                End Select
                i = i + 2
            ElseIf c = Chr(34) Then
                Exit Do
            Else
                strResult = strResult & c
                i = i + 1
            End If
        Loop
        JSON_ExtractValue = strResult

    ElseIf strFirstChar = "{" Or strFirstChar = "[" Then
        Dim lngEnd As Long
        lngEnd = JSON_FindMatchingBrace(strJSON, lngPos)
        If lngEnd > 0 Then JSON_ExtractValue = Mid$(strJSON, lngPos, lngEnd - lngPos + 1)

    Else
        Dim j As Long, strRaw As String, ch As String
        j = lngPos
        Do While j <= Len(strJSON)
            ch = Mid$(strJSON, j, 1)
            If ch = "," Or ch = "}" Or ch = "]" Then Exit Do
            strRaw = strRaw & ch
            j = j + 1
        Loop
        strRaw = Trim$(strRaw)
        Select Case LCase$(strRaw)
            Case "true": JSON_ExtractValue = True
            Case "false": JSON_ExtractValue = False
            Case "null": JSON_ExtractValue = Null
            Case Else
                If IsNumeric(strRaw) Then
                    JSON_ExtractValue = CDbl(strRaw)
                Else
                    JSON_ExtractValue = strRaw
                End If
        End Select
    End If
End Function


Private Function JSON_ArrayAllElements(strJSONArray As String) As Collection
    ' strJSONArray inkl. ytre "[" "]" - returnerer ALLE elementer i
    ' arrayet (objekter/arrays som understrenger, eller skalarverdier
    ' som tekst), i original rekkefolge.
    Dim colResult As New Collection
    If Len(strJSONArray) < 2 Then
        Set JSON_ArrayAllElements = colResult
        Exit Function
    End If
    If Left$(strJSONArray, 1) <> "[" Then
        Set JSON_ArrayAllElements = colResult
        Exit Function
    End If

    Dim lngPos As Long
    lngPos = JSON_SkipWhitespace(strJSONArray, 2)

    Do While lngPos <= Len(strJSONArray)
        If Mid$(strJSONArray, lngPos, 1) = "]" Then Exit Do

        Dim strFirstChar As String
        strFirstChar = Mid$(strJSONArray, lngPos, 1)
        Dim strElement As String, lngEnd As Long

        If strFirstChar = "{" Or strFirstChar = "[" Then
            lngEnd = JSON_FindMatchingBrace(strJSONArray, lngPos)
            If lngEnd = 0 Then Exit Do
            strElement = Mid$(strJSONArray, lngPos, lngEnd - lngPos + 1)
            lngPos = lngEnd + 1
        Else
            Dim i As Long, strResult As String, c As String
            i = lngPos
            Do While i <= Len(strJSONArray)
                c = Mid$(strJSONArray, i, 1)
                If c = "," Or c = "]" Then Exit Do
                strResult = strResult & c
                i = i + 1
            Loop
            strElement = Trim$(strResult)
            lngPos = i
        End If

        colResult.Add strElement

        lngPos = JSON_SkipWhitespace(strJSONArray, lngPos)
        If lngPos <= Len(strJSONArray) Then
            If Mid$(strJSONArray, lngPos, 1) = "," Then
                lngPos = JSON_SkipWhitespace(strJSONArray, lngPos + 1)
            End If
        End If
    Loop

    Set JSON_ArrayAllElements = colResult
End Function
