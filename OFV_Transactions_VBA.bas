Attribute VB_Name = "modOFV"
Option Explicit

' =====================================================================
' OFV Transactions API - Excel VBA-versjon
' updated 2026-09-18
'
' Portert fra ofv_transactions_lookup.py, i samme stil som den
' eksisterende Vegvesenet-makroen (CreateObject-basert HTTP, ingen
' faste VBA-referanser, Dictionary for oppslag, tabell som datakilde).
'
' Hva gjor den?
' -------------
' Leser Regnr/VIN fra en Excel-tabell, slar opp hvert unike kjoretoy mot
' OFV sitt Transactions-API, og skriver kjoretoydata, eierskiftedatoer
' og selger-/kjoperinfo tilbake i tabellen.
'
' Forutsetninger i arbeidsboken:
' - Et navngitt omrade "OFV_API" (Formler > Navnebehandling) som
'   inneholder API-nokkelen (allerede satt opp).
' - (Valgfritt) to navngitte omrader "OFV_PeriodFra" og "OFV_PeriodTil"
'   som peker til to celler formatert som dato. La cellene sta tomme for
'   a IKKE bruke periodefilteret.
' - En Excel-tabell (Sett inn > Tabell) med navnet angitt i TABLE_NAME
'   nedenfor, med kolonneoverskrifter som listet i konstantene under
'   "Kolonneoverskrifter". Se veiledningen for full liste.
' =====================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ---------------------------------------------------------------------
' Konfigurasjon
' ---------------------------------------------------------------------
Private Const TABLE_NAME As String = "tblOFV"
Private Const OFV_BASE_URL As String = "https://api.ofv.no/transactions/v1/"
Private Const OFV_MAX_RETRIES As Long = 4
Private Const OFV_RETRY_WAIT_MS As Long = 3000
Private Const OFV_PAUSE_MS As Long = 200

' Kolonneoverskrifter - ma finnes i tabellen TABLE_NAME (rekkefolge er
' likegyldig, det er kun overskriftsteksten som brukes til oppslag).
Private Const COL_REGNR As String = "Regnr"                          ' input
Private Const COL_VIN As String = "VIN"                              ' input
Private Const COL_KILDE As String = "Kilde"
Private Const COL_REGNO As String = "RegNo"
Private Const COL_CHASSIS As String = "Chassisnummer"
Private Const COL_MAKE As String = "Merke"
Private Const COL_MODEL As String = "Modell"
Private Const COL_REGTYPE As String = "RegistreringsType"
Private Const COL_FUEL As String = "Drivstoffgruppe"
Private Const COL_LEASED As String = "Leaset"
Private Const COL_USEDIMPORT As String = "Bruktimportert"
Private Const COL_FIRSTREG As String = "ForstegangsRegistrering"
Private Const COL_LASTTRANS As String = "SisteEierskifte"
Private Const COL_SELGER_TYPE As String = "SelgerEierType"
Private Const COL_SELGER_FIRMA As String = "SelgerEierFirma"
Private Const COL_SELGER_FYLKE As String = "SelgerEierFylke"
Private Const COL_SELGER_KOMMUNE As String = "SelgerEierKommune"
Private Const COL_SELGER_BRUKER_FYLKE As String = "SelgerBrukerFylke"
Private Const COL_KJOPER_TYPE As String = "KjoperEierType"
Private Const COL_KJOPER_FIRMA As String = "KjoperEierFirma"
Private Const COL_PERIODE As String = "SisteEierskifteIPeriode"
Private Const COL_STATUS As String = "Status"


' =====================================================================
' HOVEDMAKRO - kjor denne (Alt+F8, eller koble til en knapp)
' =====================================================================
Public Sub OFV_RefreshInfo()
    Dim loTable As ListObject
    Set loTable = OFV_GetTable(TABLE_NAME)
    If loTable Is Nothing Then
        MsgBox "Fant ikke tabellen '" & TABLE_NAME & "'. Opprett en Excel-tabell med" & _
               " dette navnet (Sett inn > Tabell), eller endre TABLE_NAME i toppen av modulen.", _
               vbExclamation, "OFV"
        Exit Sub
    End If
    If loTable.DataBodyRange Is Nothing Then
        MsgBox "Tabellen '" & TABLE_NAME & "' har ingen datarader.", vbInformation, "OFV"
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

    ' kolonne-indekser (header-navn -> kolonnenummer i tabellen)
    Dim objCol As Object
    Set objCol = CreateObject("Scripting.Dictionary")
    objCol.CompareMode = vbTextCompare
    Dim c As Long
    For c = 1 To loTable.ListColumns.Count
        objCol(Trim$(loTable.ListColumns(c).Name)) = c
    Next c

    If Not objCol.Exists(COL_REGNR) Or Not objCol.Exists(COL_VIN) Then
        MsgBox "Tabellen ma ha kolonnene '" & COL_REGNR & "' og '" & COL_VIN & "'.", vbExclamation, "OFV"
        Exit Sub
    End If

    ' periode (valgfritt) - to navngitte celler formatert som dato
    Dim datPeriodFra As Variant, datPeriodTil As Variant, blnPeriodActive As Boolean
    On Error Resume Next
    datPeriodFra = ThisWorkbook.Names("OFV_PeriodFra").RefersToRange.Value
    datPeriodTil = ThisWorkbook.Names("OFV_PeriodTil").RefersToRange.Value
    On Error GoTo 0
    blnPeriodActive = (IsDate(datPeriodFra) And IsDate(datPeriodTil))

    ' hent data fra tabellen til et array
    Dim varData As Variant
    varData = loTable.DataBodyRange.Value
    If Not IsArray(varData) Then
        ' tabellen har bare 1 datarad - Excel gir da ikke et array. Bygg et selv.
        ReDim varData(1 To 1, 1 To loTable.ListColumns.Count)
        For c = 1 To loTable.ListColumns.Count
            varData(1, c) = loTable.DataBodyRange.Cells(1, c).Value
        Next c
    End If

    ' bygg liste over unike kjoretoy (Regnr eller VIN, VIN har forrang)
    Dim objQueue As Object
    Set objQueue = CreateObject("Scripting.Dictionary")
    objQueue.CompareMode = vbTextCompare

    Dim r As Long, strVin As String, strReg As String, strKey As String
    For r = 1 To UBound(varData, 1)
        strVin = Trim$(Replace(CStr(varData(r, objCol(COL_VIN)) & vbNullString), " ", vbNullString))
        strReg = Trim$(Replace(CStr(varData(r, objCol(COL_REGNR)) & vbNullString), " ", vbNullString))
        strKey = OFV_BuildKey(strVin, strReg)
        If Len(strKey) > 0 Then
            If Not objQueue.Exists(strKey) Then objQueue.Add strKey, strKey
        End If
    Next r

    If objQueue.Count = 0 Then
        MsgBox "Fant ingen Regnr eller VIN i tabellen.", vbInformation, "OFV"
        Exit Sub
    End If

    Application.Cursor = xlWait
    Application.ScreenUpdating = False

    Dim objResults As Object
    Set objResults = CreateObject("Scripting.Dictionary")
    objResults.CompareMode = vbTextCompare

    Dim varKey As Variant, i As Long, t As Long, lngErr As Long
    t = objQueue.Count
    For Each varKey In objQueue.Keys
        i = i + 1
        Application.StatusBar = "OFV: Henter data (" & i & " av " & t & ") ..."

        Dim strParts() As String, blnIsVin As Boolean, strIdentifier As String
        strParts = Split(CStr(varKey), "|")
        blnIsVin = (strParts(0) = "VIN")
        strIdentifier = strParts(1)

        Dim strStatus As String
        Dim objFields As Object
        Set objFields = OFV_QueryVehicle(strApiKey, strIdentifier, blnIsVin, _
                                          blnPeriodActive, datPeriodFra, datPeriodTil, strStatus)
        objFields(COL_STATUS) = strStatus
        If strStatus <> "OK" Then lngErr = lngErr + 1
        objResults.Add varKey, objFields

        Sleep OFV_PAUSE_MS
    Next varKey

    ' skriv resultatene tilbake i arrayet
    Dim objFieldMap As Variant
    objFieldMap = OFV_GetFieldMap()

    For r = 1 To UBound(varData, 1)
        strVin = Trim$(Replace(CStr(varData(r, objCol(COL_VIN)) & vbNullString), " ", vbNullString))
        strReg = Trim$(Replace(CStr(varData(r, objCol(COL_REGNR)) & vbNullString), " ", vbNullString))
        strKey = OFV_BuildKey(strVin, strReg)
        If Len(strKey) > 0 Then
            If objResults.Exists(strKey) Then
                Dim objF As Object
                Set objF = objResults(strKey)

                Dim m As Long, strDictKey As String, strColHeader As String
                For m = LBound(objFieldMap) To UBound(objFieldMap)
                    strDictKey = objFieldMap(m)(0)
                    strColHeader = objFieldMap(m)(1)
                    If objCol.Exists(strColHeader) Then
                        If objF.Exists(strDictKey) Then
                            ' dato-kolonner far ekte Date-verdier fra OFV_DateFromISO,
                            ' eller en feiltekst/"Ingen eierskifte i perioden" - begge
                            ' skrives rett inn, Variant-tildelingen haandterer begge typer.
                            varData(r, objCol(strColHeader)) = objF(strDictKey)
                        End If
                    End If
                Next m
            End If
        End If
    Next r

    loTable.DataBodyRange.Value = varData

    ' formater dato-kolonnene (3. element i objFieldMap) som ekte datoer (DD.MM.AAAA)
    For m = LBound(objFieldMap) To UBound(objFieldMap)
        If objFieldMap(m)(2) = True Then
            strColHeader = objFieldMap(m)(1)
            If objCol.Exists(strColHeader) Then
                loTable.ListColumns(strColHeader).DataBodyRange.NumberFormat = "dd.mm.yyyy"
            End If
        End If
    Next m

    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True

    MsgBox "OFV: Ferdig. Slo opp " & t & " kjoretoy (" & lngErr & " uten treff/feil).", _
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
    ' {dictionary-nokkel i objFields, kolonneoverskrift i tabellen, er dato}
    OFV_GetFieldMap = Array( _
        Array(COL_KILDE, COL_KILDE, False), _
        Array("RegNo", COL_REGNO, False), _
        Array("ChassisNumber", COL_CHASSIS, False), _
        Array("MakeName", COL_MAKE, False), _
        Array("ModelName", COL_MODEL, False), _
        Array("RegistrationType", COL_REGTYPE, False), _
        Array("FuelGroup", COL_FUEL, False), _
        Array("IsLeased", COL_LEASED, False), _
        Array("IsUsedImported", COL_USEDIMPORT, False), _
        Array("FirstRegistrationDate", COL_FIRSTREG, True), _
        Array("LastTransactionDate", COL_LASTTRANS, True), _
        Array("FromOwnerType", COL_SELGER_TYPE, False), _
        Array("FromOwnerCompanyName", COL_SELGER_FIRMA, False), _
        Array("FromOwnerCounty", COL_SELGER_FYLKE, False), _
        Array("FromOwnerMunicipality", COL_SELGER_KOMMUNE, False), _
        Array("FromUserCounty", COL_SELGER_BRUKER_FYLKE, False), _
        Array("ToOwnerType", COL_KJOPER_TYPE, False), _
        Array("ToOwnerCompanyName", COL_KJOPER_FIRMA, False), _
        Array("PeriodTransactionDate", COL_PERIODE, True) _
    )
End Function


' =====================================================================
' Kjoretoyoppslag mot OFV Transactions-API (ett kjoretoy)
' =====================================================================
Private Function OFV_QueryVehicle(strApiKey As String, strIdentifier As String, blnIsVin As Boolean, _
                                   blnPeriodActive As Boolean, datPeriodFra As Variant, datPeriodTil As Variant, _
                                   ByRef strStatus As String) As Object
    Dim objFields As Object
    Set objFields = CreateObject("Scripting.Dictionary")
    objFields.CompareMode = vbTextCompare
    objFields(COL_KILDE) = IIf(blnIsVin, "VIN", "Regnr")
    strStatus = "OK"

    Dim strFilterKey As String
    strFilterKey = IIf(blnIsVin, "chassisNumber", "regNo")

    Dim strBody As String
    strBody = "{""filters"":{""" & strFilterKey & """:""" & OFV_JsonEscape(strIdentifier) & """}," & _
              """pagination"":{""first"":1}," & _
              """sorting"":{""orderBy"":""transactionDate"",""orderDirection"":""DESC""}}"

    Dim strResponse As String
    strResponse = OFV_PostWithRetries(strApiKey, strBody, strStatus)
    If strStatus <> "OK" Then
        Set OFV_QueryVehicle = objFields
        Exit Function
    End If

    Dim strTransactionsArray As String, strFirstTransaction As String
    strTransactionsArray = JSON_ExtractObject(strResponse, "transactions")
    strFirstTransaction = JSON_ArrayFirstElement(strTransactionsArray)

    If Len(strFirstTransaction) = 0 Then
        strStatus = "Ingen treff"
        Set OFV_QueryVehicle = objFields
        Exit Function
    End If

    objFields("RegNo") = JSON_ExtractValue(strFirstTransaction, "regNo")
    objFields("ChassisNumber") = JSON_ExtractValue(strFirstTransaction, "chassisNumber")
    objFields("MakeName") = JSON_ExtractValue(strFirstTransaction, "makeName")
    objFields("ModelName") = JSON_ExtractValue(strFirstTransaction, "modelName")
    objFields("RegistrationType") = JSON_ExtractValue(strFirstTransaction, "registrationType")
    objFields("FuelGroup") = JSON_ExtractValue(strFirstTransaction, "fuelGroup")
    objFields("IsLeased") = JSON_ExtractValue(strFirstTransaction, "isLeased")
    objFields("IsUsedImported") = JSON_ExtractValue(strFirstTransaction, "isUsedImported")
    objFields("FirstRegistrationDate") = OFV_DateFromISO(CStr(JSON_ExtractValue(strFirstTransaction, "firstRegistrationDate") & vbNullString))
    objFields("LastTransactionDate") = OFV_DateFromISO(CStr(JSON_ExtractValue(strFirstTransaction, "transactionDate") & vbNullString))

    Dim strFromObj As String, strToObj As String
    Dim strFromOwner As String, strFromUser As String, strToOwner As String, strCompanyObj As String
    strFromObj = JSON_ExtractObject(strFirstTransaction, "from")
    strToObj = JSON_ExtractObject(strFirstTransaction, "to")
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

    If blnPeriodActive Then
        Dim strPeriodBody As String
        strPeriodBody = "{""filters"":{""" & strFilterKey & """:""" & OFV_JsonEscape(strIdentifier) & """," & _
                        """transactionDateFrom"":""" & Format$(datPeriodFra, "yyyy-mm-dd") & """," & _
                        """transactionDateTo"":""" & Format$(datPeriodTil, "yyyy-mm-dd") & """}," & _
                        """pagination"":{""first"":1}," & _
                        """sorting"":{""orderBy"":""transactionDate"",""orderDirection"":""DESC""}}"

        Dim strPeriodStatus As String
        strPeriodStatus = "OK"
        Dim strPeriodResponse As String
        strPeriodResponse = OFV_PostWithRetries(strApiKey, strPeriodBody, strPeriodStatus)

        If strPeriodStatus = "OK" Then
            Dim strPeriodArr As String, strPeriodFirst As String
            strPeriodArr = JSON_ExtractObject(strPeriodResponse, "transactions")
            strPeriodFirst = JSON_ArrayFirstElement(strPeriodArr)
            If Len(strPeriodFirst) > 0 Then
                objFields("PeriodTransactionDate") = OFV_DateFromISO(CStr(JSON_ExtractValue(strPeriodFirst, "transactionDate") & vbNullString))
            Else
                objFields("PeriodTransactionDate") = "Ingen eierskifte i perioden"
            End If
        Else
            objFields("PeriodTransactionDate") = "Feil ved periodesok: " & strPeriodStatus
        End If
    End If

    Set OFV_QueryVehicle = objFields
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


Private Function OFV_GetTable(strName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, strName, vbTextCompare) = 0 Then
                Set OFV_GetTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function


' =====================================================================
' Generiske JSON-hjelpefunksjoner (haandterer nested objekter/arrays,
' i motsetning til den enklere strengsokingen i Vegvesenet-makroen -
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


Private Function JSON_ArrayFirstElement(strJSONArray As String) As String
    ' strJSONArray inkl. ytre "[" "]" - returnerer forste element
    ' (objekt/array som understreng, eller skalarverdi som tekst).
    If Len(strJSONArray) < 2 Then Exit Function
    If Left$(strJSONArray, 1) <> "[" Then Exit Function

    Dim lngPos As Long
    lngPos = JSON_SkipWhitespace(strJSONArray, 2)
    If lngPos > Len(strJSONArray) Then Exit Function
    If Mid$(strJSONArray, lngPos, 1) = "]" Then Exit Function ' tom liste

    Dim strFirstChar As String
    strFirstChar = Mid$(strJSONArray, lngPos, 1)
    If strFirstChar = "{" Or strFirstChar = "[" Then
        Dim lngEnd As Long
        lngEnd = JSON_FindMatchingBrace(strJSONArray, lngPos)
        If lngEnd > 0 Then JSON_ArrayFirstElement = Mid$(strJSONArray, lngPos, lngEnd - lngPos + 1)
    Else
        Dim i As Long, strResult As String, c As String
        i = lngPos
        Do While i <= Len(strJSONArray)
            c = Mid$(strJSONArray, i, 1)
            If c = "," Or c = "]" Then Exit Do
            strResult = strResult & c
            i = i + 1
        Loop
        JSON_ArrayFirstElement = Trim$(strResult)
    End If
End Function
