Option Explicit

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

Private Const OFV_BASE_URL As String = _
    "https://api.ofv.no/transactions/v1/"

Private Const OFV_MAX_RETRIES As Long = 4
Private Const OFV_RETRY_WAIT_MS As Long = 3000
Private Const OFV_PAUSE_MS As Long = 150
Private Const OFV_SORT_DIRECTION As String = "ASC"


'==============================================================
' NORSKE TEKSTER - BYGGES MED CHRW FOR A UNNGA o/ae
'==============================================================

Private Function OFV_BuyerTypeHeader() As String
    OFV_BuyerTypeHeader = "Kj" & ChrW(248) & "perType"
End Function

Private Function OFV_BuyerCountyHeader() As String
    OFV_BuyerCountyHeader = _
        "Kj" & ChrW(248) & "perEierFylke"
End Function

Private Function OFV_VehicleWord() As String
    OFV_VehicleWord = _
        "kj" & ChrW(248) & "ret" & ChrW(248) & "y"
End Function


'==============================================================
' HOVEDMAKRO
'==============================================================

Public Sub OFV_RefreshInfo()

    Dim wsInput As Worksheet
    Dim wsResult As Worksheet
    Dim wsOverview As Worksheet
    Dim wsControl As Worksheet
    Dim loResult As ListObject

    Dim objQueue As Object
    Dim objHitVehicles As Object
    Dim objFieldMap As Variant

    Dim colAllRows As Collection
    Dim colVehicleRows As Collection
    Dim objRow As Object

    Dim strApiKey As String
    Dim datFrom As Variant
    Dim datTo As Variant
    Dim strDateFromIso As String
    Dim strDateToIso As String
    Dim strStage As String

    Dim lngLastRowB As Long
    Dim lngLastRowC As Long
    Dim lngLastRow As Long
    Dim lngOldLastRow As Long
    Dim lngNewLastRow As Long
    Dim lngFieldCount As Long
    Dim lngOutputRows As Long
    Dim lngTransactionRows As Long
    Dim lngNoHit As Long
    Dim lngApiErrors As Long

    Dim r As Long
    Dim m As Long
    Dim i As Long
    Dim t As Long

    Dim strVin As String
    Dim strReg As String
    Dim strKey As String
    Dim strIdentifier As String
    Dim strDictKey As String
    Dim strStatus As String
    Dim strParts() As String

    Dim blnIsVin As Boolean
    Dim varKey As Variant
    Dim varValue As Variant
    Dim arrOutput() As Variant

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldCursor As Variant
    Dim blnApplicationChanged As Boolean

    On Error GoTo FatalError

    strStage = "finner arkene"

    Set wsInput = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set wsResult = ThisWorkbook.Worksheets(RESULT_SHEET)
    Set wsOverview = ThisWorkbook.Worksheets(OVERVIEW_SHEET)

    On Error Resume Next
    Set wsControl = ThisWorkbook.Worksheets(CONTROL_SHEET)
    On Error GoTo FatalError

    strStage = "leser API-nokkel og datoer"

    strApiKey = Trim$(CStr( _
        ThisWorkbook.Names("OFV_API").RefersToRange.Value))

    datFrom = ThisWorkbook.Names( _
        "OFV_DateFrom").RefersToRange.Value

    datTo = ThisWorkbook.Names( _
        "OFV_DateTo").RefersToRange.Value

    If Len(strApiKey) = 0 Then
        MsgBox "Fant ingen API-nokkel i OFV_API.", _
               vbExclamation, "OFV"
        Exit Sub
    End If

    If Not IsDate(datFrom) Or Not IsDate(datTo) Then
        MsgBox "Fyll inn gyldige datoer i Input!B2:B3.", _
               vbExclamation, "OFV"
        Exit Sub
    End If

    If CDate(datFrom) > CDate(datTo) Then
        MsgBox "Fra-dato kan ikke vaere senere enn til-dato.", _
               vbExclamation, "OFV"
        Exit Sub
    End If

    strDateFromIso = Format$(CDate(datFrom), "yyyy-mm-dd")
    strDateToIso = Format$(CDate(datTo), "yyyy-mm-dd")

    strStage = "leser kjoretoylisten"

    lngLastRowB = wsInput.Cells( _
        wsInput.Rows.Count, COL_REGNR).End(xlUp).Row

    lngLastRowC = wsInput.Cells( _
        wsInput.Rows.Count, COL_VIN).End(xlUp).Row

    lngLastRow = Application.Max(lngLastRowB, lngLastRowC)

    If lngLastRow < FIRST_ROW Then
        MsgBox "Fant ingen registreringsnummer eller VIN.", _
               vbInformation, "OFV"
        Exit Sub
    End If

    Set objQueue = CreateObject("Scripting.Dictionary")
    objQueue.CompareMode = vbTextCompare

    For r = FIRST_ROW To lngLastRow

        strVin = UCase$(Trim$(Replace( _
            CStr(wsInput.Cells(r, COL_VIN).Value & vbNullString), _
            " ", vbNullString)))

        strReg = UCase$(Trim$(Replace( _
            CStr(wsInput.Cells(r, COL_REGNR).Value & vbNullString), _
            " ", vbNullString)))

        strKey = OFV_BuildKey(strVin, strReg)

        If Len(strKey) > 0 Then
            If Not objQueue.Exists(strKey) Then
                objQueue.Add strKey, strKey
            End If
        End If

    Next r

    If objQueue.Count = 0 Then
        MsgBox "Fant ingen gyldige registreringsnummer eller VIN.", _
               vbInformation, "OFV"
        Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation
    oldCursor = Application.Cursor

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait

    blnApplicationChanged = True

    objFieldMap = OFV_GetFieldMap()
    lngFieldCount = UBound(objFieldMap) + 1

    Set colAllRows = New Collection
    Set objHitVehicles = CreateObject("Scripting.Dictionary")
    objHitVehicles.CompareMode = vbTextCompare

    t = objQueue.Count
    strStage = "henter data fra OFV"

    For Each varKey In objQueue.Keys

        i = i + 1

        Application.StatusBar = _
            "OFV: Henter data (" & i & " av " & t & ") ..."

        strParts = Split(CStr(varKey), "|")
        blnIsVin = (strParts(0) = "VIN")
        strIdentifier = strParts(1)

        Set colVehicleRows = OFV_FetchAllTransactionRows( _
            strApiKey, _
            strIdentifier, _
            blnIsVin, _
            strDateFromIso, _
            strDateToIso)

        For Each objRow In colVehicleRows

            colAllRows.Add objRow
            strStatus = OFV_VariantToString(objRow("Status"))

            If strStatus = "OK" Then

                lngTransactionRows = lngTransactionRows + 1

                If Not objHitVehicles.Exists(strIdentifier) Then
                    objHitVehicles.Add strIdentifier, True
                End If

            ElseIf Left$(strStatus, 5) = "Feil:" Then

                lngApiErrors = lngApiErrors + 1

            Else

                lngNoHit = lngNoHit + 1

            End If

        Next objRow

        Sleep OFV_PAUSE_MS

    Next varKey

    lngOutputRows = colAllRows.Count

    strStage = "finner resultat-tabellen"

    Set loResult = OFV_GetOrCreateResultTable( _
        wsResult, objFieldMap)

    lngOldLastRow = _
        loResult.Range.Row + loResult.Range.Rows.Count - 1

    If lngOutputRows > 0 Then
        lngNewLastRow = lngOutputRows + 1
    Else
        lngNewLastRow = 2
    End If

    strStage = "endrer tabellstorrelsen"

    Set loResult = OFV_ResizeResultTable( _
        wsResult, loResult, lngNewLastRow, lngFieldCount)

    strStage = "tommer gamle data"

    If Not loResult.DataBodyRange Is Nothing Then
        loResult.DataBodyRange.ClearContents
    End If

    If lngOldLastRow > lngNewLastRow Then

        With wsResult.Range( _
            wsResult.Cells(lngNewLastRow + 1, 1), _
            wsResult.Cells(lngOldLastRow, lngFieldCount))

            .ClearContents
            .ClearFormats

        End With

    End If

    'Fjerner eventuelle gamle hjelpekolonner
    wsResult.Range("W:AC").Clear

    strStage = "oppdaterer overskriftene"

    For m = LBound(objFieldMap) To UBound(objFieldMap)

        loResult.HeaderRowRange.Cells(1, m + 1).Value = _
            objFieldMap(m)(1)

    Next m

    If lngOutputRows > 0 Then

        ReDim arrOutput( _
            1 To lngOutputRows, _
            1 To lngFieldCount)

        For r = 1 To lngOutputRows

            Set objRow = colAllRows(r)

            For m = LBound(objFieldMap) To UBound(objFieldMap)

                strDictKey = CStr(objFieldMap(m)(0))

                If strDictKey = "CalculatedSellerType" Or _
                   strDictKey = "CalculatedBuyerType" Then

                    arrOutput(r, m + 1) = Empty

                ElseIf objRow.Exists(strDictKey) Then

                    varValue = objRow(strDictKey)

                    If IsNull(varValue) Or IsEmpty(varValue) Then
                        arrOutput(r, m + 1) = Empty
                    Else
                        arrOutput(r, m + 1) = varValue
                    End If

                Else
                    arrOutput(r, m + 1) = Empty
                End If

            Next m

        Next r

        strStage = "skriver resultatdata"

        loResult.DataBodyRange.Value = arrOutput

        strStage = "oppretter selger og kjoper"

        OFV_ApplyCalculatedColumns loResult

    End If

    strStage = "formaterer resultat-tabellen"

    OFV_FormatResultTable wsResult, loResult

    strStage = "oppdaterer oversikten"

    OFV_UpdateOverviewKPIs wsOverview
    OFV_RebuildOverviewPivot wsOverview, loResult

    strStage = "oppdaterer kontrollarket"

    If Not wsControl Is Nothing Then
        wsControl.Calculate
    End If

    Application.Calculation = oldCalculation

    If oldCalculation = xlCalculationManual Then

        wsResult.Calculate
        wsOverview.Calculate

        If Not wsControl Is Nothing Then
            wsControl.Calculate
        End If

    Else
        Application.CalculateFull
    End If

    Application.StatusBar = False
    Application.Cursor = oldCursor
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    blnApplicationChanged = False

    MsgBox _
        "OFV: Ferdig." & vbCrLf & vbCrLf & _
        t & " " & OFV_VehicleWord() & " lest." & vbCrLf & _
        objHitVehicles.Count & " med treff." & vbCrLf & _
        lngTransactionRows & " eierskifter funnet." & vbCrLf & _
        lngNoHit & " uten registreringer i perioden." & vbCrLf & _
        lngApiErrors & " API-feil." & vbCrLf & vbCrLf & _
        "Resultat, Oversikt og Kontroll solgte biler er oppdatert.", _
        vbInformation, "OFV"

    Exit Sub

FatalError:

    Dim lngErrorNumber As Long
    Dim strErrorDescription As String

    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description

    If blnApplicationChanged Then

        Application.StatusBar = False
        Application.Cursor = oldCursor
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        Application.Calculation = oldCalculation

    End If

    MsgBox _
        "Oppdateringen ble avbrutt." & vbCrLf & vbCrLf & _
        "Trinn: " & strStage & vbCrLf & _
        "Feil " & lngErrorNumber & ": " & strErrorDescription, _
        vbCritical, "OFV"

End Sub


'==============================================================
' FORMELKOLONNER
'==============================================================

Private Sub OFV_ApplyCalculatedColumns(ByVal lo As ListObject)

    If lo.DataBodyRange Is Nothing Then Exit Sub

    With lo.ListColumns("SelgerType").DataBodyRange

        .Formula = _
            "=IF(AND([@SelgerEierType]=""""," & _
            "[@SelgerEierFirma]=""""),""""," & _
            "IF([@SelgerEierType]=""Privat"",""Privat""," & _
            "IF([@SelgerEierFirma]<>""""," & _
            "[@SelgerEierFirma],[@SelgerEierType])))"

    End With

    With lo.ListColumns(OFV_BuyerTypeHeader()).DataBodyRange

        .Formula = _
            "=IF(AND([@KjoperEierType]=""""," & _
            "[@KjoperEierFirma]=""""),""""," & _
            "IF([@KjoperEierType]=""Privat"",""Privat""," & _
            "IF([@KjoperEierFirma]<>""""," & _
            "[@KjoperEierFirma],[@KjoperEierType])))"

    End With

    lo.ListColumns("SelgerType").DataBodyRange.Calculate
    lo.ListColumns(OFV_BuyerTypeHeader()).DataBodyRange.Calculate

End Sub


'==============================================================
' RESULTATTABELL
'==============================================================

Private Function OFV_GetOrCreateResultTable( _
    ByVal ws As Worksheet, _
    ByVal objFieldMap As Variant) As ListObject

    Dim lo As ListObject
    Dim rngTable As Range
    Dim lngFieldCount As Long
    Dim m As Long

    lngFieldCount = UBound(objFieldMap) + 1

    On Error Resume Next
    Set lo = ws.ListObjects(RESULT_TABLE)
    On Error GoTo 0

    If lo Is Nothing Then

        For m = LBound(objFieldMap) To UBound(objFieldMap)
            ws.Cells(1, m + 1).Value = objFieldMap(m)(1)
        Next m

        Set rngTable = ws.Range( _
            ws.Cells(1, 1), _
            ws.Cells(2, lngFieldCount))

        Set lo = ws.ListObjects.Add( _
            xlSrcRange, rngTable, , xlYes)

        lo.Name = RESULT_TABLE
        lo.DataBodyRange.ClearContents

    End If

    Set OFV_GetOrCreateResultTable = lo

End Function


Private Function OFV_ResizeResultTable( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject, _
    ByVal lngLastRow As Long, _
    ByVal lngFieldCount As Long) As ListObject

    Dim rngTarget As Range
    Dim strStyle As String

    Set rngTarget = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(lngLastRow, lngFieldCount))

    strStyle = lo.TableStyle

    On Error Resume Next

    If ws.FilterMode Then ws.ShowAllData
    lo.AutoFilter.ShowAllData
    lo.ShowAutoFilter = False

    Err.Clear
    lo.Resize rngTarget

    If Err.Number <> 0 Then

        Err.Clear
        lo.Unlist

        Set lo = ws.ListObjects.Add( _
            xlSrcRange, rngTarget, , xlYes)

        lo.Name = RESULT_TABLE

    End If

    On Error GoTo 0

    lo.ShowAutoFilter = True

    If Len(strStyle) > 0 Then
        lo.TableStyle = strStyle
    Else
        lo.TableStyle = "TableStyleMedium2"
    End If

    Set OFV_ResizeResultTable = lo

End Function


'==============================================================
' FORMATER RESULTATTABELL
'==============================================================

Private Sub OFV_FormatResultTable( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject)

    Dim lngLastRow As Long
    Dim r As Long
    Dim strCurrentKey As String
    Dim strPreviousKey As String

    lngLastRow = _
        lo.Range.Row + lo.Range.Rows.Count - 1

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
        .RowHeight = 30

    End With

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 9
    ws.Columns("C").ColumnWidth = 11
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 11
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 24
    ws.Columns("H").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 9
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 14
    ws.Columns("L").ColumnWidth = 11
    ws.Columns("M").ColumnWidth = 14
    ws.Columns("N").ColumnWidth = 25
    ws.Columns("O").ColumnWidth = 25
    ws.Columns("P").ColumnWidth = 13
    ws.Columns("Q").ColumnWidth = 25
    ws.Columns("R").ColumnWidth = 13
    ws.Columns("S").ColumnWidth = 13
    ws.Columns("T").ColumnWidth = 25
    ws.Columns("U").ColumnWidth = 15
    ws.Columns("V").ColumnWidth = 28

    'Skjulte stottekolonner
    ws.Columns("A:B").Hidden = True
    ws.Columns("G").Hidden = True
    ws.Columns("P:Q").Hidden = True
    ws.Columns("S:T").Hidden = True

    'Synlige resultatkolonner
    ws.Columns("C:F").Hidden = False
    ws.Columns("H:O").Hidden = False
    ws.Columns("R").Hidden = False
    ws.Columns("U:V").Hidden = False

    If lngLastRow < 2 Then Exit Sub

    With ws.Range("A2:V" & lngLastRow)

        .VerticalAlignment = xlCenter
        .Font.Color = RGB(31, 31, 31)
        .Font.Bold = False
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = False
        .RowHeight = 18

    End With

    ws.Range("K2:K" & lngLastRow).NumberFormat = "dd.mm.yyyy"
    ws.Range("L2:L" & lngLastRow).NumberFormat = "0"
    ws.Range("M2:M" & lngLastRow).NumberFormat = "dd.mm.yyyy"

    For r = 2 To lngLastRow

        strCurrentKey = Trim$(CStr( _
            ws.Cells(r, "A").Value & vbNullString))

        If r = 2 Then
            strPreviousKey = vbNullString
        Else
            strPreviousKey = Trim$(CStr( _
                ws.Cells(r - 1, "A").Value & vbNullString))
        End If

        With ws.Range("A" & r & ":V" & r)

            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone

        End With

        If r = 2 Or strCurrentKey <> strPreviousKey Then

            With ws.Range("A" & r & ":V" & r)

                .Interior.Color = RGB(217, 234, 247)
                .Font.Color = RGB(31, 31, 31)
                .Font.Bold = True

                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(91, 155, 213)
                End With

            End With

        End If

        With ws.Range("A" & r & ":V" & r) _
            .Borders(xlEdgeBottom)

            .LineStyle = xlContinuous
            .Weight = xlHairline
            .Color = RGB(217, 217, 217)

        End With

        If ws.Cells(r, "V").Value <> "OK" Then

            With ws.Cells(r, "V")
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)
                .Font.Bold = True
            End With

        End If

    Next r

End Sub


'==============================================================
' OVERSIKT OG PIVOT
'==============================================================

Private Sub OFV_UpdateOverviewKPIs(ByVal ws As Worksheet)

    ws.Range("A5").Formula = _
        "=SUMPRODUCT(--((KjoretoyInput[Regnr]<>"""")+" & _
        "(KjoretoyInput[VIN]<>"""")>0))"

    ws.Range("C5").Formula = _
        "=COUNTIF(Transaksjoner[Status],""OK"")"

    ws.Range("F5").Formula = _
        "=COUNTIF(Transaksjoner[Status],""<>OK"")"

End Sub


Private Sub OFV_RebuildOverviewPivot( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject)

    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim arrFields As Variant
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

    arrFields = Array( _
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
        OFV_BuyerTypeHeader(), _
        "Status")

    pt.ManualUpdate = True

    For i = LBound(arrFields) To UBound(arrFields)

        Set pf = pt.PivotFields(CStr(arrFields(i)))

        pf.Orientation = xlRowField
        pf.Position = i + 1

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

End Sub


'==============================================================
' IDENTIFIKATOR
'==============================================================

Private Function OFV_BuildKey( _
    ByVal strVin As String, _
    ByVal strReg As String) As String

    If Len(strVin) > 0 Then
        OFV_BuildKey = "VIN|" & UCase$(strVin)
    ElseIf Len(strReg) > 0 Then
        OFV_BuildKey = "REG|" & UCase$(strReg)
    Else
        OFV_BuildKey = vbNullString
    End If

End Function


'==============================================================
' FELTMAPPING - 22 KOLONNER A:V
'==============================================================

Private Function OFV_GetFieldMap() As Variant

    Dim fields(0 To 21) As Variant

    fields(0) = Array("Input", "Input", False)
    fields(1) = Array("Kilde", "Kilde", False)
    fields(2) = Array("RegNo", "RegNo", False)
    fields(3) = Array("ChassisNumber", "Chassisnummer", False)
    fields(4) = Array("MakeName", "Merke", False)
    fields(5) = Array("ModelName", "Modell", False)
    fields(6) = Array("RegistrationType", "RegistreringsType", False)
    fields(7) = Array("FuelGroup", "Drivstoffgruppe", False)
    fields(8) = Array("IsLeased", "Leaset", False)
    fields(9) = Array("IsUsedImported", "Bruktimportert", False)

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
        OFV_BuyerTypeHeader(), False)

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
        "ToOwnerCounty", _
        OFV_BuyerCountyHeader(), False)

    fields(21) = Array("Status", "Status", False)

    OFV_GetFieldMap = fields

End Function


'==============================================================
' HENT TRANSAKSJONER
'==============================================================

Private Function OFV_FetchAllTransactionRows( _
    ByVal strApiKey As String, _
    ByVal strIdentifier As String, _
    ByVal blnIsVin As Boolean, _
    ByVal strDateFromIso As String, _
    ByVal strDateToIso As String) As Collection

    Dim colRows As New Collection
    Dim colItems As Collection

    Dim strFilterKey As String
    Dim strCursor As String
    Dim strBody As String
    Dim strStatus As String
    Dim strResponse As String
    Dim strTransactionsArray As String
    Dim strPaginationObj As String

    Dim varItem As Variant
    Dim varHasNext As Variant
    Dim varCursor As Variant
    Dim blnHasNext As Boolean

    strFilterKey = IIf( _
        blnIsVin, "chassisNumber", "regNo")

    strCursor = vbNullString

    Do

        strBody = _
            "{""filters"":{""" & strFilterKey & """:""" & _
            OFV_JsonEscape(strIdentifier) & """," & _
            """transactionDateFrom"":""" & strDateFromIso & """," & _
            """transactionDateTo"":""" & strDateToIso & """}," & _
            """pagination"":{""first"":1000"

        If Len(strCursor) > 0 Then

            strBody = strBody & _
                ",""cursor"":""" & _
                OFV_JsonEscape(strCursor) & """"

        End If

        strBody = strBody & "}," & _
            """sorting"":{""orderBy"":""transactionDate""," & _
            """orderDirection"":""" & _
            OFV_SORT_DIRECTION & """}}"

        strResponse = OFV_PostWithRetries( _
            strApiKey, strBody, strStatus)

        If strStatus <> "OK" Then

            colRows.Add OFV_BuildEmptyFields( _
                strIdentifier, blnIsVin, strStatus)

            Set OFV_FetchAllTransactionRows = colRows
            Exit Function

        End If

        strTransactionsArray = _
            JSON_ExtractObject(strResponse, "transactions")

        Set colItems = _
            JSON_ArrayAllElements(strTransactionsArray)

        For Each varItem In colItems

            colRows.Add OFV_BuildFieldsFromTransaction( _
                strIdentifier, CStr(varItem), blnIsVin)

        Next varItem

        strPaginationObj = _
            JSON_ExtractObject(strResponse, "pagination")

        varHasNext = JSON_ExtractValue( _
            strPaginationObj, "hasNextPage")

        varCursor = JSON_ExtractValue( _
            strPaginationObj, "endCursor")

        blnHasNext = False

        Select Case VarType(varHasNext)

            Case vbBoolean
                blnHasNext = CBool(varHasNext)

            Case vbString
                blnHasNext = _
                    (LCase$(CStr(varHasNext)) = "true")

            Case vbByte, vbInteger, vbLong, _
                 vbSingle, vbDouble, vbCurrency

                blnHasNext = (varHasNext <> 0)

        End Select

        If blnHasNext Then

            If IsNull(varCursor) Or IsEmpty(varCursor) Then
                Exit Do
            End If

            strCursor = CStr(varCursor)

            If Len(strCursor) = 0 Then Exit Do

            Sleep OFV_PAUSE_MS

        Else
            Exit Do
        End If

    Loop

    If colRows.Count = 0 Then

        colRows.Add OFV_BuildEmptyFields( _
            strIdentifier, _
            blnIsVin, _
            "Ingen registreringer i perioden")

    End If

    Set OFV_FetchAllTransactionRows = colRows

End Function


Private Function OFV_BuildEmptyFields( _
    ByVal strIdentifier As String, _
    ByVal blnIsVin As Boolean, _
    ByVal strStatus As String) As Object

    Dim objFields As Object

    Set objFields = CreateObject("Scripting.Dictionary")
    objFields.CompareMode = vbTextCompare

    objFields("Input") = strIdentifier
    objFields("Kilde") = IIf(blnIsVin, "VIN", "Regnr")
    objFields("Status") = strStatus

    'Vis identifikatoren ogsa nar API-et ikke gir treff
    If blnIsVin Then
        objFields("ChassisNumber") = strIdentifier
    Else
        objFields("RegNo") = strIdentifier
    End If

    Set OFV_BuildEmptyFields = objFields

End Function


Private Function OFV_BuildFieldsFromTransaction( _
    ByVal strIdentifier As String, _
    ByVal strTxnJson As String, _
    ByVal blnIsVin As Boolean) As Object

    Dim objFields As Object
    Dim strFromObj As String
    Dim strToObj As String
    Dim strFromOwner As String
    Dim strToOwner As String
    Dim strCompanyObj As String

    Set objFields = CreateObject("Scripting.Dictionary")
    objFields.CompareMode = vbTextCompare

    objFields("Input") = strIdentifier
    objFields("Kilde") = IIf(blnIsVin, "VIN", "Regnr")

    objFields("RegNo") = _
        JSON_ExtractValue(strTxnJson, "regNo")

    objFields("ChassisNumber") = _
        JSON_ExtractValue(strTxnJson, "chassisNumber")

    objFields("MakeName") = _
        JSON_ExtractValue(strTxnJson, "makeName")

    objFields("ModelName") = _
        JSON_ExtractValue(strTxnJson, "modelName")

    objFields("RegistrationType") = _
        JSON_ExtractValue(strTxnJson, "registrationType")

    objFields("FuelGroup") = _
        JSON_ExtractValue(strTxnJson, "fuelGroup")

    objFields("IsLeased") = _
        JSON_ExtractValue(strTxnJson, "isLeased")

    objFields("IsUsedImported") = _
        JSON_ExtractValue(strTxnJson, "isUsedImported")

    objFields("FirstRegistrationDate") = _
        OFV_DateFromISO(OFV_VariantToString( _
            JSON_ExtractValue( _
                strTxnJson, "firstRegistrationDate")))

    objFields("TransactionNumber") = _
        JSON_ExtractValue(strTxnJson, "transactionNumber")

    objFields("TransactionDate") = _
        OFV_DateFromISO(OFV_VariantToString( _
            JSON_ExtractValue( _
                strTxnJson, "transactionDate")))

    strFromObj = JSON_ExtractObject(strTxnJson, "from")
    strToObj = JSON_ExtractObject(strTxnJson, "to")

    strFromOwner = JSON_ExtractObject(strFromObj, "owner")
    strToOwner = JSON_ExtractObject(strToObj, "owner")

    objFields("FromOwnerType") = _
        JSON_ExtractValue(strFromOwner, "type")

    strCompanyObj = _
        JSON_ExtractObject(strFromOwner, "companyInfo")

    objFields("FromOwnerCompanyName") = _
        JSON_ExtractValue(strCompanyObj, "name")

    objFields("FromOwnerCounty") = _
        JSON_ExtractValue(strFromOwner, "countyName")

    objFields("ToOwnerType") = _
        JSON_ExtractValue(strToOwner, "type")

    strCompanyObj = _
        JSON_ExtractObject(strToOwner, "companyInfo")

    objFields("ToOwnerCompanyName") = _
        JSON_ExtractValue(strCompanyObj, "name")

    objFields("ToOwnerCounty") = _
        JSON_ExtractValue(strToOwner, "countyName")

    objFields("Status") = "OK"

    Set OFV_BuildFieldsFromTransaction = objFields

End Function


Private Function OFV_VariantToString( _
    ByVal varValue As Variant) As String

    If IsNull(varValue) Or IsEmpty(varValue) Then
        OFV_VariantToString = vbNullString
    Else
        OFV_VariantToString = CStr(varValue)
    End If

End Function


'==============================================================
' HTTP
'==============================================================

Private Function OFV_PostWithRetries( _
    ByVal strApiKey As String, _
    ByVal strBody As String, _
    ByRef strStatus As String) As String

    Dim lngAttempt As Long
    Dim lngStatusCode As Long
    Dim strLastError As String
    Dim strResponseText As String
    Dim objHTTP As Object

    strStatus = "OK"

    For lngAttempt = 1 To OFV_MAX_RETRIES

        Set objHTTP = _
            CreateObject("MSXML2.ServerXMLHTTP.6.0")

        lngStatusCode = 0
        strResponseText = vbNullString

        On Error Resume Next

        objHTTP.Open "POST", OFV_BASE_URL, False

        objHTTP.setRequestHeader _
            "Ocp-Apim-Subscription-Key", strApiKey

        objHTTP.setRequestHeader _
            "Content-Type", "application/json"

        objHTTP.setTimeouts 10000, 10000, 30000, 30000
        objHTTP.Send strBody

        lngStatusCode = objHTTP.Status
        strResponseText = objHTTP.ResponseText

        If Err.Number <> 0 Then

            strLastError = "VBA-feil: " & Err.Description
            lngStatusCode = 0
            Err.Clear

        End If

        On Error GoTo 0

        Select Case lngStatusCode

            Case 200
                OFV_PostWithRetries = strResponseText
                Exit Function

            Case 401
                strStatus = _
                    "Feil: 401 Unauthorized - kontroller API-nokkelen"
                Exit Function

            Case 403
                strStatus = _
                    "Feil: 403 Forbidden - tilgang eller kvote"
                Exit Function

            Case 429, 500, 502, 503, 504

                strLastError = _
                    lngStatusCode & ": " & strResponseText

                Sleep OFV_RETRY_WAIT_MS * lngAttempt

            Case Else

                If lngStatusCode <> 0 Then

                    strStatus = _
                        "Feil: " & lngStatusCode & _
                        " " & strResponseText

                    Exit Function

                Else

                    Sleep OFV_RETRY_WAIT_MS * lngAttempt

                End If

        End Select

    Next lngAttempt

    strStatus = _
        "Feil: Ga opp etter " & OFV_MAX_RETRIES & _
        " forsok. Siste feil: " & strLastError

End Function


'==============================================================
' JSON-HJELPEFUNKSJONER
'==============================================================

Private Function OFV_JsonEscape( _
    ByVal strValue As String) As String

    Dim strResult As String

    strResult = Replace(strValue, "\", "\\")
    strResult = Replace( _
        strResult, Chr(34), "\" & Chr(34))

    OFV_JsonEscape = strResult

End Function


Private Function OFV_DateFromISO( _
    ByVal strISO As String) As Variant

    If Len(strISO) < 10 Then
        OFV_DateFromISO = Empty
        Exit Function
    End If

    On Error GoTo InvalidDate

    OFV_DateFromISO = DateSerial( _
        CInt(Mid$(strISO, 1, 4)), _
        CInt(Mid$(strISO, 6, 2)), _
        CInt(Mid$(strISO, 9, 2)))

    Exit Function

InvalidDate:
    OFV_DateFromISO = Empty

End Function


Private Function JSON_FindMatchingBrace( _
    ByVal strJSON As String, _
    ByVal lngOpenPos As Long) As Long

    Dim strOpen As String
    Dim strClose As String
    Dim strChar As String
    Dim lngDepth As Long
    Dim i As Long
    Dim blnInString As Boolean

    strOpen = Mid$(strJSON, lngOpenPos, 1)

    If strOpen = "{" Then
        strClose = "}"
    ElseIf strOpen = "[" Then
        strClose = "]"
    Else
        Exit Function
    End If

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

End Function


Private Function JSON_SkipWhitespace( _
    ByVal strJSON As String, _
    ByVal lngPos As Long) As Long

    Dim p As Long
    Dim strChar As String

    p = lngPos

    Do While p <= Len(strJSON)

        strChar = Mid$(strJSON, p, 1)

        If strChar = " " Or _
           strChar = vbLf Or _
           strChar = vbCr Or _
           strChar = vbTab Then

            p = p + 1

        Else
            Exit Do
        End If

    Loop

    JSON_SkipWhitespace = p

End Function


Private Function JSON_ExtractObject( _
    ByVal strJSON As String, _
    ByVal strKey As String) As String

    Dim lngKeyPos As Long
    Dim lngPos As Long
    Dim lngEnd As Long
    Dim strFirstChar As String

    If Len(strJSON) = 0 Then Exit Function

    lngKeyPos = InStr( _
        1, strJSON, _
        Chr(34) & strKey & Chr(34) & ":", _
        vbBinaryCompare)

    If lngKeyPos = 0 Then Exit Function

    lngPos = lngKeyPos + Len(strKey) + 3
    lngPos = JSON_SkipWhitespace(strJSON, lngPos)

    strFirstChar = Mid$(strJSON, lngPos, 1)

    If strFirstChar <> "{" And _
       strFirstChar <> "[" Then Exit Function

    lngEnd = JSON_FindMatchingBrace(strJSON, lngPos)

    If lngEnd = 0 Then Exit Function

    JSON_ExtractObject = Mid$( _
        strJSON, lngPos, lngEnd - lngPos + 1)

End Function


Private Function JSON_ExtractValue( _
    ByVal strJSON As String, _
    ByVal strKey As String) As Variant

    Dim lngKeyPos As Long
    Dim lngPos As Long
    Dim lngEnd As Long
    Dim i As Long
    Dim j As Long

    Dim strFirstChar As String
    Dim strResult As String
    Dim strRaw As String
    Dim c As String
    Dim nc As String
    Dim ch As String

    JSON_ExtractValue = Null

    If Len(strJSON) = 0 Then Exit Function

    lngKeyPos = InStr( _
        1, strJSON, _
        Chr(34) & strKey & Chr(34) & ":", _
        vbBinaryCompare)

    If lngKeyPos = 0 Then Exit Function

    lngPos = lngKeyPos + Len(strKey) + 3
    lngPos = JSON_SkipWhitespace(strJSON, lngPos)

    strFirstChar = Mid$(strJSON, lngPos, 1)

    If strFirstChar = Chr(34) Then

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

    ElseIf strFirstChar = "{" Or _
           strFirstChar = "[" Then

        lngEnd = JSON_FindMatchingBrace(strJSON, lngPos)

        If lngEnd > 0 Then
            JSON_ExtractValue = Mid$( _
                strJSON, lngPos, lngEnd - lngPos + 1)
        End If

    Else

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


Private Function JSON_ArrayAllElements( _
    ByVal strJSONArray As String) As Collection

    Dim colResult As New Collection
    Dim lngPos As Long
    Dim lngEnd As Long
    Dim i As Long

    Dim strFirstChar As String
    Dim strElement As String
    Dim strResult As String
    Dim c As String

    If Len(strJSONArray) < 2 Then
        Set JSON_ArrayAllElements = colResult
        Exit Function
    End If

    If Left$(strJSONArray, 1) <> "[" Then
        Set JSON_ArrayAllElements = colResult
        Exit Function
    End If

    lngPos = JSON_SkipWhitespace(strJSONArray, 2)

    Do While lngPos <= Len(strJSONArray)

        If Mid$(strJSONArray, lngPos, 1) = "]" Then Exit Do

        strFirstChar = Mid$(strJSONArray, lngPos, 1)
        strElement = vbNullString

        If strFirstChar = "{" Or strFirstChar = "[" Then

            lngEnd = JSON_FindMatchingBrace( _
                strJSONArray, lngPos)

            If lngEnd = 0 Then Exit Do

            strElement = Mid$( _
                strJSONArray, _
                lngPos, _
                lngEnd - lngPos + 1)

            lngPos = lngEnd + 1

        Else

            i = lngPos
            strResult = vbNullString

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
                lngPos = JSON_SkipWhitespace( _
                    strJSONArray, lngPos + 1)
            End If

        End If

    Loop

    Set JSON_ArrayAllElements = colResult

End Function
