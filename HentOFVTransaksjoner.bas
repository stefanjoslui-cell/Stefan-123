Option Explicit

' Enkel testmakro: henter transaksjoner (eierskifter) fra OFV Transactions API
' for hvert registreringsnummer i kolonne B, og legger dem i en tabell.
'
' Oppsett i regnearket du kjører makroen fra:
'   A1        = API-nøkkel (Ocp-Apim-Subscription-Key)
'   B1 og ned = registreringsnumre
'
' Resultatet havner i arket "Transaksjoner" (opprettes automatisk).

Public Sub HentOFVTransaksjoner()

    Const OFV_URL As String = _
        "https://data.ofv.no/transactions/v1/query"

    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim http As Object

    Dim apiKey As String
    Dim regnr As String
    Dim requestBody As String
    Dim responseText As String
    Dim httpStatus As Long

    Dim sisteRad As Long
    Dim inputRad As Long
    Dim outputRad As Long

    Dim transaksjonerJson As String
    Dim objekter As Collection
    Dim tx As String
    Dim i As Long

    Dim fraObj As String
    Dim tilObj As String
    Dim selgerObj As String
    Dim kjoperObj As String

    On Error GoTo Feilhandtering

    Set wsInput = ActiveSheet

    apiKey = Trim$(CStr(wsInput.Range("A1").Value))

    If Len(apiKey) = 0 Then
        MsgBox "Legg API-nøkkelen i celle A1.", _
               vbExclamation, "Mangler API-nøkkel"
        Exit Sub
    End If

    sisteRad = wsInput.Cells( _
        wsInput.Rows.Count, "B").End(xlUp).Row

    If sisteRad < 1 Or _
       Len(Trim$(CStr(wsInput.Range("B1").Value))) = 0 Then

        MsgBox "Legg registreringsnummer i B1 og nedover.", _
               vbExclamation, "Mangler registreringsnummer"
        Exit Sub

    End If

    Set wsOutput = FinnEllerOpprettArk( _
        ThisWorkbook, "Transaksjoner")

    wsOutput.Cells.Clear

    wsOutput.Range("A1").Value = "Regnr (søkt)"
    wsOutput.Range("B1").Value = "HTTP-status"
    wsOutput.Range("C1").Value = "Merknad"
    wsOutput.Range("D1").Value = "TransaksjonsId"
    wsOutput.Range("E1").Value = "TransaksjonsDato"
    wsOutput.Range("F1").Value = "RegistreringsType"
    wsOutput.Range("G1").Value = "Merke"
    wsOutput.Range("H1").Value = "Modell"
    wsOutput.Range("I1").Value = "Drivstoff"
    wsOutput.Range("J1").Value = "DrivstoffGruppe"
    wsOutput.Range("K1").Value = "Karosseri"
    wsOutput.Range("L1").Value = "KjøretøyGruppe"
    wsOutput.Range("M1").Value = "Girkasse"
    wsOutput.Range("N1").Value = "Leaset"
    wsOutput.Range("O1").Value = "TransaksjonsNummer"
    wsOutput.Range("P1").Value = "FørstegangsregDato"
    wsOutput.Range("Q1").Value = "SelgerFylke"
    wsOutput.Range("R1").Value = "SelgerKommune"
    wsOutput.Range("S1").Value = "KjøperFylke"
    wsOutput.Range("T1").Value = "KjøperKommune"
    wsOutput.Range("U1").Value = "Rått JSON-svar"

    wsOutput.Range("A1:U1").Font.Bold = True

    outputRad = 2

    Application.ScreenUpdating = False
    Application.StatusBar = "Starter OFV-oppslag ..."

    For inputRad = 1 To sisteRad

        regnr = Trim$(CStr( _
            wsInput.Cells(inputRad, "B").Value))

        regnr = Replace(regnr, " ", "")
        regnr = Replace(regnr, "-", "")
        regnr = UCase$(regnr)

        If Len(regnr) > 0 Then

            Application.StatusBar = _
                "Henter transaksjoner for " & regnr & " ..."

            Set http = CreateObject( _
                "WinHttp.WinHttpRequest.5.1")

            http.SetTimeouts 10000, 10000, 30000, 30000

            http.Open "POST", OFV_URL, False

            http.SetRequestHeader _
                "Content-Type", "application/json"

            http.SetRequestHeader _
                "Accept", "application/json"

            http.SetRequestHeader _
                "Ocp-Apim-Subscription-Key", apiKey

            ' NB: filteret heter "regNo" (streng) i det dokumenterte API-skjemaet,
            ' ikke "registrationNumber". pagination.first er satt høyt slik at vi
            ' får alle transaksjonene for ett regnr på én side (uten cursor-paginering).
            requestBody = _
                "{" & _
                    """filters"":{" & _
                        """regNo"":""" & JSON_Escape(regnr) & """" & _
                    "}," & _
                    """pagination"":{""first"":1000}" & _
                "}"

            On Error Resume Next

            http.Send requestBody

            If Err.Number <> 0 Then

                wsOutput.Cells(outputRad, "A").Value = regnr
                wsOutput.Cells(outputRad, "B").Value = _
                    "VBA-feil " & Err.Number
                wsOutput.Cells(outputRad, "C").Value = _
                    Err.Description

                Err.Clear

                outputRad = outputRad + 1

            Else

                httpStatus = CLng(http.Status)
                responseText = CStr(http.ResponseText)

                If httpStatus = 200 Then

                    transaksjonerJson = _
                        JSON_ExtractArrayContent( _
                            responseText, "transactions")

                    Set objekter = _
                        JSON_SplitTopLevelObjects(transaksjonerJson)

                    If objekter.Count = 0 Then

                        wsOutput.Cells(outputRad, "A").Value = regnr
                        wsOutput.Cells(outputRad, "B").Value = httpStatus
                        wsOutput.Cells(outputRad, "C").Value = _
                            "Ingen transaksjoner funnet"
                        wsOutput.Cells(outputRad, "U").Value = _
                            Left$(responseText, 32767)

                        outputRad = outputRad + 1

                    Else

                        For i = 1 To objekter.Count

                            tx = objekter(i)

                            fraObj = JSON_ExtractObject(tx, "from")
                            selgerObj = JSON_ExtractObject(fraObj, "owner")

                            tilObj = JSON_ExtractObject(tx, "to")
                            kjoperObj = JSON_ExtractObject(tilObj, "owner")

                            wsOutput.Cells(outputRad, "A").Value = regnr
                            wsOutput.Cells(outputRad, "B").Value = httpStatus
                            wsOutput.Cells(outputRad, "C").Value = "OK"

                            wsOutput.Cells(outputRad, "D").Value = _
                                JSON_ExtractString(tx, "id")
                            wsOutput.Cells(outputRad, "E").Value = _
                                JSON_ExtractString(tx, "transactionDate")
                            wsOutput.Cells(outputRad, "F").Value = _
                                JSON_ExtractString(tx, "registrationType")
                            wsOutput.Cells(outputRad, "G").Value = _
                                JSON_ExtractString(tx, "makeName")
                            wsOutput.Cells(outputRad, "H").Value = _
                                JSON_ExtractString(tx, "modelName")
                            wsOutput.Cells(outputRad, "I").Value = _
                                JSON_ExtractString(tx, "fuelName")
                            wsOutput.Cells(outputRad, "J").Value = _
                                JSON_ExtractString(tx, "fuelGroup")
                            wsOutput.Cells(outputRad, "K").Value = _
                                JSON_ExtractString(tx, "chassisName")
                            wsOutput.Cells(outputRad, "L").Value = _
                                JSON_ExtractString(tx, "vehicleGroupName")
                            wsOutput.Cells(outputRad, "M").Value = _
                                JSON_ExtractString(tx, "transmission")
                            wsOutput.Cells(outputRad, "N").Value = _
                                JSON_ExtractRaw(tx, "isLeased")
                            wsOutput.Cells(outputRad, "O").Value = _
                                JSON_ExtractRaw(tx, "transactionNumber")
                            wsOutput.Cells(outputRad, "P").Value = _
                                JSON_ExtractString(tx, "firstRegistrationDate")
                            wsOutput.Cells(outputRad, "Q").Value = _
                                JSON_ExtractString(selgerObj, "countyName")
                            wsOutput.Cells(outputRad, "R").Value = _
                                JSON_ExtractString(selgerObj, "municipalityName")
                            wsOutput.Cells(outputRad, "S").Value = _
                                JSON_ExtractString(kjoperObj, "countyName")
                            wsOutput.Cells(outputRad, "T").Value = _
                                JSON_ExtractString(kjoperObj, "municipalityName")

                            If i = 1 Then
                                wsOutput.Cells(outputRad, "U").Value = _
                                    Left$(responseText, 32767)
                            End If

                            outputRad = outputRad + 1

                        Next i

                    End If

                Else

                    wsOutput.Cells(outputRad, "A").Value = regnr
                    wsOutput.Cells(outputRad, "B").Value = httpStatus

                    Select Case httpStatus

                        Case 400
                            wsOutput.Cells(outputRad, "C").Value = _
                                "Feil i URL, filter eller JSON-body"

                        Case 401
                            wsOutput.Cells(outputRad, "C").Value = _
                                "Ikke godkjent. Kontroller API-nøkkelen"

                        Case 403
                            wsOutput.Cells(outputRad, "C").Value = _
                                "Ingen tilgang / kvote overskredet"

                        Case 404
                            wsOutput.Cells(outputRad, "C").Value = _
                                "Endepunktet ble ikke funnet"

                        Case 429
                            wsOutput.Cells(outputRad, "C").Value = _
                                "For mange kall. Rate limit nådd"

                        Case 500 To 599
                            wsOutput.Cells(outputRad, "C").Value = _
                                "Feil hos API-tjenesten"

                        Case Else
                            wsOutput.Cells(outputRad, "C").Value = _
                                http.StatusText

                    End Select

                    wsOutput.Cells(outputRad, "U").Value = _
                        Left$(responseText, 32767)

                    outputRad = outputRad + 1

                End If

            End If

            On Error GoTo Feilhandtering

            Set http = Nothing

            DoEvents

        End If

    Next inputRad

    wsOutput.Columns("A:T").AutoFit
    wsOutput.Columns("U").ColumnWidth = 80
    wsOutput.Columns("U").WrapText = True
    wsOutput.Rows(1).AutoFilter

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox _
        "OFV-uttrekket er ferdig." & vbCrLf & vbCrLf & _
        "Resultatet ligger i arket Transaksjoner.", _
        vbInformation, "Ferdig"

    Exit Sub

Feilhandtering:

    Set http = Nothing

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox _
        "Makroen stoppet." & vbCrLf & vbCrLf & _
        "Feilnummer: " & Err.Number & vbCrLf & _
        "Beskrivelse: " & Err.Description, _
        vbCritical, "Feil i OFV-uttrekk"

End Sub


Private Function FinnEllerOpprettArk( _
    ByVal wb As Workbook, _
    ByVal arknavn As String) As Worksheet

    On Error Resume Next

    Set FinnEllerOpprettArk = wb.Worksheets(arknavn)

    On Error GoTo 0

    If FinnEllerOpprettArk Is Nothing Then

        Set FinnEllerOpprettArk = _
            wb.Worksheets.Add( _
                After:=wb.Worksheets(wb.Worksheets.Count))

        FinnEllerOpprettArk.Name = arknavn

    End If

End Function


Private Function JSON_Escape( _
    ByVal tekst As String) As String

    tekst = Replace(tekst, "\", "\\")
    tekst = Replace(tekst, """", "\""")
    tekst = Replace(tekst, vbCr, "\r")
    tekst = Replace(tekst, vbLf, "\n")
    tekst = Replace(tekst, vbTab, "\t")

    JSON_Escape = tekst

End Function


Private Function JSON_Unescape( _
    ByVal tekst As String) As String

    Const midlertidig As String = Chr$(1)

    tekst = Replace(tekst, "\\", midlertidig)
    tekst = Replace(tekst, "\""", """")
    tekst = Replace(tekst, "\n", vbLf)
    tekst = Replace(tekst, "\r", vbCr)
    tekst = Replace(tekst, "\t", vbTab)
    tekst = Replace(tekst, "\/", "/")
    tekst = Replace(tekst, midlertidig, "\")

    JSON_Unescape = tekst

End Function


' Henter verdien for "key":"..." (strengfelt) et sted i json-teksten.
' Returnerer "" hvis nøkkelen mangler eller verdien ikke er en streng (f.eks. null).
Private Function JSON_ExtractString( _
    ByVal json As String, _
    ByVal key As String) As String

    Dim searchKey As String
    Dim posKey As Long
    Dim posValueStart As Long
    Dim posQuoteEnd As Long
    Dim raw As String

    searchKey = """" & key & """:"
    posKey = InStr(1, json, searchKey, vbTextCompare)

    If posKey = 0 Then
        JSON_ExtractString = ""
        Exit Function
    End If

    posValueStart = posKey + Len(searchKey)

    If Mid$(json, posValueStart, 1) <> """" Then
        JSON_ExtractString = ""
        Exit Function
    End If

    posQuoteEnd = posValueStart + 1

    Do While posQuoteEnd <= Len(json)

        If Mid$(json, posQuoteEnd, 1) = "\" Then
            posQuoteEnd = posQuoteEnd + 2
        ElseIf Mid$(json, posQuoteEnd, 1) = """" Then
            Exit Do
        Else
            posQuoteEnd = posQuoteEnd + 1
        End If

    Loop

    raw = Mid$(json, posValueStart + 1, posQuoteEnd - posValueStart - 1)
    JSON_ExtractString = JSON_Unescape(raw)

End Function


' Henter verdien for "key":<tall/true/false/null> (ikke-strengfelt).
Private Function JSON_ExtractRaw( _
    ByVal json As String, _
    ByVal key As String) As String

    Dim searchKey As String
    Dim posKey As Long
    Dim posValueStart As Long
    Dim posEnd As Long
    Dim ch As String

    searchKey = """" & key & """:"
    posKey = InStr(1, json, searchKey, vbTextCompare)

    If posKey = 0 Then
        JSON_ExtractRaw = ""
        Exit Function
    End If

    posValueStart = posKey + Len(searchKey)

    If Mid$(json, posValueStart, 1) = """" Then
        JSON_ExtractRaw = ""
        Exit Function
    End If

    posEnd = posValueStart

    Do While posEnd <= Len(json)

        ch = Mid$(json, posEnd, 1)

        If ch = "," Or ch = "}" Or ch = "]" Then
            Exit Do
        End If

        posEnd = posEnd + 1

    Loop

    JSON_ExtractRaw = Trim$(Mid$(json, posValueStart, posEnd - posValueStart))

End Function


' Henter det nøstede objektet for "key":{...}. Returnerer "" hvis det
' ikke finnes, eller hvis verdien er null.
Private Function JSON_ExtractObject( _
    ByVal json As String, _
    ByVal key As String) As String

    Dim searchKey As String
    Dim posKey As Long
    Dim posStart As Long
    Dim depth As Long
    Dim i As Long

    searchKey = """" & key & """:"
    posKey = InStr(1, json, searchKey, vbTextCompare)

    If posKey = 0 Then
        JSON_ExtractObject = ""
        Exit Function
    End If

    posStart = posKey + Len(searchKey)

    If Mid$(json, posStart, 1) <> "{" Then
        JSON_ExtractObject = ""
        Exit Function
    End If

    depth = 0

    For i = posStart To Len(json)

        Select Case Mid$(json, i, 1)

            Case "{"
                depth = depth + 1

            Case "}"
                depth = depth - 1

                If depth = 0 Then
                    JSON_ExtractObject = _
                        Mid$(json, posStart, i - posStart + 1)
                    Exit Function
                End If

        End Select

    Next i

    JSON_ExtractObject = ""

End Function


' Henter innholdet i arrayen "key":[ ... ] (uten de omsluttende hakeparentesene).
Private Function JSON_ExtractArrayContent( _
    ByVal json As String, _
    ByVal key As String) As String

    Dim searchKey As String
    Dim posKey As Long
    Dim posStart As Long
    Dim depth As Long
    Dim i As Long

    searchKey = """" & key & """:"
    posKey = InStr(1, json, searchKey, vbTextCompare)

    If posKey = 0 Then
        JSON_ExtractArrayContent = ""
        Exit Function
    End If

    posStart = InStr(posKey, json, "[")

    If posStart = 0 Then
        JSON_ExtractArrayContent = ""
        Exit Function
    End If

    depth = 0

    For i = posStart To Len(json)

        Select Case Mid$(json, i, 1)

            Case "["
                depth = depth + 1

            Case "]"
                depth = depth - 1

                If depth = 0 Then
                    JSON_ExtractArrayContent = _
                        Mid$(json, posStart + 1, i - posStart - 1)
                    Exit Function
                End If

        End Select

    Next i

    JSON_ExtractArrayContent = ""

End Function


' Deler innholdet i en array av objekter ("{...},{...},...") opp i
' hvert enkelt objekt, som en samling av tekststrenger.
Private Function JSON_SplitTopLevelObjects( _
    ByVal content As String) As Collection

    Dim result As New Collection
    Dim depth As Long
    Dim startPos As Long
    Dim i As Long
    Dim ch As String

    depth = 0
    startPos = 0

    For i = 1 To Len(content)

        ch = Mid$(content, i, 1)

        If ch = "{" Or ch = "[" Then

            If depth = 0 And ch = "{" Then
                startPos = i
            End If

            depth = depth + 1

        ElseIf ch = "}" Or ch = "]" Then

            depth = depth - 1

            If depth = 0 And startPos > 0 Then
                result.Add Mid$(content, startPos, i - startPos + 1)
                startPos = 0
            End If

        End If

    Next i

    Set JSON_SplitTopLevelObjects = result

End Function
