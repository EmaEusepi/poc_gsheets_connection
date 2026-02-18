Attribute VB_Name = "CloudCalc"
' ============================================
' Modulo VBA per Excel - Cloud Calc
' ============================================
'
' EQUIVALENTE di google_apps_script.js per Excel.
' Excel rispetta le dipendenze anche per le UDF:
' se C1 = CLOUD_CALC("multiply", B1, 2) e B1 = CLOUD_CALC("plus", A1, 10),
' Excel calcola PRIMA B1 e POI C1.
'
' INSTALLAZIONE:
' 1. Apri Excel > Alt+F11 (Editor VBA)
' 2. Menu File > Importa file... > seleziona questo file .bas
'    (oppure: Inserisci > Modulo, e incolla il codice)
' 3. Salva il file Excel come .xlsm (con macro abilitate)
' 4. Modifica API_BASE_URL qui sotto con il tuo endpoint
'
' UTILIZZO NEL FOGLIO:
' =CLOUD_CALC("plus", A1, A2)
' =CLOUD_CALC("multiply", B1, B2, B3)
' =CLOUD_CALC("average", C1:C10)
' =CLOUD_CALC("if", A1>10, "Alto", "Basso")
' =CLOUD_SUMIFS(P1:P100, N1:N100, "H Rilevate", H1:H100, "metano")
'
' NOTA: Excel usa la virgola come separatore argomenti (non il punto e virgola).
'       Se il tuo Excel usa il locale italiano, potrebbe usare ";" - dipende dalle impostazioni.

Option Explicit

' CONFIGURA IL TUO ENDPOINT QUI
Private Const API_BASE_URL As String = "http://18.153.39.218:5000/calc"

' ============================================
' FUNZIONI PUBBLICHE (usabili nel foglio)
' ============================================

'''
' Esegue calcoli su cloud tramite API.
' Equivalente a CLOUD_CALC di Google Sheets.
'
' @param operation  Nome dell'operazione (es: "plus", "if", "max")
' @param args       Argomenti variabili: valori singoli o range di celle
' @return           Risultato del calcolo
'
Public Function CLOUD_CALC(operation As Variant, ParamArray args() As Variant) As Variant
    On Error GoTo ErrorHandler

    ' --- VALIDAZIONE ---
    If IsMissing(operation) Or IsEmpty(operation) Then
        CLOUD_CALC = "#ERROR: Operazione mancante. Specificare il nome dell'operazione come primo argomento."
        Exit Function
    End If

    If VarType(operation) <> vbString Then
        CLOUD_CALC = "#ERROR: L'operazione deve essere una stringa (es: ""plus"", ""multiply"")."
        Exit Function
    End If

    ' Raccogli e appiattisci gli argomenti
    Dim flatArgs() As Variant
    Dim argCount As Long
    argCount = 0
    ReDim flatArgs(0 To 99) ' buffer iniziale

    Dim i As Long
    For i = LBound(args) To UBound(args)
        If TypeName(args(i)) = "Range" Then
            ' Range di celle: itera su ogni cella
            Dim cell As Range
            For Each cell In args(i)
                If argCount > UBound(flatArgs) Then ReDim Preserve flatArgs(0 To argCount + 99)
                flatArgs(argCount) = CellValue_(cell.value)
                argCount = argCount + 1
            Next cell
        Else
            If argCount > UBound(flatArgs) Then ReDim Preserve flatArgs(0 To argCount + 99)
            flatArgs(argCount) = CellValue_(args(i))
            argCount = argCount + 1
        End If
    Next i

    ' Ridimensiona all'effettivo
    If argCount > 0 Then
        ReDim Preserve flatArgs(0 To argCount - 1)
    Else
        ReDim flatArgs(0 To 0)
        flatArgs(0) = Empty
    End If

    ' Validazione: controlla errori Excel negli argomenti
    Dim errCheck As String
    errCheck = CheckForErrors_(flatArgs, argCount)
    If errCheck <> "" Then
        CLOUD_CALC = errCheck
        Exit Function
    End If

    ' --- COSTRUISCI PAYLOAD JSON ---
    Dim json As String
    json = "{""operation"":" & ToJson_(operation)

    If argCount > 0 Then
        json = json & ",""args"":[" & ArgsToJsonList_(flatArgs, argCount) & "]"
    Else
        json = json & ",""args"":[]"
    End If
    json = json & "}"

    ' --- CHIAMATA API ---
    CLOUD_CALC = CallApi_(json)
    Exit Function

ErrorHandler:
    CLOUD_CALC = "#ERROR: " & Err.Description
End Function

'''
' SUMIFS su cloud: somma condizionale con criteri multipli.
' Equivalente a CLOUD_SUMIFS di Google Sheets.
'
' @param sum_range       Range dei valori da sommare
' @param criteria_args   Coppie di (criteria_range, criteria): range e criterio alternati
' @return                Somma dei valori che soddisfano tutti i criteri
'
Public Function CLOUD_SUMIFS(sum_range As Range, ParamArray criteria_args() As Variant) As Variant
    On Error GoTo ErrorHandler

    ' --- VALIDAZIONE ---
    If sum_range Is Nothing Then
        CLOUD_SUMIFS = "#ERROR: sum_range mancante. Specificare il range dei valori da sommare."
        Exit Function
    End If

    Dim numCriteriaArgs As Long
    numCriteriaArgs = UBound(criteria_args) - LBound(criteria_args) + 1

    ' Servono almeno una coppia criteria_range + criteria
    If numCriteriaArgs < 2 Then
        CLOUD_SUMIFS = "#ERROR: Servono almeno criteria_range e criteria dopo sum_range."
        Exit Function
    End If

    ' Argomenti devono essere a coppie
    If numCriteriaArgs Mod 2 <> 0 Then
        CLOUD_SUMIFS = "#ERROR: Gli argomenti dopo sum_range devono essere a coppie (criteria_range, criteria)."
        Exit Function
    End If

    ' Appiattisci sum_range
    Dim flatSum() As Variant
    Dim sumCount As Long
    FlattenRange_ sum_range, flatSum, sumCount

    ' Controlla errori nel sum_range
    Dim sumErrCheck As String
    sumErrCheck = CheckForErrors_(flatSum, sumCount)
    If sumErrCheck <> "" Then
        CLOUD_SUMIFS = "#ERROR: sum_range contiene un errore (" & sumErrCheck & ")."
        Exit Function
    End If

    ' --- COSTRUISCI PAYLOAD JSON ---
    Dim json As String
    json = "{""operation"":""sumifs"",""sum_range"":[" & ArgsToJsonList_(flatSum, sumCount) & "]"
    json = json & ",""criteria_pairs"":["

    Dim pairIndex As Long
    pairIndex = 0

    Dim j As Long
    For j = LBound(criteria_args) To UBound(criteria_args) Step 2
        ' criteria_range
        If TypeName(criteria_args(j)) <> "Range" Then
            CLOUD_SUMIFS = "#ERROR: criteria_range " & (pairIndex + 1) & " deve essere un range di celle."
            Exit Function
        End If

        Dim flatCrit() As Variant
        Dim critCount As Long
        FlattenRange_ criteria_args(j), flatCrit, critCount

        ' Validazione: stessa lunghezza del sum_range
        If critCount <> sumCount Then
            CLOUD_SUMIFS = "#ERROR: criteria_range " & (pairIndex + 1) & " ha " & critCount & " elementi, ma sum_range ne ha " & sumCount & "."
            Exit Function
        End If

        ' Controlla errori nel criteria_range
        Dim critErrCheck As String
        critErrCheck = CheckForErrors_(flatCrit, critCount)
        If critErrCheck <> "" Then
            CLOUD_SUMIFS = "#ERROR: criteria_range " & (pairIndex + 1) & " contiene un errore."
            Exit Function
        End If

        ' criteria (valore singolo)
        Dim criteria As Variant
        criteria = criteria_args(j + 1)

        If pairIndex > 0 Then json = json & ","
        json = json & "{""range"":[" & ArgsToJsonList_(flatCrit, critCount) & "]"
        json = json & ",""criteria"":" & ToJson_(criteria) & "}"

        pairIndex = pairIndex + 1
    Next j

    json = json & "]}"

    ' --- CHIAMATA API ---
    CLOUD_SUMIFS = CallApi_(json)
    Exit Function

ErrorHandler:
    CLOUD_SUMIFS = "#ERROR: " & Err.Description
End Function

'''
' Elenca tutte le operazioni disponibili.
'
Public Function CLOUD_CALC_OPERATIONS() As String
    On Error GoTo ErrorHandler

    Dim url As String
    url = Replace(API_BASE_URL, "/calc", "/operations")

    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "GET", url, False
    xhr.send

    Dim responseText As String
    responseText = xhr.responseText

    ' Estrai la lista operations dal JSON
    Dim opsStart As Long
    opsStart = InStr(responseText, "[")
    Dim opsEnd As Long
    opsEnd = InStr(responseText, "]")

    If opsStart > 0 And opsEnd > opsStart Then
        Dim opsList As String
        opsList = Mid(responseText, opsStart + 1, opsEnd - opsStart - 1)
        opsList = Replace(opsList, """", "")
        CLOUD_CALC_OPERATIONS = opsList
    Else
        CLOUD_CALC_OPERATIONS = "#ERROR: Risposta non valida dal server."
    End If
    Exit Function

ErrorHandler:
    CLOUD_CALC_OPERATIONS = "#ERROR: " & Err.Description
End Function

' ============================================
' FUNZIONI PRIVATE (helper)
' ============================================

'''
' Esegue la chiamata HTTP POST all'API e restituisce il risultato.
'
Private Function CallApi_(jsonPayload As String) As Variant
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")

    xhr.Open "POST", API_BASE_URL, False
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.send jsonPayload

    Dim responseText As String
    responseText = xhr.responseText

    Dim statusCode As Long
    statusCode = xhr.Status

    ' Controlla Content-Type
    Dim contentType As String
    contentType = xhr.getResponseHeader("Content-Type")
    If InStr(1, contentType, "application/json", vbTextCompare) = 0 Then
        CallApi_ = "#ERROR: Il server non ha risposto con JSON (HTTP " & statusCode & ")."
        Exit Function
    End If

    If statusCode = 200 Then
        ' Estrai il campo "result"
        Dim result As Variant
        result = JsonGetValue_(responseText, "result")

        If IsNull(result) Then
            CallApi_ = ""
        ElseIf IsEmpty(result) Then
            CallApi_ = ""
        Else
            CallApi_ = result
        End If
    Else
        ' Estrai il campo "error"
        Dim errorMsg As Variant
        errorMsg = JsonGetValue_(responseText, "error")
        If IsEmpty(errorMsg) Then errorMsg = "HTTP " & statusCode
        CallApi_ = "#ERROR: " & CStr(errorMsg)
    End If
End Function

'''
' Converte il valore di una cella Excel in un valore utilizzabile.
' Gestisce celle vuote, errori, ecc.
'
Private Function CellValue_(value As Variant) As Variant
    If IsError(value) Then
        ' Converti errori Excel in stringhe riconoscibili
        Select Case CLng(value)
            Case CVErr(xlErrDiv0): CellValue_ = "#DIV/0!"
            Case CVErr(xlErrNA): CellValue_ = "#N/A"
            Case CVErr(xlErrName): CellValue_ = "#NAME?"
            Case CVErr(xlErrNull): CellValue_ = "#NULL!"
            Case CVErr(xlErrNum): CellValue_ = "#NUM!"
            Case CVErr(xlErrRef): CellValue_ = "#REF!"
            Case CVErr(xlErrValue): CellValue_ = "#VALUE!"
            Case Else: CellValue_ = "#ERROR"
        End Select
    ElseIf IsEmpty(value) Or value = "" Then
        CellValue_ = Null
    Else
        CellValue_ = value
    End If
End Function

'''
' Controlla se un valore e' un errore di Excel (stringa).
'
Private Function IsSheetError_(value As Variant) As Boolean
    If IsError(value) Then
        IsSheetError_ = True
        Exit Function
    End If
    If VarType(value) <> vbString Then
        IsSheetError_ = False
        Exit Function
    End If
    Dim s As String
    s = CStr(value)
    IsSheetError_ = (Left(s, 5) = "#REF!" Or Left(s, 4) = "#N/A" Or _
                     Left(s, 7) = "#VALUE!" Or Left(s, 6) = "#NULL!" Or _
                     Left(s, 5) = "#NUM!" Or Left(s, 6) = "#NAME?" Or _
                     Left(s, 7) = "#DIV/0!" Or Left(s, 6) = "#ERROR")
End Function

'''
' Controlla se un array di argomenti contiene errori.
' Restituisce stringa vuota se tutto OK, altrimenti il messaggio di errore.
'
Private Function CheckForErrors_(args() As Variant, argCount As Long) As String
    Dim i As Long
    For i = 0 To argCount - 1
        If IsSheetError_(args(i)) Then
            CheckForErrors_ = "#ERROR: L'argomento " & (i + 1) & " contiene un errore (" & CStr(args(i)) & "). Verificare le celle di input."
            Exit Function
        End If
    Next i
    CheckForErrors_ = ""
End Function

'''
' Appiattisce un Range Excel in un array monodimensionale.
'
Private Sub FlattenRange_(rng As Range, ByRef result() As Variant, ByRef count As Long)
    count = rng.Cells.count
    ReDim result(0 To count - 1)

    Dim i As Long
    i = 0
    Dim cell As Range
    For Each cell In rng
        result(i) = CellValue_(cell.value)
        i = i + 1
    Next cell
End Sub

' ============================================
' JSON HELPERS (serializzazione/parsing)
' ============================================

'''
' Converte un valore VBA in formato JSON.
'
Private Function ToJson_(value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        ToJson_ = "null"
    ElseIf VarType(value) = vbBoolean Then
        ToJson_ = IIf(value, "true", "false")
    ElseIf VarType(value) = vbString Then
        ToJson_ = """" & JsonEscape_(CStr(value)) & """"
    ElseIf IsNumeric(value) Then
        ' Forza il punto come separatore decimale (non la virgola del locale IT)
        ToJson_ = Replace(CStr(CDbl(value)), ",", ".")
    Else
        ToJson_ = """" & JsonEscape_(CStr(value)) & """"
    End If
End Function

'''
' Escape di caratteri speciali per JSON.
'
Private Function JsonEscape_(s As String) As String
    Dim result As String
    result = Replace(s, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    JsonEscape_ = result
End Function

'''
' Converte un array di argomenti in una lista JSON (senza parentesi quadre).
' Es: "1,2,""hello"",null"
'
Private Function ArgsToJsonList_(args() As Variant, argCount As Long) As String
    If argCount = 0 Then
        ArgsToJsonList_ = ""
        Exit Function
    End If

    Dim parts() As String
    ReDim parts(0 To argCount - 1)

    Dim i As Long
    For i = 0 To argCount - 1
        parts(i) = ToJson_(args(i))
    Next i

    ArgsToJsonList_ = Join(parts, ",")
End Function

'''
' Estrae il valore di un campo da una stringa JSON semplice (oggetto piatto).
' Supporta: stringhe, numeri, booleani, null.
' Restituisce Empty se il campo non esiste.
'
Private Function JsonGetValue_(jsonStr As String, key As String) As Variant
    Dim searchKey As String
    searchKey = """" & key & """"

    Dim keyPos As Long
    keyPos = InStr(1, jsonStr, searchKey)
    If keyPos = 0 Then
        JsonGetValue_ = Empty
        Exit Function
    End If

    ' Trova i due punti dopo la chiave
    Dim colonPos As Long
    colonPos = InStr(keyPos + Len(searchKey), jsonStr, ":")
    If colonPos = 0 Then
        JsonGetValue_ = Empty
        Exit Function
    End If

    ' Salta spazi dopo i due punti
    Dim valueStart As Long
    valueStart = colonPos + 1
    Do While valueStart <= Len(jsonStr) And Mid(jsonStr, valueStart, 1) = " "
        valueStart = valueStart + 1
    Loop

    If valueStart > Len(jsonStr) Then
        JsonGetValue_ = Empty
        Exit Function
    End If

    Dim firstChar As String
    firstChar = Mid(jsonStr, valueStart, 1)

    ' null
    If Mid(jsonStr, valueStart, 4) = "null" Then
        JsonGetValue_ = Null
        Exit Function
    End If

    ' boolean
    If Mid(jsonStr, valueStart, 4) = "true" Then
        JsonGetValue_ = True
        Exit Function
    End If
    If Mid(jsonStr, valueStart, 5) = "false" Then
        JsonGetValue_ = False
        Exit Function
    End If

    ' stringa (inizia con ")
    If firstChar = """" Then
        Dim strEnd As Long
        strEnd = valueStart + 1
        Do While strEnd <= Len(jsonStr)
            If Mid(jsonStr, strEnd, 1) = "\" Then
                strEnd = strEnd + 2
            ElseIf Mid(jsonStr, strEnd, 1) = """" Then
                Exit Do
            Else
                strEnd = strEnd + 1
            End If
        Loop
        Dim strVal As String
        strVal = Mid(jsonStr, valueStart + 1, strEnd - valueStart - 1)
        strVal = Replace(strVal, "\""", """")
        strVal = Replace(strVal, "\\", "\")
        strVal = Replace(strVal, "\n", vbLf)
        strVal = Replace(strVal, "\t", vbTab)
        JsonGetValue_ = strVal
        Exit Function
    End If

    ' numero
    Dim numEnd As Long
    numEnd = valueStart
    Do While numEnd <= Len(jsonStr) And InStr("0123456789.-+eE", Mid(jsonStr, numEnd, 1)) > 0
        numEnd = numEnd + 1
    Loop
    Dim numStr As String
    numStr = Mid(jsonStr, valueStart, numEnd - valueStart)
    If IsNumeric(numStr) Then
        JsonGetValue_ = CDbl(numStr)
    Else
        JsonGetValue_ = numStr
    End If
End Function
