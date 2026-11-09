Attribute VB_Name = "Extramoenia"
'==============================================================================
' Extramoenia.bas
'
' Calcolo del fatturato mensile da comunicare ai medici in extramoenia,
' distinto tra attivita' di RICOVERO (interventi) e AMBULATORIALE, con
' predisposizione di una bozza email per ciascun medico.
'
' Questo modulo NON contiene dati: solo logica. I dati dell'ente nel testo
' dell'email (ENTE_*) sono segnaposto da sostituire.
'
' Flusso di lavoro (5 macro, da eseguire in quest'ordine)
'   1. Aggiorna_Interventi     aggiorna la query Power Query degli interventi
'   2. Aggiorna_Ambulatoriale  aggiorna la query Power Query dell'ambulatoriale
'   3. PvT_Ambulatoriale       pivot: importo ambulatoriale per matricola
'   4. PvT_interventi          pivot: interventi per medico + totale da fatturare
'   5. Invia_email             una bozza email per medico, con la tabella
'                              della pivot incollata come immagine
'
' Fogli attesi (nomi modificabili nelle costanti qui sotto)
'   Anagrafica Medici   A matricola | B cognome | ... | F email (intestazione riga 1)
'   Import interventi   A MATRICOLA | B Nominativo | C Data Intervento |
'                       D Nosologico | E IMPORTO        (caricato da Power Query)
'   Import ambulatoriale  A Column4 (matricola) | B Column8 (importo)
'                       (caricato da Power Query, nomi di colonna di default)
'   PvT interventi / PvT ambulatoriale   creati dalle macro, non toccare
'==============================================================================
Option Explicit

' --- Nomi dei fogli ----------------------------------------------------------
Private Const FOGLIO_ANAGRAFICA As String = "Anagrafica Medici"
Private Const FOGLIO_INTERVENTI As String = "Import interventi"
Private Const FOGLIO_AMBULATORIALE As String = "Import ambulatoriale"
Private Const FOGLIO_PVT_INTERVENTI As String = "PvT interventi"
Private Const FOGLIO_PVT_AMBULATORIALE As String = "PvT ambulatoriale"

' --- Nomi delle tabelle pivot e delle query Power Query -----------------------
Private Const PIVOT_INTERVENTI As String = "PivotTable1"
Private Const PIVOT_AMBULATORIALE As String = "PivotTable_Amb"
Private Const QUERY_INTERVENTI As String = "Interventi"
Private Const QUERY_AMBULATORIALE As String = "Ambulatoriale"

' --- Campi usati nelle pivot ---------------------------------------------------
Private Const CAMPO_MATRICOLA As String = "MATRICOLA"
Private Const CAMPO_NOMINATIVO As String = "Nominativo"
Private Const CAMPO_NOSOLOGICO As String = "Nosologico"
Private Const CAMPO_IMPORTO As String = "IMPORTO"
Private Const CAMPO_AMB_MATRICOLA As String = "Column4"
Private Const CAMPO_AMB_IMPORTO As String = "Column8"

' --- Colonne dell'anagrafica medici -------------------------------------------
Private Const COL_MATRICOLA As String = "A"
Private Const COL_COGNOME As String = "B"
Private Const COL_EMAIL As String = "F"

' --- Email -----------------------------------------------------------------------
Private Const OGGETTO_EMAIL As String = "Fatturato prestazioni solventi"
Private Const RANGE_IMMAGINE As String = "A1:G20"   ' area della pivot incollata

' Dati dell'ente per la fatturazione: SEGNAPOSTO, da sostituire
Private Const ENTE_SDI As String = "XXXXXXX"
Private Const ENTE_PEC As String = "fatturazione@esempio.it"
Private Const ENTE_DENOMINAZIONE As String = "NOME DELL'ENTE"
Private Const ENTE_PIVA As String = "IT00000000000"
Private Const ENTE_CF As String = "00000000000"
Private Const ENTE_INDIRIZZO As String = "Via Esempio 1"
Private Const ENTE_CAP As String = "00000"
Private Const ENTE_COMUNE As String = "Comune (XX)"
Private Const ENTE_NAZIONE As String = "IT"


'==============================================================================
' 1. AGGIORNAMENTO QUERY
'    Le due query Power Query leggono file Excel esterni (interventi e
'    estrazione ambulatoriale) posti in una cartella. Le macro chiedono conferma
'    e poi aggiornano la connessione corrispondente.
'==============================================================================
Public Sub Aggiorna_Interventi()
    AggiornaQuery QUERY_INTERVENTI, _
        "Hai inserito nella cartella giusta il nuovo file degli interventi " & _
        "e lo hai rinominato in 'interventi.xlsx'?"
End Sub

Public Sub Aggiorna_Ambulatoriale()
    AggiornaQuery QUERY_AMBULATORIALE, _
        "Hai inserito nella cartella giusta la nuova estrazione ambulatoriale?" & vbCrLf & _
        "E l'hai rinominata in 'ESTRAZIONE AMBULATORIO.xlsx'?"
End Sub

Private Sub AggiornaQuery(ByVal nomeQuery As String, ByVal messaggio As String)
    Dim cn As WorkbookConnection
    Dim risposta As VbMsgBoxResult

    risposta = MsgBox(messaggio, vbYesNo + vbExclamation, "Attenzione")
    If risposta <> vbYes Then Exit Sub

    On Error GoTo GestioneErrore
    For Each cn In ThisWorkbook.Connections
        ' Le connessioni Power Query hanno il nome della query nel nome
        ' ("Query - <nome>") e nella descrizione
        If cn.Name = nomeQuery Or cn.Description Like "*" & nomeQuery & "*" Then
            cn.Refresh
            Application.CalculateUntilAsyncQueriesDone   ' attende la fine
            MsgBox "Aggiornamento completato per la query: " & nomeQuery, vbInformation
            Exit Sub
        End If
    Next cn

    MsgBox "Connessione/Query '" & nomeQuery & "' non trovata.", vbExclamation
    Exit Sub

GestioneErrore:
    MsgBox "Errore durante l'aggiornamento: " & Err.Description, vbCritical
End Sub


'==============================================================================
' 2. PIVOT AMBULATORIALE
'    Somma dell'importo ambulatoriale per matricola.
'==============================================================================
Public Sub PvT_Ambulatoriale()
    Dim wsDati As Worksheet, wsPvt As Worksheet
    Dim rngOrigine As Range
    Dim pc As PivotCache, pt As PivotTable
    Dim risposta As VbMsgBoxResult

    risposta = MsgBox("Anagrafica aggiornata?" & vbCrLf & _
        "Sono presenti tutti i medici in extramoenia, le matricole e gli indirizzi email?", _
        vbYesNo + vbExclamation, "Attenzione")
    If risposta <> vbYes Then Exit Sub

    On Error GoTo GestioneErrore
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Rigenera il foglio pivot a ogni esecuzione
    EliminaFoglioSeEsiste FOGLIO_PVT_AMBULATORIALE

    Set wsDati = ThisWorkbook.Worksheets(FOGLIO_AMBULATORIALE)
    Set rngOrigine = wsDati.UsedRange            ' tutto il database, qualunque sia la dimensione

    Set wsPvt = ThisWorkbook.Worksheets.Add
    wsPvt.Name = FOGLIO_PVT_AMBULATORIALE

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=rngOrigine, Version:=8)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("A3"), _
        TableName:=PIVOT_AMBULATORIALE, DefaultVersion:=8)

    ' Riga = matricola, valore = somma importo
    With pt.PivotFields(CAMPO_AMB_MATRICOLA)
        .Orientation = xlRowField
        .Position = 1
    End With
    pt.AddDataField pt.PivotFields(CAMPO_AMB_IMPORTO), "Somma di importo", xlSum

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Verifica che non ci siano errori, celle vuote o N/D, " & _
           "poi esegui PvT_interventi.", vbOKOnly + vbExclamation, "Attenzione"
    Exit Sub

GestioneErrore:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Errore in PvT_Ambulatoriale: " & Err.Description, vbCritical
End Sub


'==============================================================================
' 3. PIVOT INTERVENTI (RICOVERO) + TOTALE DA FATTURARE
'    Pivot con filtro sulla matricola e righe per nominativo/nosologico.
'    A fianco aggiunge:
'      D4  fatturato ambulatoriale del medico selezionato (dalla pivot 2)
'      F4  totale fattura = ambulatoriale + somma degli interventi del medico
'==============================================================================
Public Sub PvT_interventi()
    Dim wsDati As Worksheet, wsPvt As Worksheet
    Dim rngOrigine As Range
    Dim pc As PivotCache, pt As PivotTable

    On Error GoTo GestioneErrore
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    EliminaFoglioSeEsiste FOGLIO_PVT_INTERVENTI

    Set wsDati = ThisWorkbook.Worksheets(FOGLIO_INTERVENTI)
    Set rngOrigine = wsDati.UsedRange

    Set wsPvt = ThisWorkbook.Worksheets.Add
    wsPvt.Name = FOGLIO_PVT_INTERVENTI

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=rngOrigine, Version:=8)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("A3"), _
        TableName:=PIVOT_INTERVENTI, DefaultVersion:=8)

    ' Filtro (pagina): matricola. Righe: nominativo e nosologico
    With pt.PivotFields(CAMPO_MATRICOLA)
        .Orientation = xlPageField
        .Position = 1
    End With
    With pt.PivotFields(CAMPO_NOMINATIVO)
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields(CAMPO_NOSOLOGICO)
        .Orientation = xlRowField
        .Position = 2
    End With
    pt.AddDataField pt.PivotFields(CAMPO_IMPORTO), "Somma di IMPORTO", xlSum

    wsPvt.Activate
    ActiveWindow.DisplayGridlines = False

    ' B1 contiene la matricola selezionata nel filtro della pivot.
    ' Si usa .Formula (sintassi inglese): funziona con qualsiasi lingua di Excel.
    wsPvt.Range("D3").Value = "Fatturato ambulatoriale"
    wsPvt.Range("D4").Formula = "=IFERROR(VLOOKUP(B1,'" & FOGLIO_PVT_AMBULATORIALE & _
                                "'!A:B,2,FALSE),0)"
    wsPvt.Range("F3").Value = "Totale fattura"
    wsPvt.Range("F4").Formula = "=D4+SUMIF('" & FOGLIO_INTERVENTI & "'!A:A,B1,'" & _
                                FOGLIO_INTERVENTI & "'!E:E)"

    wsPvt.Range("D3").Font.Bold = True
    wsPvt.Range("D3").EntireColumn.AutoFit
    With wsPvt.Range("F3:F4")
        .Font.Bold = True
        .NumberFormat = "#,##0.00 " & ChrW(8364)      ' formato euro
        .EntireColumn.AutoFit
    End With

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Verifica che i dati siano corretti e completi, " & _
           "prima di procedere con la preparazione delle e-mail.", _
           vbOKOnly + vbExclamation, "Attenzione"
    Exit Sub

GestioneErrore:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Errore in PvT_interventi: " & Err.Description, vbCritical
End Sub


'==============================================================================
' 4. EMAIL AI MEDICI
'    Per ogni medico dell'anagrafica con matricola ed email:
'      - imposta il filtro matricola della pivot interventi
'      - copia l'area della pivot come immagine
'      - apre in Outlook un messaggio con testo + immagine
'    I messaggi vengono SOLO APERTI (.Display): l'invio e' manuale.
'    I medici senza interventi nel periodo (matricola assente dalla pivot
'    interventi) vengono saltati.
'==============================================================================
Public Sub Invia_email()
    Dim risposta As VbMsgBoxResult
    Dim wsAnag As Worksheet, wsPivot As Worksheet
    Dim pt As PivotTable, pf As PivotField
    Dim rngImmagine As Range
    Dim ultimaRiga As Long, i As Long, preparate As Long
    Dim matricola As String, cognome As String, indirizzo As String
    Dim OutApp As Object, OutMail As Object
    Dim wdDoc As Object, wdSel As Object

    risposta = MsgBox("Hai eseguito tutti i controlli?" & vbCrLf & _
        "Procedo con la creazione delle email (una per medico)?", _
        vbYesNo + vbExclamation, "Attenzione")
    If risposta <> vbYes Then Exit Sub

    On Error GoTo GestioneErrore
    Application.ScreenUpdating = False

    Set wsAnag = ThisWorkbook.Worksheets(FOGLIO_ANAGRAFICA)
    Set wsPivot = ThisWorkbook.Worksheets(FOGLIO_PVT_INTERVENTI)
    Set pt = wsPivot.PivotTables(PIVOT_INTERVENTI)
    Set pf = pt.PivotFields(CAMPO_MATRICOLA)
    Set rngImmagine = wsPivot.Range(RANGE_IMMAGINE)

    ultimaRiga = wsAnag.Cells(wsAnag.Rows.Count, COL_MATRICOLA).End(xlUp).Row

    ' Una sola istanza di Outlook: usa quella aperta oppure la avvia
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")
    Err.Clear
    On Error GoTo GestioneErrore

    If OutApp Is Nothing Then
        MsgBox "Impossibile avviare Outlook. Verifica che sia installato.", vbCritical, "Errore"
        GoTo Uscita
    End If

    wsPivot.Activate     ' la copia come immagine richiede il foglio visibile

    For i = 2 To ultimaRiga                      ' riga 1 = intestazioni
        matricola = Trim(CStr(wsAnag.Cells(i, COL_MATRICOLA).Value))
        cognome = Trim(CStr(wsAnag.Cells(i, COL_COGNOME).Value))
        indirizzo = Trim(CStr(wsAnag.Cells(i, COL_EMAIL).Value))

        ' Salta righe senza matricola o senza email
        If Len(matricola) > 0 And Len(indirizzo) > 0 Then

            ' Filtra la pivot sul medico; se la matricola non e' nella pivot, salta
            If ImpostaFiltro(pf, matricola) Then
                wsPivot.Calculate                ' aggiorna D4 / F4 per il medico
                rngImmagine.CopyPicture Appearance:=xlScreen, Format:=xlPicture

                Set OutMail = OutApp.CreateItem(0)       ' 0 = olMailItem
                With OutMail
                    .To = indirizzo
                    .Subject = OGGETTO_EMAIL
                    .Body = CorpoEmail(cognome) & vbCrLf & vbCrLf
                    .Display                              ' serve per usare WordEditor

                    ' Incolla l'immagine in fondo al corpo del messaggio
                    Set wdDoc = .GetInspector.WordEditor
                    Set wdSel = wdDoc.Application.Selection
                    wdSel.EndKey Unit:=6                  ' 6 = wdStory (fine documento)
                    wdSel.TypeParagraph
                    wdSel.Paste
                End With

                Application.CutCopyMode = False
                preparate = preparate + 1

                Set OutMail = Nothing
                Set wdDoc = Nothing
                Set wdSel = Nothing
            End If
        End If
    Next i

Uscita:
    Application.ScreenUpdating = True
    MsgBox "Procedura completata. Email preparate: " & preparate, vbInformation, "Invio email"
    Set OutApp = Nothing
    Exit Sub

GestioneErrore:
    Application.ScreenUpdating = True
    MsgBox "Errore in Invia_email: " & Err.Description, vbCritical
End Sub


'==============================================================================
' FUNZIONI DI SUPPORTO
'==============================================================================

' Testo dell'email. Le lettere accentate sono scritte con ChrW() cosi' il file
' resta in ASCII puro e si importa correttamente in qualunque editor VBA.
Private Function CorpoEmail(ByVal cognome As String) As String
    Dim a As String, e As String, i As String
    Dim t As String

    a = ChrW(224)      ' a accentata
    e = ChrW(232)      ' e accentata
    i = ChrW(236)      ' i accentata

    t = "Gentile Dott.ssa/Dott. " & cognome & "," & vbCrLf & vbCrLf
    t = t & "di seguito il dettaglio di quanto dovr" & a & " fatturare per le prestazioni " & _
            "solventi rese lo scorso mese: nella fattura dovr" & a & " essere indicato il " & _
            "mese di riferimento delle prestazioni e la tipologia di prestazione fatturata " & _
            "(ambulatoriale e/o interventi). Di seguito vengono indicati gli importi " & _
            "relativi all'attivit" & a & " di ricovero ed ambulatoriale da fatturare." & _
            vbCrLf & vbCrLf

    t = t & "La fattura dovr" & a & " essere cos" & i & " redatta: " & vbCrLf
    t = t & "Fatturazione elettronica - Dati necessari" & vbCrLf
    t = t & "- Codice Destinatario (SDI): " & ENTE_SDI & vbCrLf
    t = t & "- PEC fatturazione: " & ENTE_PEC & vbCrLf
    t = t & "(da utilizzare esclusivamente per l'invio delle fatture elettroniche)" & vbCrLf
    t = t & "Dati anagrafici" & vbCrLf
    t = t & "- Denominazione: " & ENTE_DENOMINAZIONE & vbCrLf
    t = t & "- P. IVA: " & ENTE_PIVA & vbCrLf
    t = t & "- Codice Fiscale: " & ENTE_CF & vbCrLf
    t = t & "Sede" & vbCrLf
    t = t & "- Indirizzo: " & ENTE_INDIRIZZO & vbCrLf
    t = t & "- CAP: " & ENTE_CAP & vbCrLf
    t = t & "- Comune: " & ENTE_COMUNE & vbCrLf
    t = t & "- Nazione: " & ENTE_NAZIONE & vbCrLf & vbCrLf

    t = t & "Inoltre si comunica che l'Ente non " & e & " soggetto a Split Payment in quanto " & _
            "non " & e & " un ospedale pubblico e non " & e & " quindi dotato di codice CUU." & _
            vbCrLf & vbCrLf
    t = t & "Grazie" & vbCrLf
    t = t & "Cordiali saluti"

    CorpoEmail = t
End Function

' Imposta il filtro matricola della pivot. Restituisce False se la matricola
' non compare tra gli elementi della pivot.
Private Function ImpostaFiltro(pf As PivotField, ByVal matricola As String) As Boolean
    On Error Resume Next
    pf.ClearAllFilters
    pf.CurrentPage = matricola
    ImpostaFiltro = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' Elimina il foglio con quel nome, se esiste (chiamare con DisplayAlerts = False)
Private Sub EliminaFoglioSeEsiste(ByVal nomeFoglio As String)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nomeFoglio Then
            ws.Delete
            Exit For
        End If
    Next ws
End Sub
