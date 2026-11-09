Attribute VB_Name = "RiepilogoALPI"
'==============================================================================
' RiepilogoALPI.bas
'
' Riepilogo mensile per medico a partire da un database estratto dal
' gestionale. Questo modulo NON contiene dati: solo logica.
'
' Macro incluse
'   1. Riepilogo                  - aggrega gli importi per medico (matricola)
'   2. Eliminare_valori_negativi  - utility: valori -> numeri positivi
'
' Fogli attesi nella cartella di lavoro (nomi modificabili nelle costanti):
'   "estrazione LP4"  esportazione grezza del gestionale (intestazioni in riga 1)
'   "DB Medici"       anagrafica: Chiave | Matricola | Cognome | Nome
'   "Riepilogo"       prospetto finale (matricola in col. A, dalla riga 4)
'==============================================================================
Option Explicit

' --- Nomi dei fogli ----------------------------------------------------------
Private Const FOGLIO_ESTRAZIONE As String = "estrazione LP4"
Private Const FOGLIO_MEDICI As String = "DB Medici"
Private Const FOGLIO_RISULTATO As String = "DB Risultato"
Private Const FOGLIO_PIVOT As String = "Pvt"
Private Const FOGLIO_RIEPILOGO As String = "Riepilogo"

' --- Parametri di calcolo ----------------------------------------------------
' Importo oltre il quale (in valore assoluto) su una fattura si applica il bollo.
' Le costanti sono testo perche' le formule in R1C1 vogliono il punto decimale
' anche con Excel in italiano.
Private Const SOGLIA_BOLLO As String = "77.47"
Private Const IMPORTO_BOLLO As String = "2"

' Prima riga di dati nel foglio Riepilogo (riga 3 = intestazioni)
Private Const RIGA_INIZIO As Long = 4


'==============================================================================
' 1. RIEPILOGO
'    Estrazione -> copia di lavoro -> colonne calcolate -> tabella pivot per
'    matricola -> riempimento del foglio Riepilogo.
'==============================================================================
Public Sub Riepilogo()
    Dim wsEstr As Worksheet, wsRis As Worksheet
    Dim wsPvt As Worksheet, wsRiep As Worksheet
    Dim cStorno As Long, cImporto As Long, cFattura As Long, cAgenda As Long
    Dim cCosto As Long, cBollo As Long, cChiave As Long, cMatr As Long
    Dim ultimaCol As Long, ultimaRiga As Long, ur As Long, r As Long
    Dim pc As PivotCache, pt As PivotTable
    Dim f As String

    On Error GoTo GestioneErrore
    Application.ScreenUpdating = False

    Set wsEstr = ThisWorkbook.Worksheets(FOGLIO_ESTRAZIONE)
    Set wsRiep = ThisWorkbook.Worksheets(FOGLIO_RIEPILOGO)

    ' --- 1. Rigenera i fogli di lavoro --------------------------------------
    Application.DisplayAlerts = False
    EliminaFoglioSeEsiste FOGLIO_RISULTATO
    EliminaFoglioSeEsiste FOGLIO_PIVOT
    wsEstr.Copy After:=wsEstr            ' la copia diventa il foglio attivo
    Set wsRis = ActiveSheet
    wsRis.Name = FOGLIO_RISULTATO
    Application.DisplayAlerts = True

    On Error Resume Next                 ' l'immagine (logo) potrebbe mancare
    wsRis.Shapes("Picture 1").Delete
    On Error GoTo GestioneErrore

    ' --- 2. Individua le colonne dall'intestazione (non per lettera) --------
    cStorno = TrovaColonna(wsRis, "storno")
    cImporto = TrovaColonna(wsRis, "importo")
    cFattura = TrovaColonna(wsRis, "numero_fattura")   ' prima occorrenza
    cAgenda = TrovaColonna(wsRis, "desc_agenda")
    If cStorno = 0 Or cImporto = 0 Or cFattura = 0 Or cAgenda = 0 Then
        Err.Raise vbObjectError + 1, , _
            "Intestazioni mancanti nell'estrazione " & _
            "(servono: storno, importo, numero_fattura, desc_agenda)."
    End If

    ultimaCol = wsRis.Cells(1, wsRis.Columns.Count).End(xlToLeft).Column
    ultimaRiga = wsRis.Cells(wsRis.Rows.Count, 1).End(xlUp).Row
    cCosto = ultimaCol + 1
    cBollo = ultimaCol + 2
    cChiave = ultimaCol + 3
    cMatr = ultimaCol + 4

    ' --- 3. Colonne calcolate (formule scritte su tutto l'intervallo) -------
    ' 3a. Costo corretto: gli storni vengono sottratti, il resto sommato
    wsRis.Cells(1, cCosto).Value = "Costo corretto"
    f = "=IF(RC" & cStorno & "="""",RC" & cImporto & ",-RC" & cImporto & ")"
    wsRis.Range(wsRis.Cells(2, cCosto), wsRis.Cells(ultimaRiga, cCosto)).FormulaR1C1 = f

    ' 3b. Bollo: 2 EUR una sola volta per fattura (prima riga di ogni numero)
    '     se il totale della fattura supera la soglia in valore assoluto.
    '     Richiede l'estrazione ordinata per numero fattura.
    wsRis.Cells(1, cBollo).Value = "Bollo"
    f = "=IF(RC" & cFattura & "=R[-1]C" & cFattura & ",0," & _
        "IF(OR(SUMIF(C" & cFattura & ",RC" & cFattura & ",C" & cCosto & ")>" & SOGLIA_BOLLO & "," & _
        "SUMIF(C" & cFattura & ",RC" & cFattura & ",C" & cCosto & ")<-" & SOGLIA_BOLLO & ")," & _
        IMPORTO_BOLLO & ",0))"
    wsRis.Range(wsRis.Cells(2, cBollo), wsRis.Cells(ultimaRiga, cBollo)).FormulaR1C1 = f

    ' 3c. Chiave: testo della colonna "desc_agenda" prima della parentesi
    '     (es. "COGNOME NOME (ambulatorio)" -> "COGNOME NOME")
    wsRis.Cells(1, cChiave).Value = "Chiave"
    f = "=LEFT(RC" & cAgenda & ",FIND(""("",RC" & cAgenda & ")-2)"
    wsRis.Range(wsRis.Cells(2, cChiave), wsRis.Cells(ultimaRiga, cChiave)).FormulaR1C1 = f

    ' 3d. Matricola: ricerca della chiave nell'anagrafica medici
    wsRis.Cells(1, cMatr).Value = "Matricola"
    f = "=VLOOKUP(RC" & cChiave & ",'" & FOGLIO_MEDICI & "'!C1:C4,2,FALSE)"
    wsRis.Range(wsRis.Cells(2, cMatr), wsRis.Cells(ultimaRiga, cMatr)).FormulaR1C1 = f

    ' --- 4. Tabella pivot: somma di costo corretto e bollo per matricola ----
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsRis.Range(wsRis.Cells(1, cCosto), wsRis.Cells(ultimaRiga, cMatr)))

    Set wsPvt = ThisWorkbook.Worksheets.Add(After:=wsRis)
    wsPvt.Name = FOGLIO_PIVOT
    Set pt = pc.CreatePivotTable(TableDestination:=wsPvt.Cells(3, 1), _
                                 TableName:="PivotTable1")

    With pt.PivotFields("Matricola")
        .Orientation = xlRowField
        .Position = 1
    End With
    pt.AddDataField pt.PivotFields("Costo corretto"), "Sum of Costo corretto", xlSum
    pt.AddDataField pt.PivotFields("Bollo"), "Sum of Bollo", xlSum
    pt.RowAxisLayout xlTabularRow

    ' --- 5. Foglio Riepilogo: importi online e bollo online per matricola ---
    ur = wsRiep.Cells(wsRiep.Rows.Count, 1).End(xlUp).Row
    If ur >= RIGA_INIZIO Then
        wsRiep.Range("D" & RIGA_INIZIO & ":D" & ur).ClearContents
        wsRiep.Range("F" & RIGA_INIZIO & ":F" & ur).ClearContents

        wsRiep.Range("D" & RIGA_INIZIO & ":D" & ur).FormulaR1C1 = _
            "=VLOOKUP(RC1," & FOGLIO_PIVOT & "!C1:C2,2,FALSE)"
        wsRiep.Range("F" & RIGA_INIZIO & ":F" & ur).FormulaR1C1 = _
            "=VLOOKUP(RC1," & FOGLIO_PIVOT & "!C1:C3,3,FALSE)"

        ' Medici senza movimenti nel periodo: la ricerca da' errore #N/D,
        ' si svuotano le celle (indipendente dalla lingua di Excel)
        For r = RIGA_INIZIO To ur
            If IsError(wsRiep.Cells(r, "D").Value) Then
                wsRiep.Cells(r, "D").ClearContents
                wsRiep.Cells(r, "F").ClearContents
            End If
        Next r
    End If

    wsRiep.Activate
    Application.ScreenUpdating = True
    MsgBox "Riepilogo aggiornato.", vbInformation
    Exit Sub

GestioneErrore:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Errore in Riepilogo: " & Err.Description, vbExclamation
End Sub


'==============================================================================
' 2. UTILITY - Eliminare_valori_negativi
'    Sull'intervallo selezionato: converte le formule in valori (ATTENZIONE:
'    operazione non annullabile) e trasforma ogni numero nel suo valore assoluto.
'==============================================================================
Public Sub Eliminare_valori_negativi()
    Dim rng As Range

    If TypeName(Selection) <> "Range" Then Exit Sub

    Selection.Value = Selection.Value
    For Each rng In Selection
        If WorksheetFunction.IsNumber(rng) Then
            rng.Value = Abs(rng.Value)
        End If
    Next rng
End Sub


'==============================================================================
' FUNZIONI DI SUPPORTO
'==============================================================================

' Numero di colonna dell'intestazione indicata in riga 1 (0 se non trovata)
Private Function TrovaColonna(ws As Worksheet, intestazione As String) As Long
    Dim m As Variant
    m = Application.Match(intestazione, ws.Rows(1), 0)
    If IsError(m) Then
        TrovaColonna = 0
    Else
        TrovaColonna = CLng(m)
    End If
End Function

' Elimina il foglio con quel nome, se presente
Private Sub EliminaFoglioSeEsiste(nomeFoglio As String)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nomeFoglio Then
            ws.Delete
            Exit For
        End If
    Next ws
End Sub
