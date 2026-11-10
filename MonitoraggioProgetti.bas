Attribute VB_Name = "MonitoraggioProgetti"
'==============================================================================
' MonitoraggioProgetti.bas
'
' Monitoraggio del budget dei progetti: dal budget di partenza (team mix, giornate
' e onorari per mese) al consuntivo mensile, con il calcolo del "return"
' (valore del contratto / onorari) a budget e a consuntivo.
'
' Questo modulo NON contiene dati, valori di costo o password reali:
'   - i valori orari di costo stanno nel foglio "Parametri" (non pubblicato);
'   - la password di protezione e' un segnaposto (PWD);
'   - i nomi delle linee di servizio sono segnaposto (LOS_A / LOS_B).
'
' Macro pubbliche (assegnabili a pulsanti)
'   SalvaProgetto      salva il budget di un nuovo progetto nel DB
'   RichiamaTeamMix    richiama nel foglio di compilazione il team mix di un progetto
'   ConsuntivaMese     registra le giornate effettive di un mese
'   EliminaProgetto    elimina un progetto da DB e report
'   SvuotaDB           svuota DB e report
'   GeneraReport       crea un foglio di report "fotografia" del progetto
'   Esporta            esporta il foglio attivo in una nuova cartella di lavoro
'   CancellaFoglio     elimina un foglio di report
'
' Fogli
'   Compilazione budget  maschera di input (visibile)
'   DB Progetti          database, una riga per progetto x mese x ruolo (nascosto, protetto)
'   Report Return        riepilogo return di tutti i progetti (visibile, protetto)
'   Parametri            valori orari di costo (nascosto, protetto, NON incluso nel repository)
'==============================================================================
Option Explicit

' --- Nomi dei fogli ----------------------------------------------------------
Private Const FG_BUDGET As String = "Compilazione budget"
Private Const FG_DB As String = "DB Progetti"
Private Const FG_REPORT As String = "Report Return"
Private Const FG_PARAMETRI As String = "Parametri"

' --- Protezione: SEGNAPOSTO, usare la propria password -----------------------------
Private Const PWD As String = "CAMBIA_PASSWORD"

' --- Linee di servizio (LoS): segnaposto -----------------------------------------------
' LOS_B ha un trattamento particolare: il valore orario gestionale deriva dalla
' tabella per ruolo (RANGE_COSTI_LOS_B); per le altre coincide con quello civilistico.
Private Const LOS_A As String = "LOS_A"
Private Const LOS_B As String = "LOS_B"

' --- Parametri di costo (nel foglio Parametri) -------------------------------------------
Private Const RANGE_COSTI As String = FG_PARAMETRI & "!$A:$E"          ' chiave FY&LoS&Ruolo -> valore (col. 5)
Private Const RANGE_COSTI_LOS_B As String = FG_PARAMETRI & "!$D$26:$F$33"  ' ruolo -> valore (col. 2)
Private Const ORE_GIORNATA As String = "8"

' --- Layout della maschera "Compilazione budget" --------------------------------------------
Private Const CELLA_PROGETTO As String = "B2"
Private Const CELLA_FY As String = "B3"                 ' es. FY24
Private Const CELLA_MESE_CONSUNTIVO As String = "F2"    ' numero del mese da consuntivare
Private Const CELLA_VALORE_CONTRATTO As String = "G19"
Private Const RIGA_INTESTAZIONE_MESI As Long = 5        ' "Mese 1", "Mese 2", ...
Private Const RIGA_PRIMO_RUOLO As Long = 7
Private Const RIGA_ULTIMO_RUOLO As Long = 14            ' 8 ruoli
Private Const COL_PRIMO_MESE As Long = 5                ' E: giornate del mese 1 (poi onorari in F)
Private Const N_MESI As Long = 12
Private Const RIGHE_PER_MESE As Long = 8                ' un blocco di 8 righe (ruoli) per mese nel DB

' --- Layout di "Report Return" ---------------------------------------------------------
Private Const RIGA_INTESTAZIONE_REPORT As Long = 8
Private Const PRIMA_RIGA_REPORT As Long = 9

' --- Foglio di report singolo -----------------------------------------------------------
Private Const ULTIMA_RIGA_DB As Long = 5000             ' limite per le formule matriciali


'==============================================================================
' 1. SALVA UN NUOVO PROGETTO
'    Scrive nel DB 96 righe (12 mesi x 8 ruoli) con giornate e onorari a budget
'    e aggiunge una riga al "Report Return". Poi azzera la maschera.
'==============================================================================
Public Sub SalvaProgetto()
    Dim wsB As Worksheet, wsDB As Worksheet, wsR As Worksheet
    Dim nome As String
    Dim riga As Long, m As Long, x As Long

    Set wsB = ThisWorkbook.Worksheets(FG_BUDGET)
    Set wsDB = ThisWorkbook.Worksheets(FG_DB)
    nome = NomeProgetto(wsB)

    ' --- Controlli ---------------------------------------------------------------
    If Len(nome) = 0 Then
        MsgBox "Inserisci nome progetto"
        Exit Sub
    End If
    If PrimaRigaProgetto(wsDB, nome) > 0 Then
        MsgBox T("Il progetto {e} gi{a} presente nel DB")
        Exit Sub
    End If
    If Len(wsB.Range(CELLA_FY).Value) = 0 Then
        MsgBox "Indicare il Fiscal Year con la sigla FY seguita dagli ultimi 2 numeri dell'anno d'interesse"
        Exit Sub
    End If
    If Application.WorksheetFunction.CountA(wsB.Range("B" & RIGA_PRIMO_RUOLO & ":B" & RIGA_ULTIMO_RUOLO)) = 0 Then
        MsgBox "Inserisci TeamMix di progetto"
        Exit Sub
    End If
    If wsB.Range(CELLA_VALORE_CONTRATTO).Value = 0 Then
        MsgBox "Inserisci valore del contratto"
        Exit Sub
    End If

    On Error GoTo Errore
    Application.ScreenUpdating = False
    SbloccaTutto
    Set wsR = ThisWorkbook.Worksheets(FG_REPORT)

    ' --- DB: una riga per mese e per ruolo -------------------------------------------
    riga = UltimaRiga(wsDB, "A") + 1
    For m = 1 To N_MESI
        For x = RIGA_PRIMO_RUOLO To RIGA_ULTIMO_RUOLO
            ScriviRigaBudget wsDB, riga, wsB, x, ColonnaGiornate(m)
            riga = riga + 1
        Next x
    Next m

    ' --- Report Return: una riga per progetto ------------------------------------------
    AggiungiRigaReport wsR, nome

    ' --- Azzera la maschera ------------------------------------------------------------
    AzzeraInput wsB
    wsB.Range(CELLA_PROGETTO).ClearContents
    wsB.Range(CELLA_FY).ClearContents

Uscita:
    On Error Resume Next
    BloccaTutto
    Application.ScreenUpdating = True
    Exit Sub

Errore:
    MsgBox "Errore in SalvaProgetto: " & Err.Description, vbExclamation
    Resume Uscita
End Sub

' Scrive nel DB la riga di budget di un ruolo per un mese.
' Colonne DB: A FY | B Nome progetto | C Ruolo | D LoS | E Mese | F Giornate mese |
' G Onorari budget civilistico | H Onorari budget gestionale | I Valore contratto |
' J Giornate effettive | L Numero mese | M Onorari consuntivo civilistico |
' N Onorari consuntivo gestionale | O Chiave report | P Ruolo consuntivo |
' Q LoS consuntivo | R Chiave report 2 | S Chiave report 3 (mese consuntivato)
Private Sub ScriviRigaBudget(wsDB As Worksheet, ByVal r As Long, wsB As Worksheet, _
                             ByVal rigaRuolo As Long, ByVal colGiornate As Long)
    wsDB.Cells(r, "A").Value = wsB.Range(CELLA_FY).Value
    wsDB.Cells(r, "B").Value = wsB.Range(CELLA_PROGETTO).Value
    wsDB.Cells(r, "C").Value = wsB.Cells(rigaRuolo, "A").Value
    wsDB.Cells(r, "D").Value = wsB.Cells(rigaRuolo, "B").Value
    wsDB.Cells(r, "E").Value = wsB.Cells(RIGA_INTESTAZIONE_MESI, colGiornate).Value
    wsDB.Cells(r, "F").Value = wsB.Cells(rigaRuolo, colGiornate).Value
    wsDB.Cells(r, "G").Value = wsB.Cells(rigaRuolo, colGiornate + 1).Value
    wsDB.Cells(r, "I").Value = wsB.Range(CELLA_VALORE_CONTRATTO).Value

    ' H: onorari budget gestionale
    wsDB.Cells(r, "H").Formula = "=IF(D" & r & "=""" & LOS_B & """," & _
        "VLOOKUP(C" & r & "," & RANGE_COSTI_LOS_B & ",2,FALSE)*" & ORE_GIORNATA & "*F" & r & ",G" & r & ")"
    ' L: numero del mese, ricavato da "Mese n"
    wsDB.Cells(r, "L").Formula = "=NUMBERVALUE(TRIM(RIGHT(E" & r & ",2)))"
    ' M: onorari consuntivo civilistico = giornate effettive x valore orario x ore
    wsDB.Cells(r, "M").Formula = "=IFERROR(J" & r & "*VLOOKUP(A" & r & "&Q" & r & "&P" & r & "," & _
        RANGE_COSTI & ",5,FALSE)*" & ORE_GIORNATA & ",0)"
    ' N: onorari consuntivo gestionale
    wsDB.Cells(r, "N").Formula = "=IF(Q" & r & "=""" & LOS_B & """," & _
        "VLOOKUP(P" & r & "," & RANGE_COSTI_LOS_B & ",2,FALSE)*" & ORE_GIORNATA & "*J" & r & ",M" & r & ")"
    ' O, R: chiavi per le somme dei report
    wsDB.Cells(r, "O").Formula = "=B" & r & "&E" & r
    wsDB.Cells(r, "R").Formula = "=B" & r & "&P" & r & "&Q" & r
End Sub

' Aggiunge in "Report Return" la riga del progetto, con i return a budget e a consuntivo
Private Sub AggiungiRigaReport(wsR As Worksheet, ByVal nome As String)
    Dim r As Long
    Dim d As String, rif As String

    d = "'" & FG_DB & "'!"
    r = Application.WorksheetFunction.Max(RIGA_INTESTAZIONE_REPORT, UltimaRiga(wsR, "B")) + 1
    rif = "B" & r

    wsR.Range("B" & r).Value = nome
    ' Return = valore del contratto / onorari (budget o consuntivo, civilistico o gestionale)
    wsR.Range("C" & r).Formula = "=IFERROR(VLOOKUP(" & rif & "," & d & "$B:$I,8,FALSE)/SUMIF(" & d & "$B:$B," & rif & "," & d & "$G:$G),0)"
    wsR.Range("D" & r).Formula = "=IFERROR(VLOOKUP(" & rif & "," & d & "$B:$I,8,FALSE)/SUMIF(" & d & "$B:$B," & rif & "," & d & "$H:$H),0)"
    wsR.Range("E" & r).Formula = "=IFERROR(VLOOKUP(" & rif & "," & d & "$B:$I,8,FALSE)/SUMIF(" & d & "$B:$B," & rif & "," & d & "$M:$M),0)"
    wsR.Range("F" & r).Formula = "=IFERROR(VLOOKUP(" & rif & "," & d & "$B:$I,8,FALSE)/SUMIF(" & d & "$B:$B," & rif & "," & d & "$N:$N),0)"
    ' Delta (budget - effettivo)
    wsR.Range("G" & r).Formula = "=C" & r & "-E" & r
    wsR.Range("H" & r).Formula = "=D" & r & "-F" & r
    ' Valore contratto e produzione (budget / effettiva)
    wsR.Range("I" & r).Formula = "=VLOOKUP(" & rif & "," & d & "$B:$I,8,FALSE)"
    wsR.Range("J" & r).Formula = "=SUMIF(" & d & "$B:$B," & rif & "," & d & "$H:$H)"
    wsR.Range("K" & r).Formula = "=SUMIF(" & d & "$B:$B," & rif & "," & d & "$G:$G)"
    wsR.Range("L" & r).Formula = "=SUMIF(" & d & "$B:$B," & rif & "," & d & "$N:$N)"
    wsR.Range("M" & r).Formula = "=SUMIF(" & d & "$B:$B," & rif & "," & d & "$M:$M)"

    BordiSottili wsR.Range("B" & r & ":M" & r)
End Sub


'==============================================================================
' 2. RICHIAMA IL TEAM MIX DI UN PROGETTO
'    Riporta nella maschera la linea di servizio (LoS) di ciascun ruolo, letta
'    dal primo mese del progetto nel DB (i ruoli in colonna A sono fissi).
'==============================================================================
Public Sub RichiamaTeamMix()
    Dim wsB As Worksheet, wsDB As Worksheet
    Dim nome As String
    Dim primaRiga As Long, x As Long

    Set wsB = ThisWorkbook.Worksheets(FG_BUDGET)
    Set wsDB = ThisWorkbook.Worksheets(FG_DB)
    nome = NomeProgetto(wsB)

    If Len(nome) = 0 Then
        MsgBox "Inserisci nome progetto"
        Exit Sub
    End If
    primaRiga = PrimaRigaProgetto(wsDB, nome)
    If primaRiga = 0 Then
        MsgBox T("Il progetto non {e} presente nel DB. Verificare correttezza del nome indicato")
        Exit Sub
    End If

    For x = RIGA_PRIMO_RUOLO To RIGA_ULTIMO_RUOLO
        wsB.Cells(x, "B").Value = wsDB.Cells(primaRiga + (x - RIGA_PRIMO_RUOLO), "D").Value
    Next x
End Sub


'==============================================================================
' 3. CONSUNTIVA UN MESE
'    Nel blocco di 8 righe del mese indicato (cella F2) scrive le giornate
'    effettive inserite nella maschera, il ruolo e la LoS a consuntivo e la chiave
'    "progetto + mese". Gli onorari a consuntivo si ricalcolano con le formule del DB.
'==============================================================================
Public Sub ConsuntivaMese()
    Dim wsB As Worksheet, wsDB As Worksheet
    Dim nome As String, mese As Variant
    Dim primaRiga As Long, rigaBlocco As Long, r As Long, m As Long, x As Long

    Set wsB = ThisWorkbook.Worksheets(FG_BUDGET)
    Set wsDB = ThisWorkbook.Worksheets(FG_DB)
    nome = NomeProgetto(wsB)
    mese = wsB.Range(CELLA_MESE_CONSUNTIVO).Value

    ' --- Controlli ---------------------------------------------------------------
    If Len(nome) = 0 Then
        MsgBox "Inserisci nome progetto"
        Exit Sub
    End If
    If Application.WorksheetFunction.CountA(wsB.Range("B" & RIGA_PRIMO_RUOLO & ":B" & RIGA_ULTIMO_RUOLO)) = 0 Then
        MsgBox "Richiama TeamMix di progetto"
        Exit Sub
    End If
    If Len(mese) = 0 Then
        MsgBox "Indicare il mese che si sta consuntivando"
        Exit Sub
    End If
    primaRiga = PrimaRigaProgetto(wsDB, nome)
    If primaRiga = 0 Then
        MsgBox T("Il progetto non {e} presente nel DB. Verificare correttezza del nome indicato")
        Exit Sub
    End If
    ' Mese gia' consuntivato? (la chiave "progetto + mese" e' nella colonna S)
    If Not IsError(Application.Match(nome & mese, wsDB.Range("S:S"), 0)) Then
        If MsgBox(T("Mese gi{a} consuntivato. Sovrascrivere?"), vbYesNo) = vbNo Then Exit Sub
    End If

    On Error GoTo Errore
    Application.ScreenUpdating = False
    SbloccaTutto

    For m = 1 To N_MESI
        rigaBlocco = primaRiga + (m - 1) * RIGHE_PER_MESE
        ' Si aggiorna solo il blocco il cui numero di mese (colonna L) coincide con F2
        If Not IsError(wsDB.Cells(rigaBlocco, "L").Value) Then
            If wsDB.Cells(rigaBlocco, "L").Value = mese Then
                For x = RIGA_PRIMO_RUOLO To RIGA_ULTIMO_RUOLO
                    r = rigaBlocco + (x - RIGA_PRIMO_RUOLO)
                    wsDB.Cells(r, "J").Value = wsB.Cells(x, ColonnaGiornate(m)).Value
                    wsDB.Cells(r, "P").Value = wsB.Cells(x, "A").Value
                    wsDB.Cells(r, "Q").Value = wsB.Cells(x, "B").Value
                    wsDB.Cells(r, "S").Value = nome & mese
                Next x
            End If
        End If
    Next m

    AzzeraInput wsB
    wsB.Range(CELLA_PROGETTO).ClearContents
    wsB.Range(CELLA_FY).ClearContents
    wsB.Range(CELLA_MESE_CONSUNTIVO).ClearContents

Uscita:
    On Error Resume Next
    BloccaTutto
    Application.ScreenUpdating = True
    Exit Sub

Errore:
    MsgBox "Errore in ConsuntivaMese: " & Err.Description, vbExclamation
    Resume Uscita
End Sub


'==============================================================================
' 4. ELIMINA UN PROGETTO (righe del DB e riga del Report Return)
'==============================================================================
Public Sub EliminaProgetto()
    Dim wsB As Worksheet, wsDB As Worksheet, wsR As Worksheet
    Dim nome As String
    Dim r As Long
    Dim daEliminare As Range

    Set wsB = ThisWorkbook.Worksheets(FG_BUDGET)
    Set wsDB = ThisWorkbook.Worksheets(FG_DB)
    nome = NomeProgetto(wsB)

    If Len(nome) = 0 Then
        MsgBox "Inserisci nome progetto"
        Exit Sub
    End If
    If PrimaRigaProgetto(wsDB, nome) = 0 Then
        MsgBox T("Il progetto non {e} presente nel DB. Verificare correttezza del nome indicato")
        Exit Sub
    End If
    If MsgBox("I dati del progetto indicato saranno eliminati. Vuoi proseguire?", vbYesNo) <> vbYes Then Exit Sub

    On Error GoTo Errore
    Application.ScreenUpdating = False
    SbloccaTutto
    Set wsR = ThisWorkbook.Worksheets(FG_REPORT)

    ' DB: righe del progetto (raccolte e poi eliminate in un colpo solo)
    For r = 2 To UltimaRiga(wsDB, "B")
        If wsDB.Cells(r, "B").Value = nome Then
            If daEliminare Is Nothing Then
                Set daEliminare = wsDB.Rows(r)
            Else
                Set daEliminare = Union(daEliminare, wsDB.Rows(r))
            End If
        End If
    Next r
    If Not daEliminare Is Nothing Then daEliminare.Delete Shift:=xlUp

    ' Report Return: dal basso verso l'alto per non perdere le posizioni
    For r = UltimaRiga(wsR, "B") To PRIMA_RIGA_REPORT Step -1
        If wsR.Cells(r, "B").Value = nome Then wsR.Rows(r).Delete
    Next r
    BordiSottili wsR.Range("B" & RIGA_INTESTAZIONE_REPORT & ":M" & RIGA_INTESTAZIONE_REPORT)

    wsB.Range(CELLA_PROGETTO).ClearContents

Uscita:
    On Error Resume Next
    BloccaTutto
    Application.ScreenUpdating = True
    Exit Sub

Errore:
    MsgBox "Errore in EliminaProgetto: " & Err.Description, vbExclamation
    Resume Uscita
End Sub


'==============================================================================
' 5. SVUOTA IL DATABASE
'==============================================================================
Public Sub SvuotaDB()
    Dim wsDB As Worksheet, wsR As Worksheet
    Dim ultima As Long

    If MsgBox(T("Il Database verr{a} svuotato e tutti i dati andranno persi. Vuoi proseguire?"), _
              vbYesNo) <> vbYes Then Exit Sub

    On Error GoTo Errore
    Application.ScreenUpdating = False
    SbloccaTutto
    Set wsDB = ThisWorkbook.Worksheets(FG_DB)
    Set wsR = ThisWorkbook.Worksheets(FG_REPORT)

    ultima = UltimaRiga(wsDB, "A")
    If ultima >= 2 Then wsDB.Range("A2:S" & ultima).ClearContents

    ultima = UltimaRiga(wsR, "B")
    If ultima >= PRIMA_RIGA_REPORT Then wsR.Range("B" & PRIMA_RIGA_REPORT & ":M" & ultima).Delete Shift:=xlUp
    BordiSottili wsR.Range("B" & RIGA_INTESTAZIONE_REPORT & ":M" & RIGA_INTESTAZIONE_REPORT)

Uscita:
    On Error Resume Next
    BloccaTutto
    Application.ScreenUpdating = True
    Exit Sub

Errore:
    MsgBox "Errore in SvuotaDB: " & Err.Description, vbExclamation
    Resume Uscita
End Sub


'==============================================================================
' 6. REPORT DEL SINGOLO PROGETTO
'    Crea un nuovo foglio con: return a budget / consuntivo / delta, andamento
'    mensile di giornate e produzione, consuntivo per ruolo e LoS. Il foglio e'
'    una "fotografia": a fine macro le formule diventano valori. Contiene un
'    pulsante (X) che lo elimina.
'==============================================================================
Public Sub GeneraReport()
    Dim wsB As Worksheet, wsDB As Worksheet, ws As Worksheet
    Dim nome As String, d As String, vlk As String
    Dim ruoli As Variant
    Dim i As Long, r As Long

    Set wsB = ThisWorkbook.Worksheets(FG_BUDGET)
    Set wsDB = ThisWorkbook.Worksheets(FG_DB)
    nome = NomeProgetto(wsB)

    If Len(nome) = 0 Then
        MsgBox "Inserisci nome progetto"
        Exit Sub
    End If
    If PrimaRigaProgetto(wsDB, nome) = 0 Then
        MsgBox T("Il progetto non {e} presente nel DB. Verificare correttezza del nome indicato")
        Exit Sub
    End If

    On Error GoTo Errore
    Application.ScreenUpdating = False
    ThisWorkbook.Unprotect PWD

    Set ws = ThisWorkbook.Worksheets.Add(After:=wsB)
    d = "'" & FG_DB & "'!"
    vlk = "VLOOKUP($B$4," & d & "$B:$I,8,FALSE)"

    ws.Cells.Font.Name = "Georgia"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 2.57

    ' --- Intestazione e return (B4:G7) ----------------------------------------------
    ws.Range("B4").Value = nome
    ws.Range("B5").Value = "Budget"
    ws.Range("D5").Value = "Consuntivo"
    ws.Range("F5").Value = "Delta"
    ws.Range("B6:G6").Value = Array("Return Civilistico", "Return Gestionale", _
                                    "Return Civilistico", "Return Gestionale", _
                                    "Civilistico", "Gestionale")
    ws.Range("B7").Formula = "=" & vlk & "/SUMIF(" & d & "$B:$B,$B$4," & d & "$G:$G)"
    ws.Range("C7").Formula = "=" & vlk & "/SUMIF(" & d & "$B:$B,$B$4," & d & "$H:$H)"
    ws.Range("D7").Formula = "=" & vlk & "/SUMIF(" & d & "$B:$B,$B$4," & d & "$M:$M)"
    ws.Range("E7").Formula = "=" & vlk & "/SUMIF(" & d & "$B:$B,$B$4," & d & "$N:$N)"
    ws.Range("F7").Formula = "=B7-D7"
    ws.Range("G7").Formula = "=C7-E7"
    ws.Range("B7:G7").NumberFormat = "0%"

    CentraUnisci ws.Range("B4:G4")
    CentraUnisci ws.Range("B5:C5")
    CentraUnisci ws.Range("D5:E5")
    CentraUnisci ws.Range("F5:G5")
    ws.Range("B4").Font.Size = 18: ws.Range("B4").Font.Bold = True
    ws.Range("B5:G5").Font.Size = 12
    ws.Range("B6:G6").Font.Size = 11
    ws.Range("B7:G7").Font.Size = 16
    ws.Range("B6:G7").HorizontalAlignment = xlCenter
    ws.Range("B6:G7").VerticalAlignment = xlCenter
    ws.Range("B4:G4").Interior.Color = 49407
    ws.Range("B5:C5").Interior.Color = 13434879
    ws.Range("D5:E5").Interior.Color = 13421823
    ws.Range("F5:G5").Interior.Color = 16777164
    ws.Rows(5).RowHeight = 24
    ws.Rows(6).RowHeight = 21.75
    ws.Rows(7).RowHeight = 30.75
    BordiSottili ws.Range("B4:G7")

    ' --- Mese di riferimento: ultimo mese consuntivato ------------------------------
    ws.Range("B2").Value = "Report al mese"
    ws.Range("C2").FormulaArray = "=MAX(IF(" & d & "$B$2:$B$" & ULTIMA_RIGA_DB & "=$B$4," & _
        "IF(" & d & "$S$2:$S$" & ULTIMA_RIGA_DB & "<>""""," & d & "$L$2:$L$" & ULTIMA_RIGA_DB & ")))"
    ws.Range("C2").Value = ws.Range("C2").Value
    With ws.Range("B2:C2")
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 12
    End With
    ws.Range("C2").HorizontalAlignment = xlCenter

    ' --- Andamento mensile (B10:H23) -------------------------------------------------
    ws.Range("C10").Value = "Giornate"
    ws.Range("E10").Value = "Produzione gestionale"
    ws.Range("G10").Value = "Prod. media al giorno"
    ws.Range("C11:H11").Value = Array("Budget", "Consuntivo", "Budget", "Consuntivo", "Budget", "Consuntivo")
    For i = 1 To N_MESI
        r = 11 + i
        ws.Cells(r, "B").Value = "Mese " & i
        ws.Cells(r, "C").Formula = "=SUMIF(" & d & "$O:$O,$B$4&B" & r & "," & d & "$F:$F)"   ' giornate budget
        ws.Cells(r, "D").Formula = "=SUMIF(" & d & "$O:$O,$B$4&B" & r & "," & d & "$J:$J)"   ' giornate effettive
        ws.Cells(r, "E").Formula = "=SUMIF(" & d & "$O:$O,$B$4&B" & r & "," & d & "$H:$H)"   ' produzione budget
        ws.Cells(r, "F").Formula = "=SUMIF(" & d & "$O:$O,$B$4&B" & r & "," & d & "$N:$N)"   ' produzione effettiva
        ws.Cells(r, "G").Formula = "=IFERROR(E" & r & "/C" & r & ","""")"                     ' prod. media budget
        ws.Cells(r, "H").Formula = "=IFERROR(F" & r & "/D" & r & ","""")"                     ' prod. media effettiva
    Next i
    CentraUnisci ws.Range("C10:D10")
    CentraUnisci ws.Range("E10:F10")
    CentraUnisci ws.Range("G10:H10")
    ws.Range("C10:H23").HorizontalAlignment = xlCenter
    ws.Range("E12:H23").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    ws.Range("C10:H11").Font.Bold = True
    BordiSottili ws.Range("C10:H23")
    BordiSottili ws.Range("B12:B23")

    ' --- Consuntivo per ruolo e LoS (B26:F42) ----------------------------------------
    ruoli = Array("Partner", "AP/ED/Director", "Senior Manager", "Manager", _
                  "Senior Associate", "Associate", "New Associate", "Stage")
    ws.Range("B26").Value = "Ruolo"
    ws.Range("C26").Value = "LoS"
    ws.Range("D26").Value = "Giornate consuntive"
    ws.Range("E26").Value = "Onorari consuntivi civilistico"
    ws.Range("F26").Value = "Onorari consuntivi gestionale"
    For i = 0 To UBound(ruoli)
        r = 27 + 2 * i                                    ' due righe per ruolo: LoS A e LoS B
        ws.Cells(r, "B").Value = ruoli(i)
        ws.Cells(r, "C").Value = LOS_A
        ws.Cells(r + 1, "B").Value = ruoli(i)
        ws.Cells(r + 1, "C").Value = LOS_B
    Next i
    For r = 27 To 42
        ws.Cells(r, "D").Formula = "=SUMIF(" & d & "$R:$R,$B$4&B" & r & "&C" & r & "," & d & "$J:$J)"
        ws.Cells(r, "E").Formula = "=SUMIF(" & d & "$R:$R,$B$4&B" & r & "&C" & r & "," & d & "$M:$M)"
        ws.Cells(r, "F").Formula = "=SUMIF(" & d & "$R:$R,$B$4&B" & r & "&C" & r & "," & d & "$N:$N)"
    Next r
    ws.Range("D26:F26").WrapText = True
    ws.Range("B26:F26").Font.Bold = True
    BordiSottili ws.Range("B26:F42")
    ws.Columns("B:H").AutoFit

    ' --- Fotografia: formule -> valori --------------------------------------------------
    Application.Calculate
    ws.UsedRange.Copy
    ws.UsedRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Nome del foglio: "<progetto> mese <n>" (max 31 caratteri, senza caratteri vietati)
    NominaFoglio ws, nome & " mese " & ws.Range("C2").Value

    ' Pulsante "X" per eliminare il foglio di report
    ws.Shapes.AddShape(msoShapeMathMultiply, 578.25, 15, 48.75, 48.75).OnAction = "CancellaFoglio"

    wsB.Range(CELLA_PROGETTO).ClearContents
    wsB.Range(CELLA_MESE_CONSUNTIVO).ClearContents
    ws.Range("A1").Select

Uscita:
    On Error Resume Next
    ThisWorkbook.Protect Password:=PWD, Structure:=True, Windows:=False
    Application.ScreenUpdating = True
    Exit Sub

Errore:
    MsgBox "Errore in GeneraReport: " & Err.Description, vbExclamation
    Resume Uscita
End Sub


'==============================================================================
' 7. ESPORTA IL FOGLIO ATTIVO (es. Report Return) IN UNA NUOVA CARTELLA DI LAVORO
'    Il foglio copiato viene sbloccato, senza logo, con le formule trasformate
'    in valori dalla riga 7 in giu'.
'==============================================================================
Public Sub Esporta()
    Dim ws As Worksheet
    Dim ultima As Long

    Application.ScreenUpdating = False
    ActiveSheet.Copy                         ' crea una nuova cartella di lavoro
    Set ws = ActiveSheet
    ws.Unprotect PWD

    On Error Resume Next
    ws.Shapes("Picture 1").Delete            ' logo
    On Error GoTo 0

    ultima = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If ultima >= 7 Then
        ws.Range("A7:N" & ultima).Copy
        ws.Range("A7:N" & ultima).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
    ws.Range("A1").Select
    Application.ScreenUpdating = True
End Sub


'==============================================================================
' 8. ELIMINA UN FOGLIO DI REPORT (collegata al pulsante X del report)
'    I fogli del modello non si possono eliminare.
'==============================================================================
Public Sub CancellaFoglio()
    Dim nome As String
    nome = ActiveSheet.Name

    If nome = FG_BUDGET Or nome = FG_DB Or nome = FG_REPORT Or nome = FG_PARAMETRI Then
        MsgBox "Questo foglio non puo' essere eliminato."
        Exit Sub
    End If

    ThisWorkbook.Unprotect PWD
    ActiveSheet.Delete                       ' Excel chiede conferma
    ThisWorkbook.Worksheets(FG_BUDGET).Activate
    ThisWorkbook.Protect Password:=PWD, Structure:=True, Windows:=False
End Sub


'==============================================================================
' FUNZIONI DI SUPPORTO
'==============================================================================

' Nome del progetto indicato nella maschera
Private Function NomeProgetto(wsB As Worksheet) As String
    NomeProgetto = Trim(CStr(wsB.Range(CELLA_PROGETTO).Value))
End Function

' Prima riga del DB in cui compare il progetto (0 se assente).
' Le righe di un progetto sono contigue: 12 blocchi mensili da 8 righe.
Private Function PrimaRigaProgetto(wsDB As Worksheet, ByVal nome As String) As Long
    Dim m As Variant
    m = Application.Match(nome, wsDB.Range("B:B"), 0)
    If IsError(m) Then
        PrimaRigaProgetto = 0
    Else
        PrimaRigaProgetto = CLng(m)
    End If
End Function

' Colonna delle giornate del mese m nella maschera (E, G, I, ... ; gli onorari sono a destra)
Private Function ColonnaGiornate(ByVal m As Long) As Long
    ColonnaGiornate = COL_PRIMO_MESE + 2 * (m - 1)
End Function

Private Function UltimaRiga(ws As Worksheet, ByVal col As String) As Long
    UltimaRiga = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' Azzera le giornate di tutti i mesi, il valore del contratto e il team mix
Private Sub AzzeraInput(wsB As Worksheet)
    Dim m As Long
    For m = 1 To N_MESI
        wsB.Range(wsB.Cells(RIGA_PRIMO_RUOLO, ColonnaGiornate(m)), _
                  wsB.Cells(RIGA_ULTIMO_RUOLO, ColonnaGiornate(m))).Value = 0
    Next m
    wsB.Range(CELLA_VALORE_CONTRATTO).Value = 0
    wsB.Range("B" & RIGA_PRIMO_RUOLO & ":B" & RIGA_ULTIMO_RUOLO).ClearContents
End Sub

' Rimuove le protezioni e mostra il DB
Private Sub SbloccaTutto()
    ThisWorkbook.Unprotect PWD
    With ThisWorkbook.Worksheets(FG_DB)
        .Visible = xlSheetVisible
        .Unprotect PWD
    End With
    ThisWorkbook.Worksheets(FG_REPORT).Unprotect PWD
End Sub

' Ripristina protezioni e nasconde il DB
Private Sub BloccaTutto()
    ThisWorkbook.Worksheets(FG_BUDGET).Activate
    With ThisWorkbook.Worksheets(FG_DB)
        .Protect Password:=PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True
        .EnableSelection = xlNoSelection
        .Visible = xlSheetHidden
    End With
    ThisWorkbook.Worksheets(FG_REPORT).Protect Password:=PWD, DrawingObjects:=True, _
        Contents:=True, Scenarios:=True
    ThisWorkbook.Protect Password:=PWD, Structure:=True, Windows:=False
End Sub

' Bordi sottili (esterni e interni) su un intervallo
Private Sub BordiSottili(rng As Range)
    Dim b As Variant
    For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                        xlInsideVertical, xlInsideHorizontal)
        With rng.Borders(b)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 0
        End With
    Next b
End Sub

' Unisce le celle e centra il contenuto
Private Sub CentraUnisci(rng As Range)
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    rng.Merge
End Sub

' Rinomina un foglio con un nome valido (max 31 caratteri, senza : \ / ? * [ ]);
' se il nome esiste gia' lascia quello predefinito
Private Sub NominaFoglio(ws As Worksheet, ByVal nome As String)
    Dim c As Variant
    For Each c In Array(":", "\", "/", "?", "*", "[", "]")
        nome = Replace(nome, c, "-")
    Next c
    On Error Resume Next
    ws.Name = Left(nome, 31)
    On Error GoTo 0
End Sub

' Sostituisce i segnaposto delle lettere accentate: il file resta in ASCII puro
Private Function T(ByVal testo As String) As String
    testo = Replace(testo, "{a}", ChrW(224))
    testo = Replace(testo, "{e}", ChrW(232))
    T = testo
End Function
