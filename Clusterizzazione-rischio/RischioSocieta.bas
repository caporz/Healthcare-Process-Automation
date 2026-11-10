Attribute VB_Name = "RischioSocieta"
'==============================================================================
' RischioSocieta.bas
'
' Clusterizzazione del rischio (Basso / Medio / Alto) delle societa' di un
' gruppo a partire da due questionari compilati in Excel:
'   - Modulo 1: processi svolti / esternalizzati dalla societa'
'   - Modulo 2: modello di rischio (verifica di completezza e rischio di default)
'
' La macro Salva:
'   1. controlla che i moduli siano compilati e coerenti;
'   2. conta, per quattro ambiti (corruzione, terrorismo, SSL, ambientale),
'      i processi svolti e quelli esternalizzati e li registra nel foglio DB;
'   3. scrive nel foglio "Calcolo Rischi" le formule che producono il rischio
'      per ambito, il rischio paese, il rischio HSE, i fattori interni e il
'      rischio finale;
'   4. svuota i moduli per la societa' successiva.
'
' Questo modulo NON contiene dati. Nessun nome di societa' e' scritto nel codice:
' i nomi vengono letti dalle celle dei moduli.
'==============================================================================
Option Explicit

' --- Nomi dei fogli ----------------------------------------------------------
Private Const FG_MODULO1 As String = "Modulo 1 - Processi"
Private Const FG_MODULO2 As String = "Modulo 2 - Risk Model"
Private Const FG_SOC_PAESE As String = "Associazione Soc.-Paese"
Private Const FG_PROC_AMBITI As String = "Associazione processi-ambiti"
Private Const FG_FATTORI_INTERNI As String = "Rischio Fattori interni"
Private Const FG_DB As String = "DB"
Private Const FG_CALCOLO As String = "Calcolo Rischi"

' --- Celle dei moduli ----------------------------------------------------------
Private Const CELLA_NOME_MODULO1 As String = "C2"
Private Const CELLA_NOME_MODULO2 As String = "E2"
Private Const CELLA_RISCHIO_DEFAULT As String = "E61"    ' "SI" = rischio basso per default
Private Const CELLA_CONTROLLO_MODULO1 As String = "C37"  ' 0 = modulo non compilato
Private Const CELLA_PAESE As String = "C2"               ' nel foglio dei fattori interni
Private Const TESTO_NON_COMPLETO As String = _
    "UNA O PIU' SEZIONI NON RISULTANO COMPLETE/COMPILATE, SI PREGA DI PROCEDERE ALLA COMPILAZIONE"

' --- Processi nel Modulo 1 (una riga per processo) -------------------------------
Private Const RIGA_PRIMO_PROCESSO As Long = 5
Private Const RIGA_ULTIMO_PROCESSO As Long = 36
Private Const MARCATORE As String = "a"                  ' valore che segna un processo

' --- Parametri di calcolo (testo: le formule R1C1 vogliono il punto decimale) -----
Private Const PESO_ESTERNALIZZATI As String = "0.75"     ' peso dei processi esternalizzati
Private Const SOGLIA_BASSA As String = "0.8"             ' sotto 80% del parametro = Basso
Private Const SOGLIA_ALTA As String = "1.2"              ' sopra 120% del parametro = Alto
Private Const PUNTI_BASSO As String = "0.33"
Private Const PUNTI_MEDIO As String = "0.66"
Private Const SOGLIA_PUNTI_BASSO As String = "1.65"      ' somma punti <= 1.65 = Basso
Private Const SOGLIA_PUNTI_MEDIO As String = "2.98"      ' somma punti <= 2.98 = Medio
Private Const SOGLIA_PAESE_ALTO As String = "0.68"       ' indice paese / parametro < 0.68 = Alto
Private Const SOGLIA_PAESE_MEDIO As String = "1.32"      ' < 1.32 = Medio, altrimenti Basso


'==============================================================================
' MACRO PRINCIPALE
'==============================================================================
Public Sub Salva()
    Dim wsMod1 As Worksheet, wsMod2 As Worksheet
    Dim wsDB As Worksheet, wsCalcolo As Worksheet, wsFattori As Worksheet
    Dim nome As String
    Dim riga As Long
    Dim trovato As Range

    On Error GoTo GestioneErrore
    Application.ScreenUpdating = False

    Set wsMod1 = ThisWorkbook.Worksheets(FG_MODULO1)
    Set wsMod2 = ThisWorkbook.Worksheets(FG_MODULO2)
    Set wsDB = ThisWorkbook.Worksheets(FG_DB)
    Set wsCalcolo = ThisWorkbook.Worksheets(FG_CALCOLO)
    Set wsFattori = ThisWorkbook.Worksheets(FG_FATTORI_INTERNI)

    ' --- 1. Controlli sul nome della societa' ---------------------------------
    If Len(wsMod2.Range(CELLA_NOME_MODULO2).Value) = 0 Or _
       Len(wsMod1.Range(CELLA_NOME_MODULO1).Value) = 0 Then
        MsgBox T("Inserire il nome della societ{a} nello Sheet Modulo 1 e/o nello Sheet Modulo 2")
        GoTo Uscita
    End If

    If wsMod2.Range(CELLA_NOME_MODULO2).Value <> wsMod1.Range(CELLA_NOME_MODULO1).Value Then
        MsgBox T("Il nome della societ{a} nello Sheet Modulo 1 non combacia con quello " & _
                 "inserito nello Sheet Modulo 2")
        GoTo Uscita
    End If

    nome = CStr(wsMod1.Range(CELLA_NOME_MODULO1).Value)

    ' Il nome deve esistere nell'associazione societa'-paese (ricerca esatta)
    If IsError(Application.Match(nome, ThisWorkbook.Worksheets(FG_SOC_PAESE).Range("A:A"), 0)) Then
        MsgBox "Il nome potrebbe non essere scritto correttamente: controllare eventuali " & _
               "accenti, spazi, ecc. in linea con lo sheet """ & FG_SOC_PAESE & """"
        GoTo Uscita
    End If

    ' --- 2. Societa' a rischio basso per default -----------------------------
    If wsMod2.Range(CELLA_RISCHIO_DEFAULT).Value = "SI" Then
        If MsgBox(T("La Societ{a} risulta essere a rischio basso per default. Corretto?"), _
                  vbYesNo) = vbYes Then
            riga = ProssimaRiga(wsDB, wsCalcolo)
            wsCalcolo.Range("A" & riga).Value = nome
            wsCalcolo.Range("P" & riga).Value = "Basso"
            wsDB.Range("A" & riga).Value = nome
            SvuotaModulo wsMod1
            SvuotaModulo wsMod2
        End If
        GoTo Uscita
    End If

    ' --- 3. Moduli completi? -------------------------------------------------
    If wsMod2.Range(CELLA_RISCHIO_DEFAULT).Value = TESTO_NON_COMPLETO Or _
       wsMod1.Range(CELLA_CONTROLLO_MODULO1).Value = 0 Then
        MsgBox "Lo Sheet Modulo 1 e/o lo Sheet Modulo 2 non sono stati interamente compilati"
        GoTo Uscita
    End If

    ' --- 4. Societa' gia' presente nel DB: sovrascrivere? ----------------------
    If Not IsError(Application.Match(nome, wsDB.Range("A:A"), 0)) Then
        If MsgBox(T("La Societ{a} {e} gi{a} presente nel DB. Sovrascrivere i dati?"), _
                  vbYesNo) <> vbYes Then
            GoTo Uscita
        End If
        ' Si elimina la stessa riga in entrambi i fogli, cosi' restano allineati
        Set trovato = wsCalcolo.Range("A:A").Find(What:=nome, LookAt:=xlWhole)
        If Not trovato Is Nothing Then trovato.EntireRow.Delete
        Set trovato = wsDB.Range("A:A").Find(What:=nome, LookAt:=xlWhole)
        If Not trovato Is Nothing Then trovato.EntireRow.Delete
    End If

    ' --- 5. Registrazione nel DB ---------------------------------------------
    ' La riga e' la stessa in "DB" e "Calcolo Rischi": alcune formule leggono
    ' il DB sulla stessa riga (DB!R[0]C10), quindi i due fogli restano allineati.
    riga = ProssimaRiga(wsDB, wsCalcolo)

    wsDB.Range("A" & riga).Value = nome
    ' Colonne del DB: per ogni ambito, processi svolti (col. C) e processi
    ' esternalizzati (col. D) del Modulo 1, confrontati con l'ambito (col. B..E)
    ' dell'associazione processi-ambiti.
    wsDB.Range("B" & riga).Value = ContaProcessi(wsMod1, "C", "B")   ' corruzione - svolti
    wsDB.Range("C" & riga).Value = ContaProcessi(wsMod1, "D", "B")   ' corruzione - esternalizzati
    wsDB.Range("D" & riga).Value = ContaProcessi(wsMod1, "C", "C")   ' terrorismo - svolti
    wsDB.Range("E" & riga).Value = ContaProcessi(wsMod1, "D", "C")   ' terrorismo - esternalizzati
    wsDB.Range("F" & riga).Value = ContaProcessi(wsMod1, "C", "D")   ' SSL - svolti
    wsDB.Range("G" & riga).Value = ContaProcessi(wsMod1, "D", "D")   ' SSL - esternalizzati
    wsDB.Range("H" & riga).Value = ContaProcessi(wsMod1, "C", "E")   ' ambientali - svolti
    wsDB.Range("I" & riga).Value = ContaProcessi(wsMod1, "D", "E")   ' ambientali - esternalizzati
    wsDB.Range("J" & riga).Value = wsFattori.Range(CELLA_PAESE).Value ' paese (valore fisso)

    ' --- 6. Formule del foglio "Calcolo Rischi" -----------------------------------
    ScriviFormuleRischio wsCalcolo, riga, nome

    ' --- 7. Pulizia dei moduli per la societa' successiva ----------------------------
    SvuotaModulo wsMod1
    SvuotaModulo wsMod2

Uscita:
    Application.ScreenUpdating = True
    Exit Sub

GestioneErrore:
    Application.ScreenUpdating = True
    MsgBox "Errore in Salva: " & Err.Description, vbExclamation
End Sub


'==============================================================================
' FORMULE DEL FOGLIO "CALCOLO RISCHI"
'   B..E  rischio per ambito: corruzione, terrorismo, SSL, ambientale
'   F     rischio processi (sintesi di B..E)
'   G..J  indicatori paese (indice, categoria, AML, terrorismo)
'   K     rischio paese (sintesi di G..J)
'   L     rischio HSE
'   M     rischio paese + HSE
'   N     fattori interni
'   O     rischio M + N
'   P     RISCHIO FINALE (F + O)
'==============================================================================
Private Sub ScriviFormuleRischio(ws As Worksheet, ByVal riga As Long, ByVal nome As String)
    Dim cifre As String

    ws.Range("A" & riga).Value = nome

    ' Rischio per ambito: (svolti - esternalizzati x peso) confrontato con il parametro
    ws.Range("B" & riga).FormulaR1C1 = FormulaAmbito(2, 3, 2)
    ws.Range("C" & riga).FormulaR1C1 = FormulaAmbito(4, 5, 3)
    ws.Range("D" & riga).FormulaR1C1 = FormulaAmbito(6, 7, 4)
    ws.Range("E" & riga).FormulaR1C1 = FormulaAmbito(8, 9, 5)

    ' Rischio processi: Basso = 0.33, Medio = 0.66, Alto = 1 punto per ambito
    cifre = "SUM((COUNTIF(R[0]C2:R[0]C5,""Basso"")*" & PUNTI_BASSO & ")+" & _
            "(COUNTIF(R[0]C2:R[0]C5,""Medio"")*" & PUNTI_MEDIO & ")+" & _
            "COUNTIF(R[0]C2:R[0]C5,""Alto""))"
    ws.Range("F" & riga).FormulaR1C1 = _
        "=IF(" & cifre & "<=" & SOGLIA_PUNTI_BASSO & ",""Basso""," & _
        "IF(" & cifre & "<=" & SOGLIA_PUNTI_MEDIO & ",""Medio"",""Alto""))"

    ' Indicatori paese
    ws.Range("G" & riga).FormulaR1C1 = _
        "=IF(VLOOKUP(DB!R[0]C10,CPI!C1:C2,2,FALSE)/Parametri!R8C3<" & SOGLIA_PAESE_ALTO & ",""Alto""," & _
        "IF(VLOOKUP(DB!R[0]C10,CPI!C1:C2,2,FALSE)/Parametri!R8C3<" & SOGLIA_PAESE_MEDIO & ",""Medio"",""Basso""))"
    ws.Range("H" & riga).FormulaR1C1 = "=VLOOKUP(DB!R[0]C10,SACE!C1:C9,9,FALSE)"
    ws.Range("I" & riga).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(VLOOKUP(R[0]C1,DB!C1:C10,10,FALSE),AML2017_FINAL!C1:C3,3,FALSE),"""")"
    ws.Range("J" & riga).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(VLOOKUP(R[0]C1,DB!C1:C10,10,FALSE),TERR2016_FINAL!C1:C3,3,FALSE),"""")"

    ' Rischio paese: almeno due "alto" tra G..J => alto (salvo due "basso" => medio)
    ws.Range("K" & riga).FormulaR1C1 = _
        "=IF(COUNTIF(R[0]C7:R[0]C10,""alto"")>=2,IF(COUNTIF(R[0]C7:R[0]C10,""basso"")=2,""medio"",""alto"")," & _
        "IF(COUNTIF(R[0]C7:R[0]C10,""basso"")>2,""basso"",""medio""))"

    ' Rischio HSE e sintesi paese + HSE
    ws.Range("L" & riga).FormulaR1C1 = "=VLOOKUP(R[0]C1,'Rischio HSE'!C3:C4,2,FALSE)"
    ws.Range("M" & riga).FormulaR1C1 = _
        "=IF(AND(R[0]C11=""Basso"",R[0]C12=""Assente""),""Basso""," & _
        "IF(OR(R[0]C11=""Alto"",R[0]C12=""Significativo""),""Alto"",""Medio""))"

    ' Fattori interni (valutati sulla compilazione corrente: poi diventano valore fisso)
    ws.Range("N" & riga).FormulaR1C1 = _
        "=IF(COUNTIF('" & FG_FATTORI_INTERNI & "'!R7C3,""" & MARCATORE & """)+" & _
        "COUNTIF('" & FG_FATTORI_INTERNI & "'!R12C3,""" & MARCATORE & """)>0,""Alto""," & _
        "IF(COUNTIF('" & FG_FATTORI_INTERNI & "'!R5C3:R11C3,""" & MARCATORE & """)=0,""Basso"",""Medio""))"

    ws.Range("O" & riga).FormulaR1C1 = _
        "=IF(AND(R[0]C13=""Alto"",R[0]C14=""Alto""),""Alto""," & _
        "IF(COUNTIF(R[0]C13:R[0]C14,""Basso"")>0,""Basso"",""Medio""))"

    ' Rischio finale
    ws.Range("P" & riga).FormulaR1C1 = _
        "=IF(AND(R[0]C6=""Alto"",R[0]C15=""Alto""),""Alto""," & _
        "IF((COUNTIF(R[0]C6,""Alto"")+COUNTIF(R[0]C15,""Alto""))>0,""Medio""," & _
        "IF(AND(R[0]C6=""Medio"",R[0]C15=""Medio""),""Medio"",""Basso"")))"

    ' I fattori interni dipendono dai moduli che stanno per essere svuotati:
    ' si salva il risultato come valore
    Application.Calculate
    ws.Range("N" & riga).Value = ws.Range("N" & riga).Value
End Sub

' Formula del rischio per un ambito.
'   colSvolti / colEsterni = colonne del DB (1 = A); colParametro = colonna del
'   foglio Parametri (riga 4) con il valore di riferimento dell'ambito.
Private Function FormulaAmbito(ByVal colSvolti As Long, ByVal colEsterni As Long, _
                               ByVal colParametro As Long) As String
    Dim punteggio As String
    punteggio = "(VLOOKUP(R[0]C1,DB!C1:C9," & colSvolti & ",FALSE)-" & _
                "VLOOKUP(R[0]C1,DB!C1:C9," & colEsterni & ",FALSE)*" & PESO_ESTERNALIZZATI & ")"
    FormulaAmbito = "=IF(" & punteggio & "<(Parametri!R4C" & colParametro & "*" & SOGLIA_BASSA & "),""Basso""," & _
                    "IF(" & punteggio & ">(Parametri!R4C" & colParametro & "*" & SOGLIA_ALTA & "),""Alto"",""Medio""))"
End Function


'==============================================================================
' FUNZIONI DI SUPPORTO
'==============================================================================

' Conta i processi (righe del Modulo 1) segnati nella colonna indicata (C = svolti,
' D = esternalizzati) per i quali la colonna C coincide con la cella dello stesso
' processo nell'associazione processi-ambiti (colonna dell'ambito).
' La condizione e' quella del codice originale.
Private Function ContaProcessi(wsMod1 As Worksheet, ByVal colModulo As String, _
                               ByVal colAmbito As String) As Long
    Dim wsAmbiti As Worksheet
    Dim i As Long, n As Long

    Set wsAmbiti = ThisWorkbook.Worksheets(FG_PROC_AMBITI)
    For i = RIGA_PRIMO_PROCESSO To RIGA_ULTIMO_PROCESSO
        If wsMod1.Range(colModulo & i).Value = MARCATORE And _
           wsMod1.Range("C" & i).Value = wsAmbiti.Range(colAmbito & i).Value Then
            n = n + 1
        End If
    Next i
    ContaProcessi = n
End Function

' Prossima riga libera, uguale nei due fogli (DB e Calcolo Rischi)
Private Function ProssimaRiga(wsDB As Worksheet, wsCalcolo As Worksheet) As Long
    Dim r1 As Long, r2 As Long
    r1 = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    r2 = wsCalcolo.Cells(wsCalcolo.Rows.Count, "A").End(xlUp).Row
    If r1 > r2 Then ProssimaRiga = r1 + 1 Else ProssimaRiga = r2 + 1
End Function

' Svuota le celle di input (sbloccate) di un modulo
Private Sub SvuotaModulo(ws As Worksheet)
    Dim c As Range
    For Each c In ws.UsedRange
        If c.Locked = False Then c.Value = ""
    Next c
End Sub

' Sostituisce i segnaposto delle lettere accentate: il file resta in ASCII puro
Private Function T(ByVal testo As String) As String
    testo = Replace(testo, "{a}", ChrW(224))
    testo = Replace(testo, "{e}", ChrW(232))
    T = testo
End Function
