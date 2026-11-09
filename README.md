# Riepilogo mensile per medico (VBA / Excel)

Macro VBA che parte da un **database estratto dal gestionale**, **aggrega gli importi per singolo medico** e compila un **prospetto di riepilogo mensile**.

> ⚠️ Questo repository contiene **solo codice**. Non include dati reali: niente anagrafiche, importi, nomi di medici, pazienti o indirizzi email.

---

## Cosa fa

1. Prende l'estrazione grezza del gestionale (una riga per prestazione).
2. Ne crea una **copia di lavoro** e aggiunge quattro colonne calcolate:
   | Colonna | Significato |
   |---|---|
   | `Costo corretto` | importo con segno corretto: gli **storni** vengono sottratti |
   | `Bollo` | 2 € una sola volta per fattura, se il totale della fattura supera 77,47 € (in valore assoluto) |
   | `Chiave` | nome del medico ricavato dalla descrizione dell'agenda (testo prima della parentesi) |
   | `Matricola` | matricola del medico, cercata nell'anagrafica tramite la chiave |
3. Costruisce una **tabella pivot per matricola** (somma di costo corretto e bollo).
4. Riporta i totali nel foglio **Riepilogo**, una riga per medico. I medici senza movimenti nel periodo restano vuoti.

```mermaid
flowchart LR
    A[Estrazione gestionale] --> B[Copia di lavoro]
    B --> C[Costo corretto, Bollo, Chiave, Matricola]
    D[(Anagrafica medici)] --> C
    C --> E[Pivot per matricola]
    E --> F[Foglio Riepilogo]
```

---

## Contenuto del repository

| File | Descrizione |
|---|---|
| `RiepilogoALPI.bas` | Modulo VBA con le macro |
| `README.md` | Questo documento |

### Macro incluse

| Macro | Funzione |
|---|---|
| `Riepilogo` | Esegue l'intera elaborazione descritta sopra |
| `Eliminare_valori_negativi` | Utility: sulla selezione converte le formule in valori e ogni numero nel suo valore assoluto (non annullabile) |

---

## Struttura attesa della cartella di lavoro

I nomi dei fogli sono costanti all'inizio del modulo e si cambiano in un punto solo.

| Foglio | Contenuto |
|---|---|
| `estrazione LP4` | Esportazione del gestionale, intestazioni in riga 1. Colonne necessarie: `storno`, `importo`, `numero_fattura`, `desc_agenda`. Deve essere **ordinata per numero fattura** |
| `DB Medici` | Anagrafica con 4 colonne: `Chiave`, `Matricola`, `Cognome`, `Nome` |
| `Riepilogo` | Prospetto finale: matricola in colonna A, dati dalla riga 4 (riga 3 = intestazioni). La macro compila la colonna D (importo online) e la F (bollo online) |

`DB Risultato` e `Pvt` (pivot) vengono **ricreati a ogni esecuzione**: non modificarli a mano.

Esempio fittizio di `DB Medici`:

| Chiave | Matricola | Cognome | Nome |
|---|---|---|---|
| ROSSI MARIO | 1001 | ROSSI | MARIO |
| BIANCHI LUISA | 1002 | BIANCHI | LUISA |

Nell'estrazione, `desc_agenda` deve avere la forma `ROSSI MARIO (descrizione agenda)`: la chiave è il testo che precede ` (`.

---

## Come usarla

1. Apri una cartella di lavoro Excel con i fogli sopra descritti e salvala come `.xlsm`.
2. Premi `Alt + F11` per aprire l'editor VBA.
3. `File → Importa file…` e seleziona `RiepilogoALPI.bas`.
4. Incolla la nuova estrazione nel foglio `estrazione LP4`.
5. Esegui `Riepilogo` (`Alt + F8`, oppure assegnala a un pulsante).
6. Controlla il foglio `Riepilogo`.

---

## Scelte tecniche

- Le colonne dell'estrazione sono individuate **dal nome dell'intestazione**, non dalla lettera: la macro non si rompe se il gestionale aggiunge o sposta colonne.
- Le formule sono scritte in **R1C1 su tutto l'intervallo** in un'unica istruzione, senza `Select` né `FillDown`.
- La soglia del bollo è testo (`"77.47"`) perché le formule in R1C1 richiedono il punto decimale anche con Excel in italiano.
- I medici senza movimenti si riconoscono con `IsError`, senza dipendere dalla lingua di Excel (`#N/D` / `#N/A`).
- La pivot usa solo le quattro colonne calcolate, che hanno intestazioni uniche.

---

## Limiti noti

- Se `desc_agenda` non contiene la parentesi, la formula `Chiave` restituisce errore e la riga non viene attribuita ad alcun medico.
- Il calcolo del bollo presuppone l'estrazione ordinata per numero fattura.
- Il file `Riepilogo` originale prevede anche colonne compilate a mano (versamenti, trattenute); queste non fanno parte della macro.

---

## Privacy

- Non caricare mai su GitHub file `.xlsm`/`.xlsx`, estrazioni o anagrafiche reali. Un file Excel conserva anche dati nascosti (cache della pivot, fogli nascosti, metadati).
- Per aggiungere un file d'esempio, usare solo dati inventati.
- Suggerito un `.gitignore`:

```gitignore
*.xlsx
*.xlsm
*.xls
*.csv
~$*
```

---

## Requisiti

- Microsoft Excel per Windows (VBA abilitato)

## Licenza

Scegli una licenza (ad esempio MIT) e aggiungi un file `LICENSE`.
