# Diario dei task — WordMD

Ultimo aggiornamento: 7 novembre 2026

## In corso adesso
- Provare il plugin dentro Word vero (pannello caricato, importazione e esportazione reali). Serve prima installare il certificato HTTPS di sviluppo: lo deve confermare l'utente (vedi sotto).

## In coda
- Prova nell'applicazione Word vera e pagina HTML di revisione per i controlli visivi (aspetto di stili, tabelle, diagrammi, stampa).
- Verificare in Word: note a piè di pagina inserite via OOXML, SVG nei diagrammi, stili che non si aggiornano se già presenti, apertura di un file `.mdocx` dal comando Apri di Word.
- Esportare anche le note a piè e le immagini remote nel Markdown (le immagini remote ora si scaricano solo se il sito lo permette).
- Inserire il plugin in un pacchetto da distribuire (build di produzione + manifest con indirizzo definitivo).

## Decisioni in attesa dell'utente
- Installare il certificato di sviluppo: comando `npx office-addin-dev-certs install` (Windows chiede una conferma).

## Fatti
- 7/11/2026: eliminato il vecchio abbozzo C# (cartelle `src` e `scratch_dev`, file `WordMD.slnx`), come deciso dall'utente.
- 7/11/2026: scelta l'architettura Office Add-in (Office.js), come deciso dall'utente.
- 7/11/2026: scritti i moduli: Markdown→Word (OOXML, con stili esistenti o nuovi), Word→Markdown, rendering Mermaid (SVG autonomo + PNG), profili di stile (Word classico, GitHub, Accademico elegante, Report aziendale, Stili esistenti), pannello con 4 schede, export `.md`/`.docx`/`.mdocx`, apertura di `.mdocx`.
- 7/11/2026: 14 test automatici passati (andata e ritorno Markdown→Word→Markdown con tutti gli elementi, `.docx`/`.mdocx`, dimensioni immagini). Compilazione e manifest validi.

## Decisioni già prese
- Architettura: Office Add-in in TypeScript, non VSTO.
- Il sorgente Mermaid di ogni diagramma viaggia nel testo alternativo dell'immagine in Word (titolo "WordMD Mermaid"), così resta modificabile.
- `.docx`: i diagrammi diventano immagini fisse (PNG + SVG) senza il sorgente, e gli SVG vengono salvati anche come file a fianco.
- `.mdocx`: lo stesso pacchetto Word, ma con il sorgente Mermaid nel testo alternativo; i diagrammi restano modificabili dal pannello. Word lo stampa direttamente, perché contiene già le immagini.
- Il doppio clic su `.mdocx` non si può associare da un add-in: si apre dal pannello ("Apri un file .mdocx") o con File › Apri.
- Immagini del Markdown: data URI, file scelti accanto al `.md` (percorsi relativi), indirizzi http se il sito lo permette.
- Stampa: nessuna conversione, il documento contiene immagini e stili veri; tabelle con intestazione ripetuta e righe non spezzate.

## Altri progetti aperti
- Nessuno.
