# WordMD

Add-in per Microsoft Word (Office.js, TypeScript) che importa ed esporta Markdown e inserisce diagrammi Mermaid modificabili.

Il pannello ha quattro schede: **Importa**, **Stile**, **Esporta**, **Diagramma**.

- **Importa**: file `.md` (con le immagini accanto, una cartella intera o testo incollato) convertito in OOXML con gli stili esistenti del documento o con un profilo di stile (Word classico, GitHub, Accademico elegante, Report aziendale, Stili esistenti). I blocchi `mermaid` diventano diagrammi.
- **Esporta**: `.md`, `.docx` e `.mdocx`. Nel `.docx` i diagrammi sono immagini fisse (PNG + SVG); il `.mdocx` e lo stesso pacchetto Word ma conserva il sorgente Mermaid nel testo alternativo dell'immagine, cosi i diagrammi restano modificabili dal pannello.
- **Diagramma**: inserimento e modifica di diagrammi Mermaid (SVG autonomo + PNG).

Supporto Markdown: markdown-it con note a pie di pagina, liste di definizioni, evidenziato, apice/pedice.

## Requisiti

Node.js e Word desktop con supporto WordApi 1.5. Word carica il pannello solo via HTTPS: in sviluppo servono i certificati di `office-addin-dev-certs` (`npx office-addin-dev-certs install`).

## Comandi

```
npm install
npm run dev        # server Vite su https://localhost:3000
npm start          # avvia Word con il manifest (office-addin-debugging)
npm run build      # controllo tipi + build in dist/
npm test           # test Vitest (andata e ritorno Markdown -> Word -> Markdown)
npm run validate   # validazione di manifest.xml
```

## Struttura

```
manifest.xml          manifest dell'add-in (indirizzo di sviluppo https://localhost:3000)
src/taskpane/         interfaccia del pannello (index.html, main.ts, impostazioni, autotest)
src/lib/              conversioni: mdToOoxml, ooxmlToMd, profili di stile, rendering Mermaid, pacchetti OPC/docx
tests/                test e fixture (tests/fixtures/showcase.md)
assets/               icone
scripts/              generazione delle icone
docs/task_diary.md    stato dei task
```
