import { renderMermaid } from "../lib/mermaidRender";
import { ALL_MERMAID_VARIABLES, ELEMENTS, type ElementDef, type ElementFmt, type MermaidFmt } from "../lib/profile";
import { allProfileNames, findProfile, saveState, type AppState } from "../lib/store";
import { borderField, checkField, colorField, numberField, section, selectField, textField, type Opt } from "./form";

export interface SettingsHooks {
  /** Stili personalizzati presenti nel documento aperto. */
  documentStyles(): Promise<{ paragraph: string[]; character: string[] }>;
}

const ALIGN: Opt[] = [
  { value: "left", label: "A sinistra" }, { value: "center", label: "Centrato" },
  { value: "right", label: "A destra" }, { value: "justify", label: "Giustificato" },
];

const PARAGRAPH_STYLES: Opt[] = [
  ["Normal", "Normale"], ["heading 1", "Titolo 1"], ["heading 2", "Titolo 2"], ["heading 3", "Titolo 3"],
  ["heading 4", "Titolo 4"], ["heading 5", "Titolo 5"], ["heading 6", "Titolo 6"], ["Title", "Titolo"],
  ["Subtitle", "Sottotitolo"], ["Quote", "Citazione"], ["Intense Quote", "Citazione intensa"],
  ["Block Text", "Testo blocco"], ["List Paragraph", "Paragrafo elenco"], ["List Bullet", "Elenco puntato"],
  ["List Number", "Elenco numerato"], ["Body Text", "Corpo del testo"], ["No Spacing", "Nessuna spaziatura"],
  ["HTML Preformatted", "HTML preformattato"], ["Plain Text", "Testo normale"], ["caption", "Didascalia"],
  ["footnote text", "Testo della nota a piè di pagina"],
].map(([value, label]) => ({ value, label }));

const CHARACTER_STYLES: Opt[] = [
  ["Hyperlink", "Collegamento ipertestuale"], ["Strong", "Grassetto (stile)"], ["Emphasis", "Enfasi (stile)"],
].map(([value, label]) => ({ value, label }));

const FONTS = ["Calibri", "Calibri Light", "Cambria", "Segoe UI", "Arial", "Times New Roman", "Georgia", "Garamond",
  "Verdana", "Tahoma", "Palatino Linotype", "Consolas", "Cascadia Mono", "Courier New"];

const TEXT_FIELDS = ["font", "sizePt", "color", "bold", "italic", "underline", "strike", "background"] as const;

export function mountSettings(root: HTMLElement, state: AppState, hooks: SettingsHooks): void {
  let docStyles: Awaited<ReturnType<SettingsHooks["documentStyles"]>> = { paragraph: [], character: [] };
  let previewTimer = 0;
  let previewBox: HTMLElement | null = null;

  const changed = () => {
    saveState(state);
    schedulePreview();
  };

  const render = () => {
    root.replaceChildren();
    root.append(fontList(), profileBar(), modeBar());

    const els = section("Elementi del documento", true);
    for (const def of ELEMENTS) els.body.append(elementEditor(def));
    root.append(els.root, tableSection(), imageSection(), mermaidSection());
    schedulePreview();
  };

  const fontList = (): HTMLElement => {
    const dl = document.createElement("datalist");
    dl.id = "fonts";
    for (const f of FONTS) dl.append(new Option(f));
    return dl;
  };

  // ----- profili -----

  const profileBar = (): HTMLElement => {
    const box = document.createElement("div");
    box.className = "card";
    const names = allProfileNames(state);
    const holder = { choice: state.current.name };
    const sel = selectField("Profilo di stile", holder, "choice",
      [...names.map((n) => ({ value: n.name, label: n.preset ? n.name : `${n.name} (salvato)` }))],
      () => {
        const p = findProfile(state, holder.choice);
        if (p) {
          state.current = p;
          saveState(state);
          render();
        }
      });

    const actions = document.createElement("div");
    actions.className = "row-buttons";
    const save = button("Salva con questo nome", () => {
      const name = state.current.name.trim();
      if (!name) return;
      state.custom = [...state.custom.filter((c) => c.name !== name), structuredClone(state.current)];
      saveState(state);
      render();
    });
    const del = button("Elimina", () => {
      state.custom = state.custom.filter((c) => c.name !== state.current.name);
      saveState(state);
      render();
    });
    del.disabled = !state.custom.some((c) => c.name === state.current.name);
    actions.append(save, del);
    box.append(sel, textField("Nome dello stile", state.current as never, "name", () => saveState(state)), actions);
    return box;
  };

  const modeBar = (): HTMLElement => {
    const box = document.createElement("fieldset");
    box.className = "card";
    const legend = document.createElement("legend");
    legend.textContent = "Come applicare lo stile";
    box.append(legend);
    const options: [string, string, string][] = [
      ["new", "Crea uno stile nuovo", "Word riceve stili con il nome del profilo e le impostazioni qui sotto."],
      ["existing", "Usa gli stili già presenti in Word", "Ogni elemento prende lo stile che scegli; il documento decide l'aspetto."],
    ];
    for (const [value, title, hint] of options) {
      const id = `mode-${value}`;
      const row = document.createElement("div");
      row.className = "radio";
      const input = document.createElement("input");
      input.type = "radio";
      input.name = "mode";
      input.id = id;
      input.checked = state.current.mode === value;
      input.addEventListener("change", () => {
        state.current.mode = value as "new" | "existing";
        saveState(state);
        render();
      });
      const label = document.createElement("label");
      label.htmlFor = id;
      label.innerHTML = `<strong></strong><small></small>`;
      label.querySelector("strong")!.textContent = title;
      label.querySelector("small")!.textContent = hint;
      row.append(input, label);
      box.append(row);
    }
    return box;
  };

  // ----- elementi -----

  const elementEditor = (def: ElementDef): HTMLElement => {
    const fmt = (state.current.elements[def.key] ??= {});
    const sec = section(def.label);
    if (state.current.mode === "existing") {
      const extra = def.kind === "p" ? docStyles.paragraph : docStyles.character;
      const base = def.kind === "p" ? PARAGRAPH_STYLES : CHARACTER_STYLES;
      const opts = [...base, ...extra.filter((s) => !base.some((b) => b.value === s)).map((s) => ({ value: s, label: s }))];
      sec.body.append(selectField("Stile di Word", fmt, "wordStyle", opts, changed, def.kind === "c" ? "Nessuno (formato diretto)" : undefined));
      if (def.key === "codeInline" || def.key === "del" || def.key === "mark") sec.body.append(...textFields(fmt, def));
      return sec.root;
    }
    sec.body.append(...textFields(fmt, def));
    if (def.kind === "p") sec.body.append(...paragraphFields(fmt, def));
    return sec.root;
  };

  const textFields = (fmt: ElementFmt, def: ElementDef): HTMLElement[] => {
    const f: Record<(typeof TEXT_FIELDS)[number], () => HTMLElement> = {
      font: () => textField("Carattere", fmt, "font", changed, "fonts"),
      sizePt: () => numberField("Dimensione", fmt, "sizePt", changed, { step: 0.5, min: 4, max: 96, unit: "pt" }),
      color: () => colorField("Colore del testo", fmt, "color", changed),
      bold: () => checkField("Grassetto", fmt, "bold", changed),
      italic: () => checkField("Corsivo", fmt, "italic", changed),
      underline: () => checkField("Sottolineato", fmt, "underline", changed),
      strike: () => checkField("Barrato", fmt, "strike", changed),
      background: () => colorField(def.kind === "c" ? "Evidenziazione" : "Sfondo", fmt, "background", changed, "nessuno"),
    };
    return TEXT_FIELDS.map((k) => f[k]());
  };

  const paragraphFields = (fmt: ElementFmt, def: ElementDef): HTMLElement[] => {
    void def;
    return [
      selectField("Allineamento", fmt, "align", ALIGN, changed, "Predefinito"),
      numberField("Spazio prima", fmt, "beforePt", changed, { step: 1, min: 0, unit: "pt" }),
      numberField("Spazio dopo", fmt, "afterPt", changed, { step: 1, min: 0, unit: "pt" }),
      numberField("Interlinea (1 = singola)", fmt, "lineSpacing", changed, { step: 0.05, min: 0.5, max: 4 }),
      numberField("Rientro a sinistra", fmt, "indentPt", changed, { step: 1, min: 0, unit: "pt" }),
      checkField("Interruzione di pagina prima", fmt, "pageBreakBefore", changed),
      checkField("Mantieni con il paragrafo successivo", fmt, "keepNext", changed),
      checkField("Non spezzare tra due pagine", fmt, "keepLines", changed),
      borderField("Bordo sopra", fmt, "borderTop", changed),
      borderField("Bordo a sinistra", fmt, "borderLeft", changed),
      borderField("Bordo sotto", fmt, "borderBottom", changed),
    ];
  };

  // ----- tabelle e immagini -----

  const tableSection = (): HTMLElement => {
    const t = state.current.table;
    const s = section("Tabelle");
    s.body.append(
      numberField("Larghezza", t as never, "widthPercent", changed, { step: 5, min: 20, max: 100, unit: "%" }),
      selectField("Posizione", t as never, "align", ALIGN.slice(0, 3), changed),
      selectField("Bordi", t as never, "borders", [
        { value: "all", label: "Griglia completa" }, { value: "horizontal", label: "Solo linee orizzontali" },
        { value: "outer", label: "Solo contorno" }, { value: "none", label: "Nessuno" },
      ], changed),
      numberField("Spessore dei bordi", t as never, "borderWidthPt", changed, { step: 0.25, min: 0.25, unit: "pt" }),
      colorField("Colore dei bordi", t as never, "borderColor", changed, "nero"),
      selectField("Tipo di bordo", t as never, "borderStyle", [
        { value: "single", label: "Continuo" }, { value: "dashed", label: "Tratteggiato" },
        { value: "dotted", label: "Puntinato" }, { value: "double", label: "Doppio" },
      ], changed),
      colorField("Sfondo dell'intestazione", t as never, "headerBackground", changed, "nessuno"),
      colorField("Testo dell'intestazione", t as never, "headerColor", changed),
      checkField("Intestazione in grassetto", t as never, "headerBold", changed),
      colorField("Righe alternate", t as never, "zebraBackground", changed, "nessuna"),
      numberField("Margine interno delle celle", t as never, "cellPaddingPt", changed, { step: 0.5, min: 0, unit: "pt" }),
      checkField("Ripeti l'intestazione su ogni pagina", t as never, "repeatHeader", changed),
      checkField("Non spezzare le righe tra due pagine", t as never, "keepRowsTogether", changed),
    );
    return s.root;
  };

  const imageSection = (): HTMLElement => {
    const im = state.current.image;
    const s = section("Immagini");
    s.body.append(
      numberField("Larghezza massima", im as never, "maxWidthCm", changed, { step: 0.5, min: 1, max: 40, unit: "cm" }),
      selectField("Allineamento", im as never, "align", ALIGN.slice(0, 3), changed),
      checkField("Didascalia dal testo alternativo", im as never, "captionFromAlt", changed),
      checkField("Tieni l'immagine con la didascalia", im as never, "keepWithCaption", changed),
    );
    return s.root;
  };

  // ----- Mermaid -----

  const mermaidSection = (): HTMLElement => {
    const m = state.current.mermaid;
    const s = section("Diagrammi Mermaid");
    s.body.append(
      selectField("Tema", m as never, "theme", [
        { value: "base", label: "Personalizzato (colori sotto)" }, { value: "default", label: "Predefinito" },
        { value: "neutral", label: "Neutro" }, { value: "dark", label: "Scuro" }, { value: "forest", label: "Foresta" },
      ], () => { changed(); render(); }),
    );
    if (m.theme === "base") {
      for (const v of ALL_MERMAID_VARIABLES) s.body.append(colorField(v.label, m.variables, v.key, changed, "tema"));
    }
    s.body.append(
      textField("Carattere", m as never, "fontFamily", changed, "fonts"),
      numberField("Dimensione del testo", m as never, "fontSizePx", changed, { step: 1, min: 8, max: 40, unit: "px" }),
      numberField("Spessore delle linee", m as never, "lineWidthPx", changed, { step: 0.25, min: 0.25, max: 8, unit: "px" }),
      numberField("Spessore dei bordi dei nodi", m as never, "nodeBorderPx", changed, { step: 0.25, min: 0, max: 8, unit: "px" }),
      selectField("Curvatura delle linee", m as never, "curve", [
        { value: "basis", label: "Morbida" }, { value: "linear", label: "Dritta" }, { value: "step", label: "A gradini" },
        { value: "cardinal", label: "Cardinale" }, { value: "monotoneX", label: "Monotona" },
      ], changed),
      backgroundField(m),
      numberField("Larghezza massima", m as never, "maxWidthCm", changed, { step: 0.5, min: 2, max: 40, unit: "cm" }),
      selectField("Allineamento", m as never, "align", ALIGN.slice(0, 3), changed),
      checkField("Non separare il diagramma dal testo successivo", m as never, "keepWithNext", changed),
    );
    previewBox = document.createElement("div");
    previewBox.className = "preview";
    previewBox.setAttribute("aria-label", "Anteprima dello stile dei diagrammi");
    s.body.append(previewBox);
    return s.root;
  };

  const backgroundField = (m: MermaidFmt): HTMLElement => {
    const holder = { bg: m.background === "transparent" ? undefined : m.background };
    return colorField("Sfondo del diagramma", holder, "bg", () => {
      m.background = holder.bg ?? "transparent";
      changed();
    }, "trasparente");
  };

  const SAMPLE = `flowchart LR
  A[Inizio] --> B{Scelta}
  B -->|sì| C[Passo uno]
  B -->|no| D[Passo due]
  C --> E[Fine]
  D --> E`;

  function schedulePreview(): void {
    window.clearTimeout(previewTimer);
    previewTimer = window.setTimeout(async () => {
      const box = previewBox;
      if (!box || !box.isConnected) return;
      try {
        box.innerHTML = (await renderMermaid(SAMPLE, state.current.mermaid)).svg;
      } catch (e) {
        box.textContent = e instanceof Error ? e.message : String(e);
      }
    }, 350);
  }

  const button = (label: string, onClick: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.addEventListener("click", onClick);
    return b;
  };

  render();
  void hooks.documentStyles().then((s) => {
    docStyles = s;
    if (state.current.mode === "existing") render();
  }).catch(() => undefined);
}
