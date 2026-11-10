import { checkMermaid, renderMermaid } from "../lib/mermaidRender";
import { loadState } from "../lib/store";
import { addLocalFiles, type LocalFiles } from "../lib/resources";
import { saveFiles } from "../lib/saveFiles";
import {
  exportMarkdown, exportWord, importMarkdown, insertDiagram, openWordFile, replaceSelectedDiagram, selectedDiagramCode,
} from "../lib/word";
import { mountSettings } from "./settings";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const state = loadState();
const localFiles: LocalFiles = new Map();

function say(id: string, text: string, kind: "ok" | "bad" | "" = ""): void {
  const el = $(id);
  el.textContent = text;
  el.className = "status" + (kind ? " " + kind : "");
}

/** Esegue un'azione lunga disabilitando il bottone e mostrando subito il feedback. */
async function busy(btn: HTMLButtonElement, statusId: string, label: string, job: () => Promise<void>): Promise<void> {
  btn.disabled = true;
  say(statusId, label);
  try {
    await job();
  } catch (e) {
    say(statusId, e instanceof Error ? e.message : String(e), "bad");
  } finally {
    btn.disabled = false;
  }
}

const firstLine = (s: string) => s.split("\n")[0];

function selectTab(id: string): void {
  for (const t of document.querySelectorAll<HTMLButtonElement>("[role=tab]")) {
    const on = t.id === id;
    t.setAttribute("aria-selected", String(on));
    $(t.getAttribute("aria-controls")!).hidden = !on;
  }
  if (id === "tab-import") updateSummary();
}

function initTabs(): void {
  const tabs = Array.from(document.querySelectorAll<HTMLButtonElement>("[role=tab]"));
  tabs.forEach((t, i) => {
    t.addEventListener("click", () => selectTab(t.id));
    t.addEventListener("keydown", (e) => {
      const next = e.key === "ArrowRight" ? i + 1 : e.key === "ArrowLeft" ? i - 1 : null;
      if (next === null) return;
      const target = tabs[(next + tabs.length) % tabs.length];
      selectTab(target.id);
      target.focus();
    });
  });
}

function updateSummary(): void {
  const p = state.current;
  $("style-summary").textContent =
    `${p.name} · ${p.mode === "new" ? "stile nuovo" : "stili di Word"} · diagrammi: ${p.mermaid.theme === "base" ? "personalizzati" : p.mermaid.theme}`;
}

async function documentStyles(): Promise<{ paragraph: string[]; character: string[] }> {
  return Word.run(async (ctx) => {
    const styles = ctx.document.getStyles();
    styles.load("items/nameLocal,items/builtIn,items/type");
    await ctx.sync();
    const custom = styles.items.filter((s) => !s.builtIn && !/^WordMD /.test(s.nameLocal));
    return {
      paragraph: custom.filter((s) => s.type === "Paragraph").map((s) => s.nameLocal),
      character: custom.filter((s) => s.type === "Character").map((s) => s.nameLocal),
    };
  });
}

function initImport(): void {
  const text = $<HTMLTextAreaElement>("md-text");
  const info = $("files-info");

  const take = async (files: FileList | null) => {
    if (!files?.length) return;
    const list = Array.from(files);
    addLocalFiles(localFiles, list);
    const md = list.find((f) => /\.(md|markdown|txt)$/i.test(f.name));
    if (md) {
      text.value = await md.text();
      $<HTMLInputElement>("exp-name").value = md.name.replace(/\.[^.]+$/, "");
    }
    const images = list.filter((f) => f.type.startsWith("image/")).length;
    info.textContent = `${md ? `Letto ${md.name}. ` : ""}Immagini disponibili: ${images}.`;
  };
  $<HTMLInputElement>("md-file").addEventListener("change", (e) => void take((e.target as HTMLInputElement).files));
  $<HTMLInputElement>("md-folder").addEventListener("change", (e) => void take((e.target as HTMLInputElement).files));
  $("btn-to-style").addEventListener("click", () => selectTab("tab-style"));

  const btn = $<HTMLButtonElement>("btn-import");
  btn.addEventListener("click", () => {
    const md = text.value;
    if (!md.trim()) return say("status-import", "Scegli un file o incolla del testo Markdown.", "bad");
    void busy(btn, "status-import", "Importo…", async () => {
      const r = await importMarkdown(md, state.current, localFiles, $<HTMLInputElement>("md-replace").checked);
      const problems = [
        ...r.failed.map((f) => `Diagramma ${f.index}: ${firstLine(f.message)}`),
        ...r.missingImages.map((m) => `Immagine ${m.src}: ${m.reason}`),
      ];
      const head = `Importato. Diagrammi: ${r.diagrams - r.failed.length} su ${r.diagrams}. Immagini: ${r.images}.`;
      say("status-import", problems.length ? `${head}\n${problems.join("\n")}` : head, problems.length ? "bad" : "ok");
    });
  });
}

function initExport(): void {
  const btn = $<HTMLButtonElement>("btn-export");
  btn.addEventListener("click", () => {
    const base = $<HTMLInputElement>("exp-name").value.trim().replace(/[\\/:*?"<>|]+/g, "_") || "documento";
    const fmt = document.querySelector<HTMLInputElement>("input[name=fmt]:checked")!.value;
    void busy(btn, "status-export", "Leggo il documento…", async () => {
      const r = fmt === "md" ? await exportMarkdown(base, state.current.mermaid) : await exportWord(base, fmt === "mdocx");
      const saved = await saveFiles(r.files);
      if (saved.saved === "cancelled") return say("status-export", "Esportazione annullata.");
      const where = saved.saved === "folder" ? "nella cartella scelta" : "nella cartella Download";
      let msg = `Salvati ${saved.count} file ${where}: ${r.files[0].name}`;
      if (r.files.length > 1) msg += ` e ${r.files.length - 1} altri (SVG e immagini).`;
      if (r.failed.length) msg += "\nNon esportati: " + r.failed.map((f) => `${f.fileName} (${firstLine(f.message)})`).join(", ");
      say("status-export", msg, r.failed.length ? "bad" : "ok");
    });
  });

  $<HTMLInputElement>("open-file").addEventListener("change", async (e) => {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    say("status-open", `Apro ${file.name}…`);
    try {
      const n = await openWordFile(new Uint8Array(await file.arrayBuffer()));
      say("status-open", `Aperto in una nuova finestra. Diagrammi modificabili: ${n}.`, "ok");
    } catch (err) {
      say("status-open", err instanceof Error ? err.message : String(err), "bad");
    } finally {
      input.value = "";
    }
  });
}

function initDiagram(): void {
  const code = $<HTMLTextAreaElement>("mm-code");
  const preview = $("mm-preview");
  const error = $("mm-error");
  let timer = 0;
  let version = 0;

  const refresh = async (): Promise<void> => {
    const mine = ++version;
    const source = code.value;
    const problem = source.trim() ? await checkMermaid(source) : "Scrivi il codice del diagramma.";
    if (mine !== version) return;
    error.textContent = problem ? problem.split("\n").slice(0, 3).join("\n") : "";
    if (problem) return;
    const d = await renderMermaid(source, state.current.mermaid);
    if (mine === version) preview.innerHTML = d.svg;
  };
  code.addEventListener("input", () => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => void refresh(), 400);
  });
  $("tab-diagram").addEventListener("click", () => void refresh());

  const insert = $<HTMLButtonElement>("btn-insert");
  insert.addEventListener("click", () =>
    void busy(insert, "status-diagram", "Inserisco…", async () => {
      await insertDiagram(code.value, state.current.mermaid);
      say("status-diagram", "Diagramma inserito.", "ok");
    }),
  );

  const update = $<HTMLButtonElement>("btn-update");
  update.addEventListener("click", () =>
    void busy(update, "status-diagram", "Aggiorno…", async () => {
      const done = await replaceSelectedDiagram(code.value, state.current.mermaid);
      say("status-diagram", done ? "Diagramma aggiornato." : "Seleziona prima un diagramma WordMD.", done ? "ok" : "bad");
    }),
  );

  // Selezionando un diagramma nel documento, il suo codice torna nell'editor.
  let lastSelected = "";
  const onSelection = async (): Promise<void> => {
    const src = await selectedDiagramCode();
    update.hidden = src === null;
    if (src !== null && src !== lastSelected) {
      lastSelected = src;
      code.value = src;
      selectTab("tab-diagram");
      void refresh();
    }
    if (src === null) lastSelected = "";
  };
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => void onSelection());
  void refresh();
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    document.body.textContent = "WordMD funziona solo in Microsoft Word.";
    return;
  }
  initTabs();
  initImport();
  initExport();
  initDiagram();
  mountSettings($("settings-root"), state, { documentStyles });
  updateSummary();
  if (import.meta.env.DEV) {
    void fetch("/__selftest-flag").then((r) => r.text()).then(async (t) => { if (t) await (await import("./selftest")).selfTest(t); }).catch(() => undefined);
  }
});
