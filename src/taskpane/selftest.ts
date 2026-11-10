/**
 * Autotest dentro Word vero (solo in sviluppo, con ?selftest nell'indirizzo del pannello).
 * Usa le stesse funzioni del pannello e manda l'esito a /__report del server di sviluppo.
 */
import showcase from "../../tests/fixtures/showcase.md?raw";
import { listDiagrams } from "../lib/docxTools";
import { PRESETS, existingStylesProfile } from "../lib/profile";
import { currentDocumentBytes, exportMarkdown, exportWord, importMarkdown } from "../lib/word";

const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const MD = `# Titolo

Testo con **grassetto**, un [link](https://example.com) e una nota[^1].

- uno
- due

| A | B |
|---|:-:|
| 1 | 2 |

\`\`\`mermaid
sequenceDiagram
  Alice->>Bob: Ciao
  Bob-->>Alice: Salve
\`\`\`

\`\`\`mermaid
mindmap
  root((Idea))
    Ramo A
    Ramo B
\`\`\`

![Logo](data:image/png;base64,${PNG})

[^1]: Testo della nota.
`;

type Step = { name: string; ok: boolean; detail?: string };
const steps: Step[] = [];
const rec = (name: string, ok: boolean, detail?: string) => steps.push({ name, ok, detail });

async function bodyInfo() {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    const ooxml = body.getOoxml();
    const pics = body.inlinePictures;
    pics.load("items/altTextTitle,items/altTextDescription,items/width,items/height");
    const fn = body.footnotes;
    fn.load("items");
    const tables = body.tables;
    tables.load("items/rowCount");
    const styles = ctx.document.getStyles();
    styles.load("items/nameLocal,items/builtIn");
    await ctx.sync();
    const xml = ooxml.value;
    return {
      xml,
      pics: pics.items.map((p) => ({ title: p.altTextTitle, descr: p.altTextDescription.slice(0, 40), w: Math.round(p.width), h: Math.round(p.height) })),
      footnotes: fn.items.length,
      tables: tables.items.map((t) => t.rowCount),
      wordmdStyles: styles.items.filter((s) => /^WordMD /.test(s.nameLocal)).map((s) => s.nameLocal),
    };
  });
}

async function run(): Promise<void> {
  for (const [label, profile] of [...PRESETS.map((p) => [p.id, p.make()] as const), ["existing", existingStylesProfile()] as const]) {
    try {
      const r = await importMarkdown(MD, profile, new Map(), true);
      rec(`import ${label}`, r.failed.length === 0 && r.missingImages.length === 0, JSON.stringify({ d: r.diagrams, f: r.failed, i: r.images }));
      const info = await bodyInfo();
      const mm = info.pics.filter((p) => p.title === "WordMD Mermaid");
      rec(`${label}: diagrammi nel documento`, mm.length === 2, JSON.stringify(info.pics));
      rec(`${label}: SVG conservato da Word`, info.xml.includes("svgBlip"));
      rec(`${label}: sorgente Mermaid con a capo`, mm.length > 0 && info.xml.includes("sequenceDiagram&#xA;") || info.xml.includes("sequenceDiagram&#10;") || /descr="sequenceDiagram\s*\n/.test(info.xml), info.xml.match(/descr="sequenceDiagram[^"]{0,30}/)?.[0]);
      rec(`${label}: nota a piè`, info.footnotes === 1, String(info.footnotes));
      rec(`${label}: tabella`, info.tables.length === 1 && info.tables[0] === 2, JSON.stringify(info.tables));
      rec(`${label}: stili WordMD`, profile.mode === "existing" ? !info.wordmdStyles.some((n) => n.includes("Stili esistenti")) : info.wordmdStyles.some((n) => n.startsWith(`WordMD ${profile.name} - `)), info.wordmdStyles.slice(0, 4).join(" | "));
      if (label === "GitHub") {
        const ex = await exportMarkdown("autotest", profile.mermaid);
        const md = ex.files[0].data as string;
        rec("export md", md.includes("```mermaid\nsequenceDiagram") && md.includes("[^1]: Testo della nota.") && md.includes("| A | B |"), md.slice(0, 400));
        rec("export svg", ex.files.filter((f) => f.name.endsWith(".svg")).length === 2 && ex.failed.length === 0, ex.files.map((f) => f.name).join(","));
        const mdocx = await exportWord("autotest", true);
        const bytes = mdocx.files[0].data as Uint8Array;
        rec("export mdocx", listDiagrams(bytes).length === 2 && listDiagrams(bytes).every((d) => d.svg && d.code), `bytes=${bytes.length}`);
        const docx = await exportWord("autotest", false);
        rec("export docx piatto", listDiagrams(docx.files[0].data as Uint8Array).length === 0 && docx.files.filter((f) => f.name.endsWith(".svg")).length === 2, docx.files.map((f) => f.name).join(","));
        rec("getFileAsync", (await currentDocumentBytes()).length > 1000);
      }
    } catch (e) {
      rec(`import ${label}`, false, e instanceof Error ? `${e.message}\n${e.stack?.slice(0, 400)}` : String(e));
    }
  }
}

/** Controlla che gli stili creati da Word abbiano le proprietà del profilo, e che un secondo import le aggiorni. */
async function styleChecks(): Promise<void> {
  const readH1 = async (name: string) =>
    Word.run(async (ctx) => {
      const st = ctx.document.getStyles().getByNameOrNullObject(`WordMD ${name} - Heading 1`);
      st.load("isNullObject,font/color,font/size,font/name,font/bold");
      await ctx.sync();
      return st.isNullObject ? null : { color: st.font.color, size: st.font.size, name: st.font.name, bold: st.font.bold };
    });
  const github = PRESETS.find((p) => p.id === "GitHub")!.make();
  await importMarkdown("# A", github, new Map(), true);
  const a = await readH1("GitHub");
  rec("stile Heading 1 GitHub", !!a && a.size === 24 && a.bold === true && a.name === "Segoe UI", JSON.stringify(a));
  github.elements.h1.color = "FF0000";
  github.elements.h1.sizePt = 30;
  await importMarkdown("# A", github, new Map(), true);
  const b = await readH1("GitHub");
  rec("secondo import aggiorna lo stile", !!b && b.size === 30 && /ff0000/i.test(b.color), JSON.stringify(b));
}

export async function selfTest(mode = "1"): Promise<void> {
  try {
    if (mode.startsWith("visual:")) {
      const profile = PRESETS.find((p) => p.id === mode.slice(7))?.make() ?? existingStylesProfile();
      const r = await importMarkdown(showcase, profile, new Map(), true);
      rec(`visual ${profile.name}`, r.failed.length === 0 && r.missingImages.length === 0, JSON.stringify(r));
    } else {
      await run();
      await styleChecks();
    }
  } catch (e) {
    rec("autotest", false, String(e));
  }
  await fetch("/__report", { method: "POST", body: JSON.stringify({ steps, done: true }, null, 2) });
}
