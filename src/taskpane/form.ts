/** Piccoli costruttori di campi per i pannelli delle impostazioni. Ogni campo modifica direttamente l'oggetto passato. */

export interface Opt {
  value: string;
  label: string;
}

let uid = 0;
const nextId = () => `f${++uid}`;

function wrap(label: string, control: HTMLElement, id: string, inline = false): HTMLElement {
  const row = document.createElement("div");
  row.className = inline ? "field inline" : "field";
  const l = document.createElement("label");
  l.htmlFor = id;
  l.textContent = label;
  row.append(l, control);
  return row;
}

type Obj = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any

export function textField(label: string, o: Obj, key: string, onChange: () => void, list?: string): HTMLElement {
  const id = nextId();
  const input = document.createElement("input");
  input.type = "text";
  input.id = id;
  input.value = o[key] ?? "";
  input.autocomplete = "off";
  input.spellcheck = false;
  if (list) input.setAttribute("list", list);
  input.addEventListener("input", () => {
    if (input.value) o[key] = input.value;
    else delete o[key];
    onChange();
  });
  return wrap(label, input, id);
}

export function numberField(
  label: string, o: Obj, key: string, onChange: () => void,
  opt: { step?: number; min?: number; max?: number; unit?: string } = {},
): HTMLElement {
  const id = nextId();
  const input = document.createElement("input");
  input.type = "number";
  input.id = id;
  input.inputMode = "decimal";
  input.step = String(opt.step ?? 1);
  if (opt.min !== undefined) input.min = String(opt.min);
  if (opt.max !== undefined) input.max = String(opt.max);
  input.value = o[key] === undefined ? "" : String(o[key]);
  input.addEventListener("input", () => {
    const v = parseFloat(input.value);
    if (Number.isFinite(v)) o[key] = v;
    else delete o[key];
    onChange();
  });
  return wrap(opt.unit ? `${label} (${opt.unit})` : label, input, id);
}

export function checkField(label: string, o: Obj, key: string, onChange: () => void): HTMLElement {
  const id = nextId();
  const input = document.createElement("input");
  input.type = "checkbox";
  input.id = id;
  input.checked = !!o[key];
  input.addEventListener("change", () => {
    if (input.checked) o[key] = true;
    else delete o[key];
    onChange();
  });
  const row = wrap(label, input, id, true);
  row.classList.add("check");
  row.prepend(input);
  return row;
}

export function selectField(label: string, o: Obj, key: string, options: Opt[], onChange: () => void, emptyLabel?: string): HTMLElement {
  const id = nextId();
  const sel = document.createElement("select");
  sel.id = id;
  const all = emptyLabel !== undefined ? [{ value: "", label: emptyLabel }, ...options] : options;
  for (const op of all) {
    const e = document.createElement("option");
    e.value = op.value;
    e.textContent = op.label;
    sel.append(e);
  }
  const current = o[key] ?? "";
  if (!all.some((a) => a.value === current) && current) {
    const e = document.createElement("option");
    e.value = current;
    e.textContent = current;
    sel.append(e);
  }
  sel.value = current;
  sel.addEventListener("change", () => {
    if (sel.value) o[key] = sel.value;
    else delete o[key];
    onChange();
  });
  return wrap(label, sel, id);
}

/** Colore RRGGBB (senza #) con possibilità di toglierlo ("automatico"). */
export function colorField(label: string, o: Obj, key: string, onChange: () => void, emptyText = "automatico"): HTMLElement {
  const id = nextId();
  const box = document.createElement("div");
  box.className = "color";
  const input = document.createElement("input");
  input.type = "color";
  input.id = id;
  const state = document.createElement("span");
  state.className = "color-state";
  const clear = document.createElement("button");
  clear.type = "button";
  clear.className = "ghost";
  clear.textContent = "Togli";
  clear.setAttribute("aria-label", `Togli il colore: ${label}`);
  const paint = () => {
    const v = o[key] as string | undefined;
    input.value = v ? "#" + v.toLowerCase() : "#000000";
    state.textContent = v ? "#" + v.toUpperCase() : emptyText;
    clear.hidden = !v;
    input.classList.toggle("unset", !v);
  };
  input.addEventListener("input", () => {
    o[key] = input.value.slice(1).toUpperCase();
    paint();
    onChange();
  });
  clear.addEventListener("click", () => {
    delete o[key];
    paint();
    onChange();
  });
  paint();
  box.append(input, state, clear);
  return wrap(label, box, id);
}

/** Bordo: colore + spessore in punti (0 = nessun bordo). */
export function borderField(label: string, o: Obj, key: string, onChange: () => void): HTMLElement {
  const group = document.createElement("div");
  group.className = "pair";
  const b = (o[key] ??= undefined) as { color: string; widthPt: number } | undefined;
  const tmp: Obj = { color: b?.color, widthPt: b?.widthPt ?? 0 };
  const commit = () => {
    if (tmp.widthPt > 0) o[key] = { color: tmp.color ?? "000000", widthPt: tmp.widthPt };
    else delete o[key];
    onChange();
  };
  group.append(
    numberField(`${label}: spessore`, tmp, "widthPt", commit, { step: 0.25, min: 0, unit: "pt" }),
    colorField(`${label}: colore`, tmp, "color", commit, "nero"),
  );
  return group;
}

export function section(title: string, open = false): { root: HTMLDetailsElement; body: HTMLElement } {
  const root = document.createElement("details");
  root.open = open;
  const sum = document.createElement("summary");
  sum.textContent = title;
  const body = document.createElement("div");
  body.className = "details-body";
  root.append(sum, body);
  return { root, body };
}
