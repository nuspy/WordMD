import { PRESETS, defaultProfile, normalizeProfile, type Profile } from "./profile";

const KEY = "wordmd.state.v1";

export interface AppState {
  /** Profilo attivo, con le modifiche in corso. */
  current: Profile;
  /** Profili salvati dall'utente con nome. */
  custom: Profile[];
}

/** Il pannello può non avere accesso allo storage del browser (finestre private, dati bloccati): l'app funziona comunque. */
export function loadState(): AppState {
  try {
    const raw = JSON.parse(localStorage.getItem(KEY) ?? "null") as Partial<AppState> | null;
    if (raw) {
      return {
        current: normalizeProfile(raw.current),
        custom: (raw.custom ?? []).map((p) => normalizeProfile(p)),
      };
    }
  } catch {
    /* stato illeggibile: si riparte dal profilo predefinito */
  }
  return { current: defaultProfile(), custom: [] };
}

export function saveState(state: AppState): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(state));
  } catch {
    /* storage non disponibile */
  }
}

export function allProfileNames(state: AppState): { name: string; preset: boolean }[] {
  return [
    ...PRESETS.map((p) => ({ name: p.id, preset: true })),
    ...state.custom.filter((c) => !PRESETS.some((p) => p.id === c.name)).map((c) => ({ name: c.name, preset: false })),
  ];
}

export function findProfile(state: AppState, name: string): Profile | null {
  const custom = state.custom.find((c) => c.name === name);
  if (custom) return structuredClone(custom);
  return PRESETS.find((p) => p.id === name)?.make() ?? null;
}
