import { VgmArea } from './VgmSearchService';

export type VgmRecentAccess = {
  code: string;
  name: string;
  area: VgmArea;
  url: string;
  at: number;
};

type StoredPreferences = {
  favorites: string[];
  recent: VgmRecentAccess[];
};

export default class VgmUserPreferenceService {
  private readonly key: string;

  public constructor(email: string) {
    this.key = `vgm-intranet:prefs:${(email || 'anonymous').toLowerCase()}`;
  }

  public favorites(): string[] {
    return this.read().favorites;
  }

  public isFavorite(code: string): boolean {
    return this.read().favorites.indexOf(code) !== -1;
  }

  public toggleFavorite(code: string): boolean {
    const state: StoredPreferences = this.read();
    const index: number = state.favorites.indexOf(code);
    if (index >= 0) state.favorites.splice(index,1);
    else state.favorites.unshift(code);
    state.favorites = state.favorites.slice(0,50);
    this.write(state);
    return index < 0;
  }

  public recent(): VgmRecentAccess[] {
    return this.read().recent;
  }

  public addRecent(item: VgmRecentAccess): void {
    const state: StoredPreferences = this.read();
    state.recent = state.recent.filter((entry: VgmRecentAccess) => !(entry.code === item.code && entry.area === item.area));
    state.recent.unshift({ ...item, at: Date.now() });
    state.recent = state.recent.slice(0,15);
    this.write(state);
  }

  private read(): StoredPreferences {
    try {
      const raw: string | null = window.localStorage.getItem(this.key);
      if (!raw) return { favorites: [], recent: [] };
      const parsed: StoredPreferences = JSON.parse(raw) as StoredPreferences;
      return {
        favorites: Array.isArray(parsed.favorites) ? parsed.favorites : [],
        recent: Array.isArray(parsed.recent) ? parsed.recent : []
      };
    } catch {
      return { favorites: [], recent: [] };
    }
  }

  private write(value: StoredPreferences): void {
    try { window.localStorage.setItem(this.key,JSON.stringify(value)); } catch { /* no-op */ }
  }
}
