import type { MenuList } from './MenuList';

type MenuListComponent = typeof MenuList;

/** Fetches the menu list's chunk; replaceable in tests. */
export const menuListLoader = {
  load: (): Promise<MenuListComponent> => import('./MenuList').then((module) => module.MenuList),
};

let loaded: MenuListComponent | undefined;
let loading: Promise<MenuListComponent> | undefined;

export function loadedMenuList(): MenuListComponent | undefined {
  return loaded;
}

/** Loads the menu list once; a failed load is retried by the next caller. */
export function loadMenuList(): Promise<MenuListComponent> {
  loading ??= menuListLoader.load().then(
    (component) => (loaded = component),
    (error: unknown) => { loading = undefined; throw error; },
  );
  return loading;
}

/** Loads the menu list when the browser is idle, so a first open does not wait for it. */
export function preloadMenuList(): () => void {
  if (loaded) return () => {};
  const load = () => { loadMenuList().catch(() => {}); };
  if (typeof requestIdleCallback === 'function') {
    const id = requestIdleCallback(load);
    return () => cancelIdleCallback(id);
  }
  const id = setTimeout(load, 0);
  return () => clearTimeout(id);
}

export function resetMenuList(): void {
  loaded = undefined;
  loading = undefined;
}
