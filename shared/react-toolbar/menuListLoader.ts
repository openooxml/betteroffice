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

/**
 * While an open menu waits for its list, keeps keys from reaching the page behind it: Escape and Tab call `dismiss`
 * (Tab still moves focus), every other key is stopped. Does nothing once the list has loaded.
 */
export function guardPendingMenu(dismiss: (key: 'Escape' | 'Tab') => void): () => void {
  const onKeyDown = (event: KeyboardEvent) => {
    if (loaded) return;
    if (event.key === 'Escape' || event.key === 'Tab') {
      dismiss(event.key);
      return;
    }
    event.preventDefault();
    event.stopPropagation();
  };
  document.addEventListener('keydown', onKeyDown, true);
  return () => document.removeEventListener('keydown', onKeyDown, true);
}

export function resetMenuList(): void {
  loaded = undefined;
  loading = undefined;
}
