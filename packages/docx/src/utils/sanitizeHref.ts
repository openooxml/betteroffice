import { sanitizeHref as sharedSanitizeHref } from '../../../../shared/safeHref';

export function sanitizeHref(href: string | null | undefined): string | undefined { return sharedSanitizeHref(href); }
