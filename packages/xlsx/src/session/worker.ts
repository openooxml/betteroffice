import { createScopeTransport, type SessionScope } from '../../../../shared/office-session';
import { createWorkbookSessionHost } from './host';

createWorkbookSessionHost(createScopeTransport(self as unknown as SessionScope));
