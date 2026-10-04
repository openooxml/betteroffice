import { createScopeTransport, type SessionScope } from '../../../../shared/office-session';
import { createPresentationSessionHost } from './host';

createPresentationSessionHost(createScopeTransport(self as unknown as SessionScope));
