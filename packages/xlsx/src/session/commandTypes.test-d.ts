import type {
  CommandBasis, CommandIdentity, CommandStatus, WorkbookAck, WorkbookCommand,
  WorkbookCommandFrame, WorkbookCommandSelection, WorkbookEditState, WorkbookSession,
} from '../index';
import type { WorkbookSharedAck } from './commandProtocol';

type Equal<A, B> = (<T>() => T extends A ? 1 : 2) extends (<T>() => T extends B ? 1 : 2) ? true : false;
type Assert<T extends true> = T;

export type CommandSwapCompatibility = [
  Assert<Equal<CommandIdentity, { session: string; incarnation: string }>>,
  Assert<Equal<CommandBasis, { epoch: number; version: string }>>,
  Assert<Equal<WorkbookCommand['incarnation'], string>>,
  Assert<Equal<WorkbookCommand['basis']['version'], string>>,
  Assert<Equal<CommandStatus['status'], 'applied' | 'noop' | 'refused' | 'cancelled'>>,
  Assert<Equal<CommandStatus['result'], unknown>>,
  Assert<Equal<CommandStatus<{ formatToken: string }>['result'], { formatToken: string } | undefined>>,
  Assert<Equal<WorkbookAck['version'], string>>,
  Assert<Equal<WorkbookAck['statuses'], CommandStatus[]>>,
  Assert<Equal<WorkbookAck['selection'], WorkbookCommandSelection | undefined>>,
  Assert<Equal<WorkbookAck['editState'], Partial<WorkbookEditState> | undefined>>,
  Assert<Equal<WorkbookAck['frame'], WorkbookCommandFrame | undefined>>,
  Assert<Equal<NonNullable<WorkbookSharedAck['compatibility']>['frame'], WorkbookCommandFrame | undefined>>,
  Assert<Equal<WorkbookSession['submit'], (command: WorkbookCommand) => Promise<CommandStatus>>>,
  Assert<Equal<WorkbookSession['call']['command'], (command: WorkbookCommand) => Promise<WorkbookAck>>>,
  Assert<Equal<WorkbookSession['onAck'], (listener: (ack: WorkbookAck) => void) => () => void>>,
  Assert<Equal<WorkbookSession['markPresented'], (frame: WorkbookCommandFrame |
    Pick<WorkbookCommandFrame, 'session' | 'incarnation' | 'coversThroughSeq'>) => boolean>>,
  Assert<Equal<WorkbookSession['whenPresented'], (seq: number) => Promise<void>>>,
  Assert<Equal<WorkbookSession['save'], () => Promise<Uint8Array>>>,
  Assert<Equal<WorkbookSession['editState'], Readonly<Partial<WorkbookEditState>>>>,
  Assert<Equal<WorkbookSession['selection'], WorkbookCommandSelection | undefined>>,
  Assert<Equal<Extract<WorkbookCommand, { kind: 'select' }>['payload'], { selection: WorkbookCommandSelection }>>,
  Assert<Equal<Extract<WorkbookCommand, { kind: 'setReadOnly' }>['payload'], { readOnly: boolean }>>,
];
