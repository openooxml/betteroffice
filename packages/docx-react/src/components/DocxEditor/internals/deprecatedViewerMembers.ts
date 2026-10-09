const warned = new Set<string>();

export function warnDeprecatedViewerMember(member: string, behaviour: string, twin: string): void {
  if (warned.has(member)) return;
  warned.add(member);
  console.warn(`[DocxEditor] ${member} is deprecated; ${behaviour}. Use ${twin}.`);
}

export function resetDeprecatedViewerMembersForTests(): void {
  warned.clear();
}
