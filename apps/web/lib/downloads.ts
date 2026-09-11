export function officialPackageNames(packageAccess: Record<string, string>): string[] {
  return Object.keys(packageAccess).filter((name) => name.startsWith("@betteroffice/"));
}
