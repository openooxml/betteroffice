interface ExportFace {
  id: number;
  family: string;
  bold?: boolean;
  italic?: boolean;
  bytes: Uint8Array;
}

/** Gives every face of `family` these bytes, whatever its case, except the faces `keep` holds on to. */
export function upgradeExportFaces(
  faces: Iterable<ExportFace>, family: string, bytes: Uint8Array, keep: (face: ExportFace) => boolean = () => false,
): void {
  const key = family.trim().toLowerCase();
  for (const face of faces) {
    if (face.family.trim().toLowerCase() === key && !keep(face)) face.bytes = bytes;
  }
}
