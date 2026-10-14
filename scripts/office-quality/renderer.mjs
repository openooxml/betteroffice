export function qualityRendererPlugin(format) {
  if (!['pptx', 'xlsx', 'vsdx'].includes(format))
    throw new Error('Invalid Office capture format');
  const id = 'virtual:office-quality-renderer';
  const resolved = `\0${id}`;
  return {
    name: 'office-quality-renderer',
    resolveId(source) {
      if (source === id) return resolved;
    },
    load(source) {
      if (source === resolved) return `export * from '@betteroffice/${format}';`;
    },
  };
}
