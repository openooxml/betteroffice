import type { SdtProperties } from '../types/document';

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** Build raw `w:sdtPr` for a programmatically created content control. */
export function synthesizeSdtPr(props: SdtProperties): string {
  const parts: string[] = [];
  if (props.alias) parts.push(`<w:alias w:val="${escapeXml(props.alias)}"/>`);
  if (props.tag) parts.push(`<w:tag w:val="${escapeXml(props.tag)}"/>`);
  if (props.id != null) parts.push(`<w:id w:val="${props.id}"/>`);
  if (props.lock && props.lock !== 'unlocked') parts.push(`<w:lock w:val="${props.lock}"/>`);
  if (props.placeholder) {
    parts.push(`<w:placeholder><w:docPart w:val="${escapeXml(props.placeholder)}"/></w:placeholder>`);
  }
  if (props.showingPlaceholder) parts.push('<w:showingPlcHdr/>');

  switch (props.sdtType) {
    case 'plainText':
      parts.push('<w:text/>');
      break;
    case 'date':
      parts.push(
        props.dateFormat
          ? `<w:date><w:dateFormat w:val="${escapeXml(props.dateFormat)}"/></w:date>`
          : '<w:date/>'
      );
      break;
    case 'dropDownList':
    case 'comboBox': {
      const items = (props.listItems ?? [])
        .map(
          (item) =>
            `<w:listItem w:displayText="${escapeXml(item.displayText)}" w:value="${escapeXml(item.value)}"/>`
        )
        .join('');
      parts.push(`<w:${props.sdtType} w:lastValue="">${items}</w:${props.sdtType}>`);
      break;
    }
    case 'checkbox':
      parts.push(
        `<w14:checkbox><w14:checked w14:val="${props.checked ? '1' : '0'}"/>` +
          '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
          '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>'
      );
      break;
    case 'picture':
      parts.push('<w:picture/>');
      break;
  }
  return `<w:sdtPr>${parts.join('')}</w:sdtPr>`;
}
