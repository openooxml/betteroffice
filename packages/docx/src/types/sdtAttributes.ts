import type { SdtProperties } from './document';

/** Project structured content-control properties onto the flat yrs attribute vocabulary. */
export function sdtPropsToAttrs(props: SdtProperties): Record<string, unknown> {
  return {
    sdtType: props.sdtType,
    id: props.id ?? null,
    alias: props.alias ?? null,
    tag: props.tag ?? null,
    lock: props.lock ?? null,
    placeholder: props.placeholder ?? null,
    showingPlaceholder: props.showingPlaceholder ?? false,
    dateFormat: props.dateFormat ?? null,
    listItems: props.listItems ? JSON.stringify(props.listItems) : null,
    checked: props.checked ?? null,
    dataBinding: props.dataBinding ? JSON.stringify(props.dataBinding) : null,
    rawPropertiesXml: props.rawPropertiesXml ?? null,
    rawEndPropertiesXml: props.rawEndPropertiesXml ?? null,
  };
}

/** Rebuild structured content-control properties from the flat yrs vocabulary. */
export function sdtAttrsToProps(attrs: Record<string, unknown>): SdtProperties {
  const props: SdtProperties = {
    sdtType: (attrs.sdtType as SdtProperties['sdtType']) ?? 'richText',
  };
  if (typeof attrs.id === 'number') props.id = attrs.id;
  if (attrs.alias != null) props.alias = String(attrs.alias);
  if (attrs.tag != null) props.tag = String(attrs.tag);
  if (attrs.lock != null) props.lock = attrs.lock as SdtProperties['lock'];
  if (attrs.placeholder != null) props.placeholder = String(attrs.placeholder);
  if (attrs.showingPlaceholder) props.showingPlaceholder = true;
  if (attrs.dateFormat != null) props.dateFormat = String(attrs.dateFormat);
  if (typeof attrs.listItems === 'string' && attrs.listItems) {
    try {
      props.listItems = JSON.parse(attrs.listItems) as SdtProperties['listItems'];
    } catch {
      // The raw property XML remains available for lossless serialization.
    }
  }
  if (attrs.checked != null) props.checked = attrs.checked as boolean;
  if (typeof attrs.dataBinding === 'string' && attrs.dataBinding) {
    try {
      props.dataBinding = JSON.parse(attrs.dataBinding) as SdtProperties['dataBinding'];
    } catch {
      // The raw property XML remains available for lossless serialization.
    }
  }
  if (attrs.rawPropertiesXml != null) props.rawPropertiesXml = String(attrs.rawPropertiesXml);
  if (attrs.rawEndPropertiesXml != null) {
    props.rawEndPropertiesXml = String(attrs.rawEndPropertiesXml);
  }
  return props;
}
