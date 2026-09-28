---
"@betteroffice/rust-crates": minor
---

Rust API changes that need code changes when upgrading from 0.2: `docx_parse::BlockContent` variants now hold `Arc<Paragraph>`, `Arc<Table>`, `Arc<BlockSdt>` and `Arc<RawInlineXml>`, and `MediaMap` values are `Arc<MediaFile>`, so construct and match through `Arc`. `Paragraph` (`repeated_para_id`, `para_id_attribute`, `source_ordinal`), `XmlElement` (`paragraph_ordinal`, `para_id_attribute`), `Note` (`source_ordinal`), `S13SaveRequest` (`paragraph_ids`) and `S13SelectiveSave` (`source_paragraphs`) gain public fields, so struct literals need them (or `..Default::default()` where available). `betteroffice_xlsx::UpdateOrigin` gains `Recalculation`, so exhaustive matches need an arm for it. Table operation receipts gain `changed_story_ids`.
