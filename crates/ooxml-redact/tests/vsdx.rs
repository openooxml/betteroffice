#[cfg(test)]
mod tests {
    use ooxml_redact::{Format, redact};

    const FOUNDATION: &[u8] = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
    const SECRET: &str = "CONFIDENTIALCLIENTNAME";

    fn source_parts() -> Vec<(String, Vec<u8>)> {
        ooxml_opc::unzip_parts(FOUNDATION).unwrap()
    }

    fn rewrite(parts: &mut [(String, Vec<u8>)], name: &str, from: &str, to: &str) {
        let (_, bytes) = parts.iter_mut().find(|(path, _)| path == name).unwrap();
        let xml = String::from_utf8(bytes.clone()).unwrap();
        assert!(xml.contains(from));
        *bytes = xml.replace(from, to).into_bytes();
    }

    fn redact_valid(parts: &[(String, Vec<u8>)]) -> Vec<u8> {
        let source = ooxml_opc::rezip_parts(parts).unwrap();
        vsdx_parse::parse_vsdx(&source).expect("input must parse as Visio");
        redact(&source, Format::Auto).expect("redactor accepted input")
    }

    fn assert_no_secret(output: &[u8]) {
        for (path, bytes) in ooxml_opc::unzip_parts(output).unwrap() {
            assert!(
                !String::from_utf8_lossy(&bytes).contains(SECRET),
                "secret survived in {path}"
            );
        }
    }

    #[test]
    fn relocated_parts_scrub_shape_data() {
        let mut parts = source_parts();
        rewrite(
            &mut parts,
            "visio/pages/page1.xml",
            "V='15'",
            &format!("V='{SECRET}'"),
        );
        for (path, bytes) in &mut parts {
            if path.ends_with(".xml") || path.ends_with(".rels") {
                let xml = String::from_utf8(bytes.clone()).unwrap();
                *bytes = xml
                    .replace("'visio/", "'diagram/")
                    .replace("\"visio/", "\"diagram/")
                    .replace("'/visio/", "'/diagram/")
                    .replace("\"/visio/", "\"/diagram/")
                    .into_bytes();
            }
            if let Some(tail) = path.strip_prefix("visio/") {
                *path = format!("diagram/{tail}");
            }
        }
        let output = redact_valid(&parts);
        assert_no_secret(&output);
    }

    fn custom_relationship_output() -> Vec<u8> {
        let mut parts = source_parts();
        for (path, bytes) in &mut parts {
            if path.ends_with(".xml") || path.ends_with(".rels") {
                *bytes = String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace("rId", SECRET)
                    .into_bytes();
            }
        }
        redact_valid(&parts)
    }

    #[test]
    fn relationship_declarations_scrub_author_text() {
        assert_no_secret(&custom_relationship_output());
    }

    #[test]
    fn relationship_references_stay_resolvable() {
        let output = custom_relationship_output();
        vsdx_parse::parse_vsdx(&output).expect("redacted relationship references must resolve");
    }

    #[test]
    fn report_counts_renamed_relationship_ids() {
        let (_, baseline) = ooxml_redact::redact_with_report(FOUNDATION, Format::Auto).unwrap();
        let mut parts = source_parts();
        rewrite(&mut parts, "_rels/.rels", "rId1", SECRET);
        let source = ooxml_opc::rezip_parts(&parts).unwrap();
        let (output, report) = ooxml_redact::redact_with_report(&source, Format::Auto).unwrap();
        assert_eq!(report.attributes, baseline.attributes + 1);
        assert_no_secret(&output);
    }

    #[test]
    fn standard_curve_row_types_survive() {
        let mut parts = source_parts();
        rewrite(
            &mut parts,
            "visio/pages/page1.xml",
            "<Row IX='2' Del='1'/>",
            "<Row IX='2' T='RelCubBezTo'><Cell N='X' V='1'/><Cell N='Y' V='1'/><Cell N='A' V='0.2'/><Cell N='B' V='0.4'/><Cell N='C' V='0.6'/><Cell N='D' V='0.8'/></Row>",
        );
        let output = redact_valid(&parts);
        let package = vsdx_parse::parse_vsdx(&output).unwrap();
        let shape = package.page_contents["visio/pages/page1.xml"]
            .shapes()
            .next()
            .unwrap();
        let geometry = shape
            .sections()
            .find(|section| section.name == "Geometry")
            .unwrap();
        let row = geometry.rows().find(|row| row.index == Some(2)).unwrap();
        assert_eq!(row.row_type.as_deref(), Some("RelCubBezTo"));
    }

    #[test]
    fn existing_foundation_still_parses_after_redaction() {
        let output = redact_valid(&source_parts());
        vsdx_parse::parse_vsdx(&output).unwrap();
    }

    #[test]
    fn demo_geometry_and_theme_survive_both_masking_modes() {
        let source = include_bytes!("../../../apps/demo/public/betteroffice-demo.vsdx");
        let before = vsdx_parse::parse_vsdx(source).unwrap();
        let renderer = vsdx_render::Renderer::default();
        for random_characters in [false, true] {
            let output = ooxml_redact::redact_with_options(
                source,
                Format::Auto,
                &ooxml_redact::RedactionOptions { random_characters },
            )
            .unwrap();
            let after = vsdx_parse::parse_vsdx(&output).unwrap();
            assert_eq!(before.page_part_paths, after.page_part_paths);
            for (id, theme) in &before.themes {
                assert_eq!(theme.color_scheme, after.themes[id].color_scheme);
                assert_eq!(theme.font_scheme, after.themes[id].font_scheme);
                assert_ne!(theme.name, after.themes[id].name);
            }
            for page in &before.page_part_paths {
                let a = renderer.layout_page(&before, page).unwrap();
                let b = renderer.layout_page(&after, page).unwrap();
                assert_eq!((a.width, a.height), (b.width, b.height));
                fn geometry(list: vsdx_render::VsdxDisplayList) -> Vec<vsdx_render::Primitive> {
                    fn without_text(primitives: &mut Vec<vsdx_render::Primitive>) {
                        primitives.retain(|primitive| {
                            !matches!(primitive, vsdx_render::Primitive::TextBox { .. })
                        });
                        for primitive in primitives {
                            if let vsdx_render::Primitive::Group { primitives, .. } = primitive {
                                without_text(primitives);
                            }
                        }
                    }
                    let mut primitives = list.primitives;
                    without_text(&mut primitives);
                    primitives
                }
                assert_eq!(geometry(a), geometry(b), "geometry changed on {page}");
            }
            for (_, bytes) in ooxml_opc::unzip_parts(&output).unwrap() {
                let text = String::from_utf8_lossy(&bytes);
                assert!(!text.contains("Product map"));
                assert!(!text.contains("Release flow"));
            }
        }
    }

    #[test]
    fn named_rows_remain_distinct() {
        let mut parts = source_parts();
        rewrite(
            &mut parts,
            "visio/pages/page1.xml",
            "</Shape>",
            "<Section N='Property'><Row N='Alpha'><Cell N='Value' V='private'/></Row><Row N='Bravo'><Cell N='Value' V='private'/></Row></Section></Shape>",
        );
        let output = redact_valid(&parts);
        let package = vsdx_parse::parse_vsdx(&output).unwrap();
        let shape = package.page_contents["visio/pages/page1.xml"]
            .shapes()
            .next()
            .unwrap();
        let section = shape
            .sections()
            .find(|section| section.name == "Property")
            .unwrap();
        let rows = section
            .rows()
            .map(|row| row.name.as_deref().unwrap())
            .collect::<Vec<_>>();
        assert_eq!(rows.len(), 2);
        assert_ne!(rows[0], rows[1]);
        assert!(rows.iter().all(|row| !matches!(*row, "Alpha" | "Bravo")));
    }

    #[test]
    fn unsafe_geometry_formulas_use_the_cached_number() {
        let mut parts = source_parts();
        rewrite(
            &mut parts,
            "visio/pages/page1.xml",
            "</Shape>",
            "<Cell N='Width' F='User.CONFIDENTIALCLIENTNAME' V='2.5'/></Shape>",
        );
        let output = redact_valid(&parts);
        assert_no_secret(&output);
        let package = vsdx_parse::parse_vsdx(&output).unwrap();
        let shape = package.page_contents["visio/pages/page1.xml"]
            .shapes()
            .next()
            .unwrap();
        let cell = shape.cells().find(|cell| cell.name == "Width").unwrap();
        assert_eq!(cell.formula.as_deref(), Some("2.5"));
        assert_eq!(cell.value.as_deref(), Some("2.5"));
    }

    #[test]
    fn orphan_xml_and_foreign_attributes_are_redacted() {
        let mut parts = source_parts();
        parts.push((
            "extra/data.xml".into(),
            format!(
                "<root xmlns:c='urn:customer' c:identity='{SECRET}'><value>{SECRET}</value></root>"
            )
            .into_bytes(),
        ));
        rewrite(
            &mut parts,
            "visio/pages/page1.xml",
            "Mystery='yes'",
            &format!("xmlns:c='urn:customer' c:NameU='{SECRET}'"),
        );
        assert_no_secret(&redact_valid(&parts));
    }

    #[test]
    fn extension_text_in_metadata_and_package_parts_is_redacted() {
        let mut parts = source_parts();
        for path in ["[Content_Types].xml", "_rels/.rels"] {
            let (_, bytes) = parts.iter_mut().find(|(name, _)| name == path).unwrap();
            let xml = String::from_utf8(bytes.clone()).unwrap();
            let end = xml.rfind("</").unwrap();
            let mut output = xml[..end].to_owned();
            output.push_str(&format!("<extra label='{SECRET}'>{SECRET}</extra>"));
            output.push_str(&xml[end..]);
            *bytes = output.into_bytes();
        }
        for path in [
            "docProps/core.xml",
            "docProps/app.xml",
            "visio/charts/chart1.xml",
        ] {
            parts.retain(|(name, _)| name != path);
            parts.push((
                path.into(),
                format!("<root><extra>{SECRET}</extra></root>").into_bytes(),
            ));
        }
        assert_no_secret(&redact_valid(&parts));
    }

    #[test]
    fn foreign_sections_do_not_change_shapesheet_context() {
        let mut parts = source_parts();
        parts.push(("extra/data.xml".into(), format!(
            "<root xmlns:c='urn:customer'><c:Section><c:Section>{SECRET}</c:Section></c:Section></root>"
        ).into_bytes()));
        assert_no_secret(&redact_valid(&parts));
    }

    #[test]
    fn invalid_orphan_xml_and_relationships_are_refused() {
        for xml in [
            "<root>",
            "<root/><root/>",
            "private<root/>",
            "<root/>private",
            "<!DOCTYPE root><root/>",
            "<root undeclared:attr='private'/>",
            "<root xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' r:id='missing'/>",
        ] {
            let mut parts = source_parts();
            parts.push(("extra/data.xml".into(), xml.as_bytes().to_vec()));
            let source = ooxml_opc::rezip_parts(&parts).unwrap();
            vsdx_parse::parse_vsdx(&source).unwrap();
            assert!(redact(&source, Format::Auto).is_err(), "accepted {xml}");
        }
        let mut parts = source_parts();
        parts.push(("extra/_rels/data.xml.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="duplicate" Type="urn:test" Target="../visio/document.xml"/><Relationship Id="duplicate" Type="urn:test" Target="../visio/document.xml"/></Relationships>"#.to_vec()));
        let source = ooxml_opc::rezip_parts(&parts).unwrap();
        assert!(redact(&source, Format::Auto).is_err());
    }

    #[test]
    fn committed_visio_fixtures_still_parse() {
        let fixtures =
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../vsdx-parse/tests/fixtures");
        let mut count = 0;
        for entry in std::fs::read_dir(fixtures).unwrap() {
            let path = entry.unwrap().path();
            if !matches!(
                path.extension().and_then(|value| value.to_str()),
                Some("vsdx" | "vstx")
            ) {
                continue;
            }
            let source = std::fs::read(&path).unwrap();
            let output = redact(&source, Format::Auto)
                .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
            vsdx_parse::parse_vsdx(&output).unwrap();
            count += 1;
        }
        assert!(count >= 12);
    }
}
