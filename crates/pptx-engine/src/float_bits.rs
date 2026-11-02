use std::collections::BTreeMap;

pub(crate) trait FloatBitsWalk {
    fn float_bits(&self, out: &mut Vec<u64>);
}

pub(crate) fn float_bits(value: &pptx_edit::SlideSnapshot) -> Vec<u64> {
    let mut bits = Vec::new();
    value.float_bits(&mut bits);
    bits
}

impl FloatBitsWalk for f64 {
    fn float_bits(&self, out: &mut Vec<u64>) {
        out.push(self.to_bits());
    }
}

impl<T: FloatBitsWalk> FloatBitsWalk for Option<T> {
    fn float_bits(&self, out: &mut Vec<u64>) {
        if let Some(value) = self {
            value.float_bits(out);
        }
    }
}

impl<T: FloatBitsWalk> FloatBitsWalk for Vec<T> {
    fn float_bits(&self, out: &mut Vec<u64>) {
        for value in self {
            value.float_bits(out);
        }
    }
}

impl<T: FloatBitsWalk> FloatBitsWalk for Box<T> {
    fn float_bits(&self, out: &mut Vec<u64>) {
        self.as_ref().float_bits(out);
    }
}

impl<K: FloatBitsWalk, V: FloatBitsWalk> FloatBitsWalk for BTreeMap<K, V> {
    fn float_bits(&self, out: &mut Vec<u64>) {
        for (key, value) in self {
            key.float_bits(out);
            value.float_bits(out);
        }
    }
}

impl<T: FloatBitsWalk, const N: usize> FloatBitsWalk for [T; N] {
    fn float_bits(&self, out: &mut Vec<u64>) {
        for value in self {
            value.float_bits(out);
        }
    }
}

impl FloatBitsWalk for String {
    fn float_bits(&self, _out: &mut Vec<u64>) {}
}

impl FloatBitsWalk for bool {
    fn float_bits(&self, _out: &mut Vec<u64>) {}
}

impl FloatBitsWalk for u8 {
    fn float_bits(&self, _out: &mut Vec<u64>) {}
}

impl FloatBitsWalk for u32 {
    fn float_bits(&self, _out: &mut Vec<u64>) {}
}

impl FloatBitsWalk for i32 {
    fn float_bits(&self, _out: &mut Vec<u64>) {}
}

impl FloatBitsWalk for i64 {
    fn float_bits(&self, _out: &mut Vec<u64>) {}
}

impl FloatBitsWalk for usize {
    fn float_bits(&self, _out: &mut Vec<u64>) {}
}

impl FloatBitsWalk for pptx_edit::SlideSnapshot {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            header_footer,
            title,
            hidden,
            layout,
            id,
            source_part_path,
            layout_part_path,
            name,
            notes,
            background,
            background_edited,
            shapes,
        } = self;
        header_footer.float_bits(out);
        title.float_bits(out);
        hidden.float_bits(out);
        layout.float_bits(out);
        id.float_bits(out);
        source_part_path.float_bits(out);
        layout_part_path.float_bits(out);
        name.float_bits(out);
        notes.float_bits(out);
        background.float_bits(out);
        background_edited.float_bits(out);
        shapes.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::ShapeSnapshot {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            id,
            source_id,
            kind,
            name,
            title,
            description,
            decorative,
            x,
            y,
            width,
            height,
            rotation_deg,
            flip_h,
            flip_v,
            inherited,
            hidden,
            text_box,
            aspect_ratio_locked,
            geometry,
            adjust_values,
            placeholder,
            fill,
            resolved_fill_color,
            outline,
            resolved_outline_color,
            media_part_path,
            pending_media,
            blip_effects,
            table_structure,
            graphic,
            text_stories,
            children,
        } = self;
        id.float_bits(out);
        source_id.float_bits(out);
        kind.float_bits(out);
        name.float_bits(out);
        title.float_bits(out);
        description.float_bits(out);
        decorative.float_bits(out);
        x.float_bits(out);
        y.float_bits(out);
        width.float_bits(out);
        height.float_bits(out);
        rotation_deg.float_bits(out);
        flip_h.float_bits(out);
        flip_v.float_bits(out);
        inherited.float_bits(out);
        hidden.float_bits(out);
        text_box.float_bits(out);
        aspect_ratio_locked.float_bits(out);
        geometry.float_bits(out);
        adjust_values.float_bits(out);
        placeholder.float_bits(out);
        fill.float_bits(out);
        resolved_fill_color.float_bits(out);
        outline.float_bits(out);
        resolved_outline_color.float_bits(out);
        media_part_path.float_bits(out);
        pending_media.float_bits(out);
        blip_effects.float_bits(out);
        table_structure.float_bits(out);
        graphic.float_bits(out);
        text_stories.float_bits(out);
        children.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::InheritedGeometry {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            x,
            y,
            width,
            height,
            rotation_deg,
            flip_h,
            flip_v,
        } = self;
        x.float_bits(out);
        y.float_bits(out);
        width.float_bits(out);
        height.float_bits(out);
        rotation_deg.float_bits(out);
        flip_h.float_bits(out);
        flip_v.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::StorySnapshot {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            id,
            length,
            paragraphs,
        } = self;
        id.float_bits(out);
        length.float_bits(out);
        paragraphs.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::ParagraphSnapshot {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            id,
            source_paragraphs,
            preserved_xml,
            alignment,
            level,
            bullet_json,
            runs,
        } = self;
        id.float_bits(out);
        source_paragraphs.float_bits(out);
        preserved_xml.float_bits(out);
        alignment.float_bits(out);
        level.float_bits(out);
        bullet_json.float_bits(out);
        runs.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::TextRunSnapshot {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self { text, style } = self;
        text.float_bits(out);
        style.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::TextStyle {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            bold,
            italic,
            font_size_pt,
            color,
            font_family,
            underline,
            spacing_pt,
            baseline_pct,
            kern_pt,
            caps,
        } = self;
        bold.float_bits(out);
        italic.float_bits(out);
        font_size_pt.float_bits(out);
        color.float_bits(out);
        font_family.float_bits(out);
        underline.float_bits(out);
        spacing_pt.float_bits(out);
        baseline_pct.float_bits(out);
        kern_pt.float_bits(out);
        caps.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::SlideHeaderFooter {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            footer,
            slide_number,
            date,
        } = self;
        footer.float_bits(out);
        slide_number.float_bits(out);
        date.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::SlideHeaderFooterText {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self { shape_id, text } = self;
        shape_id.float_bits(out);
        text.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::SlideDateText {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            shape_id,
            text,
            field,
        } = self;
        shape_id.float_bits(out);
        text.float_bits(out);
        field.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::SlideLayoutSnapshot {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            header_footer,
            name,
            r#type,
        } = self;
        header_footer.float_bits(out);
        name.float_bits(out);
        r#type.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::PendingMedia {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            content_type,
            base64,
        } = self;
        content_type.float_bits(out);
        base64.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::ParagraphSource {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self { id, xml } = self;
        id.float_bits(out);
        xml.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::HeaderFooterSettings {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            slide_number,
            footer,
            date,
            header,
        } = self;
        slide_number.float_bits(out);
        footer.float_bits(out);
        date.float_bits(out);
        header.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::Placeholder {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            placeholder_type,
            index,
            orientation,
            size,
        } = self;
        placeholder_type.float_bits(out);
        index.float_bits(out);
        orientation.float_bits(out);
        size.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::TableStructure {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            generation,
            story_keys,
            grid,
            columns,
            copied_columns,
            rows,
        } = self;
        generation.float_bits(out);
        story_keys.float_bits(out);
        grid.float_bits(out);
        columns.float_bits(out);
        copied_columns.float_bits(out);
        rows.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::TableRowSource {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            source_row,
            copied,
            copied_cells,
            height,
            cells,
        } = self;
        source_row.float_bits(out);
        copied.float_bits(out);
        copied_cells.float_bits(out);
        height.float_bits(out);
        cells.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::Table {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            grid,
            properties,
            rows,
        } = self;
        grid.float_bits(out);
        properties.float_bits(out);
        rows.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::TableProperties {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            first_row,
            last_row,
            first_col,
            last_col,
            band_row,
            band_col,
            style_id,
        } = self;
        first_row.float_bits(out);
        last_row.float_bits(out);
        first_col.float_bits(out);
        last_col.float_bits(out);
        band_row.float_bits(out);
        band_col.float_bits(out);
        style_id.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::TableRow {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self { height, cells } = self;
        height.float_bits(out);
        cells.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::TableCell {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            text,
            grid_span,
            row_span,
            merged,
            fill,
            picture_fill,
            borders,
        } = self;
        text.float_bits(out);
        grid_span.float_bits(out);
        row_span.float_bits(out);
        merged.float_bits(out);
        fill.float_bits(out);
        picture_fill.float_bits(out);
        borders.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::TableCellBorders {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            left,
            top,
            right,
            bottom,
        } = self;
        left.float_bits(out);
        top.float_bits(out);
        right.float_bits(out);
        bottom.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::TextBody {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            anchor,
            vertical,
            num_col,
            spc_col,
            rtl_col,
            compat_line_spacing,
            space_first_last_para,
            autofit,
            wrap,
            vertical_overflow,
            horizontal_overflow,
            inset_left,
            inset_top,
            inset_right,
            inset_bottom,
            list_style,
            default_list_style,
            paragraphs,
        } = self;
        anchor.float_bits(out);
        vertical.float_bits(out);
        num_col.float_bits(out);
        spc_col.float_bits(out);
        rtl_col.float_bits(out);
        compat_line_spacing.float_bits(out);
        space_first_last_para.float_bits(out);
        autofit.float_bits(out);
        wrap.float_bits(out);
        vertical_overflow.float_bits(out);
        horizontal_overflow.float_bits(out);
        inset_left.float_bits(out);
        inset_top.float_bits(out);
        inset_right.float_bits(out);
        inset_bottom.float_bits(out);
        list_style.float_bits(out);
        default_list_style.float_bits(out);
        paragraphs.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::TextParagraph {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            properties,
            runs,
            end_properties,
        } = self;
        properties.float_bits(out);
        runs.float_bits(out);
        end_properties.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::ParagraphProperties {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            alignment,
            level,
            margin_left,
            margin_right,
            indent,
            bullet,
            line_spacing,
            space_before,
            space_after,
            bullet_font,
            bullet_color,
            bullet_size,
            default_tab_size,
            tab_stops,
            rtl,
            default_run,
        } = self;
        alignment.float_bits(out);
        level.float_bits(out);
        margin_left.float_bits(out);
        margin_right.float_bits(out);
        indent.float_bits(out);
        bullet.float_bits(out);
        line_spacing.float_bits(out);
        space_before.float_bits(out);
        space_after.float_bits(out);
        bullet_font.float_bits(out);
        bullet_color.float_bits(out);
        bullet_size.float_bits(out);
        default_tab_size.float_bits(out);
        tab_stops.float_bits(out);
        rtl.float_bits(out);
        default_run.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::RunProperties {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            font_size_pt,
            spacing_pt,
            baseline_pct,
            kern_pt,
            bold,
            italic,
            underline,
            caps,
            font_family,
            color,
            language,
            hyperlink_relationship_id,
            effects,
        } = self;
        font_size_pt.float_bits(out);
        spacing_pt.float_bits(out);
        baseline_pct.float_bits(out);
        kern_pt.float_bits(out);
        bold.float_bits(out);
        italic.float_bits(out);
        underline.float_bits(out);
        caps.float_bits(out);
        font_family.float_bits(out);
        color.float_bits(out);
        language.float_bits(out);
        hyperlink_relationship_id.float_bits(out);
        effects.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::TextRun {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            text,
            properties,
            field_id,
            field_type,
            line_break,
        } = self;
        text.float_bits(out);
        properties.float_bits(out);
        field_id.float_bits(out);
        field_type.float_bits(out);
        line_break.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::Picture {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            base,
            aspect_ratio_locked,
            relationship_id,
            media_part_path,
            crop,
            effects,
            geometry,
            paths,
            adjust_values,
            fill,
            outline,
            shape_effects,
            style,
        } = self;
        base.float_bits(out);
        aspect_ratio_locked.float_bits(out);
        relationship_id.float_bits(out);
        media_part_path.float_bits(out);
        crop.float_bits(out);
        effects.float_bits(out);
        geometry.float_bits(out);
        paths.float_bits(out);
        adjust_values.float_bits(out);
        fill.float_bits(out);
        outline.float_bits(out);
        shape_effects.float_bits(out);
        style.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::PictureCrop {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            left,
            top,
            right,
            bottom,
        } = self;
        left.float_bits(out);
        top.float_bits(out);
        right.float_bits(out);
        bottom.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::PictureFill {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            relationship_id,
            media_part_path,
            crop,
            fill_rect,
            tile,
            effects,
        } = self;
        relationship_id.float_bits(out);
        media_part_path.float_bits(out);
        crop.float_bits(out);
        fill_rect.float_bits(out);
        tile.float_bits(out);
        effects.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::PictureTile {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self { scale_x, scale_y } = self;
        scale_x.float_bits(out);
        scale_y.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::ShapeBase {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            id,
            name,
            title,
            description,
            decorative,
            hidden,
            placeholder,
            transform,
        } = self;
        id.float_bits(out);
        name.float_bits(out);
        title.float_bits(out);
        description.float_bits(out);
        decorative.float_bits(out);
        hidden.float_bits(out);
        placeholder.float_bits(out);
        transform.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::ShapeTransform {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            x,
            y,
            width,
            height,
            rotation_deg,
            flip_h,
            flip_v,
            child_x,
            child_y,
            child_width,
            child_height,
        } = self;
        x.float_bits(out);
        y.float_bits(out);
        width.float_bits(out);
        height.float_bits(out);
        rotation_deg.float_bits(out);
        flip_h.float_bits(out);
        flip_v.float_bits(out);
        child_x.float_bits(out);
        child_y.float_bits(out);
        child_width.float_bits(out);
        child_height.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_parse::CustomGeometryPath {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            commands,
            no_fill,
            no_stroke,
        } = self;
        commands.float_bits(out);
        no_fill.float_bits(out);
        no_stroke.float_bits(out);
    }
}

impl FloatBitsWalk for ooxml_drawingml::ShapeFill {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            fill_type,
            color,
            gradient,
        } = self;
        fill_type.float_bits(out);
        color.float_bits(out);
        gradient.float_bits(out);
    }
}

impl FloatBitsWalk for ooxml_drawingml::GradientFill {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            gradient_type,
            angle,
            stops,
        } = self;
        gradient_type.float_bits(out);
        angle.float_bits(out);
        stops.float_bits(out);
    }
}

impl FloatBitsWalk for ooxml_drawingml::GradientStop {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self { position, color } = self;
        position.float_bits(out);
        color.float_bits(out);
    }
}

impl FloatBitsWalk for ooxml_drawingml::ShapeOutline {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            width,
            color,
            gradient,
            style,
            cap,
            join,
            head_end,
            tail_end,
        } = self;
        width.float_bits(out);
        color.float_bits(out);
        gradient.float_bits(out);
        style.float_bits(out);
        cap.float_bits(out);
        join.float_bits(out);
        head_end.float_bits(out);
        tail_end.float_bits(out);
    }
}

impl FloatBitsWalk for ooxml_drawingml::LineEnd {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            end_type,
            width,
            length,
        } = self;
        end_type.float_bits(out);
        width.float_bits(out);
        length.float_bits(out);
    }
}

impl FloatBitsWalk for ooxml_drawingml::ColorValue {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            rgb,
            theme_color,
            theme_tint,
            theme_shade,
            auto,
            transforms,
            luminance_modulation,
            luminance_offset,
            saturation_modulation,
            alpha,
        } = self;
        rgb.float_bits(out);
        theme_color.float_bits(out);
        theme_tint.float_bits(out);
        theme_shade.float_bits(out);
        auto.float_bits(out);
        transforms.float_bits(out);
        luminance_modulation.float_bits(out);
        luminance_offset.float_bits(out);
        saturation_modulation.float_bits(out);
        alpha.float_bits(out);
    }
}

impl FloatBitsWalk for ooxml_drawingml::ShapeEffects {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self { outer_shadow } = self;
        outer_shadow.float_bits(out);
    }
}

impl FloatBitsWalk for ooxml_drawingml::OuterShadow {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            color,
            blur_radius,
            distance,
            direction,
            rotate_with_shape,
            scale_x,
            scale_y,
            alignment,
        } = self;
        color.float_bits(out);
        blur_radius.float_bits(out);
        distance.float_bits(out);
        direction.float_bits(out);
        rotate_with_shape.float_bits(out);
        if *scale_x != 1.0 {
            scale_x.float_bits(out);
        }
        if *scale_y != 1.0 {
            scale_y.float_bits(out);
        }
        alignment.float_bits(out);
    }
}

impl FloatBitsWalk for ooxml_drawingml::ShapeStyle {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self {
            font_color,
            fill,
            line,
            effect,
            fill_disabled,
            line_disabled,
        } = self;
        font_color.float_bits(out);
        fill.float_bits(out);
        line.float_bits(out);
        effect.float_bits(out);
        fill_disabled.float_bits(out);
        line_disabled.float_bits(out);
    }
}

impl FloatBitsWalk for ooxml_drawingml::StyleReference {
    fn float_bits(&self, out: &mut Vec<u64>) {
        let Self { index, color } = self;
        index.float_bits(out);
        color.float_bits(out);
    }
}

impl FloatBitsWalk for pptx_edit::ShapeKind {
    fn float_bits(&self, _out: &mut Vec<u64>) {
        match self {
            Self::Shape | Self::Picture | Self::GraphicFrame | Self::Group => {}
        }
    }
}

impl FloatBitsWalk for pptx_parse::TextCaps {
    fn float_bits(&self, _out: &mut Vec<u64>) {
        match self {
            Self::None | Self::Small | Self::All => {}
        }
    }
}

impl FloatBitsWalk for pptx_parse::TextOverflow {
    fn float_bits(&self, _out: &mut Vec<u64>) {
        match self {
            Self::Overflow | Self::Clip | Self::Ellipsis => {}
        }
    }
}

impl FloatBitsWalk for pptx_parse::BulletFont {
    fn float_bits(&self, out: &mut Vec<u64>) {
        match self {
            Self::FollowText => {}
            Self::Typeface(value) => value.float_bits(out),
        }
    }
}

impl FloatBitsWalk for pptx_parse::BulletColor {
    fn float_bits(&self, out: &mut Vec<u64>) {
        match self {
            Self::FollowText => {}
            Self::Color(value) => value.float_bits(out),
        }
    }
}

impl FloatBitsWalk for pptx_parse::BulletSize {
    fn float_bits(&self, out: &mut Vec<u64>) {
        match self {
            Self::FollowText => {}
            Self::Percent(value) | Self::Points(value) => value.float_bits(out),
        }
    }
}

impl FloatBitsWalk for pptx_parse::Bullet {
    fn float_bits(&self, out: &mut Vec<u64>) {
        match self {
            Self::Character { value } => value.float_bits(out),
            Self::AutoNumber {
                scheme,
                start_at,
                restart,
            } => {
                scheme.float_bits(out);
                start_at.float_bits(out);
                restart.float_bits(out);
            }
            Self::None => {}
        }
    }
}

impl FloatBitsWalk for pptx_parse::TextAutofit {
    fn float_bits(&self, out: &mut Vec<u64>) {
        match self {
            Self::None | Self::Shape => {}
            Self::Normal {
                font_scale,
                line_space_reduction,
            } => {
                font_scale.float_bits(out);
                line_space_reduction.float_bits(out);
            }
        }
    }
}

impl FloatBitsWalk for pptx_parse::LineSpacing {
    fn float_bits(&self, out: &mut Vec<u64>) {
        match self {
            Self::Percent { value } | Self::Points { value } => value.float_bits(out),
        }
    }
}

impl FloatBitsWalk for pptx_parse::GraphicFrameData {
    fn float_bits(&self, out: &mut Vec<u64>) {
        match self {
            Self::Table(table) => table.float_bits(out),
            Self::Chart {
                relationship_id,
                part_path,
            } => {
                relationship_id.float_bits(out);
                part_path.float_bits(out);
            }
            Self::Diagram {
                relationship_ids,
                drawing_part_path,
            } => {
                relationship_ids.float_bits(out);
                drawing_part_path.float_bits(out);
            }
            Self::Unknown { uri, picture } => {
                uri.float_bits(out);
                picture.float_bits(out);
            }
        }
    }
}

impl FloatBitsWalk for pptx_parse::BlipEffect {
    fn float_bits(&self, out: &mut Vec<u64>) {
        match self {
            Self::BiLevel { threshold } => threshold.float_bits(out),
            Self::Grayscale => {}
            Self::Alpha { amount } => amount.float_bits(out),
            Self::Luminance {
                brightness,
                contrast,
            } => {
                brightness.float_bits(out);
                contrast.float_bits(out);
            }
            Self::Duotone { shadow, highlight } => {
                shadow.float_bits(out);
                highlight.float_bits(out);
            }
            Self::ColorChange {
                from,
                to,
                use_alpha,
            } => {
                from.float_bits(out);
                to.float_bits(out);
                use_alpha.float_bits(out);
            }
        }
    }
}

impl FloatBitsWalk for ooxml_drawingml::GeometryPathCommand {
    fn float_bits(&self, out: &mut Vec<u64>) {
        match self {
            Self::Move { x, y } | Self::Line { x, y } => {
                x.float_bits(out);
                y.float_bits(out);
            }
            Self::Quad { cpx, cpy, x, y } => {
                cpx.float_bits(out);
                cpy.float_bits(out);
                x.float_bits(out);
                y.float_bits(out);
            }
            Self::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                cp1x.float_bits(out);
                cp1y.float_bits(out);
                cp2x.float_bits(out);
                cp2y.float_bits(out);
                x.float_bits(out);
                y.float_bits(out);
            }
            Self::Close => {}
        }
    }
}

impl FloatBitsWalk for ooxml_drawingml::ColorTransform {
    fn float_bits(&self, out: &mut Vec<u64>) {
        match self {
            Self::Tint(value)
            | Self::Shade(value)
            | Self::Hue(value)
            | Self::HueMod(value)
            | Self::HueOff(value)
            | Self::Sat(value)
            | Self::SatMod(value)
            | Self::SatOff(value)
            | Self::Lum(value)
            | Self::LumMod(value)
            | Self::LumOff(value)
            | Self::Red(value)
            | Self::RedMod(value)
            | Self::RedOff(value)
            | Self::Green(value)
            | Self::GreenMod(value)
            | Self::GreenOff(value)
            | Self::Blue(value)
            | Self::BlueMod(value)
            | Self::BlueOff(value)
            | Self::Alpha(value)
            | Self::AlphaMod(value)
            | Self::AlphaOff(value) => value.float_bits(out),
            Self::Comp | Self::Inv | Self::Gray | Self::Gamma | Self::InvGamma => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use ooxml_drawingml::{
        ColorTransform, ColorValue, GeometryPathCommand, GradientFill, GradientStop, OuterShadow,
        ShapeEffects, ShapeFill, ShapeOutline, ShapeStyle, StyleReference,
    };
    use pptx_parse::{
        BlipEffect, BulletColor, BulletSize, GraphicFrameData, LineSpacing, ParagraphProperties,
        RunProperties, Table, TableCell, TableCellBorders, TableRow, TextAutofit, TextBody,
        TextParagraph, TextRun,
    };
    use serde::{Serialize, ser};

    const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

    fn assert_oracle(value: &(impl FloatBitsWalk + Serialize)) {
        let mut typed = Vec::new();
        value.float_bits(&mut typed);
        assert_eq!(typed, serde_float_bits(value).unwrap());
    }

    fn assert_snapshot_pairs(slides: &[pptx_edit::SlideSnapshot]) {
        let typed: Vec<_> = slides.iter().map(float_bits).collect();
        let oracle: Vec<_> = slides
            .iter()
            .map(|slide| serde_float_bits(slide).unwrap())
            .collect();
        assert_eq!(typed, oracle);
        for (a, (typed_a, oracle_a)) in typed.iter().zip(&oracle).enumerate() {
            for (b, (typed_b, oracle_b)) in typed.iter().zip(&oracle).enumerate() {
                assert_eq!(typed_a == typed_b, oracle_a == oracle_b, "{a}, {b}");
            }
        }
    }

    fn color(value: f64) -> ColorValue {
        let transforms: [fn(f64) -> ColorTransform; 23] = [
            ColorTransform::Tint,
            ColorTransform::Shade,
            ColorTransform::Hue,
            ColorTransform::HueMod,
            ColorTransform::HueOff,
            ColorTransform::Sat,
            ColorTransform::SatMod,
            ColorTransform::SatOff,
            ColorTransform::Lum,
            ColorTransform::LumMod,
            ColorTransform::LumOff,
            ColorTransform::Red,
            ColorTransform::RedMod,
            ColorTransform::RedOff,
            ColorTransform::Green,
            ColorTransform::GreenMod,
            ColorTransform::GreenOff,
            ColorTransform::Blue,
            ColorTransform::BlueMod,
            ColorTransform::BlueOff,
            ColorTransform::Alpha,
            ColorTransform::AlphaMod,
            ColorTransform::AlphaOff,
        ];
        ColorValue {
            transforms: transforms
                .into_iter()
                .enumerate()
                .map(|(index, transform)| transform(index as f64 + 0.125))
                .chain([
                    ColorTransform::Comp,
                    ColorTransform::Inv,
                    ColorTransform::Gray,
                    ColorTransform::Gamma,
                    ColorTransform::InvGamma,
                    ColorTransform::Alpha(value),
                ])
                .collect(),
            luminance_modulation: Some(value),
            luminance_offset: Some(0.25),
            saturation_modulation: Some(0.5),
            alpha: Some(0.75),
            ..ColorValue::default()
        }
    }

    fn fill(value: f64) -> ShapeFill {
        ShapeFill {
            fill_type: "gradient".to_owned(),
            color: Some(color(value)),
            gradient: Some(GradientFill {
                gradient_type: "linear".to_owned(),
                angle: Some(value),
                stops: vec![
                    GradientStop {
                        position: 0.125,
                        color: color(0.25),
                    },
                    GradientStop {
                        position: value,
                        color: color(0.75),
                    },
                ],
            }),
        }
    }

    fn outline(value: f64) -> ShapeOutline {
        ShapeOutline {
            width: Some(value),
            color: Some(color(0.375)),
            gradient: fill(0.625).gradient,
            ..ShapeOutline::default()
        }
    }

    fn effects(value: f64) -> ShapeEffects {
        ShapeEffects {
            outer_shadow: Some(OuterShadow {
                color: Some(color(0.875)),
                scale_x: value,
                scale_y: 1.0,
                ..OuterShadow::default()
            }),
        }
    }

    fn blip_effects(value: f64) -> Vec<BlipEffect> {
        vec![
            BlipEffect::BiLevel { threshold: value },
            BlipEffect::Grayscale,
            BlipEffect::Alpha { amount: 0.125 },
            BlipEffect::Luminance {
                brightness: 0.25,
                contrast: 0.375,
            },
            BlipEffect::Duotone {
                shadow: Some(color(0.5)),
                highlight: Some(color(0.625)),
            },
            BlipEffect::ColorChange {
                from: Some(color(0.75)),
                to: Some(color(0.875)),
                use_alpha: false,
            },
        ]
    }

    fn text_body(value: f64) -> TextBody {
        let run = RunProperties {
            font_size_pt: Some(value),
            spacing_pt: Some(1.25),
            baseline_pct: Some(2.5),
            kern_pt: Some(3.75),
            color: Some(color(0.125)),
            effects: Some(effects(0.25)),
            ..RunProperties::default()
        };
        let paragraph = ParagraphProperties {
            line_spacing: Some(LineSpacing::Percent { value: 1.125 }),
            space_before: Some(LineSpacing::Points { value: 2.25 }),
            space_after: Some(LineSpacing::Points { value }),
            bullet_color: Some(BulletColor::Color(color(0.375))),
            bullet_size: Some(BulletSize::Points(4.5)),
            default_run: Some(run.clone()),
            ..ParagraphProperties::default()
        };
        TextBody {
            autofit: Some(TextAutofit::Normal {
                font_scale: Some(value),
                line_space_reduction: Some(0.5),
            }),
            list_style: vec![paragraph.clone()],
            default_list_style: Some(Box::new(paragraph.clone())),
            paragraphs: vec![TextParagraph {
                properties: paragraph,
                runs: vec![TextRun {
                    text: "Table".to_owned(),
                    properties: run.clone(),
                    field_id: None,
                    field_type: None,
                    line_break: false,
                }],
                end_properties: Some(run),
            }],
            ..TextBody::default()
        }
    }

    fn table(value: f64) -> Table {
        Table {
            grid: vec![100, 200],
            properties: Default::default(),
            rows: vec![TableRow {
                height: 100,
                cells: vec![TableCell {
                    text: text_body(value),
                    fill: Some(fill(0.125)),
                    picture_fill: Some(pptx_parse::PictureFill {
                        relationship_id: None,
                        media_part_path: None,
                        crop: Default::default(),
                        fill_rect: Default::default(),
                        tile: Some(pptx_parse::PictureTile {
                            scale_x: value,
                            scale_y: 0.25,
                        }),
                        effects: blip_effects(0.375),
                    }),
                    borders: TableCellBorders {
                        left: Some(outline(0.5)),
                        top: Some(outline(0.625)),
                        right: Some(outline(0.75)),
                        bottom: Some(outline(0.875)),
                    },
                    ..TableCell::default()
                }],
            }],
        }
    }

    fn picture(value: f64) -> pptx_parse::Picture {
        pptx_parse::Picture {
            base: pptx_parse::ShapeBase {
                id: 1,
                name: "Picture".to_owned(),
                title: None,
                description: None,
                decorative: false,
                hidden: false,
                placeholder: None,
                transform: pptx_parse::ShapeTransform {
                    rotation_deg: value,
                    ..Default::default()
                },
            },
            aspect_ratio_locked: false,
            relationship_id: None,
            media_part_path: None,
            crop: Default::default(),
            effects: blip_effects(value),
            geometry: "custom".to_owned(),
            paths: vec![pptx_parse::CustomGeometryPath {
                commands: vec![
                    GeometryPathCommand::Move { x: value, y: 0.125 },
                    GeometryPathCommand::Line { x: 0.25, y: 0.375 },
                    GeometryPathCommand::Quad {
                        cpx: 0.5,
                        cpy: 0.625,
                        x: 0.75,
                        y: 0.875,
                    },
                    GeometryPathCommand::Cubic {
                        cp1x: 1.125,
                        cp1y: 1.25,
                        cp2x: 1.375,
                        cp2y: 1.5,
                        x: 1.625,
                        y: 1.75,
                    },
                    GeometryPathCommand::Close,
                ],
                no_fill: false,
                no_stroke: false,
            }],
            adjust_values: BTreeMap::from([("z".to_owned(), value), ("a".to_owned(), 0.25)]),
            fill: Some(fill(0.375)),
            outline: Some(outline(0.5)),
            shape_effects: Some(effects(value)),
            style: Some(Box::new(ShapeStyle {
                font_color: Some(color(0.625)),
                fill: Some(StyleReference {
                    index: 1,
                    color: Some(color(0.75)),
                }),
                line: Some(StyleReference {
                    index: 2,
                    color: Some(color(0.875)),
                }),
                effect: Some(StyleReference {
                    index: 3,
                    color: Some(color(value)),
                }),
                ..ShapeStyle::default()
            })),
        }
    }

    fn enriched_slide(source: &pptx_edit::SlideSnapshot, value: f64) -> pptx_edit::SlideSnapshot {
        let mut slide = source.clone();
        slide.background = Some(fill(value));
        let shape = &mut slide.shapes[0];
        shape.rotation_deg = value;
        shape.inherited = Some(pptx_edit::InheritedGeometry {
            rotation_deg: value,
            ..Default::default()
        });
        shape.adjust_values = BTreeMap::from([("z".to_owned(), value), ("a".to_owned(), 0.25)]);
        shape.fill = Some(fill(0.375));
        shape.outline = Some(outline(value));
        shape.blip_effects = blip_effects(0.625);
        shape.text_stories = vec![pptx_edit::StorySnapshot {
            id: "story".to_owned(),
            length: 4,
            paragraphs: vec![pptx_edit::ParagraphSnapshot {
                runs: vec![pptx_edit::TextRunSnapshot {
                    text: "Text".to_owned(),
                    style: pptx_edit::TextStyle {
                        font_size_pt: Some(value),
                        spacing_pt: Some(value),
                        baseline_pct: Some(2.5),
                        kern_pt: Some(3.75),
                        ..Default::default()
                    },
                }],
                ..Default::default()
            }],
        }];
        let mut child = shape.clone();
        child.children.clear();
        child.graphic = Some(GraphicFrameData::Table(table(value)));
        shape.children = vec![child];
        shape.graphic = Some(GraphicFrameData::Unknown {
            uri: None,
            picture: Some(Box::new(picture(value))),
        });
        slide
    }

    #[test]
    fn demo_snapshot_float_bits_match_serde_oracle() {
        let session = pptx_edit::DeckSession::open(DECK, 930).unwrap();
        let mut slides = session.snapshot().unwrap().slides;
        let source = slides[0].clone();
        for value in [
            0.0,
            -0.0,
            12.5,
            f64::from_bits(0x7ff8_0000_0000_0001),
            f64::from_bits(0x7ff8_0000_0000_0002),
            f64::INFINITY,
            f64::NEG_INFINITY,
        ] {
            let slide = enriched_slide(&source, value);
            slides.push(slide.clone());
            slides.push(slide);
        }
        let mut absent = enriched_slide(&source, 0.0);
        absent.background = None;
        absent.shapes[0].inherited = None;
        absent.shapes[0].adjust_values.clear();
        absent.shapes[0].outline = None;
        absent.shapes[0].children.clear();
        absent.shapes[0].text_stories[0].paragraphs[0].runs[0]
            .style
            .spacing_pt = None;
        slides.push(absent);
        assert_snapshot_pairs(&slides);
    }

    #[test]
    fn serde_skipped_shadow_scales_and_table_rows_match() {
        assert_oracle(&OuterShadow::default());
        for scale in [0.0, -0.0, 1.0, 2.5, f64::NAN] {
            for (scale_x, scale_y) in [(scale, 1.0), (1.0, scale), (scale, scale)] {
                assert_oracle(&OuterShadow {
                    scale_x,
                    scale_y,
                    ..OuterShadow::default()
                });
            }
        }
        let legacy = Table {
            grid: Vec::new(),
            properties: Default::default(),
            rows: vec![TableRow {
                height: 0,
                cells: vec![TableCell::from_text(text_body(-0.0))],
            }],
        };
        let mut modern = legacy.clone();
        modern.rows[0].height = 100;
        assert_oracle(&legacy);
        assert_oracle(&modern);
        assert_eq!(serde_float_bits(&legacy), serde_float_bits(&modern));
    }

    #[test]
    fn optional_fields_and_collection_boundaries_match_serde_classes() {
        let session = pptx_edit::DeckSession::open(DECK, 931).unwrap();
        let source = session.snapshot().unwrap().slides.remove(0);
        let mut first = enriched_slide(&source, 0.0);
        let style = &mut first.shapes[0].text_stories[0].paragraphs[0].runs[0].style;
        style.font_size_pt = Some(-0.0);
        style.spacing_pt = None;
        let mut second = first.clone();
        let style = &mut second.shapes[0].text_stories[0].paragraphs[0].runs[0].style;
        style.font_size_pt = None;
        style.spacing_pt = Some(-0.0);
        assert_ne!(first, second);
        assert_eq!(float_bits(&first), float_bits(&second));
        let mut third = first.clone();
        let child = third.shapes[0].children.remove(0);
        third.shapes.insert(1, child);
        let mut fourth = first.clone();
        fourth.shapes[0].adjust_values =
            BTreeMap::from([("b".to_owned(), 0.25), ("y".to_owned(), 0.0)]);
        assert_ne!(first, third);
        assert_ne!(first, fourth);
        assert_eq!(float_bits(&first), float_bits(&third));
        assert_eq!(float_bits(&first), float_bits(&fourth));
        assert_snapshot_pairs(&[first, second, third, fourth]);
    }

    #[derive(Default)]
    struct FloatBits(Vec<u64>);

    fn serde_float_bits(value: &impl Serialize) -> Option<Vec<u64>> {
        let mut bits = FloatBits::default();
        value.serialize(&mut bits).ok()?;
        Some(bits.0)
    }

    macro_rules! ignore_primitives {
        ($($method:ident($ty:ty)),* $(,)?) => {
            $(fn $method(self, _: $ty) -> Result<(), Self::Error> { Ok(()) })*
        };
    }

    impl ser::Serializer for &mut FloatBits {
        type Ok = ();
        type Error = serde::de::value::Error;
        type SerializeSeq = Self;
        type SerializeTuple = Self;
        type SerializeTupleStruct = Self;
        type SerializeTupleVariant = Self;
        type SerializeMap = Self;
        type SerializeStruct = Self;
        type SerializeStructVariant = Self;

        ignore_primitives! {
            serialize_bool(bool),
            serialize_i8(i8), serialize_i16(i16), serialize_i32(i32),
            serialize_i64(i64), serialize_i128(i128),
            serialize_u8(u8), serialize_u16(u16), serialize_u32(u32),
            serialize_u64(u64), serialize_u128(u128),
            serialize_char(char), serialize_str(&str), serialize_bytes(&[u8]),
        }

        fn serialize_f32(self, value: f32) -> Result<(), Self::Error> {
            self.0.push(u64::from(value.to_bits()));
            Ok(())
        }

        fn serialize_f64(self, value: f64) -> Result<(), Self::Error> {
            self.0.push(value.to_bits());
            Ok(())
        }

        fn serialize_none(self) -> Result<(), Self::Error> {
            Ok(())
        }

        fn serialize_some<T: ?Sized + Serialize>(self, value: &T) -> Result<(), Self::Error> {
            value.serialize(self)
        }

        fn serialize_unit(self) -> Result<(), Self::Error> {
            Ok(())
        }

        fn serialize_unit_struct(self, _: &'static str) -> Result<(), Self::Error> {
            Ok(())
        }

        fn serialize_unit_variant(
            self,
            _: &'static str,
            _: u32,
            _: &'static str,
        ) -> Result<(), Self::Error> {
            Ok(())
        }

        fn serialize_newtype_struct<T: ?Sized + Serialize>(
            self,
            _: &'static str,
            value: &T,
        ) -> Result<(), Self::Error> {
            value.serialize(self)
        }

        fn serialize_newtype_variant<T: ?Sized + Serialize>(
            self,
            _: &'static str,
            _: u32,
            _: &'static str,
            value: &T,
        ) -> Result<(), Self::Error> {
            value.serialize(self)
        }

        fn serialize_seq(self, _: Option<usize>) -> Result<Self, Self::Error> {
            Ok(self)
        }

        fn serialize_tuple(self, _: usize) -> Result<Self, Self::Error> {
            Ok(self)
        }

        fn serialize_tuple_struct(self, _: &'static str, _: usize) -> Result<Self, Self::Error> {
            Ok(self)
        }

        fn serialize_tuple_variant(
            self,
            _: &'static str,
            _: u32,
            _: &'static str,
            _: usize,
        ) -> Result<Self, Self::Error> {
            Ok(self)
        }

        fn serialize_map(self, _: Option<usize>) -> Result<Self, Self::Error> {
            Ok(self)
        }

        fn serialize_struct(self, _: &'static str, _: usize) -> Result<Self, Self::Error> {
            Ok(self)
        }

        fn serialize_struct_variant(
            self,
            _: &'static str,
            _: u32,
            _: &'static str,
            _: usize,
        ) -> Result<Self, Self::Error> {
            Ok(self)
        }

        fn collect_str<T: ?Sized + std::fmt::Display>(self, _: &T) -> Result<(), Self::Error> {
            Ok(())
        }
    }

    macro_rules! float_sequence {
        ($trait:ident, $method:ident) => {
            impl ser::$trait for &mut FloatBits {
                type Ok = ();
                type Error = serde::de::value::Error;

                fn $method<T: ?Sized + Serialize>(&mut self, value: &T) -> Result<(), Self::Error> {
                    value.serialize(&mut **self)
                }

                fn end(self) -> Result<(), Self::Error> {
                    Ok(())
                }
            }
        };
    }

    float_sequence!(SerializeSeq, serialize_element);
    float_sequence!(SerializeTuple, serialize_element);
    float_sequence!(SerializeTupleStruct, serialize_field);
    float_sequence!(SerializeTupleVariant, serialize_field);

    impl ser::SerializeMap for &mut FloatBits {
        type Ok = ();
        type Error = serde::de::value::Error;

        fn serialize_key<T: ?Sized + Serialize>(&mut self, key: &T) -> Result<(), Self::Error> {
            key.serialize(&mut **self)
        }

        fn serialize_value<T: ?Sized + Serialize>(&mut self, value: &T) -> Result<(), Self::Error> {
            value.serialize(&mut **self)
        }

        fn end(self) -> Result<(), Self::Error> {
            Ok(())
        }
    }

    macro_rules! float_struct {
        ($trait:ident) => {
            impl ser::$trait for &mut FloatBits {
                type Ok = ();
                type Error = serde::de::value::Error;

                fn serialize_field<T: ?Sized + Serialize>(
                    &mut self,
                    _: &'static str,
                    value: &T,
                ) -> Result<(), Self::Error> {
                    value.serialize(&mut **self)
                }

                fn end(self) -> Result<(), Self::Error> {
                    Ok(())
                }
            }
        };
    }

    float_struct!(SerializeStruct);
    float_struct!(SerializeStructVariant);
}
