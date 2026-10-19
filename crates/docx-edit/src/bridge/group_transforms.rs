use serde_json::{Value, json};

#[derive(Clone, Copy)]
struct Matrix {
    a: f64,
    b: f64,
    c: f64,
    d: f64,
    x: f64,
    y: f64,
}

impl Matrix {
    const IDENTITY: Self = Self {
        a: 1.0,
        b: 0.0,
        c: 0.0,
        d: 1.0,
        x: 0.0,
        y: 0.0,
    };

    fn multiply(self, other: Self) -> Self {
        Self {
            a: self.a * other.a + self.c * other.b,
            b: self.b * other.a + self.d * other.b,
            c: self.a * other.c + self.c * other.d,
            d: self.b * other.c + self.d * other.d,
            x: self.a * other.x + self.c * other.y + self.x,
            y: self.b * other.x + self.d * other.y + self.y,
        }
    }

    fn translate(x: f64, y: f64) -> Self {
        Self {
            x,
            y,
            ..Self::IDENTITY
        }
    }

    fn scale(a: f64, d: f64) -> Self {
        Self {
            a,
            d,
            ..Self::IDENTITY
        }
    }

    fn orientation(rotation: f64, flip_h: bool, flip_v: bool) -> Self {
        let (sin, cos) = rotation.to_radians().sin_cos();
        Self {
            a: cos,
            b: sin,
            c: -sin,
            d: cos,
            ..Self::IDENTITY
        }
        .multiply(Self::scale(
            if flip_h { -1.0 } else { 1.0 },
            if flip_v { -1.0 } else { 1.0 },
        ))
    }

    fn unapply(self, x: f64, y: f64) -> Option<(f64, f64)> {
        let determinant = self.a * self.d - self.b * self.c;
        if !determinant.is_finite() || determinant.abs() < f64::MIN_POSITIVE {
            return None;
        }
        Some((
            (self.d * (x - self.x) - self.c * (y - self.y)) / determinant,
            (self.a * (y - self.y) - self.b * (x - self.x)) / determinant,
        ))
    }

    fn scale_x(self) -> f64 {
        self.a.hypot(self.b).max(f64::MIN_POSITIVE)
    }

    fn scale_y(self) -> f64 {
        self.c.hypot(self.d).max(f64::MIN_POSITIVE)
    }

    fn flipped(self) -> bool {
        self.a * self.d - self.b * self.c < 0.0
    }
}

fn number(value: &Value, key: &str, default: f64) -> f64 {
    value
        .get(key)
        .and_then(Value::as_f64)
        .filter(|value| value.is_finite())
        .unwrap_or(default)
}

pub(super) fn project(shape: &mut Value) {
    let Some(root) = shape.pointer("/scene/root").cloned() else {
        return;
    };
    if !matches!(root["kind"].as_str(), Some("group" | "canvas")) {
        return;
    }
    project_node(shape, &root, Matrix::IDENTITY, (0.0, 0.0), true, 0);
}

fn project_node(
    shape: &mut Value,
    scene: &Value,
    parent: Matrix,
    parent_origin: (f64, f64),
    root: bool,
    depth: usize,
) {
    if depth > 32 {
        return;
    }
    let group = matches!(scene["kind"].as_str(), Some("group" | "canvas"));
    let transform = &shape["transform"];
    let rotation = number(transform, "rotation", 0.0) - parent.b.atan2(parent.a).to_degrees();
    let orientation = Matrix::orientation(
        rotation,
        transform["flipH"].as_bool().unwrap_or(false) != parent.flipped(),
        transform["flipV"].as_bool().unwrap_or(false),
    );
    let axes = parent.multiply(orientation);
    let scene_transform = &scene["transform"];
    let x = if root {
        number(scene_transform, "offsetX", 0.0)
    } else {
        number(&shape["offset"], "x", 0.0)
    };
    let y = if root {
        number(scene_transform, "offsetY", 0.0)
    } else {
        number(&shape["offset"], "y", 0.0)
    };
    let width = if root {
        number(
            scene_transform,
            "width",
            number(
                scene_transform,
                "childWidth",
                number(&shape["size"], "width", 0.0),
            ),
        )
    } else {
        number(&shape["size"], "width", 0.0)
    };
    let height = if root {
        number(
            scene_transform,
            "height",
            number(
                scene_transform,
                "childHeight",
                number(&shape["size"], "height", 0.0),
            ),
        )
    } else {
        number(&shape["size"], "height", 0.0)
    };
    let local_width = width
        / if group {
            parent.scale_x()
        } else {
            axes.scale_x()
        };
    let local_height = height
        / if group {
            parent.scale_y()
        } else {
            axes.scale_y()
        };
    let center = if group {
        (
            x + parent.a * local_width / 2.0 + parent.c * local_height / 2.0,
            y + parent.b * local_width / 2.0 + parent.d * local_height / 2.0,
        )
    } else {
        (
            x + axes.a * local_width / 2.0 + axes.c * local_height / 2.0,
            y + axes.b * local_width / 2.0 + axes.d * local_height / 2.0,
        )
    };
    let rendered_width = local_width * axes.scale_x();
    let rendered_height = local_height * axes.scale_y();
    let origin = (
        center.0 - rendered_width / 2.0,
        center.1 - rendered_height / 2.0,
    );
    if ![origin.0, origin.1, rendered_width, rendered_height]
        .iter()
        .all(|value| value.is_finite())
    {
        return;
    }
    if !root {
        shape["offset"] = json!({"x": origin.0 - parent_origin.0, "y": origin.1 - parent_origin.1});
        shape["size"] = json!({"width": rendered_width, "height": rendered_height});
        let sign = if axes.flipped() { -1.0 } else { 1.0 };
        shape["transform"] = json!({"rotation": (sign * axes.b).atan2(sign * axes.a).to_degrees()});
        if axes.flipped() {
            shape["transform"]["flipH"] = json!(true);
        }
    }
    if !group {
        return;
    }
    let Some((local_x, local_y)) = parent.unapply(x, y) else {
        return;
    };
    let child_width = number(scene_transform, "childWidth", local_width);
    let child_height = number(scene_transform, "childHeight", local_height);
    let nonzero = |value: f64| if value.abs() < 1e-12 { 1.0 } else { value };
    let children_matrix = parent
        .multiply(Matrix::translate(
            local_x + local_width / 2.0,
            local_y + local_height / 2.0,
        ))
        .multiply(orientation)
        .multiply(Matrix::translate(-local_width / 2.0, -local_height / 2.0))
        .multiply(Matrix::scale(
            local_width / nonzero(child_width),
            local_height / nonzero(child_height),
        ))
        .multiply(Matrix::translate(
            -number(scene_transform, "childOffsetX", 0.0),
            -number(scene_transform, "childOffsetY", 0.0),
        ));
    let Some(children) = shape.get_mut("children").and_then(Value::as_array_mut) else {
        return;
    };
    let scenes = scene["children"]
        .as_array()
        .into_iter()
        .flatten()
        .filter(|node| matches!(node["kind"].as_str(), Some("group" | "canvas" | "shape")));
    for (child, scene) in children.iter_mut().zip(scenes) {
        project_node(
            child,
            scene,
            children_matrix,
            if root { (0.0, 0.0) } else { origin },
            false,
            depth + 1,
        );
    }
}
