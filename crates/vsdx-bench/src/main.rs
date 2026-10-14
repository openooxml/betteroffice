use std::env;
use std::fs;
use std::hint::black_box;
use std::time::Instant;

use betteroffice_vsdx::Diagram;
use serde::Serialize;
use vsdx_eval::{PageShapeReferences, evaluate_cell};
use vsdx_parse::{ParseLimits, write_vsdx};
use vsdx_render::Renderer;
use vsdx_resolve::{Lookup, Resolver};

const SAMPLES: usize = 5;

#[derive(Clone, Copy, Serialize)]
#[serde(rename_all = "camelCase")]
struct Timings {
    parse_ms: f64,
    resolve_ms: f64,
    evaluate_ms: f64,
    render_ms: f64,
    save_ms: f64,
}

#[derive(Serialize)]
struct FixtureResult {
    name: String,
    #[serde(flatten)]
    timings: Timings,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let fixtures = env::args()
        .skip(1)
        .map(|argument| {
            let (name, path) = argument
                .split_once('=')
                .ok_or_else(|| format!("fixture must be NAME=PATH: {argument}"))?;
            Ok((name.to_owned(), path.to_owned()))
        })
        .collect::<Result<Vec<_>, String>>()?;
    if fixtures.is_empty() {
        return Err("usage: betteroffice-vsdx-bench NAME=PATH [... ]".into());
    }

    let mut results = Vec::with_capacity(fixtures.len());
    for (name, path) in fixtures {
        let bytes = fs::read(path)?;
        run_pipeline(&bytes)?;
        let samples = (0..SAMPLES)
            .map(|_| run_pipeline(&bytes))
            .collect::<Result<Vec<_>, _>>()?;
        results.push(FixtureResult {
            name,
            timings: median(&samples),
        });
    }
    print_results(&results)?;
    Ok(())
}

fn run_pipeline(bytes: &[u8]) -> Result<Timings, Box<dyn std::error::Error>> {
    let started = Instant::now();
    let diagram =
        Diagram::open(bytes).map_err(|error| std::io::Error::other(format!("{error:?}")))?;
    let parse_ms = elapsed_ms(started);
    let package = diagram.package();

    let started = Instant::now();
    let resolver = Resolver::new(package);
    let mut resolved_shapes = 0usize;
    let mut references = Vec::with_capacity(package.page_part_paths.len());
    for page in &package.page_part_paths {
        let page_references = PageShapeReferences::new(&resolver, page)?;
        resolved_shapes += page_references.shapes().len();
        black_box(resolver.resolve_page_connectivity(page)?);
        references.push(page_references);
    }
    black_box(resolved_shapes);
    let resolve_ms = elapsed_ms(started);

    let started = Instant::now();
    let mut evaluated_cells = 0usize;
    for references in &references {
        for (shape_id, shape) in references.shapes() {
            for (name, cell) in &shape.cells {
                let Lookup::Found(cell) = cell else {
                    continue;
                };
                let Some(expression) = cell.cell.formula.as_deref().or(cell.cell.value.as_deref())
                else {
                    continue;
                };
                black_box(evaluate_cell(
                    name,
                    expression,
                    &references.for_shape(*shape_id),
                    &ParseLimits::default(),
                ));
                evaluated_cells += 1;
            }
        }
    }
    black_box(evaluated_cells);
    let evaluate_ms = elapsed_ms(started);

    let started = Instant::now();
    let renderer = Renderer::default();
    let mut primitives = 0usize;
    for page in &package.page_part_paths {
        primitives += renderer.layout_page(package, page)?.primitives.len();
    }
    black_box(primitives);
    let render_ms = elapsed_ms(started);

    let started = Instant::now();
    black_box(write_vsdx(package)?.len());
    let save_ms = elapsed_ms(started);

    Ok(Timings {
        parse_ms,
        resolve_ms,
        evaluate_ms,
        render_ms,
        save_ms,
    })
}

fn elapsed_ms(started: Instant) -> f64 {
    started.elapsed().as_secs_f64() * 1_000.0
}

fn median(samples: &[Timings]) -> Timings {
    Timings {
        parse_ms: median_by(samples, |sample| sample.parse_ms),
        resolve_ms: median_by(samples, |sample| sample.resolve_ms),
        evaluate_ms: median_by(samples, |sample| sample.evaluate_ms),
        render_ms: median_by(samples, |sample| sample.render_ms),
        save_ms: median_by(samples, |sample| sample.save_ms),
    }
}

fn median_by(samples: &[Timings], value: impl Fn(Timings) -> f64) -> f64 {
    let mut values = samples.iter().copied().map(value).collect::<Vec<_>>();
    values.sort_by(f64::total_cmp);
    values[values.len() / 2]
}

fn print_results(results: &[FixtureResult]) -> Result<(), serde_json::Error> {
    println!(
        "{}",
        serde_json::to_string(&serde_json::json!({ "samples": SAMPLES, "fixtures": results }))?
    );
    Ok(())
}
