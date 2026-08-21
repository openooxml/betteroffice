pub fn numeric_product(source: &str, name: &str) -> u64 {
    let prefix = format!("export const {name} =");
    source
        .lines()
        .find_map(|line| line.trim().strip_prefix(&prefix))
        .unwrap_or_else(|| panic!("missing {name}"))
        .trim()
        .trim_end_matches(';')
        .split('*')
        .map(|part| part.trim().parse::<u64>().unwrap())
        .product()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn limit_scan_ignores_references_before_the_export() {
        let source = "const derived = LIMIT = 7;\nexport const LIMIT = 16 * 1024;\n";

        assert_eq!(numeric_product(source, "LIMIT"), 16 * 1024);
    }
}
