use betteroffice_xlsx::Workbook;

#[test]
fn f7_open_comparison() {
    let Ok(dir) = std::env::var("F7_OUT") else {
        return;
    };
    let mut names: Vec<_> = std::fs::read_dir(&dir)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .filter(|path| path.extension().is_some_and(|ext| ext == "xlsx"))
        .collect();
    names.sort();
    for path in names {
        let bytes = std::fs::read(&path).unwrap();
        let open = match Workbook::open(&bytes) {
            Ok(_) => "ok".to_owned(),
            Err(error) => format!("error({error})"),
        };
        let collaborative = match Workbook::open_collaborative(&bytes, 1) {
            Ok(_) => "ok".to_owned(),
            Err(error) => format!("error({error})"),
        };
        println!(
            "F7FILE {} open={open} collaborative={collaborative}",
            path.file_name().unwrap().to_string_lossy()
        );
    }
}
