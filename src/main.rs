use csv::ReaderBuilder;
use encoding_rs::WINDOWS_1252;
use rust_xlsxwriter::Workbook;
use std::error::Error;
use std::fs::File;
use std::io::{BufRead, BufReader};
use std::path::PathBuf;

fn main() -> Result<(), Box<dyn Error>> {
    let args: Vec<String> = std::env::args().collect();
    let input_path: PathBuf = args.get(1).expect("No Path provided.").into();
    let mut output_path: String = input_path
        .file_name()
        .expect("No Filename provided")
        .to_str()
        .unwrap()
        .to_string();
    output_path.push_str(".xlsx");

    let rows = parse_and_filter(&input_path)?;
    write_to_excel(&rows, &output_path)?;

    println!("Excel-Datei erstellt: {}", output_path);
    Ok(())
}

fn parse_and_filter(path: &PathBuf) -> Result<Vec<(String, f64, f64)>, Box<dyn Error>> {
    let bytes = std::fs::read(path)?;
    let (content, _, _) = encoding_rs::WINDOWS_1252.decode(&bytes);
    let content = content.into_owned();

    let mut result_rows: Vec<(String, f64, f64)> = Vec::new();
    let mut in_data_section = false;

    let mut reader = ReaderBuilder::new()
        .delimiter(b';')
        .has_headers(true)
        .from_reader(content.as_bytes());

    for record_result in reader.records() {
        let record = record_result?;

        if record.is_empty() {
            continue;
        }

        let first = clean_field(record.get(0).unwrap_or(""));

        // Start when we reach the first real sample row
        if !in_data_section {
            if first.eq_ignore_ascii_case("High") {
                in_data_section = true;
            } else {
                continue;
            }
        }

        let sample_id = first;
        let ammonium = extract_numeric_field(record.get(9).unwrap_or(""));
        let nitrat = extract_numeric_field(record.get(12).unwrap_or(""));
        println!("{:?}", ammonium);
        println!("{:?}", nitrat);

        if let (Some(a), Some(n)) = (ammonium, nitrat) {
            result_rows.push((sample_id.clone(), a, n));
        }

        // Stop after MATRIX inclusive
        if sample_id.eq_ignore_ascii_case("MATRIX") {
            break;
        }
    }

    Ok(result_rows)
}

fn clean_field(s: &str) -> String {
    s.replace("&quot;", "").replace('"', "").trim().to_string()
}

fn extract_numeric_field(s: &str) -> Option<f64> {
    let cleaned = clean_field(s).replace(',', ".");
    cleaned.parse::<f64>().ok()
}

fn write_to_excel(rows: &[(String, f64, f64)], output_path: &str) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Header
    worksheet.write_string(0, 0, "Sample ID")?;
    worksheet.write_string(0, 1, "Ammonium-N mg/L")?;
    worksheet.write_string(0, 2, "Nitrat-N mg/L")?;

    // Data
    for (i, (sample_id, ammonium_n, nitrat_n)) in rows.iter().enumerate() {
        println!("SAMPLE_ID: {sample_id}, AMMONIUM_N: {ammonium_n}, NITRAT_N: {nitrat_n}");

        let row = (i + 1) as u32;
        worksheet.write_string(row, 0, sample_id)?;
        worksheet.write_number(row, 1, *ammonium_n)?;
        worksheet.write_number(row, 2, *nitrat_n)?;
    }

    workbook.save(output_path)?;
    Ok(())
}
