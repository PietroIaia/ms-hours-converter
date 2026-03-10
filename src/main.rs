use std::fs;
use std::error::Error;
use chrono::NaiveDate;
use regex::Regex;
use rust_xlsxwriter::Workbook;

fn is_valid_date(s: &str, format: &str) -> bool {
    NaiveDate::parse_from_str(s, format).is_ok()
}

fn parse_hours(s: String) -> String {
    let parts: Vec<&str> = s.split(':').collect();
    let mut result_parts: Vec<String> = Vec::new();
    let mut part_to_be_parsed;
    // If duration is in minutes, add 0 as hour
    if parts.len() == 1 && parts[0].ends_with("min") {
        result_parts.insert(0, String::from("0"));
        part_to_be_parsed = parts[0].to_string();
    }
    // If duration is in hours, add hour
    else {
        result_parts.push(parts[0].to_string());
        part_to_be_parsed = parts[1].to_string();
    }
    // parse minutes
    part_to_be_parsed.truncate(part_to_be_parsed.len() - 3);
    let mut mins_parsed: f64 = part_to_be_parsed.parse().expect("Oh No!");
    mins_parsed = mins_parsed / 60.0 * 100.0;
    // Add minutes to result and return the value joined by "."
    result_parts.push(mins_parsed.to_string());

    return result_parts.join(".");
}    

fn write_xlsx(desktop_path: String, rows: Vec<Vec<String>>) -> Result<(), Box<dyn Error>> {
    let xlsx_path = format!("{}\\{}", desktop_path, "\\hours.xlsx");

    // Create a new workbook
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write Headers
    worksheet.write_string(0, 0, "SUMMARY")?;
    worksheet.write_string(0, 1, "HOURS")?;
    worksheet.write_string(0, 2, "DATE")?;
    worksheet.write_string(0, 3, "PROJECT")?;

    // Write every task
    for (row_idx, row) in rows.iter().enumerate() {
        let excel_row = (row_idx + 1) as u32;
        for (col_idx, value) in row.iter().enumerate() {
            worksheet.write_string(excel_row, col_idx as u16, value)?;
        }
    }

    // Save the workbook
    workbook.save(&xlsx_path)?;

    Ok(())
}

fn main() -> Result<(), Box<dyn Error>> {
    // Get Desktop full path
    let mut desktop_path = String::new();
    match dirs::home_dir() {
        Some(path) => desktop_path = path.display().to_string() + "\\Desktop",
        None => println!("Oh No!")
    }
    // Get contents from hours.txt file present in Desktop
    let content = fs::read_to_string(format!("{}\\{}", desktop_path, "\\hours.txt")).expect("Oh No!");

    // Read from file
    let mut day = String::new();
    let re = Regex::new(r"\s+|\([^)]*\)").unwrap();
    let mut task_list = vec![];
    for line in content.lines() {
        // if empty line, jump to next
        if line == "" {continue;}
        // if line is date, save it for future tasks
        if is_valid_date(line, "%Y-%m-%d") {
            day = line.to_string();
            continue;
        }
        // parse + build task row
        let parts: Vec<&str> = line.split(": ").collect();
        let project = String::from(parts[0]);
        let description = String::from(parts[1]);
        let hours = parse_hours(re.replace_all(parts[2], "").to_string());
        // Add task to list of tasks to write into CSV 
        task_list.push(vec![description, hours, day.clone(), project]);
    }
    // Write tasks into XLSX
    write_xlsx(desktop_path, task_list)?;
 
    Ok(())
}
