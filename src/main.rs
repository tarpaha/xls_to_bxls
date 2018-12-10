extern crate byteorder;
extern crate calamine;

use std::fs::File;
use std::io::Write;
use byteorder::{LittleEndian, WriteBytesExt};
use calamine::{Reader, open_workbook_auto, DataType};

#[derive(Debug)]
struct Row {
    cells: Vec<String>
}

#[derive(Debug)]
struct Sheet {
    name: String,
    max_cells_in_a_row: usize,
    rows: Vec<Row>
}

#[derive(Debug)]
struct Workbook {
    sheets: Vec<Sheet>
}

impl Workbook {
    fn from_xls(filename: &str) -> Workbook {
        let mut workbook = Workbook { sheets: Vec::new() };
        let mut xls_workbook = open_workbook_auto(filename).unwrap();
        let mut sheet_names = xls_workbook.sheet_names().to_owned();
        sheet_names.sort();
        for sheet_name in sheet_names {
            if let Some(Ok(xls_sheet)) = xls_workbook.worksheet_range(&sheet_name) {
                let mut rows = Vec::new();
                let mut max_cells_in_a_row = 0;
                for xls_row in xls_sheet.rows() {
                    let mut row = Row {
                        cells: Vec::new()
                    };
                    let mut cells_in_a_row = 0;
                    for xls_cell in xls_row {
                        let s = match xls_cell {
                            DataType::String(s) => s.clone(),
                            DataType::Int(i) => i.to_string().trim_left_matches('0').to_string(),
                            DataType::Float(f) => f.to_string().trim_left_matches('0').to_string(),
                            DataType::Bool(b) => b.to_string(),
                            _ => String::from("")
                        };
                        row.cells.push(s);
                        cells_in_a_row += 1;
                    }
                    rows.push(row);
                    if cells_in_a_row > max_cells_in_a_row {
                        max_cells_in_a_row = cells_in_a_row;
                    }
                }
                for row in &mut rows {
                    for _ in row.cells.len()..max_cells_in_a_row {
                        row.cells.push("".to_string());
                    }
                }
                workbook.sheets.push(Sheet {
                    name: sheet_name,
                    max_cells_in_a_row,
                    rows
                });
            }
        };
        workbook
    }
}

impl Workbook {
    fn write_string(&self, mut file: &File, str: &str) {
        if str.len() > 127 {
            panic!("String {} length greater than 127 chars", str);
        }
        file.write_u8(str.len() as u8).unwrap();
        file.write(str.as_bytes()).unwrap();
    }
    fn to_file(&self, filename: &str) {
        let mut file = File::create(filename).unwrap();
        file.write_u32::<LittleEndian>(0x534c5842).unwrap(); // "BXLS"
        file.write_u32::<LittleEndian>(self.sheets.len() as u32).unwrap();
        for sheet in &self.sheets {
            self.write_string(&file, &sheet.name);
            file.write_u32::<LittleEndian>(sheet.rows.len() as u32).unwrap();
            file.write_u32::<LittleEndian>(sheet.max_cells_in_a_row as u32).unwrap();
            for row in &sheet.rows {
                for cell in &row.cells {
                    self.write_string(&file, &cell);
                }
            }
        }
    }
}

use std::time::Instant;

fn main() {
    let now = Instant::now();
    let workbook = Workbook::from_xls("data/flares.xls");
    workbook.to_file("data/flares.bxls");
    let elapsed = now.elapsed();
    println!("Done in {} seconds", (elapsed.as_secs() as f64) + (elapsed.subsec_nanos() as f64 / 1000_000_000.0));
}
