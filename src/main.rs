extern crate calamine;

use calamine::{Reader, open_workbook_auto, DataType};

#[derive(Debug)]
struct Row {
    cells: Vec<String>
}

#[derive(Debug)]
struct Sheet {
    name: String,
    rows: Vec<Row>
}

#[derive(Debug)]
struct Workbook {
    sheets: Vec<Sheet>
}

impl Workbook {
    pub fn from_xls(filename: &str) -> Workbook {
        let mut workbook = Workbook { sheets: Vec::new() };
        let mut xls_workbook = open_workbook_auto(filename).expect("Cannot open file");
        let sheet_names = xls_workbook.sheet_names().to_owned();
        for sheet_name in sheet_names {
            if let Some(Ok(xls_sheet)) = xls_workbook.worksheet_range(&sheet_name) {
                let mut sheet = Sheet {
                    name: sheet_name,
                    rows:  Vec::new()
                };
                for xls_row in xls_sheet.rows() {
                    let mut row = Row {
                        cells: Vec::new()
                    };
                    for xls_cell in xls_row {
                        let s = match xls_cell {
                            DataType::String(s) => s.clone(),
                            DataType::Int(i) => i.to_string(),
                            DataType::Float(f) => f.to_string(),
                            DataType::Bool(b) => b.to_string(),
                            _ => String::from("")
                        };
                        row.cells.push(s);
                    }
                    sheet.rows.push(row);
                }
                workbook.sheets.push(sheet);
            }
        };
        workbook
    }
}

fn main() {
    let workbook = Workbook::from_xls("data/flares.xls");
    println!("{:?}", workbook);
}
