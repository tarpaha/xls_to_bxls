extern crate calamine;

use calamine::{Reader, open_workbook_auto};

fn main() {
    let mut workbook = open_workbook_auto("data/flares.xls").expect("Cannot open file");
    let sheet_names = workbook.sheet_names().to_owned();
    for sheet in sheet_names {
        println!("{:?}", sheet);
        if let Some(Ok(range)) = workbook.worksheet_range(&sheet) {
            for row in range.rows() {
                println!("    {:?}", row);
            }
        }
    }
}
