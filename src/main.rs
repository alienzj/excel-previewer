extern crate calamine;
extern crate madato;
extern crate serde_derive;

use calamine::{open_workbook_auto, DataType, Reader};
use madato::types::{ErroredTable, NamedTable, RenderOptions, TableRow, KVFilter};

fn read_excel_to_named_tables(
    filename: String,
    sheet_name: Option<String>,
) -> Vec<Result<NamedTable<String, String>, ErroredTable>> {
    // opens a new workbook
    let mut workbook = open_workbook_auto(filename).expect("Cannot open file");

    let sheet_names = if let Some(sheet_name) = sheet_name {
        workbook
            .sheet_names()
            .to_owned()
            .into_iter()
            .filter(|n| *n == sheet_name)
            .to_owned()
            .collect::<Vec<_>>()
    } else {
        workbook.sheet_names().to_owned()
    };

    // println!("{:?}", sheet_names);

    let sheets: Vec<Result<NamedTable<String, String>, ErroredTable>> = sheet_names
        .iter()
        .map(|name| {
            let maybe_sheet = workbook.worksheet_range(name);
            match maybe_sheet {
                None => Err((name.clone(), format!("sheet {} is empty", name))),
                Some(Err(err)) => Err((name.clone(), format!("{}", err))),
                Some(Ok(sheet)) => Ok((name.clone(), {
                    let first_row: Vec<(usize, String)> = sheet
                        .rows()
                        .next()
                        .expect("Missing data in the sheet")
                        .iter()
                        .enumerate()
                        .map(|(i, c)| match c {
                            DataType::Empty => (i, format!("NULL{}", i)),
                            _ => (i, c.to_string()),
                        })
                        .collect();

                    sheet
                        .rows()
                        .skip(1)
                        .map(|row| {
                            first_row
                                .iter()
                                .map(|(i, col)| ((**col).to_string(), md_santise(&row[*i])))
                                .collect::<TableRow<String, String>>()
                        })
                        .collect::<Vec<_>>()
                })),
            }
        })
        .collect::<Vec<_>>();

    sheets
}

fn list_sheet_names(filename: String) -> Result<Vec<String>, String> {
    let workbook = open_workbook_auto(filename).expect("Could not open the file");
    Ok(workbook.sheet_names().to_owned())
}

fn md_santise(data: &DataType) -> String {
    data.to_string()
        .replace("|", "\\|")
        .replace("\r\r", "<br/>")
        .replace("\n", "<br/>")
        .replace("\r", "<br/>")
}

fn get_sheet_names(filename: String) {
    for s in list_sheet_names(filename).unwrap() {
        println!("{}", s);
    }
}

fn spreadsheet_to_md(
    filename: String,
    render_options: &Option<RenderOptions>,
) -> Result<String, String> {
    let results =
        read_excel_to_named_tables(filename, render_options.clone().and_then(|r| r.sheet_name));
 
    if results.len() == 1 {
        Ok(madato::named_table_to_md(
            results[0].clone(),
            false,
            render_options,
        ))
    } else if results.len() > 1 {
        Ok(results
            .iter()
            .map(|table_result| {
                madato::named_table_to_md(table_result.clone(), true, &render_options.clone())
            })
            .collect::<Vec<String>>()
            .join("\n\n"))
    } else {
        Err(String::from("Sorry, no results"))
    }
}

fn main() -> Result<(), String> {
    let filename = String::from("/home/alienzj/projects/excel-previewer/data/TabS7-Oral_antiSMASH.xlsx");
    let render_options = Some(RenderOptions {
        headings: None,
        sheet_name: None,
        filters: None,
    });

    let output_string = spreadsheet_to_md(filename, &render_options);

    match output_string {
        Ok(markdown) => {
            println!("{}", markdown);
            Ok(())
        }
        Err(err) => Err(err),
    }
}
