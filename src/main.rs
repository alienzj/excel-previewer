extern crate calamine;
extern crate comrak;
extern crate html_minifier;
extern crate htmlescape;
extern crate madato;

#[macro_use]
extern crate lazy_static_include;

#[macro_use]
extern crate lazy_static;

use calamine::{open_workbook_auto, DataType, Reader};
use comrak::{markdown_to_html, ComrakOptions};
use html_minifier::HTMLMinifier;
use htmlescape::*;
use madato::types::{ErroredTable, NamedTable, RenderOptions, TableRow};

use std::env;

lazy_static_include_str!(MARKDOWN_CSS, "resources/github-markdown.css");

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

fn _list_sheet_names(filename: String) -> Result<Vec<String>, String> {
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

fn _get_sheet_names(filename: String) {
    for s in _list_sheet_names(filename).unwrap() {
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

fn html_render(markdown_html: &str, title: &str) -> Result<String, String> {
    let mut minifier = HTMLMinifier::new();

    minifier
        .digest("<!DOCTYPE html>")
        .map_err(|err| err.to_string())?;
    minifier.digest("<html>").map_err(|err| err.to_string())?;

    minifier.digest("<head>").map_err(|err| err.to_string())?;
    minifier
        .digest("<meta charset=UTF-8>")
        .map_err(|err| err.to_string())?;
    minifier
        .digest("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, shrink-to-fit=no\">")
        .map_err(|err| err.to_string())?;

    minifier.digest("<title>").map_err(|err| err.to_string())?;
    minifier
        .digest(&encode_minimal(title))
        .map_err(|err| err.to_string())?;
    minifier.digest("</title>").map_err(|err| err.to_string())?;

    minifier.digest("<style>").map_err(|err| err.to_string())?;
    minifier
        .digest(&encode_minimal(&MARKDOWN_CSS))
        .map_err(|err| err.to_string())?;
    minifier.digest("</style>").map_err(|err| err.to_string())?;

    minifier.digest("</head>").map_err(|err| err.to_string())?;

    minifier.digest("<body>").map_err(|err| err.to_string())?;

    minifier
        .digest("<article class=\"markdown-body\">")
        .map_err(|err| err.to_string())?;
    minifier
        .digest(&markdown_html)
        .map_err(|err| err.to_string())?;
    minifier
        .digest("</article>")
        .map_err(|err| err.to_string())?;

    minifier.digest("</body>").map_err(|err| err.to_string())?;

    minifier.digest("</html>").map_err(|err| err.to_string())?;

    let minified_html = minifier.get_html();

    Ok(minified_html)
}

fn main() -> Result<(), String> {
    let args: Vec<String> = env::args().collect();
    let filename = &args[1];

    let render_options = Some(RenderOptions {
        headings: None,
        sheet_name: None,
        filters: None,
    });

    let output_string = spreadsheet_to_md(filename.to_string(), &render_options);

    match output_string {
        Ok(markdown) => {
            let mut options = ComrakOptions::default();
            options.unsafe_ = true;
            options.ext_autolink = true;
            options.ext_description_lists = true;
            options.ext_footnotes = true;
            options.ext_strikethrough = true;
            options.ext_superscript = true;
            options.ext_table = true;
            options.ext_tagfilter = true;
            options.ext_tasklist = true;
            options.hardbreaks = true;

            let html = markdown_to_html(&markdown[..], &options);
            let html_rendered = html_render(&html, "hello");
            match html_rendered {
                Ok(html_output) => {
                    println!("{}", html_output);
                    Ok(())
                }
                Err(err) => Err(err),
            }
        }
        Err(err) => Err(err),
    }
}
