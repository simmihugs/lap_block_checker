use calamine::{open_workbook, Reader, Xlsx};
use clap::Parser;
use colored::Colorize;
use serde::{Deserialize, Serialize};

fn is_numeric(x: &String) -> bool {
    x.parse::<f64>().is_ok()
}

enum Element {
    Program(Vec<String>),
    Trailer(Vec<String>),
    Commercial(Vec<String>),
}

impl Element {
    fn print(&self) {
        match self {
            Element::Program(_v) => println!("{}", "program   ".black().on_green()),
            Element::Commercial(_v) => println!("{}", "commercial".black().on_yellow()),
            Element::Trailer(_v) => println!("{}", "trailer   ".black().on_red()),
        }
    }
}

#[derive(Clone, Serialize, Deserialize, Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    #[arg(short, long)]
    filename: String,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Args = Args::parse();

    let mut excel: Xlsx<_> = open_workbook(args.filename)?;
    let mut list: Vec<Element> = vec![];

    let sheets = excel.sheet_names().to_owned();

    if let Ok(r) = excel.worksheet_range(&sheets[0]) {
        for row in r.rows() {
            let is_program: bool = row[6].to_string().contains("HD");
            let parts = row[8]
                .to_string()
                .split(" ")
                .map(|x| x.to_string())
                .collect::<Vec<String>>();
            let is_commercial: bool = {
                if parts.len() > 1 {
                    !is_numeric(&parts[0]) && is_numeric(&parts[1])
                } else {
                    false
                }
            };
            let is_trailer: bool = {
                let is_a_commercial: bool = row[6].to_string().starts_with("WS");
                !is_a_commercial && is_numeric(&parts[0]) && !is_numeric(&parts[1])
            };
            if is_program {
                list.push(Element::Program(vec![
                    row[1].to_string(),
                    row[6].to_string(),
                    row[8].to_string(),
                ]));
            } else if is_commercial {
                list.push(Element::Commercial(vec![
                    row[1].to_string(),
                    row[6].to_string(),
                    row[8].to_string(),
                ]));
            } else if is_trailer {
                match &list.last() {
                    Some(Element::Trailer(..)) => (),
                    _ => {
                        list.push(Element::Trailer(vec![
                            row[1].to_string(),
                            row[6].to_string(),
                            row[8].to_string(),
                        ]));
                    }
                }
            }
        }
    }
    for l in list {
        l.print();
    }

    Ok(())
}
