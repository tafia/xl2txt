#[macro_use]
extern crate error_chain;
extern crate calamine;

mod errors;

use std::path::{PathBuf, Path};
use std::fs;
use std::io::{BufWriter, Write};
use errors::Result;
use calamine::{Sheets, Range};

fn main() {
    let file = ::std::env::args().skip(1).next().expect("USAGE: xl2txt file");
    run(file.into()).unwrap();
}

fn run(file: PathBuf) -> Result<()> {

    let paths = XlPaths::new(file)?;
    let mut xl = Sheets::open(&paths.orig)?;

    // defined names
    {
        let mut f = BufWriter::new(fs::File::create(paths.names)?);
        writeln!(f, "| Name | Formula |")?;
        writeln!(f, "|------|---------|")?;
        for &(ref name, ref formula) in xl.defined_names()? {
            writeln!(f, "| {} | {} |", name, formula)?;
        }
    }

    // sheets
    let sheets = xl.sheet_names()?;
    for s in sheets {
        write_range(paths.data.join(format!("{}.md", &s)), xl.worksheet_range(&s)?)?;
        write_range(paths.formula.join(format!("{}.md", &s)), xl.worksheet_formula(&s)?)?;
    }

    // vba
    if !xl.has_vba() {
        return Ok(());
    }

    let mut vba = xl.vba_project()?;
    let vba = vba.to_mut();
    for module in vba.get_module_names() {
        let mut m = fs::File::create(paths.vba.join(format!("{}.vb", module)))?;
        write!(m, "{}", vba.get_module(module)?)?;
    }
    {
        let mut f = BufWriter::new(fs::File::create(paths.refs)?);
        writeln!(f, "| Name | Description | Path |")?;
        writeln!(f, "|------|-------------|------|")?;
        for r in vba.get_references() {
            writeln!(f, "| {} | {} | {} |", r.name, r.description, r.path.display())?;
        }
    }

    Ok(())
}

struct XlPaths {
    orig: PathBuf,
    data: PathBuf,
    formula: PathBuf,
    vba: PathBuf,
    refs: PathBuf,
    names: PathBuf,
}

impl XlPaths {

    fn new(orig: PathBuf) -> Result<XlPaths> {

        if !orig.exists() {
            bail!("Cannot find {}", orig.display());
        }

        if !orig.is_file() {
            bail!("{} is not a file", orig.display());
        }
        
        match orig.extension().and_then(|e| e.to_str()) {
            Some("xls") | Some("xlsx") | Some("xlsb") | Some("ods") => (),
            Some(e) => bail!("Unrecognized extension: {}", e),
            None => bail!("Expecting an excel file, couln't find an extension"),
        }

        let root: PathBuf = orig.parent()
            .map_or::<PathBuf, _>(".{}".into(), |p| p.into())
            .join(format!(".{}", &*orig.file_name().unwrap().to_string_lossy()));

        if !root.exists() {
            fs::create_dir_all(&root)?;
        }
        let data = root.join("sheets");
        if !data.exists() {
            fs::create_dir(&data)?;
        }
        let vba = root.join("vba");
        if !vba.exists() {
            fs::create_dir(&vba)?;
        }
        let formula = root.join("formula");
        if !formula.exists() {
            fs::create_dir(&formula)?;
        }

        Ok(XlPaths {
            orig: orig,
            data: data,
            formula: formula,
            vba: vba,
            refs: root.join("refs.md"),
            names: root.join("names.txt"),
        })
    }
}

fn write_range<P, T>(path: P, range: Range<T>) -> Result<()> 
    where P: AsRef<Path>, 
          T: PartialEq + Clone + Default + ::std::fmt::Debug,
{
    if range.is_empty() {
        return Ok(());
    }

    let mut f = BufWriter::new(fs::File::create(path.as_ref())?);

    // next rows: table data
    let mut is_first = true;
    for row in range.rows().skip(1) {
        for c in row {
            write!(f, "| {:?} ", c)?;
        }
        writeln!(f, "|")?;
        if is_first {
            // first row: consider as header
            for _ in &range[0] {
                write!(f, "|---")?;
            }
            writeln!(f, "|")?;
            is_first = false;
        }
    }
    Ok(())
}
