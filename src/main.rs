#[macro_use]
extern crate error_chain;
extern crate calamine;

mod errors;

use std::path::{PathBuf, Path};
use std::fs;
use std::io::{BufWriter, Write};
use errors::Result;
use calamine::{Sheets, Range, CellType};

fn main() {
    let mut args = ::std::env::args();
    let file = args.by_ref().skip(1).next().expect("USAGE: xl2txt file [root]");
    let root = args.next().map(|r| r.into());
    run(file.into(), root).unwrap();
}

fn run(file: PathBuf, root: Option<PathBuf>) -> Result<()> {

    let paths = XlPaths::new(file, root)?;
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
    fn new(orig: PathBuf, root: Option<PathBuf>) -> Result<XlPaths> {

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

        let root_next = format!(".{}", &*orig.file_name().unwrap().to_string_lossy());
        let root = root.unwrap_or_else(|| orig.parent().map_or(".".into(), |p| p.into()))
            .join(root_next);

        if root.exists() {
            fs::remove_dir_all(&root)?;
        }
        fs::create_dir_all(&root)?;
        let data = root.join("data");
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
            names: root.join("names.md"),
        })
    }
}

fn write_range<P, T>(path: P, range: Range<T>) -> Result<()> 
    where P: AsRef<Path>, 
          T: CellType + ::std::fmt::Display,
{
    if range.is_empty() {
        return Ok(());
    }

    let mut f = BufWriter::new(fs::File::create(path.as_ref())?);

    let ((srow, scol), (_, ecol)) = (range.start(), range.end());
    write!(f, "|   ")?;
    for c in scol..ecol+1 {
        write!(f, "| {} ", get_column(c))?;
    }
    writeln!(f, "|")?;
    for _ in scol..ecol+2 {
        write!(f, "|---")?;
    }
    writeln!(f, "|")?;

    // next rows: table data
    let srow = srow as usize + 1;
    for (i, row) in range.rows().enumerate() {
        write!(f, "| {} ", srow + i)?;
        for c in row {
            write!(f, "| {} ", c)?;
        }
        writeln!(f, "|")?;
    }
    Ok(())
}

fn get_column(mut col: u32) -> String {
    let mut buf = String::new();
    if col < 26 {
        buf.push((b'A' + col as u8) as char);
    } else {
        let mut rev = String::new();
        while col >= 26 {
            let c = col % 26;
            rev.push((b'A' + c as u8) as char);
            col -= c;
            col /= 26;
        }
        buf.extend(rev.chars().rev());
    }
    buf
}
