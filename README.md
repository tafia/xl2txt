# xl2txt

Convert Excel / OpenDocument SpreadSheets files into text files

```sh
xl2txt /my/excel/filename.xlsx
```

This command will per default create a `/my/excel/.filename` directory which contains:
- `data/sheet_name.md` files for cells value. [example](https://github.com/tafia/xl2txt/blob/master/tests/.issues.xlsb/data/datatypes.md)
- `formula/sheet_name.md` files for cells formula [example](https://github.com/tafia/xl2txt/blob/master/tests/.issues.xlsb/formula/datatypes.md)
- `refs.md` file for all references. [example](https://github.com/tafia/xl2txt/blob/master/tests/.issues.xlsb/refs.md)
- `names.md` file for all defined names. [example](https://github.com/tafia/xl2txt/blob/master/tests/.issues.xlsb/names.md)
- `vba/module_name.vb` files for each module. [example](https://github.com/tafia/xl2txt/blob/master/tests/.issues.xlsb/vba/testVBA.vb)

Internally it relies heavily on [calamine](https://github.com/tafia/calamine) crate.

Supports all kind of excel files (xls, xlsx, xlsm, xlsb, xla, xlam) and ods files.
