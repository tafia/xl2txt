# xl2txt

Convert Excel / OpenDocument SpreadSheets files into text files

```sh
xl2txt /my/excel/filename.xlsx
```

This command will per default create a `/my/excel/.filename` directory which contains:
- parsed data from every worksheet in the `data` sub directory
- parsed formula from every worksheet in the `formula` sub directory
- references used in vba in `refs.md` file
- defined names used within the workbook into `names.md` file
- vba modules in `vba` sub directory
