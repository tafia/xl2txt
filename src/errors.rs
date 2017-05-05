//! `Error` management module

#![allow(missing_docs)]

error_chain! {
    foreign_links {
        Io(::std::io::Error);
        Sheets(::calamine::Error);
    }

//     errors {
//         InvalidExtension(ext: String) {
//             description("invalid extension")
//             display("invalid extension: '{}'", ext)
//         }
//         CellOutOfRange(try_pos: (u32, u32), min_pos: (u32, u32)) {
//             description("no cell found at this position")
//             display("there is no cell at position '{:?}'. Minimum position is '{:?}'",
//                     try_pos, min_pos)
//         }
//         WorksheetName(name: String) {
//             description("invalid worksheet name")
//             display("invalid worksheet name: '{}'", name)
//         }
//         WorksheetIndex(idx: usize) {
//             description("invalid worksheet index")
//             display("invalid worksheet index: {}", idx)
//         }
//     }
}
