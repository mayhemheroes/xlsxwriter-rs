#![no_main]
use libfuzzer_sys::fuzz_target;
use xlsxwriter::*;
use xlsxwriter::format::FormatColor;
use std::str;
use std::ffi::CString;

const MAX_STRING_SIZE: usize = 10;

fuzz_target!(|data: &[u8]| {
    let wk = Workbook::new("target/simple1.xlsx");
    match wk {
        Ok(workbook) => {
            let mut sheet1 = workbook.add_worksheet(None).expect("Failed");
            let mut idx = 0;
            while idx + 5 < data.len() {
                let x =  unsafe{std::ptr::read(&data[idx])} as u16;
                let y =  unsafe{std::ptr::read(&data[idx+2])} as u16;
                let mut str_end = idx + 4 + MAX_STRING_SIZE;
                if str_end > data.len() {
                    str_end = data.len();
                }

                let str_to_add = str::from_utf8(&data[idx + 4..str_end]);

                match str_to_add {
                    Ok(string_add) => {
                        // Theres a defect here in the original repository
                        // It unwraps this CString::new object unsafely
                        // I check that string is ok before passing it on
                        match CString::new(string_add) {
                            Ok(..)=>{
                                sheet1.write_string(x.into(), y, string_add, Some(Format::new().set_font_color(FormatColor::Red))).expect("ERROR");
                            },
                            Err(..)=>()
                        }
                    }
                    Err(..)=>()
                }

                idx = str_end;
            }
        },
        Err(_) => ()
    };
});
