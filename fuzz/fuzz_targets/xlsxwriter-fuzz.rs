#![no_main]
use libfuzzer_sys::fuzz_target;
use xlsxwriter::*;
use std::str;

const MAX_STRING_SIZE: usize = 10;

fuzz_target!(|data: &[u8]| {
    let workbook = Workbook::new("target/simple1.xlsx");
    let format1 = workbook.add_format().set_font_color(FormatColor::Red);
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
                sheet1.write_string(x.into(), y, string_add, Some(&format1));
            }
            Err(..)=>()
        }

        idx = str_end;
    }
});
