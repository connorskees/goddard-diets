use calamine::{open_workbook, DataType, Xlsx};

// use xlsxwriter::Workbook;

use goddard_diets::DataFrame;
use simple_excel_writer::{Row as ExcelRow, Workbook};

fn main() {
    let workbook: Xlsx<_> =
        open_workbook("/Users/user/Downloads/All guests_Connor.xlsx").expect("Cannot open file");

    let df = DataFrame::from_xlsx(workbook);

    let attendees_with_guest = df.filter_by_col("haveguest", |has_guest| {
        has_guest == &DataType::String("X".to_owned())
    });

    let all_guests = attendees_with_guest
        .subselect_cols(&[
            "tablenumber",
            "first",
            "last",
            "descrip",
            "host",
            "guestfirst",
            "guestlast",
            "guestdietary",
        ])
        .map_col("descrip", |row| {
            DataType::String(format!("Guest of {} {}", row["first"], row["last"]))
        })
        .drop_columns(&["first", "last"])
        .rename_column("guestfirst", "first")
        .rename_column("guestlast", "last")
        .rename_column("guestdietary", "dietary")
        .filter_by_col("dietary", |dietary| dietary != &DataType::Empty);

    let mut attendees_with_allergies = df
        .subselect_cols(&["tablenumber", "descrip", "host", "first", "last", "dietary"])
        .filter_by_col("dietary", |dietary| dietary != &DataType::Empty);

    attendees_with_allergies.concat(all_guests);

    let mut wb = Workbook::create_in_memory();

    let mut sheet = wb.create_sheet("SheetName");

    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;

        sw.append_row(ExcelRow::from_iter([1.0_f64, 2.0, 3.0].into_iter()))?;

        Ok(())
        // sw.append_row(row!["Name", "Title", "Success", "XML Remark"])
        // sw.append_row(row![
        //     "Amy",
        //     (),
        //     true,
        //     "<xml><tag>\"Hello\" & 'World'</tag></xml>"
        // ])?;
        // sw.append_blank_rows(2);
        // sw.append_row(row!["Tony", blank!(30), "retired"])
    })
    .expect("write excel error!");

    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;

        for row in attendees_with_allergies.arr.rows() {
            let mut excel_row = ExcelRow::new();

            for col in row.columns().into_iter().next().unwrap() {
                dbg!(col); // into_iter().collect::<Vec<_>>()
            }
            // sw.append_row(ExcelRow::from_iter(row.into_iter()))?;
        }

        Ok(())
        // sw.append_row(row!["Name", "Title", "Success", "XML Remark"])
        // sw.append_row(row![
        //     "Amy",
        //     (),
        //     true,
        //     "<xml><tag>\"Hello\" & 'World'</tag></xml>"
        // ])?;
        // sw.append_blank_rows(2);
        // sw.append_row(row!["Tony", blank!(30), "retired"])
    })
    .expect("write excel error!");

    // (wb.close().unwrap().unwrap());

    // let workbook = Workbook::new("simple1.xlsx");

    // let mut sheet1 = workbook.add_worksheet(None).unwrap();

    // for (col_idx, col) in attendees_with_allergies
    //     .arr
    //     .columns()
    //     .into_iter()
    //     .enumerate()
    // {
    //     let row = col.rows().into_iter().next().unwrap();

    //     for (row_idx, row) in row.into_iter().enumerate() {
    //         dbg!(row_idx, col_idx);
    //         match row {
    //             DataType::String(text) => sheet1
    //                 .write_string(row_idx as u32, col_idx as u16, text, None)
    //                 .unwrap(),
    //             &DataType::Float(number) => sheet1
    //                 .write_number(row_idx as u32, col_idx as u16, number, None)
    //                 .unwrap(),
    //             DataType::Empty => continue,
    //             d => todo!("{:?}", d),
    //         }
    //     }
    // }
}
