use std::{
    io::{BufRead, BufReader, Seek},
    ops::Index,
};

use calamine::{open_workbook, DataType, Reader, Xlsx};
use ndarray::{Array2, Axis};
use wasm_bindgen::prelude::*;
// use xlsxwriter::Workbook;
use simple_excel_writer::sheet::ToCellValue;
use simple_excel_writer::{Row as ExcelRow, Workbook};

#[derive(Clone)]
pub struct DataFrame {
    pub arr: Array2<DataType>,
    cols: usize,
    header: Vec<DataType>,
}

pub struct Row<'a> {
    header: &'a [DataType],
    row: Vec<DataType>,
}

impl<'a> Index<&str> for Row<'a> {
    type Output = DataType;

    fn index(&self, index: &str) -> &Self::Output {
        let col_idx = self
            .header
            .iter()
            .position(|header| header == &DataType::String(index.to_owned()))
            .unwrap();

        return &self.row[col_idx];
    }
}

impl DataFrame {
    pub fn from_xlsx(mut workbook: Xlsx<impl BufRead + Seek>) -> Self {
        let sheet_names = workbook.sheet_names();

        assert_eq!(sheet_names.len(), 1);

        let sheet_name = &sheet_names[0].clone();

        let sheet = workbook
            .worksheet_range(sheet_name)
            .unwrap()
            .unwrap()
            .rows()
            .map(|x| x.to_vec())
            .collect::<Vec<Vec<_>>>();

        let row_count = sheet.len();
        let col_count = sheet[0].len();

        let flat = sheet.into_iter().flatten().collect();

        DataFrame::new(Array2::from_shape_vec((row_count, col_count), flat).unwrap())
    }

    pub fn new(arr: Array2<DataType>) -> Self {
        let header = arr.row(0).to_vec();

        Self::with_header(header, arr)
    }

    pub fn with_header(header: Vec<DataType>, arr: Array2<DataType>) -> Self {
        Self {
            cols: arr.ncols(),
            header,
            arr,
        }
    }

    pub fn concat(&mut self, other: Self) -> &mut Self {
        assert_eq!(self.header, other.header);

        self.arr.append(Axis(0), other.arr.view()).unwrap();

        self
    }

    fn idx_for_column(&self, col: &str) -> usize {
        self.header
            .iter()
            .position(|header| header == &DataType::String(col.to_owned()))
            .unwrap()
    }

    pub fn filter_by_col(&self, col: &str, predicate: impl Fn(&DataType) -> bool) -> DataFrame {
        let col_idx = self.idx_for_column(col);

        let filtered_rows = self
            .arr
            .outer_iter()
            .filter(|row| predicate(&row[col_idx]))
            .flatten()
            .cloned()
            .collect::<Vec<_>>();

        let df =
            Array2::from_shape_vec((filtered_rows.len() / self.cols, self.cols), filtered_rows)
                .unwrap();

        return DataFrame::with_header(self.header.clone(), df);
    }

    pub fn drop_columns(&self, cols: &[&str]) -> DataFrame {
        let new_cols = self
            .header
            .iter()
            .map(|c| match c {
                DataType::String(col) => col.as_str(),
                _ => panic!(),
            })
            .filter(|col| !cols.contains(col))
            .collect::<Vec<_>>();

        return self.subselect_cols(new_cols.as_slice());
    }

    pub fn rename_column(&mut self, original: &str, new: &str) -> &mut DataFrame {
        let col_idx = self.idx_for_column(original);

        self.header[col_idx] = DataType::String(new.to_owned());

        self
    }

    pub fn map_col(&self, col: &str, f: impl Fn(Row<'_>) -> DataType) -> DataFrame {
        let col_idx = self.idx_for_column(col);

        let mut df = self.clone();

        for mut row in df.arr.rows_mut() {
            row[col_idx] = f(Row {
                row: row.to_vec(),
                header: &self.header,
            });
        }

        df
    }

    pub fn subselect_cols(&self, cols: &[&str]) -> DataFrame {
        let header_idxs = cols
            .into_iter()
            .map(|desired_header| self.idx_for_column(desired_header))
            .collect::<Vec<_>>();

        let rows = self
            .arr
            .outer_iter()
            .map(|row| {
                header_idxs
                    .iter()
                    .map(|&idx| row[idx].clone())
                    .collect::<Vec<_>>()
            })
            .flatten()
            .collect::<Vec<_>>();

        let df = Array2::from_shape_vec((rows.len() / header_idxs.len(), header_idxs.len()), rows)
            .unwrap();

        return DataFrame::with_header(
            cols.into_iter()
                .map(|&col| DataType::String(col.to_owned()))
                .collect(),
            df,
        );
    }
}

#[wasm_bindgen]
extern "C" {
    pub fn alert(s: &str);
}

#[wasm_bindgen]
pub fn main(buffer: &[u8]) -> Vec<u8> {
    let workbook = Xlsx::new(BufReader::new(std::io::Cursor::new(buffer))).unwrap();

    // match workbook {
    //     Ok(..) => {}
    //     Err(e) => {
    //         unsafe { alert(&format!("{:?}", e)) };
    //     }
    // }

    // // open_workbook("/Users/user/Downloads/All guests_Connor.xlsx").expect("Cannot open file");

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

    // unsafe { alert(&attendees_with_allergies.arr.len().to_string()) }

    // // let workbook = Workbook::new("simple1.xlsx");
    let mut wb = Workbook::create_in_memory();

    let mut sheet = wb.create_sheet("SheetName");

    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;

        for row in attendees_with_allergies.arr.rows() {
            let mut excel_row = ExcelRow::new();

            for col in row.columns().into_iter().next().unwrap() {
                // dbg!(col)
                match col {
                    DataType::String(text) => excel_row.add_cell(text.as_str()),
                    &DataType::Float(number) => excel_row.add_cell(number),
                    DataType::Empty => {
                        excel_row.add_empty_cells(1);
                    }
                    d => todo!("{:?}", d),
                }
            }

            sw.append_row(excel_row)?;
        }

        Ok(())
    })
    .expect("write excel error!");

    // // let mut sheet1 = workbook.add_worksheet(None).unwrap();

    // // for (col_idx, col) in attendees_with_allergies
    // //     .arr
    // //     .columns()
    // //     .into_iter()
    // //     .enumerate()
    // // {
    // //     let row = col.rows().into_iter().next().unwrap();

    // //     for (row_idx, row) in row.into_iter().enumerate() {
    // //         dbg!(row_idx, col_idx);
    // //         match row {
    // //             DataType::String(text) => sheet1
    // //                 .write_string(row_idx as u32, col_idx as u16, text, None)
    // //                 .unwrap(),
    // //             &DataType::Float(number) => sheet1
    // //                 .write_number(row_idx as u32, col_idx as u16, number, None)
    // //                 .unwrap(),
    // //             DataType::Empty => continue,
    // //             d => todo!("{:?}", d),
    // //         }
    // //     }
    // // }

    // unsafe { alert(&format!("{:?}", wb.close())) };
    return wb.close().unwrap().unwrap();
}
