// TODO - todo
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2024, John McNamara, jmcnamara@cpan.org

use pyo3::prelude::*;
use pyo3::types::PyFloat;
use pyo3::types::PyString;
use rust_xlsxwriter::Workbook as RustWorkbook;
use rust_xlsxwriter::Worksheet as RustWorksheet;

/// TODO
#[pymodule]
fn xlsxwriter_lite(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_class::<Workbook>()?;
    m.add_class::<Worksheet>()?;
    Ok(())
}

#[pyclass]
pub struct Workbook {
    path: String,
    //worksheets: Vec<Worksheet>,
    rust_workbook: RustWorkbook,
}

#[pymethods]
impl Workbook {
    /// TODO
    ///
    /// # Errors
    ///
    #[new]
    pub fn new(path: &PyString) -> PyResult<Workbook> {
        Ok(Workbook {
            path: path.to_string(),
            rust_workbook: RustWorkbook::new(),
            //worksheets: vec![],
        })
    }

    /// TODO
    ///
    /// # Errors
    ///
    pub fn add_worksheet(&mut self) -> PyResult<Worksheet> {
        let worksheet = Worksheet::new();
        // let mut worksheet = Worksheet::new().unwrap();
        //self.worksheets.push(worksheet);
        //let worksheet = self.worksheets.last_mut().unwrap();
        Ok(worksheet)
    }

    /// TODO
    ///
    /// # Errors
    ///
    pub fn push_worksheet(&mut self, worksheet: &mut Worksheet) -> PyResult<()> {
        let mut tmp = Worksheet::new();
        std::mem::swap(&mut tmp.rust_worksheet, &mut worksheet.rust_worksheet);
        self.rust_workbook.push_worksheet(tmp.rust_worksheet);
        Ok(())
    }

    /// TODO
    ///
    /// # Errors
    ///
    pub fn close(&mut self) -> PyResult<()> {
        self.rust_workbook.save(&self.path).unwrap();
        Ok(())
    }
}

#[pyclass]
pub struct Worksheet {
    pub(crate) rust_worksheet: RustWorksheet,
}

#[pymethods]
impl Worksheet {
    /// TODO
    ///
    /// # Errors
    ///
    #[new]
    pub fn new() -> Worksheet {
        Worksheet {
            rust_worksheet: RustWorksheet::new(),
        }
    }

    pub fn write_number(&mut self, row: u32, col: u16, num: &PyFloat) {
        self.rust_worksheet
            .write_number(row, col, num.value())
            .unwrap();
    }
}
