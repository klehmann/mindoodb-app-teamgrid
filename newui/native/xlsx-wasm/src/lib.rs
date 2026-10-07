//! WASM facade over the vendored GenOffice xlsx sidecar (read path only).
//!
//! Every method answers with the same JSON the native sidecar puts in a
//! response's `result` (serialized by the very same serde types), so the
//! TypeScript side can reuse GenOffice's `WorkbookFile` /
//! `WorkbookRangeResult` shapes. JSON travels as a string: one
//! `JSON.parse` is far cheaper than walking a big object through
//! wasm-bindgen. Errors throw a JS `Error` whose `code` matches the
//! sidecar's error codes (`invalid_request`, `io_error`, `workbook_error`).

use base64::Engine;
use serde::Serialize;
use wasm_bindgen::prelude::*;
use xlsx_sidecar::{CellRange, SidecarError, WorkbookSessions};

fn to_js_error(error: SidecarError) -> JsValue {
    let code = match error {
        SidecarError::InvalidRequest(_) => "invalid_request",
        SidecarError::Io(_) => "io_error",
        SidecarError::Workbook(_) => "workbook_error",
        SidecarError::Cancelled => "cancelled",
    };
    let js_error = js_sys::Error::new(&error.to_string());
    let _ = js_sys::Reflect::set(&js_error, &"code".into(), &code.into());
    js_error.into()
}

fn to_json<T: Serialize>(value: Result<T, SidecarError>) -> Result<String, JsValue> {
    let value = value.map_err(to_js_error)?;
    serde_json::to_string(&value).map_err(|error| to_js_error(error.into()))
}

/// One engine holds any number of open workbook sessions, keyed by the
/// `sessionId` that `open` returns.
#[wasm_bindgen]
pub struct XlsxEngine {
    sessions: WorkbookSessions,
}

impl Default for XlsxEngine {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen]
impl XlsxEngine {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            sessions: WorkbookSessions::new(),
        }
    }

    /// Opens an .xlsx from its bytes. Returns the sidecar `open` result
    /// (WorkbookFile-shaped JSON incl. `sessionId`).
    pub fn open(
        &mut self,
        bytes: Vec<u8>,
        name: Option<String>,
        locale: Option<String>,
        #[wasm_bindgen(js_name = shortDateFormat)] short_date_format: Option<String>,
    ) -> Result<String, JsValue> {
        to_json(self.sessions.open_bytes(
            bytes,
            name.as_deref().unwrap_or("workbook.xlsx"),
            locale.as_deref().unwrap_or("zh"),
            short_date_format.as_deref(),
        ))
    }

    /// Sidecar `read_range` result (WorkbookRangeResult-shaped JSON). The
    /// first read of a sheet indexes it completely (no threads in WASM), so
    /// `indexingComplete` is always true.
    #[wasm_bindgen(js_name = readRange)]
    pub fn read_range(
        &mut self,
        #[wasm_bindgen(js_name = sessionId)] session_id: &str,
        #[wasm_bindgen(js_name = sheetId)] sheet_id: &str,
        #[wasm_bindgen(js_name = startRow)] start_row: usize,
        #[wasm_bindgen(js_name = endRow)] end_row: usize,
        #[wasm_bindgen(js_name = startColumn)] start_column: usize,
        #[wasm_bindgen(js_name = endColumn)] end_column: usize,
    ) -> Result<String, JsValue> {
        let range = CellRange {
            start_row,
            end_row,
            start_column,
            end_column,
        };
        to_json(self.sessions.read_range(session_id, sheet_id, &range))
    }

    /// Sidecar `read_formula_cells` result.
    #[wasm_bindgen(js_name = readFormulaCells)]
    pub fn read_formula_cells(
        &mut self,
        #[wasm_bindgen(js_name = sessionId)] session_id: &str,
        #[wasm_bindgen(js_name = sheetId)] sheet_id: &str,
    ) -> Result<String, JsValue> {
        to_json(self.sessions.read_formula_cells(session_id, sheet_id))
    }

    /// Sidecar `read_media` result: `{ mediaType, base64 }` for a visual or
    /// cell-image id.
    #[wasm_bindgen(js_name = readMedia)]
    pub fn read_media(
        &self,
        #[wasm_bindgen(js_name = sessionId)] session_id: &str,
        #[wasm_bindgen(js_name = visualId)] visual_id: &str,
    ) -> Result<String, JsValue> {
        to_json(self.sessions.read_media(session_id, visual_id))
    }

    /// The raw image bytes behind a visual or cell-image id (for a Blob URL
    /// without a base64 round trip on the JS side).
    #[wasm_bindgen(js_name = readMediaBytes)]
    pub fn read_media_bytes(
        &self,
        #[wasm_bindgen(js_name = sessionId)] session_id: &str,
        #[wasm_bindgen(js_name = visualId)] visual_id: &str,
    ) -> Result<Vec<u8>, JsValue> {
        let media = self
            .sessions
            .read_media(session_id, visual_id)
            .map_err(to_js_error)?;
        base64::engine::general_purpose::STANDARD
            .decode(media.base64)
            .map_err(|error| to_js_error(SidecarError::Workbook(error.to_string())))
    }

    pub fn close(
        &mut self,
        #[wasm_bindgen(js_name = sessionId)] session_id: &str,
    ) -> Result<(), JsValue> {
        self.sessions.close(session_id).map_err(to_js_error)
    }
}
