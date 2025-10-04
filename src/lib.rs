#[cfg(target_arch = "wasm32")]
use wasm_bindgen::prelude::*;
#[cfg(target_arch = "wasm32")]
use wasm_bindgen::JsValue;

pub mod errors;
pub mod imagesize;
mod template;
pub mod utils;

// 重新导出常用的类型和函数
pub use errors::XlsxError;
pub use imagesize::get_image_dimensions;
pub use utils::{to_column_index, to_column_name, timestamp_to_excel_date, excel_date_to_timestamp};

/// 当 `console_error_panic_hook` 功能启用时，我们可以调用 `set_panic_hook` 函数
/// 至少一次在初始化过程中，以便在 panic 时获得更好的错误消息。
pub fn set_panic_hook() {
    #[cfg(feature = "console_error_panic_hook")]
    console_error_panic_hook::set_once();
}

// WASM 平台：导出 WASM 兼容的渲染函数
#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
pub fn render_template(
    zip_bytes: Vec<u8>,
    data_json: &str,
) -> Result<JsValue, JsValue> {
    let data: serde_json::Value = serde_json::from_str(data_json)
            .map_err(|e| JsValue::from_str(&format!("JSON Parse Error: {e}")))?;

    // 调用模板渲染函数
    let result = template::render_template(zip_bytes, &data)
        .map_err(|e| JsValue::from_str(&e.to_string()))?;

    // 返回结果
    Ok(JsValue::from(result))
}

// WASM 平台：导出工具函数
#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
pub fn wasm_get_image_dimensions(data: Vec<u8>) -> JsValue {
    match imagesize::get_image_dimensions(&data) {
        Some((width, height)) => {
            let obj = js_sys::Object::new();
            js_sys::Reflect::set(&obj, &"width".into(), &width.into()).unwrap();
            js_sys::Reflect::set(&obj, &"height".into(), &height.into()).unwrap();
            obj.into()
        }
        None => JsValue::NULL,
    }
}

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
pub fn wasm_to_column_name(current: &str, increment: u32) -> String {
    utils::to_column_name(current, increment)
}

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
pub fn wasm_to_column_index(col_name: &str) -> u32 {
    utils::to_column_index(col_name)
}

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
pub fn wasm_timestamp_to_excel_date(timestamp_ms: i64) -> f64 {
    utils::timestamp_to_excel_date(timestamp_ms)
}

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
pub fn wasm_excel_date_to_timestamp(excel_date: f64) -> Option<i64> {
    utils::excel_date_to_timestamp(excel_date)
}

// 非 WASM 平台：直接导出原生 Rust 函数
#[cfg(not(target_arch = "wasm32"))]
pub use template::render_template;
