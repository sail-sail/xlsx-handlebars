//! 错误类型定义

use thiserror::Error;

/// XLSX 处理错误类型
#[derive(Error, Debug)]
pub enum XlsxError {
    #[error("Invalid Zip Format")]
    InvalidZipFormat,
    #[error("{0}")]
    TemplateRenderError(String),
}
