use serde_json::Value;
use std::{io::{Cursor, Read, Write}, sync::{Arc, Mutex}};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};
use std::collections::HashMap;
use crate::{utils::{to_column_name, merge_handlebars_in_xml, register_basic_helpers, post_process_xml, replace_shared_strings_in_sheet, validate_xlsx_format}, XlsxError};
use crate::imagesize::get_image_dimensions;
use uuid::Uuid;

use handlebars::{Handlebars, RenderErrorReason};

/// 用于标记需要删除的行的 UUID
/// 配合 {{removeRow}} helper 使用
const REMOVE_ROW_KEY: &str = "|e5nBk+z4RMKqlyBo+xQ48A-remove-row|";

/// 用于标记数字类型的 UUID
/// 配合 {{num aa}} helper 使用
const TO_NUMBER_KEY: &str = "|e5nBk+z4RMKqlyBo+xQ48A-num|";

/// 用于标记公式类型的 UUID
/// 配合 {{formula "=SUM(A1:B1)"}} helper 使用
const TO_FORMULA_KEY: &str = "|e5nBk+z4RMKqlyBo+xQ48A-formula|";

/// 图片信息结构
#[derive(Debug, Clone)]
struct ImageInfo {
    col: u32,             // 列号（1-based）
    row: u32,             // 行号（1-based）
    base64_data: String,  // base64 图片数据
    width: Option<u32>,   // 用户指定宽度（像素）
    height: Option<u32>,  // 用户指定高度（像素）
    rid: String,          // 唯一的关系 ID（使用 UUID 避免冲突）
}

pub fn render_template(
  zip_bytes: Vec<u8>,
  data: &Value,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  
  // 首先验证输入是否为有效的 XLSX 文件
  validate_xlsx_format(&zip_bytes)?;
  
  // 创建一个 Cursor 来读取 zip 字节
  let cursor = Cursor::new(zip_bytes);
  let mut archive = ZipArchive::new(cursor)?;
  
  // 存储解压缩的文件内容
  let files: Arc<Mutex<HashMap<String, Vec<u8>>>> = Arc::new(Mutex::new(HashMap::new()));
  
  // 解压缩所有文件
  for i in 0..archive.len() {
    let mut file = archive.by_index(i)?;
    let file_name = file.name().to_string();
    
    // 跳过目录项
    if file_name.ends_with('/') {
      continue;
    }
    
    // 删除掉 xl/calcChain.xml 文件
    if file_name == "xl/calcChain.xml" {
      continue;
    }
    
    let mut contents = Vec::new();
    file.read_to_end(&mut contents)?;
    files.lock().unwrap().insert(file_name, contents);
  }
  
  // dbg!(files.lock().unwrap().keys());
  
  // 处理 sharedStrings.xml 文件
  // 把 sst 标签中的 si 标签 解析出来放到数组中, 其中的 si 标签换成 is 标签
  let mut shared_strings = Vec::new();
  {
    let file_name = "xl/sharedStrings.xml";
    let contents = files.lock().unwrap().remove(file_name);
    if let Some(contents) = contents {
      let xml_content = String::from_utf8(contents.clone())?;
      let mut start = 0;
      while let Some(si_start) = xml_content[start..].find("<si>") {
        let abs_start = start + si_start;
        if let Some(si_end) = xml_content[abs_start..].find("</si>") {
          let abs_end = abs_start + si_end + "</si>".len();
          let si_xml = &xml_content[abs_start..abs_end];
          // 将 si 标签替换为 is 标签
          let is_xml = si_xml
            .replace("<si>", "<is>")
            .replace("</si>", "</is>");
          shared_strings.push(is_xml);
          start = abs_end;
        } else {
          break;
        }
      }
      let xml_content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>"#.to_string();
      let contents = xml_content.into_bytes();
      files.lock().unwrap().insert(file_name.to_string(), contents);
    }
  }
  
  let mut handlebars = Handlebars::new();
      
  handlebars.set_strict_mode(false); // 允许未定义的变量
  
  register_basic_helpers(&mut handlebars)?;
  
  let data1 = Arc::new(Mutex::new(data.clone()));
  let data2 = Arc::clone(&data1);
  
  handlebars.register_helper("set_data", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    if let Some(key) = h.param(0).and_then(|v| v.value().as_str())
      && let Some(value) = h.param(1) {
        let mut data2 = data2.lock().unwrap();
        data2[key] = value.value().clone();
      }
    Ok(())
  }));
  
  // sheet_name
  let sheet_name = Arc::new(Mutex::new(String::new()));
  let sheet_name2 = Arc::clone(&sheet_name);
  
  // 行号偏移量
  let row_offset: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
  let row_offset2 = Arc::clone(&row_offset);
  let row_offset3 = Arc::clone(&row_offset);
  let row_offset4 = Arc::clone(&row_offset);
  let row_offset5 = Arc::clone(&row_offset);
  let row_offset6 = Arc::clone(&row_offset);
  let row_offset_for_remove = Arc::clone(&row_offset);  // 用于 removeRow helper
  
  // row_offset_plus 接收参数, 每次调用加上参数的值
  handlebars.register_helper("row_offset_plus", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    if let Some(value) = h.param(0).and_then(|v| v.value().as_u64()) {
      let mut offset = row_offset2.lock().unwrap();
      *offset += u32::try_from(value).unwrap();
    }
    Ok(())
  }));
  
  handlebars.register_helper("row_offset_reset", Box::new(move |_: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let mut offset = row_offset.lock().unwrap();
    *offset = 0;
    Ok(())
  }));
  
  // get_row_offset 获取当前行号偏移量
  handlebars.register_helper("get_row_offset", Box::new(move |_: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let offset = row_offset6.lock().unwrap();
    out.write(&offset.to_string())?;
    Ok(())
  }));
  
  // 当前行号
  let row_inline = Arc::new(Mutex::new(1u32));
  let row_inline2 = Arc::clone(&row_inline);
  let row_inline3 = Arc::clone(&row_inline);
  let row_inline4 = Arc::clone(&row_inline);
  let row_inline5 = Arc::clone(&row_inline);
  
  // 设置当前行号
  handlebars.register_helper("set_row_inline", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    if let Some(value) = h.param(0).and_then(|v| v.value().as_u64()) {
      let mut row_inline = row_inline2.lock().unwrap();
      *row_inline = u32::try_from(value).expect("set_row_inline too large for u32");
    }
    Ok(())
  }));
  
  // 获取计算后的 最终当前行号 _r = 当前行号 row_inline + 行号偏移量 row_offset
  handlebars.register_helper("_r", Box::new(move |_h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let row_inline = row_inline3.lock().unwrap();
    let row_offset = row_offset3.lock().unwrap();
    let r = *row_inline + *row_offset;
    out.write(&r.to_string())?;
    Ok(())
  }));
  
  // 列号偏移量
  let col_offset: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
  let col_offset2 = Arc::clone(&col_offset);
  let col_offset3 = Arc::clone(&col_offset);
  let col_offset4 = Arc::clone(&col_offset);
  let col_offset5 = Arc::clone(&col_offset);
  let col_offset6 = Arc::clone(&col_offset);
  
  handlebars.register_helper("col_offset_plus", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    if let Some(value) = h.param(0).and_then(|v| v.value().as_u64()) {
      let mut offset = col_offset2.lock().unwrap();
      *offset += u32::try_from(value).unwrap();
    }
    Ok(())
  }));
  
  handlebars.register_helper("col_offset_reset", Box::new(move |_: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let mut offset = col_offset.lock().unwrap();
    *offset = 0;
    out.write("")?;
    Ok(())
  }));
  
  // get_col_offset 获取当前列号偏移量
  handlebars.register_helper("get_col_offset", Box::new(move |_: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let offset = col_offset6.lock().unwrap();
    out.write(&offset.to_string())?;
    Ok(())
  }));
  
  // 当前列号
  let col_inline = Arc::new(Mutex::new(1u32));
  let col_inline2 = Arc::clone(&col_inline);
  let col_inline3 = Arc::clone(&col_inline);
  let col_inline4 = Arc::clone(&col_inline);
  let col_inline5 = Arc::clone(&col_inline);
  
  // 设置当前列号
  handlebars.register_helper("set_col_inline", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    if let Some(value) = h.param(0).and_then(|v| v.value().as_u64()) {
      let mut col_inline = col_inline2.lock().unwrap();
      *col_inline = u32::try_from(value).expect("set_col_inline too large for u32");
    }
    Ok(())
  }));
  
  // 获取计算后的 最终当前列号 _c = 当前列号 col_inline + 列号偏移量 col_offset
  handlebars.register_helper("_c", Box::new(move |_h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let col_inline = col_inline3.lock().unwrap();
    let col_offset = col_offset3.lock().unwrap();
    let c_num = *col_inline + *col_offset;
    let c_str = to_column_name("A", c_num - 1); // 列号从 1 开始, 需要减 1
    out.write(&c_str)?;
    Ok(())
  }));
  
  // handlebars.register_helper("_cr", Box::new(move |_h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
  //   let col_inline = col_inline4.lock().unwrap();
  //   let col_offset = col_offset4.lock().unwrap();
  //   let row_inline = row_inline4.lock().unwrap();
  //   let row_offset = row_offset4.lock().unwrap();
  //   let c_num = *col_inline + *col_offset;
  //   let c_str = to_column_name("A", c_num - 1); // 列号从 1 开始, 需要减 1
  //   let r_num = *row_inline + *row_offset;  // 行号需要加上偏移量
  //   out.write(&format!("{c_str}{r_num}"))?;
  //   Ok(())
  // }));
  
  // 上面的 _cr helper 在之前的逻辑上, 加入2个参数, 第一个参数是初始列号比如 B, 第二个参数是行号比如 10
  // 如果这两个参数都存在, 则使用这两个参数计算最终的列号和行号
  // 如果参数不存在, 则使用当前的列号和行号, 保持之前的逻辑
  handlebars.register_helper("_cr", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let col_inline = col_inline4.lock().unwrap();
    let col_offset = col_offset4.lock().unwrap();
    let row_inline = row_inline4.lock().unwrap();
    let row_offset = row_offset4.lock().unwrap();
    
    // 处理第一个参数: 初始列号
    let c_str = if let Some(param) = h.param(0) {
      if param.value().is_string() {
        let col_name = param.value().as_str().unwrap();
        // 将列名转换为列索引
        let col_index = crate::utils::to_column_index(col_name);
        // 计算最终列索引
        let final_col_index = col_index + *col_offset;
        // 转换回列名
        to_column_name("A", final_col_index.saturating_sub(1)) // 列号从 1 开始, 需要减 1
      } else if param.value().is_number() {
        let col_num = param.value().as_u64().unwrap_or(1) as u32;
        let final_col_num = col_num + *col_offset;
        to_column_name("A", final_col_num.saturating_sub(1)) // 列号从 1 开始, 需要减 1
      } else {
        // 非法类型则使用当前列号
        let c_num = *col_inline + *col_offset;
        to_column_name("A", c_num.saturating_sub(1)) // 列号从 1 开始, 需要减 1
      }
    } else {
      // 没有参数则使用当前列号
      let c_num = *col_inline + *col_offset;
      to_column_name("A", c_num.saturating_sub(1)) // 列号从 1 开始, 需要减 1
    };
    
    // 处理第二个参数: 行号
    let r_num = if let Some(param) = h.param(1).and_then(|v| v.value().as_u64()) {
      let r = param as u32 + *row_offset; // 行号需要加上偏移量
      r
    } else {
      *row_inline + *row_offset // 使用当前行号
    };
    
    out.write(&format!("{c_str}{r_num}"))?;
    
    Ok(())
  }));
  
  // 标记删除行的 helper
  // 用法: {{#each items}}...{{else}}<row><c><v>{{removeRow}}</v></c></row>{{/each}}
  // 重要: 会减少 row_offset，确保后续行号正确
  handlebars.register_helper("removeRow", Box::new(move |_: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    // 减少行偏移量，因为这一行会被删除
    // 这样后续行号会自动减1，避免出现空白行
    let mut offset = row_offset_for_remove.lock().unwrap();
    if *offset > 0 {
      *offset -= 1;
    }
    out.write(REMOVE_ROW_KEY)?;
    Ok(())
  }));
  
  // 标记数字类型的 helper
  // 用法: <c r="{{_cr}}"><v>{{num some_value}}</v></c>
  handlebars.register_helper("num", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    out.write(TO_NUMBER_KEY)?; // 先写入标记，后续处理时替换
    if let Some(param) = h.param(0) {
      if param.value().is_number() {
        out.write(&param.value().to_string())?;
      } else if param.value().is_string() {
        let s = param.value().as_str().unwrap();
        if let Ok(n) = s.parse::<f64>() {
          out.write(&n.to_string())?;
        } else {
          out.write("0")?; // 解析失败则输出 0
        }
      } else {
        out.write("0")?; // 其他类型则输出 0
      }
    } else {
      out.write("0")?; // 没有参数则输出 0
    }
    Ok(())
  }));
  
  // 标记公式类型的 helper
  // 用法: <c r="{{_cr}}"><f>{{formula "=SUM(A1:B1)"}}</f></c>
  handlebars.register_helper("formula", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    out.write(TO_FORMULA_KEY)?; // 先写入标记，后续处理时替换
    if let Some(param) = h.param(0) {
      if param.value().is_string() {
        let formula = param.value().as_str().unwrap();
        out.write(formula)?;
      } else {
        out.write("")?; // 非字符串则输出空
      }
    } else {
      out.write("")?; // 没有参数则输出空
    }
    Ok(())
  }));
  
  // 字符串拼接 helper
  // 用法: {{concat "=SUM(" (_c) "1:" (_c) "10)"}}
  // 或者: {{formula (concat "=SUM(" (_c) "1:" (_c) "10)")}}
  // 可以接受任意数量的参数，将它们全部拼接成一个字符串
  handlebars.register_helper("concat", Box::new(|h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let mut result = String::new();
    
    // 遍历所有参数并拼接
    for param in h.params() {
      let value = param.value();
      
      // 根据类型转换为字符串
      if value.is_string() {
        result.push_str(value.as_str().unwrap());
      } else if value.is_number() {
        result.push_str(&value.to_string());
      } else if value.is_boolean() {
        result.push_str(&value.to_string());
      } else if value.is_null() {
        // null 不添加任何内容
      } else {
        // 其他类型（对象、数组等）转为 JSON 字符串
        result.push_str(&value.to_string());
      }
    }
    
    out.write(&result)?;
    Ok(())
  }));
  
  // 列名转换 helper - 将列索引转换为列名，支持偏移量
  // 用法: {{toColumnName "A" 5}} -> "F"
  // 用法: {{toColumnName (_c) 3}} -> 当前列向右偏移 3 列的列名
  handlebars.register_helper("toColumnName", Box::new(|h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    if let Some(current_col) = h.param(0) {
      let current_str = if current_col.value().is_string() {
        current_col.value().as_str().unwrap().to_string()
      } else if current_col.value().is_number() {
        // 如果传入的是数字，先转换为列名
        let col_num = current_col.value().as_u64().unwrap_or(1) as u32;
        to_column_name("A", col_num.saturating_sub(1))
      } else {
        "A".to_string()
      };
      
      let increment = if let Some(inc_param) = h.param(1) {
        inc_param.value().as_u64().unwrap_or(0) as u32
      } else {
        0
      };
      
      let result = to_column_name(&current_str, increment);
      out.write(&result)?;
    } else {
      out.write("A")?; // 默认返回 A
    }
    Ok(())
  }));
  
  // 列索引转换 helper - 将列名转换为列索引（1-based）
  // 用法: {{toColumnIndex "A"}} -> 1
  // 用法: {{toColumnIndex "Z"}} -> 26
  // 用法: {{toColumnIndex "AA"}} -> 27
  handlebars.register_helper("toColumnIndex", Box::new(|h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    if let Some(col_name) = h.param(0) {
      let col_str = if col_name.value().is_string() {
        col_name.value().as_str().unwrap()
      } else {
        "A" // 默认值
      };
      
      let index = crate::utils::to_column_index(col_str);
      out.write(&index.to_string())?;
    } else {
      out.write("1")?; // 默认返回 1
    }
    Ok(())
  }));
  
  // 合并单元格 mergeCells: [ "C4:D5", "F4:G4" ]
  let merge_cells: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
  let merge_cells2 = Arc::clone(&merge_cells);
  
  // 超链接信息收集（按 sheet 分组）
  let hyperlinks_by_sheet: Arc<Mutex<HashMap<String, Vec<crate::utils::HyperlinkInfo>>>> = Arc::new(Mutex::new(HashMap::new()));
  let hyperlinks_by_sheet2 = Arc::clone(&hyperlinks_by_sheet);
  let sheet_name_for_hyperlink = Arc::clone(&sheet_name);
  
  // 图片信息收集（按 sheet 分组）
  let images_by_sheet: Arc<Mutex<HashMap<String, Vec<ImageInfo>>>> = Arc::new(Mutex::new(HashMap::new()));
  let images_by_sheet2 = Arc::clone(&images_by_sheet);
  let sheet_name3 = Arc::clone(&sheet_name);
  
  // 注册 mergeCell helper - 用于收集需要合并的单元格范围
  // 用法: {{mergeCell "C4:D5"}} 或 {{mergeCell (concat (_c) (_r) ":" (toColumnName (_c) 3) (_r))}}
  handlebars.register_helper("mergeCell", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    if let Some(ref_value) = h.param(0) {
      if let Some(ref_str) = ref_value.value().as_str() {
        // 简单验证格式：应该包含冒号分隔符
        if ref_str.contains(':') {
          let mut cells = merge_cells2.lock().unwrap();
          cells.push(ref_str.to_string());
        }
      }
    }
    Ok(())
  }));
  
  // 注册 hyperlink helper - 用于在 Excel 中添加超链接
  // 新用法: {{hyperlink (_cr) "Sheet2!A1" "链接文本"}}
  // 参数1: ref - 单元格引用（如 "A26"），通常使用 (_cr) 自动计算
  // 参数2: location - 链接目标（必需）
  // 参数3: display - 显示文本（可选，默认为空）
  handlebars.register_helper("hyperlink", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    // 获取 ref 参数（单元格引用）
    let ref_cell = h.param(0).and_then(|v| v.value().as_str());
    if ref_cell.is_none() || ref_cell.unwrap().is_empty() {
      return Ok(()); // 没有 ref，直接返回
    }
    let ref_cell = ref_cell.unwrap().to_string();
    
    // 获取 location 参数
    let location = h.param(1).and_then(|v| v.value().as_str());
    if location.is_none() || location.unwrap().is_empty() {
      return Ok(()); // 没有链接目标，直接返回
    }
    let location = location.unwrap();
    
    // 获取 display 参数（可选）
    let display = h.param(2)
      .and_then(|v| v.value().as_str())
      .unwrap_or("")
      .to_string();
    
    // 获取当前 sheet 名称
    let current_sheet = sheet_name_for_hyperlink.lock().unwrap().clone();
    
    if !current_sheet.is_empty() {
      // 添加超链接信息
      hyperlinks_by_sheet2
        .lock().unwrap()
        .entry(current_sheet)
        .or_insert_with(Vec::new)
        .push(crate::utils::HyperlinkInfo {
          ref_cell,
          location: location.to_string(),
          display,
        });
    }
    
    Ok(()) // 不输出任何内容
  }));
  
  // 注册 img helper - 用于在 Excel 中插入图片
  // 用法: {{img "base64数据" 100 200}} 或 {{img image.data image.width image.height}}
  handlebars.register_helper("img", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    // 获取参数
    let base64_data = h.param(0).and_then(|v| v.value().as_str());
    if base64_data.is_none() || base64_data.unwrap().is_empty() {
      return Ok(()); // 没有图片数据，直接返回
    }
    let base64_data = base64_data.unwrap();
    
    let width = h.param(1).and_then(|v| v.value().as_u64()).map(|w| w as u32);
    let height = h.param(2).and_then(|v| v.value().as_u64()).map(|h| h as u32);
    
    // 获取当前单元格位置
    let col = *col_inline5.lock().unwrap() + *col_offset5.lock().unwrap();
    let row = *row_inline5.lock().unwrap() + *row_offset5.lock().unwrap();
    
    // 获取当前 sheet 名称
    let current_sheet = sheet_name3.lock().unwrap().clone();
    
    if !current_sheet.is_empty() {
      // 生成唯一的 rid（使用 UUID 避免与模板中现有 ID 冲突）
      let rid = Uuid::new_v4().to_string().replace("-", "");
      let rid = format!("rId{}", &rid[..16]); // 例如: rId1234567890abcdef
      
      // 添加图片信息
      images_by_sheet2
        .lock().unwrap()
        .entry(current_sheet)
        .or_insert_with(Vec::new)
        .push(ImageInfo {
          col,
          row,
          base64_data: base64_data.to_string(),
          width,
          height,
          rid,
        });
    }
    
    Ok(()) // 不输出任何内容
  }));
  
  // 用于收集需要删除的工作表路径
  let sheets_to_delete: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
  let sheets_to_delete2 = Arc::clone(&sheets_to_delete);
  let sheet_name4 = Arc::clone(&sheet_name);
  
  // 删除当前工作表的 helper
  // 用法: {{deleteCurrentSheet}} - 标记删除当前正在渲染的工作表
  // 注意: 
  // 1. 不能删除工作簿中的最后一个工作表（Excel 至少需要一个工作表）
  // 2. 会自动清理相关的 rels、drawings 等文件
  // 3. 建议在条件渲染中使用，如 {{#if shouldDelete}}{{deleteCurrentSheet}}{{/if}}
  handlebars.register_helper("deleteCurrentSheet", Box::new(move |_: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let current_sheet = sheet_name4.lock().unwrap().clone();
    if !current_sheet.is_empty() {
      sheets_to_delete2.lock().unwrap().push(current_sheet);
    }
    Ok(())
  }));
  
  // 用于收集需要重命名的工作表（sheet_path -> new_name）
  let sheets_to_rename: Arc<Mutex<HashMap<String, String>>> = Arc::new(Mutex::new(HashMap::new()));
  let sheets_to_rename2 = Arc::clone(&sheets_to_rename);
  let sheet_name5 = Arc::clone(&sheet_name);
  
  // 重命名当前工作表的 helper
  // 用法: {{setCurrentSheetName "新名称"}} 或 {{setCurrentSheetName (concat department.name " - " year)}}
  // 注意:
  // 1. 工作表名称不能包含：\ / ? * [ ]
  // 2. 名称长度不能超过 31 个字符
  // 3. 不能与现有工作表重名（会自动处理）
  handlebars.register_helper("setCurrentSheetName", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    if let Some(new_name) = h.param(0).and_then(|v| v.value().as_str()) {
      let current_sheet = sheet_name5.lock().unwrap().clone();
      if !current_sheet.is_empty() && !new_name.is_empty() {
        // 过滤非法字符并限制长度
        let clean_name: String = new_name
          .chars()
          .filter(|c| !matches!(c, '\\' | '/' | '?' | '*' | '[' | ']'))
          .take(31)
          .collect();
        
        if !clean_name.is_empty() {
          sheets_to_rename2.lock().unwrap().insert(current_sheet, clean_name);
        }
      }
    }
    Ok(())
  }));
  
  // 用于收集需要隐藏的工作表（sheet_path -> hide_type）
  let sheets_to_hide: Arc<Mutex<HashMap<String, String>>> = Arc::new(Mutex::new(HashMap::new()));
  let sheets_to_hide2 = Arc::clone(&sheets_to_hide);
  let sheet_name6 = Arc::clone(&sheet_name);
  
  // 隐藏当前工作表的 helper
  // 用法: {{hideCurrentSheet}} 或 {{hideCurrentSheet "veryHidden"}}
  // 隐藏级别:
  // - "hidden" (默认): 普通隐藏，用户可以通过右键菜单取消隐藏
  // - "veryHidden": 超级隐藏，需要 VBA 或属性编辑器才能取消隐藏
  handlebars.register_helper("hideCurrentSheet", Box::new(move |h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, _out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let hide_type = h.param(0)
      .and_then(|v| v.value().as_str())
      .unwrap_or("hidden");
    
    // 验证隐藏类型
    let hide_type = match hide_type {
      "veryHidden" => "veryHidden",
      _ => "hidden", // 默认为普通隐藏
    };
    
    let current_sheet = sheet_name6.lock().unwrap().clone();
    if !current_sheet.is_empty() {
      sheets_to_hide2.lock().unwrap().insert(current_sheet, hide_type.to_string());
    }
    Ok(())
  }));
  
  // 遍历 sheet.xml 找到所有 t="s" 的 c 标签, 把 v 标签中的数字替换成对应的字符串
  // 例如: <c r="A1" t="s"><v>0</v></c> 替换成 <c r="A1" t="inlineStr"><is><t>字符串内容</t></is></c>
  {
    let mut files = files.lock().unwrap();
    // 收集所有 sheet 文件名并排序
    let mut sheet_names: Vec<String> = files.keys()
      .filter(|name| name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml"))
      .cloned()
      .collect();
    sheet_names.sort();

    for sheet_name in sheet_names {
      if let Some(contents) = files.get_mut(&sheet_name) {
        // 设置全变量 sheet_name
        *sheet_name2.lock().unwrap() = sheet_name.clone();
        
        let xml_content = std::str::from_utf8(contents)?;
        let xml_content = "{{row_offset_reset}}".to_string() + xml_content;
        
        // 提取并移除模板中已有的 mergeCells 和 hyperlinks 标签
        // 这些静态的合并范围和超链接会被转换为 helper 调用
        let (xml_content, static_merge_refs, static_hyperlinks) = crate::utils::extract_and_remove_merge_cells_and_hyperlinks(&xml_content)?;
        
        // 在 sharedStrings 中注入 helper 调用
        // 因为 replace_shared_strings 会替换整个 <v> 内容，所以要在 sharedStrings 里注入
        let mut shared_strings_modified = shared_strings.clone();
        crate::utils::inject_helpers_into_shared_strings(
            &xml_content,
            &mut shared_strings_modified,
            &static_merge_refs,
            &static_hyperlinks,
        )?;
        
        // 第一步：替换 sharedStrings，将 t="s" 转换为 t="inlineStr"
        let xml_content = replace_shared_strings_in_sheet(&xml_content, &shared_strings_modified)?;
        
        // 第二步：合并被分割的 handlebars 语法
        let xml_content = merge_handlebars_in_xml(xml_content)?;
        
        // 渲染模板
        let mut xml_content = handlebars.render_template(
          &xml_content,
          &*data1.lock().map_err(|e| Box::new(std::io::Error::other(format!("Failed to lock data: {e}"))))?,
        ).map_err(|e| {
          let reason: &RenderErrorReason = e.reason();
          XlsxError::TemplateRenderError(reason.to_string())
        })?;
        
        // 后处理：删除标记行、转换数字类型、转换公式类型等
        if xml_content.contains(REMOVE_ROW_KEY) || xml_content.contains(TO_NUMBER_KEY) || xml_content.contains(TO_FORMULA_KEY) {
          let remove_key = if xml_content.contains(REMOVE_ROW_KEY) { Some(REMOVE_ROW_KEY) } else { None };
          let number_key = if xml_content.contains(TO_NUMBER_KEY) { Some(TO_NUMBER_KEY) } else { None };
          let formula_key = if xml_content.contains(TO_FORMULA_KEY) { Some(TO_FORMULA_KEY) } else { None };
          
          // 获取合并单元格信息
          let merge_refs = merge_cells.lock().unwrap().clone();
          
          // 获取超链接信息
          let hyperlinks_map = hyperlinks_by_sheet.lock().unwrap();
          let sheet_hyperlinks = hyperlinks_map.get(&sheet_name);
          
          xml_content = post_process_xml(
            &xml_content,
            remove_key,
            number_key,
            formula_key,
            if merge_refs.is_empty() { None } else { Some(&merge_refs) },
            sheet_hyperlinks.map(|v| v.as_slice()),
          )?;
        }
        
        *contents = xml_content.into_bytes();
      }
    }
    
    // 处理图片插入
    let images_map = images_by_sheet.lock().unwrap();
    if !images_map.is_empty() {
      process_images(&mut files, &images_map)?;
    }
    
    // 处理工作表删除
    let sheets_to_delete_list = sheets_to_delete.lock().unwrap().clone();
    if !sheets_to_delete_list.is_empty() {
      delete_sheets(&mut files, &sheets_to_delete_list)?;
    }
    
    // 处理工作表重命名
    let sheets_to_rename_map = sheets_to_rename.lock().unwrap().clone();
    if !sheets_to_rename_map.is_empty() {
      rename_sheets(&mut files, &sheets_to_rename_map)?;
    }
    
    // 处理工作表隐藏
    let sheets_to_hide_map = sheets_to_hide.lock().unwrap().clone();
    if !sheets_to_hide_map.is_empty() {
      hide_sheets(&mut files, &sheets_to_hide_map)?;
    }
  }
  
  // Extract files from Arc<Mutex<_>>
  let files = Arc::try_unwrap(files).map_err(|_| Box::new(std::io::Error::other("Failed to unwrap Arc")))?.into_inner().map_err(|e| Box::new(std::io::Error::other(format!("Failed to get inner value: {e:?}"))))?;
  
  // 重新压缩文件
  let mut output = Vec::new();
  {
    let cursor = Cursor::new(&mut output);
    let mut zip_writer = ZipWriter::new(cursor);
    
    for entry in files {
      let (file_name, contents): (String, Vec<u8>) = entry;
      let options = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .compression_level(Some(6)); // 设置压缩级别
      
      zip_writer.start_file(file_name, options)?;
      zip_writer.write_all(&contents)?;
    }
    
    zip_writer.finish()?;
  }
  
  Ok(output)
}

/// 处理图片插入：为每个 sheet 生成 drawing.xml 和 _rels 文件，保存图片到 media
fn process_images(
  files: &mut HashMap<String, Vec<u8>>,
  images_map: &HashMap<String, Vec<ImageInfo>>,
) -> Result<(), Box<dyn std::error::Error>> {
  use base64::Engine;
  
  let mut image_counter = 1; // 全局图片计数器
  
  for (sheet_path, images) in images_map {
    if images.is_empty() {
      continue;
    }
    
    // 从 "xl/worksheets/sheet1.xml" 提取 sheet 编号
    let sheet_num: u32 = sheet_path
      .trim_start_matches("xl/worksheets/sheet")
      .trim_end_matches(".xml")
      .parse()
      .unwrap_or(1);
    
    // 生成 drawing.xml
    let drawing_path = format!("xl/drawings/drawing{}.xml", sheet_num);
    let drawing_xml = generate_drawing_xml(images, &mut image_counter)?;
    files.insert(drawing_path, drawing_xml.into_bytes());
    
    // 生成 drawing.xml.rels
    let drawing_rels_path = format!("xl/drawings/_rels/drawing{}.xml.rels", sheet_num);
    let drawing_rels = generate_drawing_rels(images);
    files.insert(drawing_rels_path, drawing_rels.into_bytes());
    
    // 生成 sheet.xml.rels（建立 sheet 到 drawing 的关系）
    let sheet_rels_path = format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_num);
    let sheet_rels = generate_sheet_rels(sheet_num);
    files.insert(sheet_rels_path, sheet_rels.into_bytes());
    
    // 在 sheet.xml 中添加 <drawing r:id="rId1" /> 引用
    if let Some(sheet_content) = files.get_mut(sheet_path) {
      let mut xml = String::from_utf8(sheet_content.clone())?;
      
      // 在 </worksheet> 之前插入 <drawing> 标签
      if !xml.contains("<drawing") {
        xml = xml.replace("</worksheet>", "  <drawing r:id=\"rId1\" />\n</worksheet>");
        *sheet_content = xml.into_bytes();
      }
    }
    
    // 保存所有图片到 xl/media/
    for img_info in images.iter() {
      let image_data = base64::engine::general_purpose::STANDARD
        .decode(&img_info.base64_data)
        .map_err(|e| format!("Failed to decode base64 image: {}", e))?;
      
      // 使用 rid 作为文件名，确保唯一性
      let image_path = format!("xl/media/{}.png", img_info.rid);
      files.insert(image_path, image_data);
    }
  }
  
  // 更新 [Content_Types].xml 添加 PNG 类型和 drawing 类型
  if let Some(content_types) = files.get_mut("[Content_Types].xml") {
    let mut xml = String::from_utf8(content_types.clone())?;
    
    // 添加 PNG 扩展类型
    if !xml.contains("Extension=\"png\"") {
      xml = xml.replace(
        "</Types>",
        "  <Default Extension=\"png\" ContentType=\"image/png\"/>\n</Types>",
      );
    }
    
    // 为每个 drawing.xml 添加 Override 声明
    for (sheet_path, images) in images_map {
      if !images.is_empty() {
        let sheet_num: u32 = sheet_path
          .trim_start_matches("xl/worksheets/sheet")
          .trim_end_matches(".xml")
          .parse()
          .unwrap_or(1);
        
        let drawing_part_name = format!("/xl/drawings/drawing{}.xml", sheet_num);
        if !xml.contains(&drawing_part_name) {
          xml = xml.replace(
            "</Types>",
            &format!(
              "  <Override PartName=\"{}\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>\n</Types>",
              drawing_part_name
            ),
          );
        }
      }
    }
    
    *content_types = xml.into_bytes();
  }
  
  Ok(())
}

/// 生成 drawing.xml 内容
fn generate_drawing_xml(
  images: &[ImageInfo],
  image_counter: &mut usize,
) -> Result<String, Box<dyn std::error::Error>> {
  let mut xml = String::from(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
"#,
  );
  
  for img_info in images.iter() {
    // 解码图片数据以获取实际尺寸
    use base64::Engine;
    let image_data = base64::engine::general_purpose::STANDARD
      .decode(&img_info.base64_data)
      .map_err(|e| format!("Failed to decode base64: {}", e))?;
    
    let (actual_width, actual_height) = get_image_dimensions(&image_data)
      .ok_or("Failed to detect image dimensions")?;
    
    // 使用用户指定尺寸或实际尺寸
    let width_px = img_info.width.unwrap_or(actual_width);
    let height_px = img_info.height.unwrap_or(actual_height);
    
    // 转换为 EMU (1 px = 9525 EMU)
    let width_emu = width_px as i64 * 9525;
    let height_emu = height_px as i64 * 9525;
    
    // 使用 oneCellAnchor 模式：只指定起始位置和绝对尺寸，不受单元格大小限制
    let from_col = img_info.col - 1; // 转换为 0-based
    let from_row = img_info.row - 1;
    
    xml.push_str(&format!(
      r#"  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>{}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>{}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="{}" cy="{}"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="{}" name="Picture {}"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{}"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="{}" cy="{}"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
"#,
      from_col,      // from col
      from_row,      // from row
      width_emu,     // ext cx (绝对宽度)
      height_emu,    // ext cy (绝对高度)
      *image_counter, // cNvPr id
      *image_counter, // Picture name
      &img_info.rid, // rId (使用 UUID 生成的唯一 ID)
      width_emu,     // xfrm ext cx
      height_emu,    // xfrm ext cy
    ));
    
    *image_counter += 1;
  }
  
  xml.push_str("</xdr:wsDr>");
  Ok(xml)
}

/// 生成 drawing.xml.rels 内容
fn generate_drawing_rels(images: &[ImageInfo]) -> String {
  let mut xml = String::from(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#,
  );
  
  for img_info in images {
    xml.push_str(&format!(
      r#"  <Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{}.png"/>
"#,
      img_info.rid, img_info.rid
    ));
  }
  
  xml.push_str("</Relationships>");
  xml
}

/// 生成 sheet.xml.rels 内容（建立 sheet 到 drawing 的关系）
fn generate_sheet_rels(sheet_num: u32) -> String {
  format!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{}.xml"/>
</Relationships>"#,
    sheet_num
  )
}

/// 删除指定的工作表及其关键文件
/// 
/// 删除工作表包括以下步骤：
/// 1. 从 workbook.xml 中删除 <sheet> 节点
/// 2. 从 workbook.xml.rels 中删除对应的 Relationship
/// 3. 删除 worksheet 文件本身 (xl/worksheets/sheet{N}.xml)
/// 4. 删除相关的 rels 文件 (xl/worksheets/_rels/sheet{N}.xml.rels)
/// 5. 从 [Content_Types].xml 中删除 worksheet 的 Override 声明
/// 
/// 注意：
/// - 不能删除最后一个工作表，Excel 工作簿必须至少包含一个工作表
/// - 不删除 drawing 文件，避免潜在的引用关系问题，且图片数据不敏感
fn delete_sheets(
  files: &mut HashMap<String, Vec<u8>>,
  sheets_to_delete: &[String],
) -> Result<(), Box<dyn std::error::Error>> {
  if sheets_to_delete.is_empty() {
    return Ok(());
  }
  
  // 1. 解析 workbook.xml 获取所有工作表信息
  let workbook_path = "xl/workbook.xml";
  let workbook_content = files.get(workbook_path)
    .ok_or("workbook.xml not found")?;
  let mut workbook_xml = String::from_utf8(workbook_content.clone())?;
  
  // 2. 解析 workbook.xml.rels 获取关系映射
  let workbook_rels_path = "xl/_rels/workbook.xml.rels";
  let workbook_rels_content = files.get(workbook_rels_path)
    .ok_or("workbook.xml.rels not found")?;
  let mut workbook_rels_xml = String::from_utf8(workbook_rels_content.clone())?;
  
  // 3. 统计总工作表数量
  let total_sheets = workbook_xml.matches("<sheet ").count();
  
  // 4. 检查是否会删除所有工作表
  if sheets_to_delete.len() >= total_sheets {
    return Err(Box::new(std::io::Error::other(
      "Cannot delete all worksheets. Excel workbook must contain at least one worksheet."
    )));
  }
  
  // 5. 对每个要删除的工作表进行处理
  for sheet_path in sheets_to_delete {
    // 从路径提取 sheet 编号: "xl/worksheets/sheet1.xml" -> "1"
    let sheet_num: u32 = match sheet_path
      .trim_start_matches("xl/worksheets/sheet")
      .trim_end_matches(".xml")
      .parse() {
        Ok(num) => num,
        Err(_) => continue, // 无法解析编号，跳过此工作表
      };
    
    // 5.1 从 workbook.xml.rels 中找到对应的 rId
    let rels_target = format!("worksheets/sheet{}.xml", sheet_num);
    let mut rid = String::new();
    
    // 查找并删除对应的 Relationship
    if let Some(rel_start) = workbook_rels_xml.find(&format!("Target=\"{}\"", rels_target)) {
      // 向前查找 Id="rIdXXX"
      let before = &workbook_rels_xml[..rel_start];
      if let Some(id_start) = before.rfind("Id=\"") {
        let id_part = &workbook_rels_xml[id_start + 4..];
        if let Some(id_end) = id_part.find('"') {
          rid = id_part[..id_end].to_string();
        }
      }
      
      // 删除整个 Relationship 节点
      if let Some(node_start) = before.rfind("<Relationship ") {
        if let Some(node_end) = workbook_rels_xml[rel_start..].find("/>") {
          let full_end = rel_start + node_end + 2;
          // 删除节点，包括前后的空白
          let mut delete_start = node_start;
          let mut delete_end = full_end;
          
          // 删除前面的空白和换行
          while delete_start > 0 && matches!(workbook_rels_xml.as_bytes()[delete_start - 1], b' ' | b'\t' | b'\r' | b'\n') {
            delete_start -= 1;
          }
          
          // 删除后面的空白和换行（保留一个换行）
          while delete_end < workbook_rels_xml.len() && matches!(workbook_rels_xml.as_bytes()[delete_end], b' ' | b'\t') {
            delete_end += 1;
          }
          if delete_end < workbook_rels_xml.len() && workbook_rels_xml.as_bytes()[delete_end] == b'\n' {
            delete_end += 1;
          }
          
          workbook_rels_xml.replace_range(delete_start..delete_end, "");
        }
      }
    }
    
    // 5.2 从 workbook.xml 中删除对应的 <sheet> 节点
    if !rid.is_empty() {
      let sheet_pattern = format!("r:id=\"{}\"", rid);
      if let Some(sheet_pos) = workbook_xml.find(&sheet_pattern) {
        // 向前查找 <sheet 标签的开始
        let before = &workbook_xml[..sheet_pos];
        if let Some(tag_start) = before.rfind("<sheet ") {
          // 向后查找 /> 或 </sheet>
          let after = &workbook_xml[sheet_pos..];
          if let Some(tag_end) = after.find("/>") {
            let full_end = sheet_pos + tag_end + 2;
            
            // 删除节点，包括前后的空白
            let mut delete_start = tag_start;
            let mut delete_end = full_end;
            
            // 删除前面的空白和换行
            while delete_start > 0 && matches!(workbook_xml.as_bytes()[delete_start - 1], b' ' | b'\t' | b'\r' | b'\n') {
              delete_start -= 1;
            }
            
            // 删除后面的空白和换行（保留一个换行）
            while delete_end < workbook_xml.len() && matches!(workbook_xml.as_bytes()[delete_end], b' ' | b'\t') {
              delete_end += 1;
            }
            if delete_end < workbook_xml.len() && workbook_xml.as_bytes()[delete_end] == b'\n' {
              delete_end += 1;
            }
            
            workbook_xml.replace_range(delete_start..delete_end, "");
          }
        }
      }
    }
    
    // 5.3 删除工作表文件本身
    files.remove(sheet_path);
    
    // 5.4 删除相关的 rels 文件
    let sheet_rels = format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_num);
    files.remove(&sheet_rels);
    
    // 注意：不删除 xl/drawings/drawing{N}.xml 和 xl/drawings/_rels/drawing{N}.xml.rels
    // 原因：避免潜在的引用关系问题，且图片形状数据不敏感，保留不影响 Excel 显示
    
    // 5.5 从 [Content_Types].xml 中删除 worksheet 的 Override 声明
    if let Some(content_types) = files.get_mut("[Content_Types].xml") {
      let mut ct_xml = String::from_utf8(content_types.clone())?;
      
      // 只删除 worksheet 的 Override 声明
      let worksheet_override = format!(
        "  <Override PartName=\"/xl/worksheets/sheet{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n",
        sheet_num
      );
      ct_xml = ct_xml.replace(&worksheet_override, "");
      
      // 不删除 drawing 的 Override，保留 drawing 文件
      
      *content_types = ct_xml.into_bytes();
    }
  }
  
  // 6. 更新修改后的文件
  files.insert(workbook_path.to_string(), workbook_xml.into_bytes());
  files.insert(workbook_rels_path.to_string(), workbook_rels_xml.into_bytes());
  
  Ok(())
}

/// 重命名指定的工作表
/// 
/// 重命名工作表包括以下步骤：
/// 1. 在 workbook.xml 中找到对应的 <sheet> 节点
/// 2. 修改 name 属性为新名称
/// 
/// 注意事项：
/// - 工作表名称会自动过滤非法字符：\ / ? * [ ]
/// - 名称长度会自动限制在 31 个字符以内
/// - 如果新名称与现有工作表重名，会自动添加数字后缀
fn rename_sheets(
  files: &mut HashMap<String, Vec<u8>>,
  sheets_to_rename: &HashMap<String, String>,
) -> Result<(), Box<dyn std::error::Error>> {
  if sheets_to_rename.is_empty() {
    return Ok(());
  }
  
  // 1. 解析 workbook.xml
  let workbook_path = "xl/workbook.xml";
  let workbook_content = files.get(workbook_path)
    .ok_or("workbook.xml not found")?;
  let mut workbook_xml = String::from_utf8(workbook_content.clone())?;
  
  // 2. 收集所有现有的工作表名称（用于检测重名）
  let mut existing_names: Vec<String> = Vec::new();
  let mut start = 0;
  while let Some(name_pos) = workbook_xml[start..].find("<sheet ") {
    let abs_pos = start + name_pos;
    if let Some(name_start) = workbook_xml[abs_pos..].find("name=\"") {
      let name_abs_start = abs_pos + name_start + 6; // "name=\"".len()
      if let Some(name_end) = workbook_xml[name_abs_start..].find('"') {
        let name = workbook_xml[name_abs_start..name_abs_start + name_end].to_string();
        existing_names.push(name);
        start = name_abs_start + name_end;
      } else {
        break;
      }
    } else {
      break;
    }
  }
  
  // 3. 对每个要重命名的工作表进行处理
  for (sheet_path, new_name) in sheets_to_rename {
    // 从路径提取 sheet 编号
    let sheet_num: u32 = match sheet_path
      .trim_start_matches("xl/worksheets/sheet")
      .trim_end_matches(".xml")
      .parse() {
        Ok(num) => num,
        Err(_) => continue,
      };
    
    // 生成唯一的新名称（如果重名则添加后缀）
    let mut final_name = new_name.clone();
    let mut counter = 1;
    while existing_names.contains(&final_name) {
      // 限制名称+后缀的总长度不超过 31
      let suffix = format!(" ({})", counter);
      let max_base_len = 31 - suffix.len();
      let base = if new_name.len() > max_base_len {
        &new_name[..max_base_len]
      } else {
        new_name
      };
      final_name = format!("{}{}", base, suffix);
      counter += 1;
      
      // 防止无限循环
      if counter > 100 {
        final_name = format!("Sheet{}", sheet_num);
        break;
      }
    }
    
    // 在 workbook.xml 中查找并替换工作表名称
    // 需要找到对应 sheet 编号的 <sheet> 节点
    let sheet_id_pattern = format!("sheetId=\"{}\"", sheet_num);
    if let Some(sheet_id_pos) = workbook_xml.find(&sheet_id_pattern) {
      // 向前查找 <sheet 标签的开始
      let before = &workbook_xml[..sheet_id_pos];
      if let Some(tag_start) = before.rfind("<sheet ") {
        // 向后查找 />
        let after = &workbook_xml[sheet_id_pos..];
        if let Some(tag_end) = after.find("/>") {
          let full_end = sheet_id_pos + tag_end + 2;
          let sheet_tag = &workbook_xml[tag_start..full_end];
          
          // 在这个标签中查找并替换 name 属性
          if let Some(name_start) = sheet_tag.find("name=\"") {
            let name_abs_start = tag_start + name_start + 6;
            if let Some(name_end) = workbook_xml[name_abs_start..].find('"') {
              let name_abs_end = name_abs_start + name_end;
              let old_name = workbook_xml[name_abs_start..name_abs_end].to_string();
              
              // 替换名称
              workbook_xml.replace_range(name_abs_start..name_abs_end, &final_name);
              
              // 更新 existing_names 列表
              if let Some(pos) = existing_names.iter().position(|n| n == &old_name) {
                existing_names[pos] = final_name.clone();
              }
            }
          }
        }
      }
    }
  }
  
  // 4. 更新 workbook.xml
  files.insert(workbook_path.to_string(), workbook_xml.into_bytes());
  
  Ok(())
}

/// 隐藏指定的工作表
/// 
/// 隐藏工作表包括以下步骤：
/// 1. 在 workbook.xml 中找到对应的 <sheet> 节点
/// 2. 添加或修改 state 属性为 "hidden" 或 "veryHidden"
/// 
/// 隐藏级别：
/// - "hidden": 普通隐藏，用户可以通过右键菜单 → 取消隐藏
/// - "veryHidden": 超级隐藏，需要 VBA 代码或属性编辑器才能取消隐藏
/// 
/// 注意：至少要保留一个可见的工作表，否则 Excel 会报错
fn hide_sheets(
  files: &mut HashMap<String, Vec<u8>>,
  sheets_to_hide: &HashMap<String, String>,
) -> Result<(), Box<dyn std::error::Error>> {
  if sheets_to_hide.is_empty() {
    return Ok(());
  }
  
  // 1. 解析 workbook.xml
  let workbook_path = "xl/workbook.xml";
  let workbook_content = files.get(workbook_path)
    .ok_or("workbook.xml not found")?;
  let mut workbook_xml = String::from_utf8(workbook_content.clone())?;
  
  // 2. 统计总工作表数量和已隐藏的数量
  let total_sheets = workbook_xml.matches("<sheet ").count();
  
  // 3. 检查是否会隐藏所有工作表
  if sheets_to_hide.len() >= total_sheets {
    return Err(Box::new(std::io::Error::other(
      "Cannot hide all worksheets. Excel workbook must have at least one visible worksheet."
    )));
  }
  
  // 4. 对每个要隐藏的工作表进行处理
  for (sheet_path, hide_type) in sheets_to_hide {
    // 从路径提取 sheet 编号
    let sheet_num: u32 = match sheet_path
      .trim_start_matches("xl/worksheets/sheet")
      .trim_end_matches(".xml")
      .parse() {
        Ok(num) => num,
        Err(_) => continue,
      };
    
    // 在 workbook.xml 中查找对应的 <sheet> 节点
    let sheet_id_pattern = format!("sheetId=\"{}\"", sheet_num);
    if let Some(sheet_id_pos) = workbook_xml.find(&sheet_id_pattern) {
      // 向前查找 <sheet 标签的开始
      let before = &workbook_xml[..sheet_id_pos];
      if let Some(tag_start) = before.rfind("<sheet ") {
        // 向后查找 />
        let after = &workbook_xml[sheet_id_pos..];
        if let Some(tag_end) = after.find("/>") {
          let full_end = sheet_id_pos + tag_end + 2;
          let sheet_tag = &workbook_xml[tag_start..full_end];
          
          // 检查是否已经有 state 属性
          if sheet_tag.contains("state=") {
            // 已有 state 属性，替换它
            if let Some(state_start) = sheet_tag.find("state=\"") {
              let state_abs_start = tag_start + state_start + 7; // "state=\"".len()
              if let Some(state_end) = workbook_xml[state_abs_start..].find('"') {
                let state_abs_end = state_abs_start + state_end;
                workbook_xml.replace_range(state_abs_start..state_abs_end, hide_type);
              }
            }
          } else {
            // 没有 state 属性，在 /> 之前添加
            let insert_pos = full_end - 2; // 在 /> 的 / 之前
            workbook_xml.insert_str(insert_pos, &format!(" state=\"{}\"", hide_type));
          }
        }
      }
    }
  }
  
  // 5. 更新 workbook.xml
  files.insert(workbook_path.to_string(), workbook_xml.into_bytes());
  
  Ok(())
}
