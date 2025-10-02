use serde_json::Value;
use std::{io::{Cursor, Read, Write}, sync::{Arc, Mutex}};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};
use std::collections::HashMap;
use crate::{utils::{excel_column_name, merge_handlebars_in_xml, register_basic_helpers, post_process_xml, replace_shared_strings_in_sheet, validate_xlsx_format}, XlsxError};

use handlebars::{Handlebars, RenderErrorReason};
// use uuid::Uuid;
// use base64::{Engine as _, engine::general_purpose};

/// 用于标记需要删除的行的 UUID
/// 配合 {{removeRow}} helper 使用
const REMOVE_ROW_KEY: &str = "|e5nBk+z4RMKqlyBo+xQ48A-remove-row|";

/// 用于标记数字类型的 UUID
/// 配合 {{num aa}} helper 使用
const TO_NUMBER_KEY: &str = "|e5nBk+z4RMKqlyBo+xQ48A-num|";

/// 用于标记公式类型的 UUID
/// 配合 {{formula "=SUM(A1:B1)"}} helper 使用
const TO_FORMULA_KEY: &str = "|e5nBk+z4RMKqlyBo+xQ48A-formula|";

pub fn render_handlebars(
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
  
  // 行号偏移量
  let row_offset: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
  let row_offset2 = Arc::clone(&row_offset);
  let row_offset3 = Arc::clone(&row_offset);
  let row_offset4 = Arc::clone(&row_offset);
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
  
  // 当前行号
  let row_inline = Arc::new(Mutex::new(1u32));
  let row_inline2 = Arc::clone(&row_inline);
  let row_inline3 = Arc::clone(&row_inline);
  let row_inline4 = Arc::clone(&row_inline);
  
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
  
  // 当前列号
  let col_inline = Arc::new(Mutex::new(1u32));
  let col_inline2 = Arc::clone(&col_inline);
  let col_inline3 = Arc::clone(&col_inline);
  let col_inline4 = Arc::clone(&col_inline);
  
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
    let c_str = excel_column_name("A", c_num - 1); // 列号从 1 开始, 需要减 1
    out.write(&c_str)?;
    Ok(())
  }));
  
  handlebars.register_helper("_cr", Box::new(move |_h: &handlebars::Helper, _: &Handlebars, _: &handlebars::Context, _: &mut handlebars::RenderContext, out: &mut dyn handlebars::Output| -> handlebars::HelperResult {
    let col_inline = col_inline4.lock().unwrap();
    let col_offset = col_offset4.lock().unwrap();
    let row_inline = row_inline4.lock().unwrap();
    let row_offset = row_offset4.lock().unwrap();
    let c_num = *col_inline + *col_offset;
    let c_str = excel_column_name("A", c_num - 1); // 列号从 1 开始, 需要减 1
    let r_num = *row_inline + *row_offset;  // 行号需要加上偏移量
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
        let xml_content = std::str::from_utf8(contents)?;
        let xml_content = "{{row_offset_reset}}".to_string() + xml_content;
        let xml_content = replace_shared_strings_in_sheet(&xml_content, &shared_strings)?;
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
          xml_content = post_process_xml(&xml_content, remove_key, number_key, formula_key)?;
        }
        
        *contents = xml_content.into_bytes();
      }
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
