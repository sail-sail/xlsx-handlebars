use serde_json::Value;
use std::{io::{Cursor, Read, Write}, sync::{Arc, Mutex}};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};
use std::collections::HashMap;
use crate::{utils::{excel_column_name, merge_handlebars_in_xml, register_basic_helpers, post_process_xml, replace_shared_strings_in_sheet, validate_xlsx_format}, XlsxError};
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
  
  // sheet_name
  let sheet_name = Arc::new(Mutex::new(String::new()));
  let sheet_name2 = Arc::clone(&sheet_name);
  
  // 行号偏移量
  let row_offset: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
  let row_offset2 = Arc::clone(&row_offset);
  let row_offset3 = Arc::clone(&row_offset);
  let row_offset4 = Arc::clone(&row_offset);
  let row_offset5 = Arc::clone(&row_offset);
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
        excel_column_name("A", col_num.saturating_sub(1))
      } else {
        "A".to_string()
      };
      
      let increment = if let Some(inc_param) = h.param(1) {
        inc_param.value().as_u64().unwrap_or(0) as u32
      } else {
        0
      };
      
      let result = excel_column_name(&current_str, increment);
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
      
      let index = crate::utils::excel_column_index(col_str);
      out.write(&index.to_string())?;
    } else {
      out.write("1")?; // 默认返回 1
    }
    Ok(())
  }));
  
  // 合并单元格 mergeCells: [ "C4:D5", "F4:G4" ]
  let merge_cells: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
  let merge_cells2 = Arc::clone(&merge_cells);
  
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
          
          // 获取合并单元格信息
          let merge_refs = merge_cells.lock().unwrap().clone();
          
          xml_content = post_process_xml(
            &xml_content,
            remove_key,
            number_key,
            formula_key,
            if merge_refs.is_empty() { None } else { Some(&merge_refs) },
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
