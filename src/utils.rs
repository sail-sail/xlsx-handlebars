use quick_xml::{Reader, Writer, events::Event};

use std::io::{Cursor, Write};
use crate::errors::XlsxError;

/// 验证 XLSX 文件格式
/// 检查文件是否为有效的 ZIP 格式，并包含必需的 XLSX 文件结构
pub(crate) fn validate_xlsx_format(file_data: &[u8]) -> Result<(), XlsxError> {
    // 检查文件大小
    if file_data.len() < 22 {
        return Err(XlsxError::InvalidZipFormat);
    }
    
    // 检查 ZIP 文件签名
    // ZIP 文件的签名通常是 0x504B0304 (PK..) 或 0x504B0506 (PK.. 空文件)
    // 或者 0x504B0708 (PK.. 分割压缩包)
    let signature = u32::from_le_bytes([
        file_data[0], file_data[1], file_data[2], file_data[3]
    ]);
    
    match signature {
        0x04034b50 | 0x06054b50 | 0x08074b50 => {
            // 有效的 ZIP 签名
        },
        _ => return Err(XlsxError::InvalidZipFormat),
    }
    
    Ok(())
}

// /// XML 转义符转换成正常字符
// pub fn xml_escape_to_normal(xml_content: String) -> String {
//     xml_content
//         .replace("&lt;", "<")
//         .replace("&gt;", ">")
//         .replace("&amp;", "&")
//         .replace("&quot;", "\"")
//         .replace("&apos;", "'")
// }

/**
 * 1. 合并所有的 handlebars, 因为 handlebars 语法内部的字体颜色啥的是没有意义的, 内部的xml标签没有意义合并成普通文本
 * 2. 寻找所有的 {{#each 循环, 看看第1个each是否跨行, 是否跨列, 如果跨行了, 那全局map存储第1个each是row类型的, 就跨行的,  否则只有跨列的, 那就是 c 类型的
 *
 * 3. 区分好所有each之后, 再次遍历标签, 这个时候, 我就能区分不同的循环类型, 就可以对行号和列号的累加做累加逻辑了
 */
#[derive(Debug, PartialEq, Eq, Clone)]
pub(crate) enum EachType {
    Row, // 按行循环
    Col, // 按列循环
    None, // 即非行也非列循环
}

/// Each 块的信息，包含类型和变量名
#[derive(Debug, Clone)]
pub(crate) struct EachBlockInfo {
    each_type: EachType,
    var_name: String, // {{#each 后面的变量名
    start_row: Option<u32>, // {{#each 时的行号
    end_row: Option<u32>,   // {{/each}} 时的行号
    start_col: Option<u32>, // {{#each 时的列号
    end_col: Option<u32>,   // {{/each}} 时的列号
}

/// 合并被XML标签分割的Handlebars语法
/// 
/// 这个函数会识别被XML标签分割的 Handlebars 表达式并将其合并。
/// 例如: `<w:t>{</w:t><w:t>{name</w:t><w:t>}</w:t><w:t>}</w:t>` 
/// 会被合并为: `{{name}}`
pub(crate) fn merge_handlebars_in_xml(xml_content: String) -> Result<String, Box<dyn std::error::Error>> {
    // 快速检查：如果内容中没有大括号，直接返回原内容
    if !xml_content.contains('{') {
        return Ok(xml_content);
    }
    
    let mut each_block_stack = Vec::<EachBlockInfo>::new();
    
    // 创建XML阅读器和写入器
    let mut reader = Reader::from_str(&xml_content);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    
    // 用于累积文本内容的缓冲区
    let mut text_buffer = String::new();
    
    // 跟踪大括号状态
    let mut brace_count = 0;        // 当前大括号的数量
    let mut in_handlebars = false;  // 是否在完整的 handlebars 表达式中 (如 {{...}})
    
    // 当前行号(用于跟踪 each 块的行范围)
    let mut current_row: u32 = 0;
    
    // 当前列号(用于跟踪 each 块的列范围)
    let mut current_col: u32 = 0;
    
    loop {
        match reader.read_event_into(&mut buf) {
            // 处理文本节点
            Ok(Event::Text(ref e)) => {
                let text = std::str::from_utf8(e)?;
                
                // 先将 XML 转义符转换回正常字符
                // 这样 Handlebars 才能正确解析语法，例如：
                // {{formula &quot;=SUM(A1:B1)&quot;}} -> {{formula "=SUM(A1:B1)"}}
                let text = text
                    .replace("&lt;", "<")
                    .replace("&gt;", ">")
                    .replace("&amp;", "&")
                    .replace("&quot;", "\"")
                    .replace("&apos;", "'");
                
                // 逐字符分析文本，统计大括号
                for ch in text.chars() {
                    if ch == '{' {
                        brace_count += 1;
                        // 当遇到连续的两个 { 时，标记进入 handlebars 表达式
                        if brace_count >= 2 {
                            in_handlebars = true;
                        }
                    } else if ch == '}' {
                        if brace_count > 0 {
                            brace_count -= 1;
                        }
                        // 当大括号数量归零时，标记退出 handlebars 表达式
                        if brace_count == 0 {
                            in_handlebars = false;
                        }
                    }
                }
                
                // 将当前文本添加到缓冲区
                text_buffer.push_str(&text);
                
                // 如果不在 handlebars 表达式中且大括号已平衡，输出缓冲的文本
                // 否则继续累积文本，等待 handlebars 表达式完整
                if !in_handlebars && brace_count == 0 && !text_buffer.is_empty() {
                    // 如果 text_buffer 中包含 {{#each 或 {{/each}}
                    if text_buffer.contains("{{#each") {
                        // 可能包含多个 {{#each，需要逐个处理
                        let mut remaining = text_buffer.as_str();
                        while let Some(start_idx) = remaining.find("{{#each ") {
                            // 提取 {{#each 后面的变量名
                            // 例如: "{{#each projects}}" -> "projects"
                            // 或者: "{{#each items}}" -> "items"
                            let after_each = &remaining[start_idx + 8..]; // 跳过 "{{#each "
                            let var_name = after_each
                                .split('}')
                                .next()
                                .unwrap_or("")
                                .trim()
                                .to_string();
                            
                            // 将 EachBlockInfo 压入栈，记录开始行号和列号
                            each_block_stack.push(EachBlockInfo {
                                each_type: EachType::None, // 先占位，后续可以根据需要更新
                                var_name,
                                start_row: Some(current_row), // 记录当前行号
                                end_row: None,
                                start_col: Some(current_col), // 记录当前列号
                                end_col: None,
                            });
                            
                            // 继续查找下一个 {{#each
                            remaining = &after_each[1..];
                        }
                    }
                    if text_buffer.contains("{{/each}}") {
                        let count = text_buffer.matches("{{/each}}").count();
                        // 弹出对应的 each_block_stack
                        for _ in 0..count {
                            if let Some(mut block_info) = each_block_stack.pop() {
                                // 记录结束行号和列号
                                block_info.end_row = Some(current_row);
                                block_info.end_col = Some(current_col);
                                
                                // 使用提取的变量名
                                let _var_name = block_info.var_name;
                                
                                // 计算每个循环项的行偏移量
                                let row_offset_per_item = if let (Some(start), Some(end)) = (block_info.start_row, block_info.end_row) {
                                    end.saturating_sub(start)
                                } else {
                                    0
                                };
                                
                                // 计算每个循环项的列偏移量
                                let col_offset_per_item = if let (Some(start), Some(end)) = (block_info.start_col, block_info.end_col) {
                                    end.saturating_sub(start)
                                } else {
                                    0
                                };
                                
                                // 每个 block_info 对应一个 {{/each}} 标签, 每个 {{/each}} 标签前面
                                // 加上偏移量（循环结束后多出来的行数或列数）
                                if block_info.each_type == EachType::Row {
                                    // 如果是 Row 类型的 each, 则在 text_buffer 前面加上 row_offset_plus
                                    text_buffer = format!("{{{{row_offset_plus {row_offset_per_item}}}}}{text_buffer}");
                                } else if block_info.each_type == EachType::Col {
                                    // 如果是 Col 类型的 each, 则在 text_buffer 前面加上 col_offset_plus
                                    text_buffer = format!("{{{{col_offset_plus {col_offset_per_item}}}}}{text_buffer}");
                                }
                            } else {
                                break;
                            }
                        }
                    }
                    // 使用 from_escaped 避免 Writer 重复转义 (例如 " 变成 &quot;)
                    writer.write_event(Event::Text(quick_xml::events::BytesText::from_escaped(&text_buffer)))?;
                    text_buffer.clear();
                }
            }
            
            // 处理开始标签 (如 <w:t>)
            Ok(Event::Start(ref e)) => {
                // 如果在 handlebars 表达式中，跳过XML标签，只保留文本内容
                if !in_handlebars && brace_count == 0 {
                    // 先输出之前缓冲的文本
                    if !text_buffer.is_empty() {
                      // 使用 from_escaped 避免 Writer 重复转义
                      writer.write_event(Event::Text(quick_xml::events::BytesText::from_escaped(&text_buffer)))?;
                      text_buffer.clear();
                    }
                    let tag_name = e.name().as_ref().to_vec();
                    if tag_name == b"row" {
                      // 从 row 标签的 r 属性中提取行号
                      for attr in e.attributes().flatten() {
                        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                        if key == "r" {
                          let value = std::str::from_utf8(&attr.value).unwrap_or("0");
                          current_row = value.parse::<u32>().unwrap_or(0);
                          break;
                        }
                      }
                      
                      // 修改 each_block_stack 的最后一个元素的类型为 Row
                      if let Some(last) = each_block_stack.last_mut()
                        && (last.each_type == EachType::None || last.each_type == EachType::Col) {
                          last.each_type = EachType::Row;
                        }
                      // 创建新的开始标签，用于修改属性
                      let mut new_start = e.borrow();
                      new_start.clear_attributes(); // 清除现有属性
                      
                      for attr in e.attributes() {
                        let attr = attr?;
                        let key = std::str::from_utf8(attr.key.as_ref())?;
                        if key == "r" {
                          // 如果在 each 块内 的 c 标签，更新 r 属性的行号
                          let value = std::str::from_utf8(&attr.value)?;
                          let row_num = value.parse::<u32>().unwrap_or(0);
                          let value = format!("{{{{col_offset_reset}}}}{{{{set_row_inline {row_num}}}}}{{{{_r}}}}");
                          new_start.push_attribute((key.as_bytes(), value.as_bytes()));
                        } else {
                          // 保持其他属性不变
                          new_start.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
                        }
                      }
                      // 输出修改后的开始标签
                      writer.write_event(Event::Start(new_start))?;
                    } else if tag_name == b"c" {
                      // 从 c 标签的 r 属性中提取列号
                      for attr in e.attributes().flatten() {
                        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                        if key == "r" {
                          let value = std::str::from_utf8(&attr.value).unwrap_or("");
                          // 从 E7 中提取列字母部分
                          let r_char: String = value.chars().take_while(|c| c.is_alphabetic()).collect();
                          current_col = excel_column_index(&r_char);
                          break;
                        }
                      }
                      
                      // 修改 each_block_stack 的最后一个元素的类型为 Col
                      if let Some(last) = each_block_stack.last_mut()
                        && last.each_type == EachType::None {
                          last.each_type = EachType::Col;
                        }
                      // 创建新的开始标签，用于修改属性
                      let mut new_start = e.borrow();
                      new_start.clear_attributes(); // 清除现有属性
                      
                      for attr in e.attributes() {
                        let attr = attr?;
                        let key = std::str::from_utf8(attr.key.as_ref())?;
                        if key == "r" {
                          // 如果在 each 块内 的 c 标签，更新 r 属性的列号
                          let value = std::str::from_utf8(&attr.value)?;
                          // 从 E7 中提取列字母部分
                          let r_char: String = value.chars().take_while(|c| c.is_alphabetic()).collect();
                          let col_inline = excel_column_index(&r_char);
                          // println!("{r_char} -> {col_inline}");
                          let value = format!("{{{{set_col_inline {col_inline}}}}}{{{{_cr}}}}");
                          new_start.push_attribute((key.as_bytes(), value.as_bytes()));
                        } else {
                          // 保持其他属性不变
                          new_start.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
                        }
                      }
                      // 输出修改后的开始标签
                      writer.write_event(Event::Start(new_start))?;
                    } else {
                      // 输出原始开始标签
                      writer.write_event(Event::Start(e.clone()))?;
                    }
                }
                // 在 handlebars 表达式中时，忽略XML标签，这样就实现了"合并"效果
            }
            
            // 处理结束标签 (如 </w:t>)
            Ok(Event::End(ref e)) => {
                // 如果在 handlebars 表达式中，跳过XML标签
                if !in_handlebars && brace_count == 0 {
                    // 先输出之前缓冲的文本
                    if !text_buffer.is_empty() {
                        // 使用 from_escaped 避免 Writer 重复转义
                        writer.write_event(Event::Text(quick_xml::events::BytesText::from_escaped(&text_buffer)))?;
                        text_buffer.clear();
                    }
                    // let tag_name = e.name().as_ref().to_vec();
                    // 输出结束标签
                    writer.write_event(Event::End(e.clone()))?;
                }
            }
            
            // 处理自闭合标签 (如 <w:br/>)
            Ok(Event::Empty(ref e)) => {
                if !in_handlebars && brace_count == 0 {
                    if !text_buffer.is_empty() {
                        // 使用 from_escaped 避免 Writer 重复转义
                        writer.write_event(Event::Text(quick_xml::events::BytesText::from_escaped(&text_buffer)))?;
                        text_buffer.clear();
                    }
                    writer.write_event(Event::Empty(e.clone()))?;
                }
            }
            
            // 处理其他XML事件（注释、CDATA、处理指令等）
            Ok(Event::Eof) => break,
            Ok(event) => {
                if !in_handlebars && brace_count == 0 {
                    if !text_buffer.is_empty() {
                        // 使用 from_escaped 避免 Writer 重复转义
                        writer.write_event(Event::Text(quick_xml::events::BytesText::from_escaped(&text_buffer)))?;
                        text_buffer.clear();
                    }
                    writer.write_event(event)?;
                }
            }
            
            // XML解析错误
            Err(e) => return Err(format!("XML解析错误 at position {}: {:?}", reader.buffer_position(), e).into()),
        }
        buf.clear();
    }
    
    // 输出剩余的文本缓冲（如果有的话）
    if !text_buffer.is_empty() {
        // 使用 from_escaped 避免 Writer 重复转义
        writer.write_event(Event::Text(quick_xml::events::BytesText::from_escaped(&text_buffer)))?;
    }
    
    // 将结果转换为字符串返回
    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}

/// 找到所有 t="s" 的 c 标签, 把 v 标签中的数字替换成对应的字符串
/// 例如: <c r="A1" t="s"><v>0</v></c> 替换成 <c r="A1" t="inlineStr"><is><t>字符串内容</t></is></c>
pub(crate) fn replace_shared_strings_in_sheet(
  sheet_xml: &str,
  shared_strings: &[String]
) -> Result<String, Box<dyn std::error::Error>> {
  
  let mut reader = Reader::from_str(sheet_xml);
  let mut writer = Writer::new(Cursor::new(Vec::new()));
  let mut buf = Vec::new();
  
  // 跟踪当前状态
  let mut in_shared_string_cell = false;  // 是否在 t="s" 的 c 标签内
  let mut current_cell_attrs: Vec<(Vec<u8>, Vec<u8>)> = Vec::new(); // 当前 c 标签的属性
  let mut shared_string_v_content = String::new();      // v 标签的内容
  
  // 跟踪任意 c 标签内的状态
  let mut in_cell = false;           // 是否在任意 c 标签内
  let mut found_f_tag = false;           // 在当前 c 标签内是否找到了 f 标签
  let mut in_v_tag = false;          // 是否在任意 v 标签内
  
  loop {
    match reader.read_event_into(&mut buf) {
      // 处理开始标签
      Ok(Event::Start(ref e)) => {
        let tag_name = e.name().as_ref().to_vec();
        
        if tag_name == b"c" {
          // 进入任意 c 标签
          in_cell = true;
          found_f_tag = false; // 重置f标签标志
          
          // 检查是否有 t="s" 属性
          let mut has_shared_string = false;
          let mut attrs = Vec::new();
          
          for attr in e.attributes() {
            let attr = attr?;
            let key = std::str::from_utf8(attr.key.as_ref())?;
            let value = std::str::from_utf8(&attr.value)?;
            
            if key == "t" && value == "s" {
              has_shared_string = true;
              // 将 t="s" 替换为 t="inlineStr"
              attrs.push((attr.key.as_ref().to_vec(), b"inlineStr".to_vec()));
            } else {
              attrs.push((attr.key.as_ref().to_vec(), attr.value.to_vec()));
            }
          }
          
          if has_shared_string {
            in_shared_string_cell = true;
            current_cell_attrs = attrs;
          } else {
            // 正常输出 c 标签
            writer.write_event(Event::Start(e.clone()))?;
          }
        } else if tag_name == b"f" && in_cell {
          // 在任意 c 标签内遇到 f 标签，标记找到了f标签
          found_f_tag = true;
          // 正常输出 f 标签
          writer.write_event(Event::Start(e.clone()))?;
        } else if tag_name == b"v" && in_cell {
          in_v_tag = true;
          if in_shared_string_cell {
            // 在 shared string cell 内处理 v 标签
            if found_f_tag {
              // 如果已经找到了 f 标签，跳过 v 标签
            } else {
              shared_string_v_content.clear();
              // 不输出 v 标签，因为我们要替换成 is 标签
            }
          } else if found_f_tag {
            // 在普通 c 标签内，如果已经找到了 f 标签，跳过 v 标签
            // 不输出，不处理
          } else {
            // 在普通 c 标签内，没有 f 标签，正常输出 v 标签
            writer.write_event(Event::Start(e.clone()))?;
          }
        } else if in_shared_string_cell {
          // 在 shared string cell 内的其他标签，正常输出
          writer.write_event(Event::Start(e.clone()))?;
        } else {
          // 不在 shared string cell 内的标签，正常输出
          writer.write_event(Event::Start(e.clone()))?;
        }
      }
      
      // 处理结束标签
      Ok(Event::End(ref e)) => {
        let tag_name = e.name().as_ref().to_vec();
        
        if tag_name == b"c" {
          if in_shared_string_cell {
            // 输出修改后的 shared string c 标签及其内容
            let mut new_start = quick_xml::events::BytesStart::new("c");
            for (key, value) in &current_cell_attrs {
              new_start.push_attribute((key.as_slice(), value.as_slice()));
            }
            writer.write_event(Event::Start(new_start))?;
            
            // 如果有 v 标签内容且没有f标签，替换成对应的字符串
            if !shared_string_v_content.is_empty() && !found_f_tag
              && let Ok(index) = shared_string_v_content.parse::<usize>()
                && index < shared_strings.len() {
                  // 解析 shared_strings[index] 并作为 XML 事件插入，避免被转义
                  let si_content = &shared_strings[index];
                  let si_content = replace_shared_string_si_with_handlebars(si_content)?;
                  let mut is_reader = Reader::from_str(&si_content);
                  let mut is_buf = Vec::new();
                  loop {
                    match is_reader.read_event_into(&mut is_buf) {
                      Ok(Event::Eof) => break,
                      Ok(ev) => writer.write_event(ev)?,
                      Err(e) => return Err(format!("shared_string parse error: {:?}", e).into()),
                    }
                    is_buf.clear();
                  }
                }
            
            writer.write_event(Event::End(e.clone()))?;
            
            // 重置 shared string cell 状态
            in_shared_string_cell = false;
            current_cell_attrs.clear();
            shared_string_v_content.clear();
          } else {
            // 普通 c 标签结束，正常输出
            writer.write_event(Event::End(e.clone()))?;
          }
          
          // 重置 c 标签状态
          in_cell = false;
          found_f_tag = false;
        } else if tag_name == b"f" && in_cell {
          // 输出 f 结束标签
          writer.write_event(Event::End(e.clone()))?;
        } else if tag_name == b"v" && in_v_tag {
          in_v_tag = false;
          if in_shared_string_cell {
            // shared string cell 内的 v 标签结束
            if found_f_tag {
              // 如果已经找到了 f 标签，跳过 v 结束标签
            }
          } else if found_f_tag {
            // 普通 c 标签内，如果已经找到了 f 标签，跳过 v 结束标签
            // 不输出，不处理
          } else {
            // 在普通 c 标签内，没有 f 标签，正常输出 v 结束标签
            writer.write_event(Event::End(e.clone()))?;
          }
        } else if in_shared_string_cell {
          // 在 shared string cell 内的其他结束标签，正常输出
          writer.write_event(Event::End(e.clone()))?;
        } else {
          // 不在 shared string cell 内的结束标签，正常输出
          writer.write_event(Event::End(e.clone()))?;
        }
      }
      
      // 处理文本节点
      Ok(Event::Text(ref e)) => {
        if in_v_tag && found_f_tag {
          // 在有f标签的c标签内的v标签中，跳过文本内容
          // 不输出，不处理
        } else if in_v_tag && in_shared_string_cell {
          // 在 shared string cell 内的 v 标签中，收集内容
          let text = std::str::from_utf8(e)?;
          shared_string_v_content.push_str(text);
        } else {
          // 其他所有情况，正常输出文本
          writer.write_event(Event::Text(e.clone()))?;
        }
      }
      
      // 处理自闭合标签
      Ok(Event::Empty(ref e)) => {
        if !in_shared_string_cell {
          writer.write_event(Event::Empty(e.clone()))?;
        }
      }
      
      // 处理其他XML事件
      Ok(Event::Eof) => break,
      Ok(event) => {
        if !in_shared_string_cell {
          writer.write_event(event)?;
        }
      }
      
      // XML解析错误
      Err(e) => return Err(format!("XML解析错误 at position {}: {:?}", reader.buffer_position(), e).into()),
    }
    buf.clear();
  }
  
  // 将结果转换为字符串返回
  let result = writer.into_inner().into_inner();
  Ok(String::from_utf8(result)?)
}

pub(crate) fn replace_shared_string_si_with_handlebars(
  si_xml: &str
) -> Result<String, Box<dyn std::error::Error>> {
  
  let mut reader = Reader::from_str(si_xml);
  let mut writer = Writer::new(Cursor::new(Vec::new()));
  
  // 跟踪当前状态
  let mut in_r_tag = false;          // 是否在 r 标签内
  let mut in_t_tag = false;          // 是否在 t 标签内
  let mut t_text_content = String::new(); // 当前 t 标签内的文本内容
  let mut t_events: Vec<Event<'_>> = Vec::new(); // 当前 t 标签内的所有XML事件
  
  loop {
    match reader.read_event() {
      // 处理开始标签
      Ok(Event::Start(ref e)) => {
        let tag_name = e.name().as_ref().to_vec();
        
        if tag_name == b"r" {
          in_r_tag = true;
          writer.write_event(Event::Start(e.clone()))?;
        } else if tag_name == b"t" {
          in_t_tag = true;
          t_text_content.clear();
          t_events.clear();
          // 不立即输出 t 开始标签，先收集内容
        } else if in_t_tag {
          // 在 t 标签内，收集事件
          t_events.push(Event::Start(e.clone()));
        } else {
          // 直接输出
          writer.write_event(Event::Start(e.clone()))?;
        }
      }
      
      // 处理结束标签
      Ok(Event::End(ref e)) => {
        let tag_name = e.name().as_ref().to_vec();
        
        if tag_name == b"t" {
          in_t_tag = false;
          
          // 检查收集到的文本内容是否包含 Handlebars 循环语法
          let contains_each_start = t_text_content.contains("{{#each");
          let contains_each_end = t_text_content.contains("{{/each");
          
          if (contains_each_start || contains_each_end) && !in_r_tag {
            // 如果包含 each 语法且不在 r 标签内，需要包裹 r 标签
            writer.write_event(Event::Start(quick_xml::events::BytesStart::new("r")))?;
            writer.write_event(Event::Start(quick_xml::events::BytesStart::new("t")))?;
            
            // 输出收集的事件
            for t_event in t_events.iter() {
              writer.write_event(t_event.clone())?;
            }
            
            writer.write_event(Event::End(e.clone()))?; // t 结束标签
            writer.write_event(Event::End(quick_xml::events::BytesEnd::new("r")))?; // r 结束标签
          } else {
            // 正常输出 t 标签及其内容
            writer.write_event(Event::Start(quick_xml::events::BytesStart::new("t")))?;
            
            // 输出收集的事件
            for t_event in t_events.iter() {
              writer.write_event(t_event.clone())?;
            }
            
            writer.write_event(Event::End(e.clone()))?;
          }
        } else if tag_name == b"r" {
          in_r_tag = false;
          writer.write_event(Event::End(e.clone()))?;
        } else if in_t_tag {
          // 在 t 标签内，收集结束事件
          t_events.push(Event::End(e.clone()));
        } else {
          writer.write_event(Event::End(e.clone()))?;
        }
      }
      
      // 处理文本节点
      Ok(Event::Text(ref e)) => {
        if in_t_tag {
          let text = std::str::from_utf8(e)?;
          // 收集 t 标签内的文本内容
          t_text_content.push_str(text);
          // 收集文本事件
          t_events.push(Event::Text(e.clone()));
        } else {
          writer.write_event(Event::Text(e.clone()))?;
        }
      }
      
      // 处理自闭合标签
      Ok(Event::Empty(ref e)) => {
        if in_t_tag {
          t_events.push(Event::Empty(e.clone()));
        } else {
          writer.write_event(Event::Empty(e.clone()))?;
        }
      }
      
      // 处理其他XML事件
      Ok(Event::Eof) => break,
      Ok(event) => {
        if in_t_tag {
          t_events.push(event);
        } else {
          writer.write_event(event)?;
        }
      }
      // XML解析错误
      Err(e) => return Err(format!("XML解析错误 at position {}: {:?}", reader.buffer_position(), e).into()),
    }
  }
  
  // 将结果转换为字符串返回
  let result = writer.into_inner().into_inner();
  Ok(String::from_utf8(result)?)
}



/// 注册基础的 Handlebars helper 函数
pub(crate) fn register_basic_helpers(handlebars: &mut handlebars::Handlebars) -> Result<(), Box<dyn std::error::Error>> {
    use handlebars::handlebars_helper;
    use serde_json::Value;
    
    // 注册 eq helper (相等比较)
    handlebars_helper!(eq: |x: Value, y: Value| x == y);
    handlebars.register_helper("eq", Box::new(eq));
    
    // 注册 ne helper (不等比较)  
    handlebars_helper!(ne: |x: Value, y: Value| x != y);
    handlebars.register_helper("ne", Box::new(ne));
    
    // 注册 gt helper (大于)
    handlebars_helper!(gt: |x: i64, y: i64| x > y);
    handlebars.register_helper("gt", Box::new(gt));
    
    // 注册 lt helper (小于)
    handlebars_helper!(lt: |x: i64, y: i64| x < y);
    handlebars.register_helper("lt", Box::new(lt));
    
    // 注册 upper helper (转大写)
    handlebars_helper!(upper: |s: String| s.to_uppercase());
    handlebars.register_helper("upper", Box::new(upper));
    
    // 注册 lower helper (转小写)
    handlebars_helper!(lower: |s: String| s.to_lowercase());
    handlebars.register_helper("lower", Box::new(lower));
    
    // 注册 len helper (数组/字符串长度)
    handlebars_helper!(len: |x: Value| {
        match x {
            Value::Array(arr) => arr.len(),
            Value::String(s) => s.chars().count(),
            Value::Object(obj) => obj.len(),
            _ => 0
        }
    });
    handlebars.register_helper("len", Box::new(len));
    
    Ok(())
}

/// 在 Excel 的 sheet.xml 中列名
/// 传入当前列名和一个增量，返回新的列名
/// 用于生成 Excel 列名，如 A, B, ..., Z, AA, AB, ..., ZZ, AAA, ...
pub(crate) fn excel_column_name(current: &str, increment: u32) -> String {
  let mut col_index = 0;
  for (i, ch) in current.chars().rev().enumerate() {
    let ch_val = (ch as u8 - b'A' + 1) as u32;
    col_index += ch_val * 26_u32.pow(i as u32);
  }
  
  col_index += increment;
  
  let mut new_col_name = String::new();
  let mut n = col_index;
  while n > 0 {
    let rem = (n - 1) % 26;
    new_col_name.push((b'A' + rem as u8) as char);
    n = (n - 1) / 26;
  }
  
  new_col_name.chars().rev().collect()
}

/// 在 Excel 的 sheet.xml 中列名
/// 传入当前列名传入字母，返回对应的列索引 (1-based)
/// 用于生成 Excel 列名，如 A, B, ..., Z, AA, AB, ..., ZZ, AAA, ...
pub(crate) fn excel_column_index(col_name: &str) -> u32 {
  let mut col_index = 0;
  for (i, ch) in col_name.chars().rev().enumerate() {
    let ch_val = (ch as u8 - b'A' + 1) as u32;
    col_index += ch_val * 26_u32.pow(i as u32);
  }
  col_index
}

#[cfg(test)]
mod tests {
  use super::*;
  
  #[test]
  fn test_excel_column_name() {
    assert_eq!(excel_column_name("A", 0), "A");
    assert_eq!(excel_column_name("A", 1), "B");
    assert_eq!(excel_column_name("Z", 1), "AA");
    assert_eq!(excel_column_name("AA", 1), "AB");
    assert_eq!(excel_column_name("AZ", 1), "BA");
    assert_eq!(excel_column_name("ZZ", 1), "AAA");
    assert_eq!(excel_column_name("AAA", 26), "ABA");
  }
  
  #[test]
  fn test_replace_shared_string_si_with_handlebars() {
    // 测试包含 {{#each 且已经在 r 标签内的情况 - 应该保持原样
    let input_with_each_in_r = r#"<si>
  <r>
    <t>a</t>
  </r>
  <r>
    <rPr>
      <sz val="11" />
      <color rgb="FFFF0000" />
    </rPr>
    <t>{{#each projects}}</t>
  </r>
  <phoneticPr fontId="1" type="noConversion" />
</si>"#;
    
    let result = replace_shared_string_si_with_handlebars(input_with_each_in_r).unwrap();
    println!("输入包含 {{#each 且在 r 标签内的结果:");
    println!("{}", result);
    
    // 因为已经在 r 标签内，所以应该保持原样
    assert!(result.contains("{{#each projects}}"));
    assert!(result.contains("<r>"));
    assert!(result.contains("<t>{{#each projects}}</t>"));
    
    // 测试包含 {{#each 但不在 r 标签内的情况 - 应该被包裹
    let input_with_each_not_in_r = r#"<si>
  <t>{{#each projects}}</t>
  <phoneticPr fontId="1" type="noConversion" />
</si>"#;
    
    let result2 = replace_shared_string_si_with_handlebars(input_with_each_not_in_r).unwrap();
    println!("输入包含 {{#each 且不在 r 标签内的结果:");
    println!("{}", result2);
    // 应该被包裹在 r 标签内
    assert!(result2.contains("{{#each projects}}"));
    assert!(result2.contains("<r><t>{{#each projects}}</t></r>"));
    
    // 测试包含 {{/each}} 且不在 r 标签内的情况
    let input_with_end_each_not_in_r = r#"<si>
  <t>{{/each}}</t>
</si>"#;
    
    let result3 = replace_shared_string_si_with_handlebars(input_with_end_each_not_in_r).unwrap();
    println!("输入包含 {{/each}} 且不在 r 标签内的结果:");
    println!("{}", result3);
    // 应该被包裹在 r 标签内
    assert!(result3.contains("{{/each}}"));
    assert!(result3.contains("<r><t>{{/each}}</t></r>"));
    
    // 测试不包含 each 的情况 - 应该保持原样
    let input_normal = r#"<si>
  <r>
    <t>normal text</t>
  </r>
  <r>
    <rPr>
      <sz val="11" />
    </rPr>
    <t>{{name}}</t>
  </r>
</si>"#;
    
    let result4 = replace_shared_string_si_with_handlebars(input_normal).unwrap();
    println!("输入不包含 each 的结果:");
    println!("{}", result4);
    // 应该保持 r 标签结构
    assert!(result4.contains("<r>"));
    assert!(result4.contains("{{name}}"));
  }
  
  #[test]
  fn test_excel_column_index() {
    assert_eq!(excel_column_index("A"), 1);
    assert_eq!(excel_column_index("E"), 5);
    assert_eq!(excel_column_index("Z"), 26);
    assert_eq!(excel_column_index("AA"), 27);
    assert_eq!(excel_column_index("AZ"), 52);
    assert_eq!(excel_column_index("BA"), 53);
    assert_eq!(excel_column_index("ZZ"), 702);
    assert_eq!(excel_column_index("AAA"), 703);
  }
  
  #[test]
  fn test_excel_column_name_and_index() {
    let test_cases = vec![
      ("A", 1),
      ("Z", 26),
      ("AA", 27),
      ("AZ", 52),
      ("BA", 53),
      ("ZZ", 702),
      ("AAA", 703),
      ("AAB", 704),
      ("ABC", 731),
      ("ZZZ", 18278),
    ];
    
    for (col_name, expected_index) in test_cases {
      let index = excel_column_index(col_name);
      assert_eq!(index, expected_index, "Column name to index failed for {}", col_name);
      
      let name = excel_column_name(col_name, 0);
      assert_eq!(name, col_name, "Column name identity failed for {}", col_name);
      
      let name_plus_one = excel_column_name(col_name, 1);
      let index_plus_one = excel_column_index(&name_plus_one);
      assert_eq!(index_plus_one, expected_index + 1, "Column name to index failed for {} + 1", col_name);
    }
  }
  
}

/// 删除包含指定标记的整个 row 行
/// 
/// 这个函数用于删除 XLSX sheet 中包含特定 UUID 标记的整行。
/// 通常配合 `{{removeRow}}` helper 使用，用于清理 `{{#each}}{{else}}` 产生的空白行。
/// 
/// # 参数
/// * `xml_content` - sheet.xml 的 XML 内容
/// * `target_uuid` - 要查找和删除的 UUID 标记
/// ```
pub(crate) fn post_process_xml(
    xml_content: &str, 
    remove_row_key: Option<&str>,
    to_number_key: Option<&str>,
    to_formula_key: Option<&str>
) -> Result<String, Box<dyn std::error::Error>> {
    let mut reader = Reader::from_str(xml_content);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    
    let mut current_row_content = String::new();
    let mut in_row = false;
    let mut row_depth = 0;
    
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if e.name().as_ref() == b"row" {
                    in_row = true;
                    row_depth += 1;
                    current_row_content.clear();
                    current_row_content.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    // 添加属性
                    for attr in e.attributes().flatten() {
                        current_row_content.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    current_row_content.push('>');
                } else if in_row {
                    current_row_content.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    // 添加属性
                    for attr in e.attributes().flatten() {
                        current_row_content.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    current_row_content.push('>');
                } else {
                    writer.write_event(Event::Start(e.clone()))?;
                }
            }
            Ok(Event::End(ref e)) => {
                if e.name().as_ref() == b"row" && in_row {
                    row_depth -= 1;
                    if row_depth == 0 {
                        current_row_content.push_str(&format!("</{}>", String::from_utf8_lossy(e.name().as_ref())));
                        
                        // 检查当前行是否需要删除
                        let should_remove = if let Some(key) = remove_row_key {
                            current_row_content.contains(key)
                        } else {
                            false
                        };
                        
                        if !should_remove {
                            // 处理数字类型转换
                            let mut processed_content = if let Some(num_key) = to_number_key {
                                process_number_cells(&current_row_content, num_key)?
                            } else {
                                current_row_content.clone()
                            };
                            
                            // 处理公式类型转换
                            if let Some(formula_key) = to_formula_key {
                                processed_content = process_formula_cells(&processed_content, formula_key)?;
                            }
                            
                            // 写入处理后的行
                            writer.get_mut().write_all(processed_content.as_bytes())?;
                        }
                        // 如果包含删除标记，则跳过整行
                        
                        in_row = false;
                        current_row_content.clear();
                    } else {
                        current_row_content.push_str(&format!("</{}>", String::from_utf8_lossy(e.name().as_ref())));
                    }
                } else if in_row {
                    current_row_content.push_str(&format!("</{}>", String::from_utf8_lossy(e.name().as_ref())));
                } else {
                    writer.write_event(Event::End(e.clone()))?;
                }
            }
            Ok(Event::Text(ref e)) => {
                if in_row {
                    let text = std::str::from_utf8(e)?;
                    current_row_content.push_str(text);
                } else {
                    writer.write_event(Event::Text(e.clone()))?;
                }
            }
            Ok(Event::Empty(ref e)) => {
                if in_row {
                    current_row_content.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    for attr in e.attributes().flatten() {
                        current_row_content.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    current_row_content.push_str("/>");
                } else {
                    writer.write_event(Event::Empty(e.clone()))?;
                }
            }
            Ok(Event::Comment(ref e)) => {
                if !in_row {
                    writer.write_event(Event::Comment(e.clone()))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(event) => {
                if !in_row {
                    writer.write_event(event)?;
                }
            }
            Err(e) => return Err(format!("XML解析错误 at position {}: {:?}", reader.buffer_position(), e).into()),
        }
        buf.clear();
    }
    
    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}

/// 处理行内容中的数字类型单元格
/// 将包含 to_number_key 标记的单元格转换为数字格式
/// 提取 <is> 标签内的文本，转换为 <v>数值</v> 格式
fn process_number_cells(row_content: &str, to_number_key: &str) -> Result<String, Box<dyn std::error::Error>> {
    // 如果不包含数字标记，直接返回
    if !row_content.contains(to_number_key) {
        return Ok(row_content.to_string());
    }
    
    // 使用 XML 解析器来准确处理
    let mut reader = Reader::from_str(row_content);
    let mut output = String::new();
    let mut buf = Vec::new();
    
    let mut in_cell = false;
    let mut cell_attrs = Vec::new();
    let mut cell_content = String::new();
    
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if e.name().as_ref() == b"c" {
                    in_cell = true;
                    cell_attrs.clear();
                    cell_content.clear();
                    
                    // 保存所有属性
                    for attr in e.attributes().flatten() {
                        cell_attrs.push((
                            String::from_utf8_lossy(attr.key.as_ref()).to_string(),
                            String::from_utf8_lossy(&attr.value).to_string()
                        ));
                    }
                } else if in_cell {
                    // 收集单元格内的内容
                    cell_content.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    for attr in e.attributes().flatten() {
                        cell_content.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    cell_content.push('>');
                } else {
                    // 非单元格内容，直接输出
                    output.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    for attr in e.attributes().flatten() {
                        output.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    output.push('>');
                }
            }
            Ok(Event::End(ref e)) => {
                if e.name().as_ref() == b"c" && in_cell {
                    // 单元格结束，处理并输出
                    cell_content.push_str("</c>");
                    
                    // 检查内容是否包含数字标记
                    if cell_content.contains(to_number_key) {
                        // 提取 <is> 标签内的所有 <t> 文本
                        let text_value = extract_text_from_is(&cell_content, to_number_key)?;
                        
                        // 重新构建单元格，移除 t 属性
                        output.push_str("<c");
                        for (key, value) in &cell_attrs {
                            if key != "t" {  // 移除 t 属性
                                output.push_str(&format!(" {}=\"{}\"", key, value));
                            }
                        }
                        output.push('>');
                        
                        // 添加 <v> 标签包含提取的数值
                        output.push_str(&format!("<v>{}</v>", text_value));
                        output.push_str("</c>");
                    } else {
                        // 非数字单元格，原样输出
                        output.push_str("<c");
                        for (key, value) in &cell_attrs {
                            output.push_str(&format!(" {}=\"{}\"", key, value));
                        }
                        output.push('>');
                        
                        let content_without_tags = cell_content
                            .strip_prefix("<c")
                            .and_then(|s| s.find('>').map(|pos| &s[pos+1..]))
                            .unwrap_or(&cell_content);
                        let content_without_tags = content_without_tags
                            .strip_suffix("</c>")
                            .unwrap_or(content_without_tags);
                        output.push_str(content_without_tags);
                        output.push_str("</c>");
                    }
                    
                    in_cell = false;
                } else if in_cell {
                    // 单元格内的结束标签
                    cell_content.push_str(&format!("</{}>", String::from_utf8_lossy(e.name().as_ref())));
                } else {
                    // 非单元格内容
                    output.push_str(&format!("</{}>", String::from_utf8_lossy(e.name().as_ref())));
                }
            }
            Ok(Event::Text(ref e)) => {
                let text = std::str::from_utf8(e)?;
                if in_cell {
                    cell_content.push_str(text);
                } else {
                    output.push_str(text);
                }
            }
            Ok(Event::Empty(ref e)) => {
                if in_cell {
                    cell_content.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    for attr in e.attributes().flatten() {
                        cell_content.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    cell_content.push_str("/>");
                } else {
                    output.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    for attr in e.attributes().flatten() {
                        output.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    output.push_str("/>");
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {
                // 其他事件跳过
            }
            Err(e) => return Err(format!("处理数字单元格时 XML 解析错误: {:?}", e).into()),
        }
        buf.clear();
    }
    
    Ok(output)
}

/// 从 <is> 标签内提取所有 <t> 标签的文本内容，并移除数字标记
fn extract_text_from_is(cell_content: &str, to_number_key: &str) -> Result<String, Box<dyn std::error::Error>> {
    // cell_content 包含完整的单元格内容，可能格式不完整
    // 我们需要找到 <is> 标签并提取其中的文本
    
    // 首先尝试找到 <is> 标签的位置
    if let Some(is_start) = cell_content.find("<is") {
        if let Some(is_end) = cell_content[is_start..].find("</is>") {
            // 提取 <is>...</is> 部分
            let is_content = &cell_content[is_start..is_start + is_end + 5]; // +5 for "</is>"
            
            // 解析这个片段
            let mut reader = Reader::from_str(is_content);
            reader.config_mut().check_end_names = false; // 不严格检查标签匹配
            let mut buf = Vec::new();
            let mut result = String::new();
            let mut in_t = false;
            
            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) => {
                        if e.name().as_ref() == b"t" {
                            in_t = true;
                        }
                    }
                    Ok(Event::End(ref e)) => {
                        if e.name().as_ref() == b"t" {
                            in_t = false;
                        }
                    }
                    Ok(Event::Text(ref e)) => {
                        if in_t {
                            let text = std::str::from_utf8(e)?;
                            result.push_str(text);
                        }
                    }
                    Ok(Event::Eof) => break,
                    Ok(_) => {}
                    Err(e) => {
                        // 如果解析失败，尝试简单的字符串搜索
                        eprintln!("警告: XML 解析失败，使用简单方法提取: {:?}", e);
                        return extract_text_simple(is_content, to_number_key);
                    }
                }
                buf.clear();
            }
            
            // 移除数字标记
            let result = result.replace(to_number_key, "");
            return Ok(result);
        }
    }
    
    // 如果没有找到 <is> 标签，尝试简单方法
    extract_text_simple(cell_content, to_number_key)
}

/// 使用简单的字符串方法提取文本（备用方案）
fn extract_text_simple(content: &str, to_number_key: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut result = String::new();
    
    // 查找所有 <t>...</t> 标签
    let mut pos = 0;
    while let Some(t_start) = content[pos..].find("<t>") {
        let abs_start = pos + t_start + 3; // +3 for "<t>"
        if let Some(t_end) = content[abs_start..].find("</t>") {
            let abs_end = abs_start + t_end;
            result.push_str(&content[abs_start..abs_end]);
            pos = abs_end + 4; // +4 for "</t>"
        } else {
            break;
        }
    }
    
    // 移除数字标记
    let result = result.replace(to_number_key, "");
    Ok(result)
}

/// 处理行内容中的公式类型单元格
/// 将包含 to_formula_key 标记的单元格转换为公式格式
/// 提取 <is> 标签内的文本，转换为 <f>公式</f> 格式
fn process_formula_cells(row_content: &str, to_formula_key: &str) -> Result<String, Box<dyn std::error::Error>> {
    // 如果不包含公式标记，直接返回
    if !row_content.contains(to_formula_key) {
        return Ok(row_content.to_string());
    }
    
    // 使用 XML 解析器来准确处理
    let mut reader = Reader::from_str(row_content);
    let mut output = String::new();
    let mut buf = Vec::new();
    
    let mut in_cell = false;
    let mut cell_attrs = Vec::new();
    let mut cell_content = String::new();
    
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if e.name().as_ref() == b"c" {
                    in_cell = true;
                    cell_attrs.clear();
                    cell_content.clear();
                    
                    // 保存所有属性
                    for attr in e.attributes().flatten() {
                        cell_attrs.push((
                            String::from_utf8_lossy(attr.key.as_ref()).to_string(),
                            String::from_utf8_lossy(&attr.value).to_string()
                        ));
                    }
                } else if in_cell {
                    // 收集单元格内的内容
                    cell_content.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    for attr in e.attributes().flatten() {
                        cell_content.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    cell_content.push('>');
                } else {
                    // 非单元格内容，直接输出
                    output.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    for attr in e.attributes().flatten() {
                        output.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    output.push('>');
                }
            }
            Ok(Event::End(ref e)) => {
                if e.name().as_ref() == b"c" && in_cell {
                    // 单元格结束，处理并输出
                    cell_content.push_str("</c>");
                    
                    // 检查内容是否包含公式标记
                    if cell_content.contains(to_formula_key) {
                        // 提取 <is> 或 <f> 标签内的公式文本
                        let formula_text = extract_formula_from_cell(&cell_content, to_formula_key)?;
                        
                        // 重新构建单元格，移除 t 属性
                        output.push_str("<c");
                        for (key, value) in &cell_attrs {
                            if key != "t" {  // 移除 t 属性
                                output.push_str(&format!(" {}=\"{}\"", key, value));
                            }
                        }
                        output.push('>');
                        
                        // 添加 <f> 标签包含公式
                        output.push_str(&format!("<f>{}</f>", formula_text));
                        output.push_str("</c>");
                    } else {
                        // 非公式单元格，原样输出
                        output.push_str("<c");
                        for (key, value) in &cell_attrs {
                            output.push_str(&format!(" {}=\"{}\"", key, value));
                        }
                        output.push('>');
                        
                        let content_without_tags = cell_content
                            .strip_prefix("<c")
                            .and_then(|s| s.find('>').map(|pos| &s[pos+1..]))
                            .unwrap_or(&cell_content);
                        let content_without_tags = content_without_tags
                            .strip_suffix("</c>")
                            .unwrap_or(content_without_tags);
                        output.push_str(content_without_tags);
                        output.push_str("</c>");
                    }
                    
                    in_cell = false;
                } else if in_cell {
                    // 单元格内的结束标签
                    cell_content.push_str(&format!("</{}>", String::from_utf8_lossy(e.name().as_ref())));
                } else {
                    // 非单元格内容
                    output.push_str(&format!("</{}>", String::from_utf8_lossy(e.name().as_ref())));
                }
            }
            Ok(Event::Text(ref e)) => {
                let text = std::str::from_utf8(e)?;
                if in_cell {
                    cell_content.push_str(text);
                } else {
                    output.push_str(text);
                }
            }
            Ok(Event::Empty(ref e)) => {
                if in_cell {
                    cell_content.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    for attr in e.attributes().flatten() {
                        cell_content.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    cell_content.push_str("/>");
                } else {
                    output.push_str(&format!("<{}", String::from_utf8_lossy(e.name().as_ref())));
                    for attr in e.attributes().flatten() {
                        output.push_str(&format!(" {}=\"{}\"", 
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)));
                    }
                    output.push_str("/>");
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {
                // 其他事件跳过
            }
            Err(e) => return Err(format!("处理公式单元格时 XML 解析错误: {:?}", e).into()),
        }
        buf.clear();
    }
    
    Ok(output)
}

/// 从单元格内容中提取公式文本
/// 可能来自 <is><t>标记公式</t></is> 或 <f>标记公式</f> 标签
fn extract_formula_from_cell(cell_content: &str, to_formula_key: &str) -> Result<String, Box<dyn std::error::Error>> {
    // 首先尝试从 <is> 标签提取（类似数字的处理）
    if let Some(is_start) = cell_content.find("<is") {
        if let Some(is_end) = cell_content[is_start..].find("</is>") {
            let is_content = &cell_content[is_start..is_start + is_end + 5];
            
            let mut reader = Reader::from_str(is_content);
            reader.config_mut().check_end_names = false;
            let mut buf = Vec::new();
            let mut result = String::new();
            let mut in_t = false;
            
            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) => {
                        if e.name().as_ref() == b"t" {
                            in_t = true;
                        }
                    }
                    Ok(Event::End(ref e)) => {
                        if e.name().as_ref() == b"t" {
                            in_t = false;
                        }
                    }
                    Ok(Event::Text(ref e)) => {
                        if in_t {
                            let text = std::str::from_utf8(e)?;
                            result.push_str(text);
                        }
                    }
                    Ok(Event::Eof) => break,
                    Ok(_) => {}
                    Err(e) => {
                        eprintln!("警告: XML 解析失败，使用简单方法提取: {:?}", e);
                        return extract_formula_simple(cell_content, to_formula_key);
                    }
                }
                buf.clear();
            }
            
            let result = result.replace(to_formula_key, "");
            return Ok(result);
        }
    }
    
    // 尝试从 <f> 标签提取
    if let Some(f_start) = cell_content.find("<f>") {
        if let Some(f_end) = cell_content[f_start + 3..].find("</f>") {
            let formula = &cell_content[f_start + 3..f_start + 3 + f_end];
            let formula = formula.replace(to_formula_key, "");
            return Ok(formula);
        }
    }
    
    // 备用简单方法
    extract_formula_simple(cell_content, to_formula_key)
}

/// 使用简单的字符串方法提取公式文本（备用方案）
fn extract_formula_simple(content: &str, to_formula_key: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut result = String::new();
    
    // 首先尝试从 <f> 标签提取
    if let Some(f_start) = content.find("<f>") {
        if let Some(f_end) = content[f_start + 3..].find("</f>") {
            result = content[f_start + 3..f_start + 3 + f_end].to_string();
        }
    }
    
    // 如果没有找到，从 <t> 标签提取
    if result.is_empty() {
        let mut pos = 0;
        while let Some(t_start) = content[pos..].find("<t>") {
            let abs_start = pos + t_start + 3;
            if let Some(t_end) = content[abs_start..].find("</t>") {
                let abs_end = abs_start + t_end;
                result.push_str(&content[abs_start..abs_end]);
                pos = abs_end + 4;
            } else {
                break;
            }
        }
    }
    
    // 移除公式标记
    let result = result.replace(to_formula_key, "");
    Ok(result)
}
