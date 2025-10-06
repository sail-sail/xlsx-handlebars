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

/// 超链接信息结构
#[derive(Debug, Clone)]
pub(crate) struct HyperlinkInfo {
    pub ref_cell: String,     // 单元格引用，如 "A26"
    pub location: String,     // 链接目标，如 "被链接的工作表!A1"
    pub display: String,      // 显示文本（可选）
}

/// 从 sheet XML 中提取并移除 mergeCells 和 hyperlinks 标签
/// 
/// 这个函数会：
/// 1. 找到并移除 <mergeCells> 标签及其内容，提取合并单元格范围
/// 2. 找到并移除 <hyperlinks> 标签及其内容，提取超链接信息
/// 3. 返回去除标签后的 XML、合并范围列表和超链接列表
/// 
/// 注意：提取的范围是静态的，不包含行号/列号偏移
/// 需要在渲染过程中通过 helper 动态添加偏移后的范围
pub(crate) fn extract_and_remove_merge_cells_and_hyperlinks(
    sheet_xml: &str
) -> Result<(String, Vec<String>, Vec<HyperlinkInfo>), Box<dyn std::error::Error>> {
    let mut merge_refs = Vec::new();
    let mut hyperlinks = Vec::new();
    let mut result_xml = sheet_xml.to_string();
    
    // 1. 提取并移除 mergeCells 标签
    if let Some(start) = result_xml.find("<mergeCells") {
        let after_start = &result_xml[start..];
        
        if let Some(end) = after_start.find("</mergeCells>") {
            // 完整标签: <mergeCells>...</mergeCells>
            let merge_cells_content = &after_start[..end + "</mergeCells>".len()];
            
            // 提取所有 ref 属性
            let mut pos = 0;
            while let Some(ref_pos) = merge_cells_content[pos..].find("ref=\"") {
                let abs_ref_pos = pos + ref_pos + 5;
                if let Some(quote_pos) = merge_cells_content[abs_ref_pos..].find('"') {
                    let ref_value = &merge_cells_content[abs_ref_pos..abs_ref_pos + quote_pos];
                    merge_refs.push(ref_value.to_string());
                    pos = abs_ref_pos + quote_pos;
                } else {
                    break;
                }
            }
            
            // 移除整个 mergeCells 标签
            result_xml = format!("{}{}", &result_xml[..start], &result_xml[start + merge_cells_content.len()..]);
        } else if let Some(end) = after_start.find("/>") {
            // 自闭合标签: <mergeCells ... />
            let merge_cells_content = &after_start[..end + "/>".len()];
            
            // 提取所有 ref 属性
            let mut pos = 0;
            while let Some(ref_pos) = merge_cells_content[pos..].find("ref=\"") {
                let abs_ref_pos = pos + ref_pos + 5;
                if let Some(quote_pos) = merge_cells_content[abs_ref_pos..].find('"') {
                    let ref_value = &merge_cells_content[abs_ref_pos..abs_ref_pos + quote_pos];
                    merge_refs.push(ref_value.to_string());
                    pos = abs_ref_pos + quote_pos;
                } else {
                    break;
                }
            }
            
            // 移除整个 mergeCells 标签
            result_xml = format!("{}{}", &result_xml[..start], &result_xml[start + merge_cells_content.len()..]);
        }
    }
    
    // 2. 提取并移除 hyperlinks 标签
    if let Some(start) = result_xml.find("<hyperlinks") {
        let after_start = &result_xml[start..];
        
        if let Some(end) = after_start.find("</hyperlinks>") {
            // 完整标签: <hyperlinks>...</hyperlinks>
            let hyperlinks_content = &after_start[..end + "</hyperlinks>".len()];
            
            // 提取所有 hyperlink 节点
            let mut pos = 0;
            while let Some(link_start) = hyperlinks_content[pos..].find("<hyperlink ") {
                let abs_link_start = pos + link_start;
                if let Some(link_end) = hyperlinks_content[abs_link_start..].find("/>") {
                    let link_tag = &hyperlinks_content[abs_link_start..abs_link_start + link_end + 2];
                    
                    // 提取 ref 属性
                    let ref_cell = if let Some(ref_start) = link_tag.find("ref=\"") {
                        let ref_value_start = ref_start + 5;
                        if let Some(ref_end) = link_tag[ref_value_start..].find('"') {
                            link_tag[ref_value_start..ref_value_start + ref_end].to_string()
                        } else {
                            String::new()
                        }
                    } else {
                        String::new()
                    };
                    
                    // 提取 location 属性
                    let location = if let Some(loc_start) = link_tag.find("location=\"") {
                        let loc_value_start = loc_start + 10;
                        if let Some(loc_end) = link_tag[loc_value_start..].find('"') {
                            link_tag[loc_value_start..loc_value_start + loc_end].to_string()
                        } else {
                            String::new()
                        }
                    } else {
                        String::new()
                    };
                    
                    // 提取 display 属性（可选）
                    let display = if let Some(disp_start) = link_tag.find("display=\"") {
                        let disp_value_start = disp_start + 9;
                        if let Some(disp_end) = link_tag[disp_value_start..].find('"') {
                            link_tag[disp_value_start..disp_value_start + disp_end].to_string()
                        } else {
                            String::new()
                        }
                    } else {
                        String::new()
                    };
                    
                    if !ref_cell.is_empty() && !location.is_empty() {
                        hyperlinks.push(HyperlinkInfo {
                            ref_cell,
                            location,
                            display,
                        });
                    }
                    
                    pos = abs_link_start + link_end + 2;
                } else {
                    break;
                }
            }
            
            // 移除整个 hyperlinks 标签
            result_xml = format!("{}{}", &result_xml[..start], &result_xml[start + hyperlinks_content.len()..]);
        } else if let Some(end) = after_start.find("/>") {
            // 自闭合标签: <hyperlinks ... /> (不常见，但处理一下)
            let hyperlinks_content = &after_start[..end + "/>".len()];
            result_xml = format!("{}{}", &result_xml[..start], &result_xml[start + hyperlinks_content.len()..]);
        }
    }
    
    Ok((result_xml, merge_refs, hyperlinks))
}

/// 在 sharedStrings 数组中注入 helper 调用
/// 通过查找单元格的 sharedString 索引，然后在对应的 shared_strings[index] 前面插入 helper
pub(crate) fn inject_helpers_into_shared_strings(
    xml_content: &str,
    shared_strings: &mut Vec<String>,
    merge_refs: &[String],
    hyperlinks: &[HyperlinkInfo],
) -> Result<(), Box<dyn std::error::Error>> {
    // 处理 mergeCells
    for merge_ref in merge_refs {
        if let Some(colon_pos) = merge_ref.find(':') {
            let start_cell = &merge_ref[..colon_pos];
            let end_cell = &merge_ref[colon_pos + 1..];
            
            // 解析结束单元格的列号和行号
            let end_col = end_cell.chars().take_while(|c| c.is_alphabetic()).collect::<String>();
            let end_row = end_cell.chars().skip_while(|c| c.is_alphabetic()).collect::<String>();
            
            // 查找起始单元格并获取其 sharedString 索引
            let cell_pattern = format!("<c r=\"{}\"", start_cell);
            if let Some(cell_start) = xml_content.find(&cell_pattern) {
                let cell_section = &xml_content[cell_start..];
                
                // 查找 <v> 标签中的索引值
                if let Some(v_start) = cell_section.find("<v>") {
                    if let Some(v_end) = cell_section[v_start + 3..].find("</v>") {
                        let index_str = &cell_section[v_start + 3..v_start + 3 + v_end];
                        if let Ok(index) = index_str.parse::<usize>() {
                            if index < shared_strings.len() {
                                // 构造 helper 调用
                                let helper_call = format!(
                                    "{{{{mergeCell (concat (_cr) \":\" (_cr \"{}\" {}))}}}}",
                                    end_col, end_row
                                );
                                
                                // 在 sharedString 内容的 <t> 标签内部前面插入 helper
                                let original = &shared_strings[index];
                                if let Some(t_start) = original.find("<t>") {
                                    let insert_pos = t_start + 3;
                                    let mut modified = original.to_string();
                                    modified.insert_str(insert_pos, &helper_call);
                                    shared_strings[index] = modified;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    // 处理 hyperlinks
    for link in hyperlinks {
        // 查找单元格并获取其 sharedString 索引
        let cell_pattern = format!("<c r=\"{}\"", link.ref_cell);
        if let Some(cell_start) = xml_content.find(&cell_pattern) {
            let cell_section = &xml_content[cell_start..];
            
            // 查找 <v> 标签中的索引值
            if let Some(v_start) = cell_section.find("<v>") {
                if let Some(v_end) = cell_section[v_start + 3..].find("</v>") {
                    let index_str = &cell_section[v_start + 3..v_start + 3 + v_end];
                    if let Ok(index) = index_str.parse::<usize>() {
                        if index < shared_strings.len() {
                            // 构造 helper 调用
                            let helper_call = if link.display.is_empty() {
                                format!("{{{{hyperlink (_cr) \"{}\" \"\"}}}}", link.location)
                            } else {
                                format!("{{{{hyperlink (_cr) \"{}\" \"{}\"}}}}", link.location, link.display)
                            };
                            
                            // 在 sharedString 内容的 <t> 标签内部前面插入 helper
                            let original = &shared_strings[index];
                            if let Some(t_start) = original.find("<t>") {
                                let insert_pos = t_start + 3;
                                let mut modified = original.to_string();
                                modified.insert_str(insert_pos, &helper_call);
                                shared_strings[index] = modified;
                            }
                        }
                    }
                }
            }
        }
    }
    
    Ok(())
}

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
                          current_col = to_column_index(&r_char);
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
                          let col_inline = to_column_index(&r_char);
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
    
    // 注册 add helper (加法)
    handlebars_helper!(add: |x: i64, y: i64| x + y);
    handlebars.register_helper("add", Box::new(add));
    
    // 注册 sub helper (减法)
    handlebars_helper!(sub: |x: i64, y: i64| x - y);
    handlebars.register_helper("sub", Box::new(sub));
    
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
pub fn to_column_name(current: &str, increment: u32) -> String {
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
pub fn to_column_index(col_name: &str) -> u32 {
  let mut col_index = 0;
  for (i, ch) in col_name.chars().rev().enumerate() {
    let ch_val = (ch as u8 - b'A' + 1) as u32;
    col_index += ch_val * 26_u32.pow(i as u32);
  }
  col_index
}

/// 将时间戳（毫秒）转换为 Excel 日期序列号
/// 
/// Excel 使用从 1900年1月1日开始的序列号来表示日期。
/// 由于 Excel 的历史 bug（将 1900 年视为闰年），需要特殊处理。
/// 
/// # 参数
/// * `timestamp_ms` - Unix 时间戳（毫秒），从 1970-01-01 00:00:00 UTC 开始
/// 
/// # 返回
/// Excel 日期序列号（浮点数）
/// 
/// # 示例
/// ```rust
/// use xlsx_handlebars::timestamp_to_excel_date;
/// 
/// // 2024-01-01 00:00:00 UTC
/// let timestamp = 1704067200000i64;
/// let excel_date = timestamp_to_excel_date(timestamp);
/// println!("Excel date number: {}", excel_date);
/// ```
pub fn timestamp_to_excel_date(timestamp_ms: i64) -> f64 {
    // Excel 的基准日期是 1900-01-01，但实际上从 1899-12-30 开始计数
    // 25569 是从 1900-01-01 到 1970-01-01 的天数
    const EXCEL_EPOCH_OFFSET: i64 = 25569;
    const MS_PER_DAY: i64 = 86400000;
    
    let mut val_tmp = timestamp_ms + EXCEL_EPOCH_OFFSET * MS_PER_DAY;
    
    // Excel 错误地将 1900 年视为闰年，需要调整
    // 60 代表 1900-02-29（不存在的日期）
    if val_tmp <= 60 * MS_PER_DAY {
        val_tmp += MS_PER_DAY;
    } else {
        val_tmp += 2 * MS_PER_DAY;
    }
    
    val_tmp as f64 / MS_PER_DAY as f64
}

/// 将 Excel 日期序列号转换为 Unix 时间戳（毫秒）
/// 
/// Excel 使用从 1900年1月1日开始的序列号来表示日期。
/// 此函数将 Excel 序列号转换回 Unix 时间戳。
/// 
/// # 参数
/// * `excel_date` - Excel 日期序列号
/// 
/// # 返回
/// * `Some(timestamp_ms)` - Unix 时间戳（毫秒），从 1970-01-01 00:00:00 UTC 开始
/// * `None` - 如果输入的序列号无效（如负数或等于60）
/// 
/// # 示例
/// ```rust
/// use xlsx_handlebars::excel_date_to_timestamp;
/// 
/// // Excel 序列号 45294.0 表示 2024-01-01
/// let excel_date = 45294.0;
/// if let Some(timestamp) = excel_date_to_timestamp(excel_date) {
///     println!("Unix timestamp: {}", timestamp);
/// }
/// ```
pub fn excel_date_to_timestamp(excel_date: f64) -> Option<i64> {
    const MS_PER_DAY: i64 = 86400000;
    const EXCEL_EPOCH_OFFSET: i64 = 25569;
    
    let mut val = excel_date;
    
    // 处理 Excel 的 1900 年闰年 bug（反向操作）
    if val < 60.0 {
        val -= 1.0;
    } else if val > 60.0 {
        val -= 2.0;
    }
    
    // 60 代表不存在的日期 1900-02-29
    if val < 0.0 || excel_date == 60.0 {
        return None;
    }
    
    // 转换为时间戳
    let timestamp = (val * MS_PER_DAY as f64).round() as i64 - (EXCEL_EPOCH_OFFSET * MS_PER_DAY);
    
    Some(timestamp)
}

#[cfg(test)]
mod tests {
  use super::*;
  
  #[test]
  fn test_excel_column_name() {
    assert_eq!(to_column_name("A", 0), "A");
    assert_eq!(to_column_name("A", 1), "B");
    assert_eq!(to_column_name("Z", 1), "AA");
    assert_eq!(to_column_name("AA", 1), "AB");
    assert_eq!(to_column_name("AZ", 1), "BA");
    assert_eq!(to_column_name("ZZ", 1), "AAA");
    assert_eq!(to_column_name("AAA", 26), "ABA");
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
    assert_eq!(to_column_index("A"), 1);
    assert_eq!(to_column_index("E"), 5);
    assert_eq!(to_column_index("Z"), 26);
    assert_eq!(to_column_index("AA"), 27);
    assert_eq!(to_column_index("AZ"), 52);
    assert_eq!(to_column_index("BA"), 53);
    assert_eq!(to_column_index("ZZ"), 702);
    assert_eq!(to_column_index("AAA"), 703);
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
      let index = to_column_index(col_name);
      assert_eq!(index, expected_index, "Column name to index failed for {}", col_name);
      
      let name = to_column_name(col_name, 0);
      assert_eq!(name, col_name, "Column name identity failed for {}", col_name);
      
      let name_plus_one = to_column_name(col_name, 1);
      let index_plus_one = to_column_index(&name_plus_one);
      assert_eq!(index_plus_one, expected_index + 1, "Column name to index failed for {} + 1", col_name);
    }
  }
  
  #[test]
  fn test_excel_date_conversion() {
    // 测试 2024-01-01 00:00:00 UTC
    // Unix timestamp: 1704067200000 ms
    let timestamp_2024 = 1704067200000i64;
    let excel_date = timestamp_to_excel_date(timestamp_2024);
    
    // Excel 中 2024-01-01 的序列号应该是 45294 (包含 Excel 1900 bug 调整)
    assert!((excel_date - 45294.0).abs() < 0.001, "Excel date for 2024-01-01 should be ~45294, got {}", excel_date);
    
    // 反向转换
    if let Some(timestamp) = excel_date_to_timestamp(excel_date) {
      // 允许一些精度损失（毫秒级别）
      assert!((timestamp - timestamp_2024).abs() < 1000, "Timestamp mismatch: expected {}, got {}", timestamp_2024, timestamp);
    } else {
      panic!("Failed to convert Excel date back to timestamp");
    }
    
    // 测试 1970-01-01 00:00:00 UTC (Unix epoch)
    let timestamp_1970 = 0i64;
    let excel_date_1970 = timestamp_to_excel_date(timestamp_1970);
    
    // Excel 中 1970-01-01 的序列号应该是 25571 (25569 + 2 for bug)
    assert!((excel_date_1970 - 25571.0).abs() < 0.001, "Excel date for 1970-01-01 should be ~25571, got {}", excel_date_1970);
    
    // 测试边界情况：1900-02-28 (序列号 59)
    let excel_date_59 = 59.0;
    assert!(excel_date_to_timestamp(excel_date_59).is_some());
    
    // 测试无效日期：1900-02-29 (序列号 60，不存在的日期)
    let excel_date_60 = 60.0;
    assert!(excel_date_to_timestamp(excel_date_60).is_none(), "Excel date 60 (1900-02-29) should be invalid");
    
    // 测试负数（无效）
    assert!(excel_date_to_timestamp(-1.0).is_none(), "Negative Excel date should be invalid");
  }
  
  #[test]
  fn test_extract_and_remove_merge_cells_and_hyperlinks() {
    // 测试包含完整 mergeCells 标签的情况
    let input_with_merge = r#"<?xml version="1.0"?>
<worksheet>
  <sheetData>
    <row r="1">
      <c r="A1"><v>Test</v></c>
    </row>
  </sheetData>
  <mergeCells count="2">
    <mergeCell ref="A1:B1"/>
    <mergeCell ref="C2:D3"/>
  </mergeCells>
  <pageMargins left="0.7" right="0.7"/>
</worksheet>"#;
    
    let (result_xml, merge_refs, hyperlinks) = extract_and_remove_merge_cells_and_hyperlinks(input_with_merge).unwrap();
    
    // 验证合并范围被正确提取
    assert_eq!(merge_refs.len(), 2);
    assert_eq!(merge_refs[0], "A1:B1");
    assert_eq!(merge_refs[1], "C2:D3");
    assert_eq!(hyperlinks.len(), 0);
    
    // 验证 mergeCells 标签被移除
    assert!(!result_xml.contains("<mergeCells"));
    assert!(!result_xml.contains("</mergeCells>"));
    assert!(!result_xml.contains("mergeCell"));
    
    // 验证其他内容保持不变
    assert!(result_xml.contains("<sheetData>"));
    assert!(result_xml.contains("<pageMargins"));
    
    // 测试包含 hyperlinks 的情况
    let input_with_hyperlinks = r#"<?xml version="1.0"?>
<worksheet>
  <sheetData>
    <row r="1">
      <c r="A1"><v>Link</v></c>
    </row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="A1" location="Sheet2!A1" display="Go to Sheet2"/>
    <hyperlink ref="B2" location="https://example.com" display="Example"/>
  </hyperlinks>
</worksheet>"#;
    
    let (result_xml2, merge_refs2, hyperlinks2) = extract_and_remove_merge_cells_and_hyperlinks(input_with_hyperlinks).unwrap();
    assert_eq!(merge_refs2.len(), 0);
    assert_eq!(hyperlinks2.len(), 2);
    assert_eq!(hyperlinks2[0].ref_cell, "A1");
    assert_eq!(hyperlinks2[0].location, "Sheet2!A1");
    assert_eq!(hyperlinks2[0].display, "Go to Sheet2");
    assert_eq!(hyperlinks2[1].ref_cell, "B2");
    assert_eq!(hyperlinks2[1].location, "https://example.com");
    assert!(!result_xml2.contains("<hyperlinks"));
    
    // 测试同时包含 mergeCells 和 hyperlinks 的情况
    let input_both = r#"<?xml version="1.0"?>
<worksheet>
  <sheetData/>
  <mergeCells count="1">
    <mergeCell ref="A1:B2"/>
  </mergeCells>
  <hyperlinks>
    <hyperlink ref="C3" location="Sheet1!A1" display="Link"/>
  </hyperlinks>
</worksheet>"#;
    
    let (result_xml3, merge_refs3, hyperlinks3) = extract_and_remove_merge_cells_and_hyperlinks(input_both).unwrap();
    assert_eq!(merge_refs3.len(), 1);
    assert_eq!(hyperlinks3.len(), 1);
    assert!(!result_xml3.contains("mergeCells"));
    assert!(!result_xml3.contains("hyperlinks"));
  }

  
}

/// 删除包含指定标记的整个 row 行
/// 
/// 这个函数用于删除 XLSX sheet 中包含特定 UUID 标记的整行。
/// 通常配合 `{{removeRow}}` helper 使用，用于清理 `{{#each}}{{else}}` 产生的空白行。
/// 
/// # 参数
/// * `xml_content` - sheet.xml 的 XML 内容
/// * `remove_key` - 要查找和删除的行标记
/// * `to_number_key` - 数字类型转换标记
/// * `to_formula_key` - 公式类型转换标记
/// * `merge_cells` - 需要合并的单元格范围列表
/// ```
pub(crate) fn post_process_xml(
    xml_content: &str, 
    remove_key: Option<&str>,
    to_number_key: Option<&str>,
    to_formula_key: Option<&str>,
    merge_cells: Option<&[String]>,
    hyperlinks: Option<&[HyperlinkInfo]>,
) -> Result<String, Box<dyn std::error::Error>> {
    let mut reader = Reader::from_str(xml_content);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    
    let mut current_row_content = String::new();
    let mut in_row = false;
    let mut row_depth = 0;
    let mut hyperlinks_inserted = false; // 标记是否已插入 hyperlinks
    
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
                    // 检查是否是 pageMargins 开始标签，如果是则先插入 hyperlinks
                    if e.name().as_ref() == b"pageMargins" && !hyperlinks_inserted {
                        hyperlinks_inserted = true;
                        
                        // 先插入 hyperlinks（如果有）
                        if let Some(links) = hyperlinks {
                            if !links.is_empty() {
                                use uuid::Uuid;
                                
                                // 生成 hyperlinks XML
                                let hyperlinks_xml = links.iter()
                                    .map(|link| {
                                        let uuid = Uuid::new_v4();
                                        let uuid_str = format!("{{{}}}", uuid.to_string().to_uppercase());
                                        
                                        // 构造超链接标签
                                        if link.display.is_empty() {
                                            format!(
                                                "<hyperlink ref=\"{}\" location=\"{}\" xr:uid=\"{}\"/>",
                                                link.ref_cell, link.location, uuid_str
                                            )
                                        } else {
                                            format!(
                                                "<hyperlink ref=\"{}\" location=\"{}\" display=\"{}\" xr:uid=\"{}\"/>",
                                                link.ref_cell, link.location, link.display, uuid_str
                                            )
                                        }
                                    })
                                    .collect::<Vec<_>>()
                                    .join("");
                                
                                // 写入 hyperlinks（带命名空间属性）
                                let hyperlinks_tag = format!(
                                    "<hyperlinks xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\">{}</hyperlinks>",
                                    hyperlinks_xml
                                );
                                writer.get_mut().write_all(hyperlinks_tag.as_bytes())?;
                            }
                        }
                    }
                    
                    writer.write_event(Event::Start(e.clone()))?;
                }
            }
            Ok(Event::End(ref e)) => {
                if e.name().as_ref() == b"row" && in_row {
                    row_depth -= 1;
                    if row_depth == 0 {
                        current_row_content.push_str(&format!("</{}>", String::from_utf8_lossy(e.name().as_ref())));
                        
                        // 检查当前行是否需要删除
                        let should_remove = if let Some(key) = remove_key {
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
                    // 检查是否是 sheetData 结束标签
                    if e.name().as_ref() == b"sheetData" {
                        // 先输出 sheetData 结束标签
                        writer.write_event(Event::End(e.clone()))?;
                        
                        // 如果有合并单元格信息，插入 mergeCells 标签
                        if let Some(refs) = merge_cells {
                            if !refs.is_empty() {
                                // 去重处理
                                let mut unique_refs: Vec<String> = refs.to_vec();
                                unique_refs.sort();
                                unique_refs.dedup();
                                
                                // 生成 mergeCells XML
                                let merge_cells_xml = format!(
                                    "<mergeCells count=\"{}\">{}</mergeCells>",
                                    unique_refs.len(),
                                    unique_refs.iter()
                                        .map(|r| format!("<mergeCell ref=\"{}\"/>", r))
                                        .collect::<Vec<_>>()
                                        .join("")
                                );
                                
                                // 写入 mergeCells
                                writer.get_mut().write_all(merge_cells_xml.as_bytes())?;
                            }
                        }
                    } else {
                        writer.write_event(Event::End(e.clone()))?;
                    }
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
                    // 检查是否是 pageMargins 自闭合标签，如果是则先插入 hyperlinks
                    if e.name().as_ref() == b"pageMargins" && !hyperlinks_inserted {
                        hyperlinks_inserted = true;
                        
                        // 先插入 hyperlinks（如果有）
                        if let Some(links) = hyperlinks {
                            if !links.is_empty() {
                                use uuid::Uuid;
                                
                                // 生成 hyperlinks XML
                                let hyperlinks_xml = links.iter()
                                    .map(|link| {
                                        let uuid = Uuid::new_v4();
                                        let uuid_str = format!("{{{}}}", uuid.to_string().to_uppercase());
                                        
                                        // 构造超链接标签
                                        if link.display.is_empty() {
                                            format!(
                                                "<hyperlink ref=\"{}\" location=\"{}\" xr:uid=\"{}\"/>",
                                                link.ref_cell, link.location, uuid_str
                                            )
                                        } else {
                                            format!(
                                                "<hyperlink ref=\"{}\" location=\"{}\" display=\"{}\" xr:uid=\"{}\"/>",
                                                link.ref_cell, link.location, link.display, uuid_str
                                            )
                                        }
                                    })
                                    .collect::<Vec<_>>()
                                    .join("");
                                
                                // 写入 hyperlinks（带命名空间属性）
                                let hyperlinks_tag = format!(
                                    "<hyperlinks xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\">{}</hyperlinks>",
                                    hyperlinks_xml
                                );
                                writer.get_mut().write_all(hyperlinks_tag.as_bytes())?;
                            }
                        }
                    }
                    
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
            Err(e) => return Err(format!("XML Error at position {}: {:?}", reader.buffer_position(), e).into()),
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
