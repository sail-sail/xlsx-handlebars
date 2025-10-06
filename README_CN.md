# xlsx-handlebars

[![Crates.io](https://img.shields.io/crates/v/xlsx-handlebars.svg)](https://crates.io/crates/xlsx-handlebars)
[![Documentation](https://docs.rs/xlsx-handlebars/badge.svg)](https://docs.rs/xlsx-handlebars)
[![License](https://img.shields.io/crates/l/xlsx-handlebars.svg)](https://github.com/sail-sail/xlsx-handlebars#license)

| 中文文档 | [English](README.md) | [Demo](https://sail-sail.github.io/xlsx-handlebars-demo/)  

一个用于处理 XLSX 文件 Handlebars 模板的 Rust 库，支持多平台使用：
- 🦀 Rust 原生
- 🌐 WebAssembly (WASM)
- 📦 npm 包
- 🟢 Node.js
- 🦕 Deno
- 🌍 浏览器端
- 📋 JSR (JavaScript Registry)

## 功能特性

- ⚡ **极致性能**：2.12秒渲染10万行数据（约4.7万行/秒）- 比 Python 快 14-28倍，比 JavaScript 快 7-14倍
- ✅ **智能合并**：自动处理被 XML 标签分割的 Handlebars 语法
- ✅ **XLSX 验证**：内置文件格式验证，确保输入文件有效
- ✅ **Handlebars 支持**：完整的模板引擎，支持变量、条件、循环、Helper 函数
- ✅ **跨平台**：Rust 原生 + WASM 支持多种运行时
- ✅ **TypeScript**：完整的类型定义和智能提示
- ✅ **零依赖**：WASM 二进制文件，无外部依赖
- ✅ **错误处理**：详细的错误信息和类型安全的错误处理

## 安装

### Rust

```bash
cargo add xlsx-handlebars
```

### npm

```bash
npm install xlsx-handlebars
```

### Deno

```typescript
import { render_template, init } from "jsr:@sail/xlsx-handlebars";
```

## 使用示例

### Rust

```rust
use xlsx_handlebars::render_template;
use serde_json::json;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 读取 XLSX 模板文件
    let template_bytes = std::fs::read("template.xlsx")?;
    
    // 准备数据
    let data = json!({
        "name": "张三",
        "company": "ABC科技有限公司",
        "position": "软件工程师",
        "projects": [
            {"name": "项目A", "status": "已完成"},
            {"name": "项目B", "status": "进行中"}
        ],
        "has_bonus": true,
        "bonus_amount": 5000
    });
    
    // 渲染模板
    let result = render_template(template_bytes, &data)?;
    
    // 保存结果
    std::fs::write("output.xlsx", result)?;
    
    Ok(())
}
```

### JavaScript/TypeScript (Node.js)

```javascript
import init, { render_template } from "xlsx-handlebars";
import fs from 'fs';

async function processTemplate() {
    // 初始化 WASM 模块
    await init();
    
    // 读取模板文件
    const templateBytes = fs.readFileSync("template.xlsx");
    
    // 准备数据
    const data = {
        name: "李明",
        company: "XYZ技术有限公司",
        position: "高级开发工程师",
        projects: [
            { name: "E-commerce平台", status: "已完成" },
            { name: "移动端APP", status: "开发中" }
        ],
        has_bonus: true,
        bonus_amount: 8000
    };
    
    // 渲染模板
    const result = render_template(templateBytes, JSON.stringify(data));
    
    // 保存结果
    fs.writeFileSync('output.xlsx', new Uint8Array(result));
}

processTemplate().catch(console.error);
```

### Deno

```typescript
import init, { render_template } from "https://deno.land/x/xlsx_handlebars/mod.ts";

async function processTemplate() {
    // 初始化 WASM 模块
    await init();
    
    // 读取模板文件
    const templateBytes = await Deno.readFile("template.xlsx");
    
    // 准备数据
    const data = {
        name: "王小明",
        department: "研发部",
        projects: [
            { name: "智能客服系统", status: "已上线" },
            { name: "数据可视化平台", status: "开发中" }
        ]
    };
    
    // 渲染模板
    const result = render_template(templateBytes, JSON.stringify(data));
    
    // 保存结果
    await Deno.writeFile("output.xlsx", new Uint8Array(result));
}

if (import.meta.main) {
    await processTemplate();
}
```

### 浏览器端

```html
<!DOCTYPE html>
<html>
<head>
    <title>XLSX Handlebars 示例</title>
</head>
<body>
    <input type="file" id="fileInput" accept=".xlsx">
    <button onclick="processFile()">处理模板</button>
    
    <script type="module">
        import init, { render_template } from './pkg/xlsx_handlebars.js';
        
        // 初始化 WASM
        await init();
        
        window.processFile = async function() {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];
            
            if (!file) return;
            
            const arrayBuffer = await file.arrayBuffer();
            const templateBytes = new Uint8Array(arrayBuffer);
            
            const data = {
                name: "张三",
                company: "示例公司"
            };
            
            try {
                const result = render_template(templateBytes, JSON.stringify(data));
                
                // 下载结果
                const blob = new Blob([new Uint8Array(result)], {
                    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'processed.xlsx';
                a.click();
            } catch (error) {
                console.error('处理失败:', error);
            }
        };
    </script>
</body>
</html>
```

## 模板语法

### 基础变量替换

```handlebars
员工姓名: {{name}}
公司: {{company}}
职位: {{position}}
```

### 条件渲染

```handlebars
{{#if has_bonus}}
奖金: ¥{{bonus_amount}}
{{else}}
无奖金
{{/if}}

{{#unless is_intern}}
正式员工
{{/unless}}
```

### 循环渲染

```handlebars
项目经历:
{{#each projects}}
- {{name}}: {{description}} ({{status}})
{{/each}}

技能列表:
{{#each skills}}
{{@index}}. {{this}}
{{/each}}
```

### Helper 函数

内置的 Helper 函数：

```handlebars
<!-- 基础 helper -->
{{upper name}}           <!-- 转大写 -->
{{lower company}}        <!-- 转小写 -->
{{len projects}}         <!-- 数组长度 -->
{{#if (eq status "completed")}}已完成{{/if}}    <!-- 相等比较 -->
{{#if (gt score 90)}}优秀{{/if}}               <!-- 大于比较 -->
{{#if (lt age 30)}}年轻{{/if}}                 <!-- 小于比较 -->

<!-- 字符串拼接 -->
{{concat "你好" " " "世界"}}                    <!-- 字符串拼接 -->
{{concat "总计: " count}}                      <!-- 混合字符串和变量 -->

<!-- Excel 专用 helper -->
{{num employee.salary}}                         <!-- 标记单元格为数字类型 -->
{{formula "=SUM(A1:B1)"}}                      <!-- 静态 Excel 公式 -->
{{formula (concat "=SUM(" (_c) "1:" (_c) "10)")}} <!-- 使用当前列的动态公式 -->
{{mergeCell "C4:D5"}}                          <!-- 合并单元格 C4 到 D5 -->
{{img logo.data 100 100}}                       <!-- 插入图片，指定宽高 -->

<!-- 列名转换 helper -->
{{toColumnName "A" 5}}                          <!-- A + 5 偏移 = F -->
{{toColumnName (_c) 3}}                         <!-- 当前列向右偏移 3 列 -->
{{toColumnIndex "AA"}}                          <!-- AA 列的索引 = 27 -->
```

#### Excel 公式 Helper

**静态公式**:
```handlebars
<!-- 在 Excel 单元格中 -->
{{formula "=SUM(A1:B1)"}}
{{formula "=AVERAGE(C2:C10)"}}
{{formula "=IF(D1>100,\"高\",\"低\")"}}
```

**使用 `concat` 的动态公式**:
```handlebars
<!-- 动态行引用 -->
{{formula (concat "=A" (_r) "*B" (_r))}}

<!-- 动态列引用 -->
{{formula (concat "=SUM(" (_c) "2:" (_c) "10)")}}

<!-- 复杂动态公式 -->
{{formula (concat "=IF(" (_cr) ">100,\"高\",\"低\")")}}
```

**可用的位置 helper**:
- `(_c)` - 当前列字母 (A, B, C, ...)
- `(_r)` - 当前行号 (1, 2, 3, ...)
- `(_cr)` - 当前单元格引用 (A1, B2, C3, ...)

#### 列名转换 Helper

**`toColumnName`** - 将列名或列索引转换为新的列名，支持偏移量：

```handlebars
<!-- 基础用法：从指定列名开始偏移 -->
{{toColumnName "A" 0}}     <!-- A (无偏移) -->
{{toColumnName "A" 5}}     <!-- F (A + 5) -->
{{toColumnName "Z" 1}}     <!-- AA (Z + 1) -->

<!-- 配合当前列使用 -->
{{toColumnName (_c) 3}}    <!-- 当前列向右偏移 3 列 -->

<!-- 动态公式中的应用 -->
{{formula (concat "=SUM(" (_c) "1:" (toColumnName (_c) 3) "1)")}}
<!-- 示例：如果当前列是 B，生成公式 =SUM(B1:E1) -->
```

**`toColumnIndex`** - 将列名转换为列索引（1-based）：

```handlebars
{{toColumnIndex "A"}}      <!-- 1 -->
{{toColumnIndex "Z"}}      <!-- 26 -->
{{toColumnIndex "AA"}}     <!-- 27 -->
{{toColumnIndex "AB"}}     <!-- 28 -->
```

#### 合并单元格 Helper

**`mergeCell`** - 标记需要合并的单元格范围：

```handlebars
<!-- 静态合并单元格 -->
{{mergeCell "C4:D5"}}      <!-- 合并 C4 到 D5 区域 -->
{{mergeCell "F4:G4"}}      <!-- 合并 F4 到 G4 区域 -->

<!-- 动态合并单元格：从当前位置合并 -->
{{mergeCell (concat (_c) (_r) ":" (toColumnName (_c) 3) (_r))}}
<!-- 示例：如果当前在 B5，合并 B5:E5（向右合并4列） -->

<!-- 动态合并单元格：跨行跨列 -->
{{mergeCell (concat (_c) (_r) ":" (toColumnName (_c) 2) (add (_r) 2))}}
<!-- 示例：如果当前在 C3，合并 C3:E5（3列×3行的区域） -->

<!-- 在循环中动态合并 -->
{{#each sections}}
  {{mergeCell (concat "A" (add @index 2) ":D" (add @index 2))}}
  <!-- 为每个 section 合并一行的 A-D 列 -->
{{/each}}
```

**注意事项**：
- `mergeCell` 不产生输出，仅收集合并信息
- 合并范围格式必须是 `起始单元格:结束单元格`（如 `"A1:B2"`）
- 相同的合并范围会自动去重
- 合并信息会在渲染完成后自动添加到 Excel 文件中

#### 超链接 Helper

**`hyperlink`** - 在 Excel 单元格中添加超链接：

```handlebars
<!-- 基础用法：链接到其他工作表 -->
{{hyperlink (_cr) "Sheet2!A1" "查看详情"}}

<!-- 链接到外部网址（需在模板中预设） -->
{{hyperlink (_cr) "https://example.com" "访问网站"}}

<!-- 动态链接 -->
{{#each items}}
  {{hyperlink (_cr) (concat "详情!" name) name}}
{{/each}}
```

**参数说明**：
- 第一个参数：单元格引用，通常使用 `(_cr)` 获取当前单元格
- 第二个参数：链接目标（工作表引用或 URL）
- 第三个参数：显示文本（可选）

**注意事项**：
- `hyperlink` 不产生输出，仅收集超链接信息
- 超链接会在渲染完成后自动添加到 Excel 文件中
- 支持工作表内部引用（如 `"Sheet2!A1"`）
- 外部链接需要在模板 Excel 文件中预先配置关系

#### 数字类型 Helper

使用 `{{num value}}` 确保单元格在 Excel 中被识别为数字：

```handlebars
<!-- 不使用 num: 当作文本处理 -->
{{employee.salary}}

<!-- 使用 num: 当作数字处理 -->
{{num employee.salary}}
```

特别适用于以下场景：
- 值可能是字符串但应当作数字处理
- 需要确保 Excel 中的数字格式正确
- 需要在公式中使用该值

#### 图片插入 Helper

**`img`** - 在 Excel 中插入 base64 编码的图片：

```handlebars
<!-- 基础用法：插入图片并使用原始尺寸 -->
{{img logo.data}}

<!-- 指定宽度和高度（单位：像素） -->
{{img photo.data 150 200}}

<!-- 使用数据中的尺寸 -->
{{img image.data image.width image.height}}
```

**特性**：
- ✅ 支持 PNG、JPEG、WebP、BMP、TIFF、GIF 等常见图片格式
- ✅ 自动检测图片实际尺寸
- ✅ 可选指定宽度和高度（像素）
- ✅ 图片定位在当前单元格位置
- ✅ 图片不受单元格大小限制，保持比例
- ✅ 支持同一 sheet 插入多张图片
- ✅ 支持多个 sheet 各自插入图片
- ✅ 使用 UUID 避免 ID 冲突

**完整示例**：

```javascript
// 在 JavaScript 中准备图片数据
import fs from 'fs';

const imageBuffer = fs.readFileSync('logo.png');
const base64Image = imageBuffer.toString('base64');

const data = {
  company: {
    logo: base64Image,
    name: "科技公司"
  },
  products: [
    {
      name: "产品A",
      photo: base64Image,
      width: 120,
      height: 120
    },
    {
      name: "产品B", 
      photo: base64Image,
      width: 100,
      height: 100
    }
  ]
};

// 在模板中使用
```

```handlebars
<!-- Excel 模板示例 -->
公司Logo: {{img company.logo 100 50}}

产品列表:
{{#each products}}
产品名: {{name}}
图片: {{img photo width height}}
{{/each}}
```

**使用技巧**：
- 如果只指定宽度，高度会等比例缩放
- 如果只指定高度，宽度会等比例缩放
- 如果都不指定，使用图片原始尺寸
- 图片会放置在调用 `{{img}}` 的单元格位置
- base64 数据不包含 `data:image/png;base64,` 前缀，只需要纯 base64 字符串

#### 工作表管理 Helpers

**`deleteCurrentSheet`** - 删除当前正在渲染的工作表：

```handlebars
<!-- 基础用法 -->
{{deleteCurrentSheet}}

<!-- 条件删除 -->
{{#if shouldDelete}}
  {{deleteCurrentSheet}}
{{/if}}

<!-- 删除非活跃工作表 -->
{{#unless isActive}}
  {{deleteCurrentSheet}}
{{/unless}}
```

**特性**：
- ✅ 从工作簿中移除工作表及其关系
- ✅ 清理相关文件（rels、content types）
- ✅ 保留 drawing 文件（安全考虑）
- ✅ 不能删除最后一个工作表（Excel 要求）
- ✅ 延迟执行，所有渲染完成后统一删除

**`setCurrentSheetName`** - 重命名当前工作表：

```handlebars
<!-- 静态名称 -->
{{setCurrentSheetName "销售报表"}}

<!-- 动态名称 -->
{{setCurrentSheetName (concat department.name " - " year "年")}}

<!-- 基于循环的命名 -->
{{#each departments}}
  {{setCurrentSheetName (concat "部门" @index " - " name)}}
{{/each}}
```

**特性**：
- ✅ 自动过滤非法字符：`\ / ? * [ ]`
- ✅ 自动限制长度为 31 个字符
- ✅ 自动处理重名，添加数字后缀
- ✅ 支持动态名称生成

**`hideCurrentSheet`** - 隐藏当前工作表：

```handlebars
<!-- 普通隐藏（用户可通过右键取消隐藏） -->
{{hideCurrentSheet}}
{{hideCurrentSheet "hidden"}}

<!-- 超级隐藏（需要 VBA 才能取消隐藏） -->
{{hideCurrentSheet "veryHidden"}}

<!-- 条件隐藏 -->
{{#unless (eq userRole "admin")}}
  {{hideCurrentSheet "veryHidden"}}
{{/unless}}
```

**隐藏级别**：
- `hidden` - 普通隐藏，用户可通过 Excel 右键菜单取消隐藏
- `veryHidden` - 超级隐藏，需要 VBA 或属性编辑器才能取消隐藏

**特性**：
- ✅ 不能隐藏所有工作表（Excel 要求至少一个可见）
- ✅ 两种隐藏级别：普通隐藏和超级隐藏
- ✅ 适用于权限控制和敏感数据保护

**常见使用场景**：

```handlebars
<!-- 多语言报表：删除未使用的语言工作表 -->
{{#if (ne language "zh-CN")}}
  {{deleteCurrentSheet}}
{{/if}}

<!-- 动态部门报表：按部门重命名工作表 -->
{{setCurrentSheetName (concat department.name " 报表")}}

<!-- 权限控制：对普通用户隐藏管理员工作表 -->
{{#unless (eq userRole "admin")}}
  {{hideCurrentSheet "veryHidden"}}
{{/unless}}

<!-- 条件工作流：根据状态删除、重命名或隐藏 -->
{{#if (eq status "inactive")}}
  {{deleteCurrentSheet}}
{{else}}
  {{setCurrentSheetName (concat "活跃 - " name)}}
  {{#if isInternal}}
    {{hideCurrentSheet}}
  {{/if}}
{{/if}}
```

### 复杂示例

```handlebars
=== 员工报告 ===

基本信息:
姓名: {{employee.name}}
部门: {{employee.department}}
职位: {{employee.position}}
入职时间: {{employee.hire_date}}

{{#if employee.has_bonus}}
💰 奖金: ¥{{employee.bonus_amount}}
{{/if}}

项目经历 (共{{len projects}}个):
{{#each projects}}
{{@index}}. {{name}}
   描述: {{description}}
   状态: {{status}}
   团队规模: {{team_size}}人
   
{{/each}}

技能评估:
{{#each skills}}
- {{name}}: {{level}}/10 ({{years}}年经验)
{{/each}}

在表格中若需要删除一整行, 只需要在任意单元格上添加:
{{removeRow}}


{{#if (gt performance.score 90)}}
🎉 绩效评级: 优秀
{{else if (gt performance.score 80)}}
👍 绩效评级: 良好
{{else}}
📈 绩效评级: 需改进
{{/if}}
```

## 构建和开发

### 构建 WASM 包

```bash
# 构建所有目标
npm run build

# 或分别构建
npm run build:web    # 浏览器版本
npm run build:npm    # Node.js 版本 
npm run build:jsr    # Deno 版本
```

### 运行示例

```bash
# Rust 示例
cargo run --example rust_example

# Node.js 示例
node examples/node_example.js

# Deno 示例  
deno run --allow-read --allow-write examples/deno_example.ts

# 浏览器示例
cd tests/npm_test
node serve.js
# 然后在浏览器中打开 http://localhost:8080
# 选择 examples/template.xlsx 文件测试
```

## 工具函数

xlsx-handlebars 提供了一系列实用工具函数，帮助你更高效地处理 Excel 相关操作。

### 获取图片尺寸

从原始图片数据中检测图片尺寸，无需依赖完整的图片处理库。

```rust
use xlsx_handlebars::get_image_dimensions;

// 读取图片文件
let image_data = std::fs::read("logo.png")?;

// 获取尺寸
if let Some((width, height)) = get_image_dimensions(&image_data) {
    println!("图片尺寸: {}x{}", width, height);
} else {
    println!("不支持的图片格式");
}
```

**支持的格式**：
- PNG
- JPEG
- WebP (VP8, VP8L, VP8X)
- BMP
- TIFF (II/MM 字节序)
- GIF (87a/89a)

### Excel 列名转换

在 Excel 中进行列名和列索引之间的转换。

```rust
use xlsx_handlebars::{to_column_name, to_column_index};

// 列名递增
assert_eq!(to_column_name("A", 0), "A");
assert_eq!(to_column_name("A", 1), "B");
assert_eq!(to_column_name("Z", 1), "AA");
assert_eq!(to_column_name("AA", 1), "AB");

// 列名转索引 (1-based)
assert_eq!(to_column_index("A"), 1);
assert_eq!(to_column_index("Z"), 26);
assert_eq!(to_column_index("AA"), 27);
assert_eq!(to_column_index("BA"), 53);
```

**JavaScript/TypeScript 示例**：

```javascript
import { wasm_to_column_name, wasm_to_column_index } from 'xlsx-handlebars';

// 列名递增
console.log(wasm_to_column_name("A", 1));  // "B"
console.log(wasm_to_column_name("Z", 1));  // "AA"

// 列名转索引
console.log(wasm_to_column_index("AA"));   // 27
console.log(wasm_to_column_index("BA"));   // 53
```

### Excel 日期转换

在 Unix 时间戳和 Excel 日期序列号之间转换。Excel 使用从 1900-01-01 开始的序列号表示日期。

```rust
use xlsx_handlebars::{timestamp_to_excel_date, excel_date_to_timestamp};

// 时间戳转 Excel 日期
let timestamp = 1704067200000i64;  // 2024-01-01 00:00:00 UTC
let excel_date = timestamp_to_excel_date(timestamp);
println!("Excel 日期序列号: {}", excel_date);  // 45294.0

// Excel 日期转时间戳
if let Some(ts) = excel_date_to_timestamp(45294.0) {
    println!("时间戳: {}", ts);  // 1704067200000
}
```

**JavaScript/TypeScript 示例**：

```javascript
import { 
    wasm_timestamp_to_excel_date, 
    wasm_excel_date_to_timestamp 
} from 'xlsx-handlebars';

// 日期转 Excel 序列号
const date = new Date('2024-01-01T00:00:00Z');
const excelDate = wasm_timestamp_to_excel_date(date.getTime());
console.log('Excel 日期:', excelDate);  // 45294.0

// Excel 序列号转日期
const timestamp = wasm_excel_date_to_timestamp(45294.0);
if (timestamp !== null) {
    const convertedDate = new Date(timestamp);
    console.log('日期:', convertedDate.toISOString());
}
```

**常见使用场景**：

```rust
// 在模板中使用前验证图片尺寸
let image_data = std::fs::read("photo.jpg")?;
match get_image_dimensions(&image_data) {
    Some((w, h)) if w <= 1000 && h <= 1000 => {
        println!("有效图片: {}x{}", w, h);
        // 继续进行模板渲染
    }
    Some((w, h)) => {
        eprintln!("图片过大: {}x{} (最大 1000x1000)", w, h);
    }
    None => {
        eprintln!("不支持的图片格式");
    }
}
```

```rust
// 动态生成单元格引用
let start_col = "B";
let num_cols = 5;
for i in 0..num_cols {
    let col_name = to_column_name(start_col, i);
    let col_index = to_column_index(&col_name);
    println!("列 {}: 名称={}, 索引={}", i, col_name, col_index);
}
```

```rust
// 在模板数据中包含日期
use serde_json::json;

let date_timestamp = 1704067200000i64;  // 2024-01-01
let excel_date = timestamp_to_excel_date(date_timestamp);

let data = json!({
    "report_date": excel_date,
    "employee": {
        "name": "张三",
        "hire_date": timestamp_to_excel_date(1609459200000i64)  // 2021-01-01
    }
});
```

```rust
// 批量处理图片
for file in &["logo.png", "banner.jpg", "icon.gif"] {
    let data = std::fs::read(file)?;
    match get_image_dimensions(&data) {
        Some((w, h)) => println!("{}: {}x{}", file, w, h),
        None => eprintln!("{}: 不支持的格式", file),
    }
}
```

这些工具函数帮助你：
- ✅ 在插入前验证图片尺寸
- ✅ 动态生成单元格引用和公式
- ✅ 处理 Excel 日期格式
- ✅ 避免加载笨重的外部库
- ✅ 同时支持 Rust 和 JavaScript/TypeScript

## 技术特性

## 性能和兼容性

### 极致性能表现 ⚡

xlsx-handlebars 凭借 Rust 实现了**业界顶尖的性能表现**：

| 数据量 | 处理耗时 | 吞吐量 |
|--------|---------|--------|
| 1,000 行 | ~0.02秒 | 实时生成报表 |
| 10,000 行 | ~0.21秒 | 在线导出 |
| 100,000 行 | ~2.12秒 | 批量处理 |
| 1,000,000 行 | ~21秒 | 大数据报表 |

**性能对比** (处理10万行数据)：

| 技术栈 | 耗时 | 与 xlsx-handlebars 对比 |
|-------|------|------------------------|
| **xlsx-handlebars (Rust)** | **2.12秒** | **1倍 (基准)** ⭐ |
| Python (openpyxl) | 30-60秒 | 慢 14-28倍 |
| JavaScript (xlsx.js) | 15-30秒 | 慢 7-14倍 |
| Java (Apache POI) | 8-15秒 | 慢 3-7倍 |
| C# (EPPlus) | 5-10秒 | 慢 2-4倍 |

**为什么这么快？**
- 🦀 **Rust 零成本抽象**：编译期优化，无运行时开销
- 🔄 **流式架构**：直接在内存中处理 ZIP 条目，避免文件 I/O
- ⚡ **事件驱动 XML 解析**：使用 quick-xml 高效解析，无需构建完整 DOM 树
- 🎯 **单次遍历渲染**：一次迭代完成所有模板替换

### 兼容性

- **零拷贝**: Rust 和 WASM 之间高效的内存管理
- **流式处理**: 适合处理大型 XLSX 文件
- **跨平台**: 支持 Windows、macOS、Linux、Web
- **现代浏览器**: 支持所有支持 WASM 的现代浏览器

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE-MIT](LICENSE-MIT) 文件。

## 支持

- 📚 [文档](https://docs.rs/xlsx-handlebars)
- 🐛 [问题反馈](https://github.com/sail-sail/xlsx-handlebars/issues)
- 💬 [讨论](https://github.com/sail-sail/xlsx-handlebars/discussions)

---

<div align="center">
  <p>
    <strong>xlsx-handlebars</strong> - 让 XLSX 模板处理变得简单高效
  </p>
  <p>
    <a href="https://github.com/sail-sail/xlsx-handlebars">⭐ 给项目点个星</a>
    ·
    <a href="https://github.com/sail-sail/xlsx-handlebars/issues">🐛 报告问题</a>
    ·
    <a href="https://github.com/sail-sail/xlsx-handlebars/discussions">💬 参与讨论</a>
  </p>
</div>


## 捐赠鼓励支持此项目,支付宝扫码:
![捐赠鼓励支持此项目](https://www.ejsexcel.com/alipay.jpg)
