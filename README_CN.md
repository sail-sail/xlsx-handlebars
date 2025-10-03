# xlsx-handlebars

[![Crates.io](https://img.shields.io/crates/v/xlsx-handlebars.svg)](https://crates.io/crates/xlsx-handlebars)
[![Documentation](https://docs.rs/xlsx-handlebars/badge.svg)](https://docs.rs/xlsx-handlebars)
[![License](https://img.shields.io/crates/l/xlsx-handlebars.svg)](https://github.com/sail-sail/xlsx-handlebars#license)

中文文档 | [English](README.md)

一个用于处理 XLSX 文件 Handlebars 模板的 Rust 库，支持多平台使用：
- 🦀 Rust 原生
- 🌐 WebAssembly (WASM)
- 📦 npm 包
- 🟢 Node.js
- 🦕 Deno
- 🌍 浏览器端
- 📋 JSR (JavaScript Registry)

## 功能特性

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
import { render, init } from "jsr:@sail/xlsx-handlebars";
```

## 使用示例

### Rust

```rust
use xlsx_handlebars::render_handlebars;
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
    let result = render_handlebars(template_bytes, &data)?;
    
    // 保存结果
    std::fs::write("output.xlsx", result)?;
    
    Ok(())
}
```

### JavaScript/TypeScript (Node.js)

```javascript
import { render, init } from 'xlsx-handlebars';
import fs from 'fs';

async function processTemplate() {
    // 初始化 WASM 模块
    await init();
    
    // 读取模板文件
    const templateBytes = fs.readFileSync('template.xlsx');
    
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
    const result = render(templateBytes, JSON.stringify(data));
    
    // 保存结果
    fs.writeFileSync('output.xlsx', new Uint8Array(result));
}

processTemplate().catch(console.error);
```

### Deno

```typescript
import { render, init } from "https://deno.land/x/xlsx_handlebars/mod.ts";

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
    const result = render(templateBytes, JSON.stringify(data));
    
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
        import { render, init } from './pkg/xlsx_handlebars.js';
        
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
                const result = render(templateBytes, JSON.stringify(data));
                
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

## 错误处理

库提供了详细的错误类型和消息：

### Rust

```rust
use xlsx_handlebars::{render_handlebars, XlsxError};

match render_handlebars(template_bytes, &data) {
    Ok(result) => {
        println!("处理成功！");
        std::fs::write("output.xlsx", result)?;
    }
    Err(e) => match e.downcast_ref::<XlsxError>() {
        Some(XlsxError::InvalidZipFormat) => {
            eprintln!("错误: 无效的 XLSX 文件格式");
        }
        _ => {
            eprintln!("其他错误: {}", e);
        }
    }
}
```

### JavaScript/TypeScript

```javascript
try {
    const result = render(templateBytes, JSON.stringify(data));
    console.log('处理成功！');
} catch (error) {
    console.error('处理失败:', error);
}
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

## 技术特性

### 智能合并算法

该库的核心创新是智能合并被 XML 标签分割的 Handlebars 语法。在 XLSX 文件中，当用户输入模板语法时，Word 可能会将其拆分成多个 XML 标签

## 性能和兼容性

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
![捐赠鼓励支持此项目](https://ejsexcel.com/alipay.jpg)
