# xlsx-handlebars

[![Crates.io](https://img.shields.io/crates/v/xlsx-handlebars.svg)](https://crates.io/crates/xlsx-handlebars)
[![Documentation](https://docs.rs/xlsx-handlebars/badge.svg)](https://docs.rs/xlsx-handlebars)
[![License](https://img.shields.io/crates/l/xlsx-handlebars.svg)](https://github.com/sail-sail/xlsx-handlebars#license)

English | [中文文档](README_CN.md)

A Rust library for processing XLSX files with Handlebars templates, supporting multiple platforms:
- 🦀 Rust Native
- 🌐 WebAssembly (WASM)
- 📦 npm Package
- 🟢 Node.js
- 🦕 Deno
- 🌍 Browser
- 📋 JSR (JavaScript Registry)

## Features

- ✅ **Smart Merge**: Automatically handles Handlebars syntax split by XML tags
- ✅ **XLSX Validation**: Built-in file format validation to ensure valid input files
- ✅ **Handlebars Support**: Full template engine with variables, conditions, loops, and Helper functions
- ✅ **Cross-Platform**: Rust native + WASM support for multiple runtimes
- ✅ **TypeScript**: Complete type definitions and IntelliSense
- ✅ **Zero Dependencies**: WASM binary with no external dependencies
- ✅ **Error Handling**: Detailed error messages and type-safe error handling

## Installation

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

## Usage Examples

### Rust

```rust
use xlsx_handlebars::render_handlebars;
use serde_json::json;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Read XLSX template file
    let template_bytes = std::fs::read("template.xlsx")?;
    
    // Prepare data
    let data = json!({
        "name": "John Doe",
        "company": "ABC Tech Inc.",
        "position": "Software Engineer",
        "projects": [
            {"name": "Project A", "status": "Completed"},
            {"name": "Project B", "status": "In Progress"}
        ],
        "has_bonus": true,
        "bonus_amount": 5000
    });
    
    // Render template
    let result = render_handlebars(template_bytes, &data)?;
    
    // Save result
    std::fs::write("output.xlsx", result)?;
    
    Ok(())
}
```

### JavaScript/TypeScript (Node.js)

```javascript
import { render, init } from 'xlsx-handlebars';
import fs from 'fs';

async function processTemplate() {
    // Initialize WASM module
    await init();
    
    // Read template file
    const templateBytes = fs.readFileSync('template.xlsx');
    
    // Prepare data
    const data = {
        name: "Jane Smith",
        company: "XYZ Technology Ltd.",
        position: "Senior Developer",
        projects: [
            { name: "E-commerce Platform", status: "Completed" },
            { name: "Mobile App", status: "In Development" }
        ],
        has_bonus: true,
        bonus_amount: 8000
    };
    
    // Render template
    const result = render(templateBytes, JSON.stringify(data));
    
    // Save result
    fs.writeFileSync('output.xlsx', new Uint8Array(result));
}

processTemplate().catch(console.error);
```

### Deno

```typescript
import { render, init } from "https://deno.land/x/xlsx_handlebars/mod.ts";

async function processTemplate() {
    // Initialize WASM module
    await init();
    
    // Read template file
    const templateBytes = await Deno.readFile("template.xlsx");
    
    // Prepare data
    const data = {
        name: "Alice Johnson",
        department: "R&D",
        projects: [
            { name: "AI Customer Service", status: "Live" },
            { name: "Data Visualization Platform", status: "In Development" }
        ]
    };
    
    // Render template
    const result = render(templateBytes, JSON.stringify(data));
    
    // Save result
    await Deno.writeFile("output.xlsx", new Uint8Array(result));
}

if (import.meta.main) {
    await processTemplate();
}
```

### Browser

```html
<!DOCTYPE html>
<html>
<head>
    <title>XLSX Handlebars Example</title>
</head>
<body>
    <input type="file" id="fileInput" accept=".xlsx">
    <button onclick="processFile()">Process Template</button>
    
    <script type="module">
        import { render, init } from './pkg/xlsx_handlebars.js';
        
        // Initialize WASM
        await init();
        
        window.processFile = async function() {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];
            
            if (!file) return;
            
            const arrayBuffer = await file.arrayBuffer();
            const templateBytes = new Uint8Array(arrayBuffer);
            
            const data = {
                name: "John Doe",
                company: "Example Company"
            };
            
            try {
                const result = render(templateBytes, JSON.stringify(data));
                
                // Download result
                const blob = new Blob([new Uint8Array(result)], {
                    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'processed.xlsx';
                a.click();
            } catch (error) {
                console.error('Processing failed:', error);
            }
        };
    </script>
</body>
</html>
```

## Template Syntax

### Basic Variable Substitution

```handlebars
Employee Name: {{name}}
Company: {{company}}
Position: {{position}}
```

### Conditional Rendering

```handlebars
{{#if has_bonus}}
Bonus: ${{bonus_amount}}
{{else}}
No Bonus
{{/if}}

{{#unless is_intern}}
Full-time Employee
{{/unless}}
```

### Loop Rendering

```handlebars
Project Experience:
{{#each projects}}
- {{name}}: {{description}} ({{status}})
{{/each}}

Skills:
{{#each skills}}
{{@index}}. {{this}}
{{/each}}
```

### Helper Functions

Built-in Helper functions:

```handlebars
<!-- Basic helpers -->
{{upper name}}           <!-- Convert to uppercase -->
{{lower company}}        <!-- Convert to lowercase -->
{{len projects}}         <!-- Array length -->
{{#if (eq status "completed")}}Completed{{/if}}    <!-- Equality comparison -->
{{#if (gt score 90)}}Excellent{{/if}}              <!-- Greater than comparison -->
{{#if (lt age 30)}}Young{{/if}}                    <!-- Less than comparison -->

<!-- String concatenation -->
{{concat "Hello" " " "World"}}                     <!-- String concatenation -->
{{concat "Total: " count}}                         <!-- Mix strings and variables -->

<!-- Excel-specific helpers -->
{{num employee.salary}}                            <!-- Mark cell as number type -->
{{formula "=SUM(A1:B1)"}}                         <!-- Static Excel formula -->
{{formula (concat "=SUM(" (_c) "1:" (_c) "10)")}} <!-- Dynamic formula with current column -->
```

#### Excel Formula Helpers

**Static Formula**:
```handlebars
<!-- In Excel cell -->
{{formula "=SUM(A1:B1)"}}
{{formula "=AVERAGE(C2:C10)"}}
{{formula "=IF(D1>100,\"High\",\"Low\")"}}
```

**Dynamic Formula with `concat`**:
```handlebars
<!-- Dynamic row reference -->
{{formula (concat "=A" (_r) "*B" (_r))}}

<!-- Dynamic column reference -->
{{formula (concat "=SUM(" (_c) "2:" (_c) "10)")}}

<!-- Complex dynamic formula -->
{{formula (concat "=IF(" (_cr) ">100,\"High\",\"Low\")")}}
```

**Available position helpers**:
- `(_c)` - Current column letter (A, B, C, ...)
- `(_r)` - Current row number (1, 2, 3, ...)
- `(_cr)` - Current cell reference (A1, B2, C3, ...)

#### Number Type Helper

Use `{{num value}}` to ensure a cell is treated as a number in Excel:

```handlebars
<!-- Without num: treated as text -->
{{employee.salary}}

<!-- With num: treated as number -->
{{num employee.salary}}
```

This is especially useful when:
- The value might be a string but should be treated as a number
- You want to ensure proper number formatting in Excel
- You need the value to work in formulas

### Complex Example

```handlebars
=== Employee Report ===

Basic Information:
Name: {{employee.name}}
Department: {{employee.department}}
Position: {{employee.position}}
Hire Date: {{employee.hire_date}}

{{#if employee.has_bonus}}
💰 Bonus: ${{employee.bonus_amount}}
{{/if}}

Project Experience (Total {{len projects}}):
{{#each projects}}
{{@index}}. {{name}}
   Description: {{description}}
   Status: {{status}}
   Team Size: {{team_size}} people
   
{{/each}}

Skills Assessment:
{{#each skills}}
- {{name}}: {{level}}/10 ({{years}} years of experience)
{{/each}}

To remove an entire row in a table, simply add to any cell:
{{removeRow}}


{{#if (gt performance.score 90)}}
🎉 Performance Rating: Excellent
{{else if (gt performance.score 80)}}
👍 Performance Rating: Good
{{else}}
📈 Performance Rating: Needs Improvement
{{/if}}
```

## Error Handling

The library provides detailed error types and messages:

### Rust

```rust
use xlsx_handlebars::{render_handlebars, XlsxError};

match render_handlebars(template_bytes, &data) {
    Ok(result) => {
        println!("Processing successful!");
        std::fs::write("output.xlsx", result)?;
    }
    Err(e) => match e.downcast_ref::<XlsxError>() {
        Some(XlsxError::InvalidZipFormat) => {
            eprintln!("Error: Invalid XLSX file format");
        }
        _ => {
            eprintln!("Other error: {}", e);
        }
    }
}
```

### JavaScript/TypeScript

```javascript
try {
    const result = render(templateBytes, JSON.stringify(data));
    console.log('Processing successful!');
} catch (error) {
    console.error('Processing failed:', error);
}
```

## Build and Development

### Build WASM Package

```bash
# Build all targets
npm run build

# Or build separately
npm run build:web    # Browser version
npm run build:npm    # Node.js version 
npm run build:jsr    # Deno version
```

### Run Examples

```bash
# Rust example
cargo run --example rust_example

# Node.js example
node examples/node_example.js

# Deno example  
deno run --allow-read --allow-write examples/deno_example.ts

# Browser example
cd tests/npm_test
node serve.js
# Then open http://localhost:8080 in your browser
# Select examples/template.xlsx file to test
```

## Technical Features

### Smart Merge Algorithm

The core innovation of this library is the smart merge of Handlebars syntax split by XML tags. In XLSX files, when users input template syntax, Excel may split it into multiple XML tags.

## Performance and Compatibility

- **Zero-Copy**: Efficient memory management between Rust and WASM
- **Streaming**: Suitable for processing large XLSX files
- **Cross-Platform**: Supports Windows, macOS, Linux, Web
- **Modern Browsers**: Supports all modern browsers with WASM support

## License

This project is licensed under the MIT License - see the [LICENSE-MIT](LICENSE-MIT) file for details.

## Support

- 📚 [Documentation](https://docs.rs/xlsx-handlebars)
- 🐛 [Issues](https://github.com/sail-sail/xlsx-handlebars/issues)
- 💬 [Discussions](https://github.com/sail-sail/xlsx-handlebars/discussions)

---

<div align="center">
  <p>
    <strong>xlsx-handlebars</strong> - Making XLSX template processing simple and efficient
  </p>
  <p>
    <a href="https://github.com/sail-sail/xlsx-handlebars">⭐ Star the project</a>
    ·
    <a href="https://github.com/sail-sail/xlsx-handlebars/issues">🐛 Report Issues</a>
    ·
    <a href="https://github.com/sail-sail/xlsx-handlebars/discussions">💬 Join Discussions</a>
  </p>
</div>


## Support this project with a donation via Alipay:
![Support this project with a donation](https://ejsexcel.com/alipay.jpg)
