// Node.js example for xlsx-handlebars utility functions
// node examples/utils_example.js
const fs = require("node:fs");
const path = require("node:path");

async function utilsExample() {
    try {
        // 导入 WASM 模块
        const { 
            default: init, 
            wasm_to_column_name,
            wasm_to_column_index,
            wasm_timestamp_to_excel_date,
            wasm_excel_date_to_timestamp,
            wasm_get_image_dimensions
        } = await import('../pkg-npm/xlsx_handlebars.js');
        
        // 初始化 WASM
        const wasmPath = path.join(__dirname, '../pkg-npm/xlsx_handlebars_bg.wasm');
        const wasmBuffer = fs.readFileSync(wasmPath);
        await init(wasmBuffer);
        
        console.log('🔧 xlsx-handlebars 工具函数示例\n');
        
        // 1. Excel 列名转换
        console.log('📊 Excel 列名转换示例：');
        console.log(`   列名 "A" + 0 = "${wasm_to_column_name("A", 0)}"`);
        console.log(`   列名 "A" + 1 = "${wasm_to_column_name("A", 1)}"`);
        console.log(`   列名 "Z" + 1 = "${wasm_to_column_name("Z", 1)}"`);
        console.log(`   列名 "AA" + 1 = "${wasm_to_column_name("AA", 1)}"`);
        console.log(`   列名 "AZ" + 1 = "${wasm_to_column_name("AZ", 1)}"`);
        console.log();
        
        // 2. Excel 列名转索引
        console.log('🔢 Excel 列名转索引示例：');
        console.log(`   列名 "A" = 索引 ${wasm_to_column_index("A")}`);
        console.log(`   列名 "Z" = 索引 ${wasm_to_column_index("Z")}`);
        console.log(`   列名 "AA" = 索引 ${wasm_to_column_index("AA")}`);
        console.log(`   列名 "BA" = 索引 ${wasm_to_column_index("BA")}`);
        console.log(`   列名 "ZZ" = 索引 ${wasm_to_column_index("ZZ")}`);
        console.log();
        
        // 3. 日期转 Excel 序列号
        console.log('📅 日期转 Excel 序列号示例：');
        const date1 = new Date('2024-01-01T00:00:00Z');
        const timestamp1 = BigInt(date1.getTime());
        const excelDate1 = wasm_timestamp_to_excel_date(timestamp1);
        console.log(`   日期: ${date1.toISOString()}`);
        console.log(`   时间戳: ${timestamp1}`);
        console.log(`   Excel 序列号: ${excelDate1}`);
        console.log();
        
        const date2 = new Date('2025-10-04T00:00:00Z');
        const timestamp2 = BigInt(date2.getTime());
        const excelDate2 = wasm_timestamp_to_excel_date(timestamp2);
        console.log(`   日期: ${date2.toISOString()}`);
        console.log(`   时间戳: ${timestamp2}`);
        console.log(`   Excel 序列号: ${excelDate2}`);
        console.log();
        
        // 4. Excel 序列号转日期
        console.log('🔄 Excel 序列号转日期示例：');
        const excelNum1 = 45294.0;
        const convertedTimestamp1 = wasm_excel_date_to_timestamp(excelNum1);
        if (convertedTimestamp1 !== null && convertedTimestamp1 !== undefined) {
            const convertedDate1 = new Date(Number(convertedTimestamp1));
            console.log(`   Excel 序列号: ${excelNum1}`);
            console.log(`   转换后时间戳: ${convertedTimestamp1}`);
            console.log(`   转换后日期: ${convertedDate1.toISOString()}`);
        }
        console.log();
        
        const excelNum2 = 25571.0; // 1970-01-01
        const convertedTimestamp2 = wasm_excel_date_to_timestamp(excelNum2);
        if (convertedTimestamp2 !== null && convertedTimestamp2 !== undefined) {
            const convertedDate2 = new Date(Number(convertedTimestamp2));
            console.log(`   Excel 序列号: ${excelNum2}`);
            console.log(`   转换后时间戳: ${convertedTimestamp2}`);
            console.log(`   转换后日期: ${convertedDate2.toISOString()}`);
        }
        console.log();
        
        // 5. 图片尺寸获取
        console.log('🖼️  图片尺寸获取示例：');
        const imagePath = path.join(__dirname, 'image.png');
        if (fs.existsSync(imagePath)) {
            const imageBuffer = fs.readFileSync(imagePath);
            const imageArray = new Uint8Array(imageBuffer);
            const dimensions = wasm_get_image_dimensions(imageArray);
            
            if (dimensions !== null && dimensions !== undefined) {
                console.log(`   图片路径: ${imagePath}`);
                console.log(`   宽度: ${dimensions.width}px`);
                console.log(`   高度: ${dimensions.height}px`);
            } else {
                console.log(`   ⚠️  无法识别图片格式`);
            }
        } else {
            console.log(`   ⚠️  图片文件不存在: ${imagePath}`);
        }
        console.log();
        
        // 6. 实用场景示例：动态生成 Excel 单元格引用
        console.log('💡 实用场景：动态生成单元格引用');
        const startCol = "B";
        const numCols = 5;
        console.log(`   从列 "${startCol}" 开始生成 ${numCols} 列：`);
        for (let i = 0; i < numCols; i++) {
            const colName = wasm_to_column_name(startCol, i);
            const colIndex = wasm_to_column_index(colName);
            console.log(`   [${i}] 列名: ${colName}, 索引: ${colIndex}`);
        }
        console.log();
        
        console.log('✅ 所有工具函数示例执行完成！');
        
    } catch (error) {
        console.error('❌ 错误:', error);
        console.error('\n提示: 请确保先运行 `npm run build:npm` 构建 WASM 包');
    }
}

// 运行示例
utilsExample().catch(console.error);
