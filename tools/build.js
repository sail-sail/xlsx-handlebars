#!/usr/bin/env node

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

console.log('🚀 Building xlsx-handlebars for all platforms...');

function checkAndInstallWasmBindgen() {
  console.log('\n🔍 Checking wasm-bindgen-cli installation...');
  try {
    // 检查是否已安装 wasm-bindgen
    execSync('wasm-bindgen --version', { stdio: 'pipe' });
    console.log('✓ wasm-bindgen-cli is already installed');
    return;
  } catch (error) {
    // 未安装，尝试使用 cargo-binstall 安装
    console.log('⚠ wasm-bindgen-cli not found, attempting to install...');
    
    try {
      // 先检查 cargo-binstall 是否已安装
      execSync('cargo binstall --version', { stdio: 'pipe' });
      console.log('✓ cargo-binstall found, using it to install wasm-bindgen-cli...');
      execSync('cargo binstall -y wasm-bindgen-cli', { stdio: 'inherit' });
      console.log('✅ wasm-bindgen-cli installed successfully via cargo-binstall');
    } catch (binstallError) {
      // cargo-binstall 也没有，使用 cargo install
      console.log('⚠ cargo-binstall not found, falling back to cargo install...');
      console.log('💡 Tip: Install cargo-binstall for faster future installations:');
      console.log('   cargo install cargo-binstall');
      execSync('cargo install wasm-bindgen-cli', { stdio: 'inherit' });
      console.log('✅ wasm-bindgen-cli installed successfully via cargo install');
    }
  }
}

function runCommand(command, description) {
  console.log(`\n📦 ${description}...`);
  try {
    execSync(command, { stdio: 'inherit' });
    console.log(`✅ ${description} completed successfully`);
  } catch (error) {
    console.error(`❌ ${description} failed!`);
    process.exit(1);
  }
}

function copyFile(src, dest, description) {
  try {
    if (fs.existsSync(src)) {
      fs.copyFileSync(src, dest);
      console.log(`✓ Copied ${description}`);
    } else {
      console.warn(`⚠ ${src} not found, skipping`);
    }
  } catch (error) {
    console.error(`❌ Failed to copy ${description}:`, error.message);
  }
}

function removeGitignoreFiles(directories) {
  console.log('\n🧹 Cleaning up auto-generated .gitignore files...');
  directories.forEach(dir => {
    const gitignorePath = path.join(dir, '.gitignore');
    try {
      if (fs.existsSync(gitignorePath)) {
        fs.unlinkSync(gitignorePath);
        console.log(`✓ Removed ${gitignorePath}`);
      } else {
        console.log(`✓ No .gitignore found in ${dir}`);
      }
    } catch (error) {
      console.error(`❌ Failed to remove ${gitignorePath}:`, error.message);
    }
  });
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

function addMainFieldToPackageJson(packageJsonPath) {
  try {
    if (fs.existsSync(packageJsonPath)) {
      const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
      
      // 检查是否已经有 main 字段
      if (!packageJson.main && packageJson.module) {
        // 在 module 字段前面添加 main 字段
        const { module, ...rest } = packageJson;
        const updatedPackageJson = {
          ...rest,
          main: module, // 使用相同的文件作为 main 入口
          module,
        };
        
        fs.writeFileSync(packageJsonPath, JSON.stringify(updatedPackageJson, null, 2) + '\n');
        console.log(`✓ Added "main" field to ${path.basename(packageJsonPath)}`);
      } else if (packageJson.main) {
        console.log(`✓ "main" field already exists in ${path.basename(packageJsonPath)}`);
      }
    }
  } catch (error) {
    console.error(`❌ Failed to update ${packageJsonPath}:`, error.message);
  }
}

function removeUnnecessaryFiles() {
  console.log('\n🧹 Cleaning up unnecessary files...');
  
  // 删除 pkg-jsr 下的 package.json
  const jsrFilesToRemove = ['pkg-jsr/package.json'];
  jsrFilesToRemove.forEach(filePath => {
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
        console.log(`✓ Removed ${filePath}`);
      } else {
        console.log(`✓ ${filePath} not found, skipping`);
      }
    } catch (error) {
      console.error(`❌ Failed to remove ${filePath}:`, error.message);
    }
  });
}

// 检查并安装 wasm-bindgen-cli
checkAndInstallWasmBindgen();

// 构建 Rust 库
runCommand('cargo build --release', 'Building Rust library');

// 构建 npm 包
runCommand('wasm-pack build --target web --out-dir pkg-npm', 'Building WASM for npm');

// 为 npm 包添加 main 字段以支持 CommonJS
console.log('\n🔧 Updating npm package.json...');
addMainFieldToPackageJson('pkg-npm/package.json');

// 构建 JSR 包
runCommand('wasm-pack build --target web --out-dir pkg-jsr', 'Building JSR package WASM files');

// 使用 wasm-opt 进一步优化 WASM 文件
console.log('\n⚡ Optimizing WASM files with wasm-opt...');
try {
  execSync('wasm-opt --version', { stdio: 'pipe' });
  
  // 优化 npm 包的 WASM
  if (fs.existsSync('pkg-npm/xlsx_handlebars_bg.wasm')) {
    const npmWasmSize = fs.statSync('pkg-npm/xlsx_handlebars_bg.wasm').size;
    runCommand('wasm-opt -Oz pkg-npm/xlsx_handlebars_bg.wasm -o pkg-npm/xlsx_handlebars_bg.wasm', 
               'Optimizing npm WASM');
    const npmOptimizedSize = fs.statSync('pkg-npm/xlsx_handlebars_bg.wasm').size;
    const npmSavings = ((1 - npmOptimizedSize / npmWasmSize) * 100).toFixed(1);
    console.log(`  ✓ npm WASM: ${(npmWasmSize / 1024).toFixed(1)}KB → ${(npmOptimizedSize / 1024).toFixed(1)}KB (saved ${npmSavings}%)`);
  }
  
  // 优化 JSR 包的 WASM
  if (fs.existsSync('pkg-jsr/xlsx_handlebars_bg.wasm')) {
    const jsrWasmSize = fs.statSync('pkg-jsr/xlsx_handlebars_bg.wasm').size;
    runCommand('wasm-opt -Oz pkg-jsr/xlsx_handlebars_bg.wasm -o pkg-jsr/xlsx_handlebars_bg.wasm', 
               'Optimizing JSR WASM');
    const jsrOptimizedSize = fs.statSync('pkg-jsr/xlsx_handlebars_bg.wasm').size;
    const jsrSavings = ((1 - jsrOptimizedSize / jsrWasmSize) * 100).toFixed(1);
    console.log(`  ✓ JSR WASM: ${(jsrWasmSize / 1024).toFixed(1)}KB → ${(jsrOptimizedSize / 1024).toFixed(1)}KB (saved ${jsrSavings}%)`);
  }
} catch (error) {
  console.warn('⚠ wasm-opt not found, skipping additional optimization');
  console.warn('  Install with: cargo install wasm-opt');
}

// 确保 pkg-jsr 目录存在
ensureDir('pkg-jsr');

// 复制必要文件到 pkg-jsr 目录
console.log('\n📋 Copying additional files to JSR package...');
copyFile('LICENSE-MIT', 'pkg-jsr/LICENSE-MIT', 'LICENSE-MIT');

console.log('✅ JSR package files are ready');

// 清理自动生成的 .gitignore 文件
removeGitignoreFiles(['pkg-npm', 'pkg-jsr']);

// 删除不必要的文件
removeUnnecessaryFiles();

console.log('\n🎉 All builds completed successfully!');
console.log('\n📁 Output directories:');
console.log('  - Rust: target/release/');
console.log('  - npm: pkg-npm/ (supports both Node.js and browsers)');
console.log('  - JSR: pkg-jsr/ (supports Deno and Node.js)');
console.log('\n🚀 Ready for publishing!');
console.log('\n📦 Publishing commands:');
console.log('  npm run publish');
