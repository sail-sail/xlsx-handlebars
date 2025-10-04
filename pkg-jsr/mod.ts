/**
 * Deno module entry point for xlsx-handlebars
 * 
 * @example
 * ```typescript
 * import { render, init } from "https://deno.land/x/xlsx_handlebars/mod.ts";
 * 
 * // Initialize WASM module
 * await init();
 * 
 * const templateBytes = await Deno.readFile("template.xlsx");
 * const data = { name: "张三", company: "ABC公司" };
 * const result = render(templateBytes, JSON.stringify(data));
 * 
 * await Deno.writeFile("output.xlsx", new Uint8Array(result));
 * ```
 */

// Import the WASM module and its functions
import init, { render_template } from "./xlsx_handlebars.js";

// Re-export the main functions
export { render_template, init };
export default init;

/**
 * Deno-specific utility functions
 */
export class XlsxHandlebarsUtils {
  /**
   * Load WASM module for Deno environment
   */
  static async initWasm(): Promise<void> {
    // Auto-initialize WASM in Deno
    const { default: init } = await import("./pkg-npm/xlsx_handlebars.js");
    await init();
  }

  /**
   * Read XLSX file from file system
   */
  static async readXlsxFile(filePath: string): Promise<Uint8Array> {
    return await Deno.readFile(filePath);
  }

  /**
   * Write XLSX file to file system
   */
  static async writeXlsxFile(filePath: string, data: Uint8Array): Promise<void> {
    await Deno.writeFile(filePath, data);
  }

  /**
   * Check if file exists
   */
  static async fileExists(filePath: string): Promise<boolean> {
    try {
      await Deno.stat(filePath);
      return true;
    } catch {
      return false;
    }
  }
}
