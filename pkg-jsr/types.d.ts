/**
 * Type definitions for xlsx-handlebars
 */

export declare class XlsxHandlebars {
  /**
   * Creates a new XlsxHandlebars instance
   */
  constructor();

  /**
   * Free the memory used by this instance
   */
  free(): void;

  /**
   * Load a XLSX template file
   * @param bytes - The XLSX file as a Uint8Array
   */
  load_template(bytes: Uint8Array): void;

  /**
   * Render the template with the given data
   * @param data_json - The data as a JSON string
   * @returns The rendered XLSX file as a Uint8Array
   */
  render(data_json: string): Uint8Array;

  /**
   * Get the list of variables in the template
   * @returns A JSON string containing the list of variables
   */
  get_template_variables(): string;
}

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
  readonly memory: WebAssembly.Memory;
  readonly __wbg_xlsxhandlebars_free: (a: number, b: number) => void;
  readonly xlsxhandlebars_new: () => number;
  readonly xlsxhandlebars_load_template: (a: number, b: number, c: number) => [number, number];
  readonly xlsxhandlebars_render: (a: number, b: number, c: number) => [number, number, number, number];
  readonly xlsxhandlebars_get_template_variables: (a: number) => [number, number, number, number];
  readonly __wbindgen_free: (a: number, b: number, c: number) => void;
  readonly __wbindgen_malloc: (a: number, b: number) => number;
  readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
  readonly __wbindgen_export_3: WebAssembly.Table;
  readonly __externref_table_dealloc: (a: number) => void;
  readonly __wbindgen_start: () => void;
}

/**
 * Initialize the WASM module
 * @param module_or_path - The WASM module or path to it
 * @returns A promise that resolves to the initialized output
 */
export default function init(module_or_path?: InitInput | Promise<InitInput>): Promise<InitOutput>;
