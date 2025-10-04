/* tslint:disable */
/* eslint-disable */
export function render_template(zip_bytes: Uint8Array, data_json: string): any;
export function wasm_get_image_dimensions(data: Uint8Array): any;
export function wasm_to_column_name(current: string, increment: number): string;
export function wasm_to_column_index(col_name: string): number;
export function wasm_timestamp_to_excel_date(timestamp_ms: bigint): number;
export function wasm_excel_date_to_timestamp(excel_date: number): bigint | undefined;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
  readonly memory: WebAssembly.Memory;
  readonly render_template: (a: number, b: number, c: number, d: number) => [number, number, number];
  readonly wasm_get_image_dimensions: (a: number, b: number) => any;
  readonly wasm_to_column_name: (a: number, b: number, c: number) => [number, number];
  readonly wasm_to_column_index: (a: number, b: number) => number;
  readonly wasm_timestamp_to_excel_date: (a: bigint) => number;
  readonly wasm_excel_date_to_timestamp: (a: number) => [number, bigint];
  readonly __wbindgen_exn_store: (a: number) => void;
  readonly __externref_table_alloc: () => number;
  readonly __wbindgen_export_2: WebAssembly.Table;
  readonly __wbindgen_malloc: (a: number, b: number) => number;
  readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
  readonly __wbindgen_free: (a: number, b: number, c: number) => void;
  readonly __externref_table_dealloc: (a: number) => void;
  readonly __wbindgen_start: () => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;
/**
* Instantiates the given `module`, which can either be bytes or
* a precompiled `WebAssembly.Module`.
*
* @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
*
* @returns {InitOutput}
*/
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
* If `module_or_path` is {RequestInfo} or {URL}, makes a request and
* for everything else, calls `WebAssembly.instantiate` directly.
*
* @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
*
* @returns {Promise<InitOutput>}
*/
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
