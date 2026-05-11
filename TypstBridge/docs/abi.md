# TypstBridge ABI Draft

This document defines the planned C ABI between the Rust Typst bridge and .NET.

The ABI is intentionally C-compatible. Rust-owned memory must always be released by Rust-provided free functions. C# must never free Rust allocations directly.

## Design Rules

- All string inputs are UTF-8 with explicit lengths.
- All returned buffers and strings are owned by the bridge.
- The caller must release returned results with `typst_bridge_free_result`.
- Panics must never cross the FFI boundary.
- Compile results must support one output item for PDF and one output item per page for PNG/SVG.

## Status Codes

```c
typedef enum typst_bridge_status {
    TYPST_BRIDGE_OK = 0,
    TYPST_BRIDGE_ERR_INVALID_ARGUMENT = 1,
    TYPST_BRIDGE_ERR_COMPILE = 2,
    TYPST_BRIDGE_ERR_RENDER = 3,
    TYPST_BRIDGE_ERR_IO = 4,
    TYPST_BRIDGE_ERR_PANIC = 5,
    TYPST_BRIDGE_ERR_UNSUPPORTED = 6,
    TYPST_BRIDGE_ERR_INTERNAL = 255
} typst_bridge_status;
```

## Output Format

```c
typedef enum typst_bridge_output_format {
    TYPST_BRIDGE_OUTPUT_PDF = 1,
    TYPST_BRIDGE_OUTPUT_PNG = 2,
    TYPST_BRIDGE_OUTPUT_SVG = 3
} typst_bridge_output_format;
```

## Compile Request

```c
typedef struct typst_bridge_compile_request {
    uint32_t abi_version;

    const char* source_utf8;
    uintptr_t source_len;

    const char* working_dir_utf8;
    uintptr_t working_dir_len;

    const char* root_file_name_utf8;
    uintptr_t root_file_name_len;

    const char** font_paths_utf8;
    uintptr_t font_paths_count;

    typst_bridge_output_format output_format;
    double ppi;
    uint32_t flags;
} typst_bridge_compile_request;
```

`working_dir_utf8` is required for PPTX-generated sources because the converter emits relative asset paths such as `assets/...`.

## Output Item

```c
typedef struct typst_bridge_output_item {
    uint32_t page_index;
    char* file_name_utf8;
    uintptr_t file_name_len;
    uint8_t* data;
    uintptr_t data_len;
} typst_bridge_output_item;
```

Expected behavior:

- PDF returns one item with `page_index = 0`.
- PNG returns one item per page, ordered by page index.
- SVG returns one item per page, ordered by page index.

## Diagnostic Item

```c
typedef struct typst_bridge_diagnostic {
    uint32_t severity;
    char* message_utf8;
    uintptr_t message_len;
    char* file_utf8;
    uintptr_t file_len;
    uint32_t line;
    uint32_t column;
} typst_bridge_diagnostic;
```

Severity values:

```c
#define TYPST_BRIDGE_DIAG_ERROR 1
#define TYPST_BRIDGE_DIAG_WARNING 2
#define TYPST_BRIDGE_DIAG_INFO 3
```

## Compile Result

```c
typedef struct typst_bridge_compile_result {
    typst_bridge_status status;

    typst_bridge_output_item* outputs;
    uintptr_t outputs_count;

    typst_bridge_diagnostic* diagnostics;
    uintptr_t diagnostics_count;

    char* message_utf8;
    uintptr_t message_len;
} typst_bridge_compile_result;
```

## Functions

```c
uint32_t typst_bridge_abi_version(void);
const char* typst_bridge_version(void);
typst_bridge_status typst_bridge_probe(void);

typst_bridge_compile_result* typst_bridge_compile(
    const typst_bridge_compile_request* request
);

void typst_bridge_free_result(
    typst_bridge_compile_result* result
);

void typst_bridge_free_string(
    char* value
);

const char* typst_bridge_last_error_message(void);
```

## Acceptance Criteria

- Invalid requests return structured errors where possible.
- Panics are caught and returned as `TYPST_BRIDGE_ERR_PANIC`.
- Repeated compile/free cycles do not crash.
- Output signatures are correct: `%PDF`, PNG magic bytes, and SVG/XML text.
