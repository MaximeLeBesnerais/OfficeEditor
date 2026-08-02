# OfficeEditor Examples

Each example page shows the **input → transformation → output**. The console runner exercises the core numbered examples; later sections may be standalone, tested library snippets where explicitly noted.

## Quick Start

```bash
cd examples
dotnet run
```

All outputs are written to `examples/output/`.

---

## Example Index

| Format | Examples | Output Folder |
|--------|----------|---------------|
| [DOCX](Docx/README.md) | Basic creation, Markdown → DOCX, JSON instructions, Variable detection, Mail merge | `output/docx/` |
| [PPTX](Pptx/README.md) | Slides, Tables, Variable merge, **→ PDF**, **→ Typst source**, **→ PNG thumbnails** | `output/pptx/` |
| [XLSX](Xlsx/README.md) | Worksheets, Formulas, Cross-sheet references, Variable detection, Mail merge, JSON instructions, Layout & merges, Read back | `output/xlsx/` |

## Project Structure

```
examples/
├── Program.cs              # Runs all examples
├── Docx/
│   ├── README.md           # DOCX input/output docs
│   ├── sample.md           # Markdown source for conversion example
│   └── instructions/
│       └── sample.json     # JSON instruction set example
├── Pptx/
│   ├── README.md           # PPTX input/output docs
│   └── instructions/
│       └── sample.json     # JSON instruction set example
├── Xlsx/
│   ├── README.md           # XLSX input/output docs
│   └── instructions/
│       └── sample.json     # JSON instruction set example
├── Shared/
│   └── data.json           # Sample merge data
└── output/                 # Generated files (created on run)
    ├── docx/
    ├── pptx/
    └── xlsx/
```

## Shared Data

`Shared/data.json` — sample data used by mail-merge examples:

```json
{
  "companyName": "Acme Corporation",
  "address": "123 Business Avenue, Suite 100",
  "city": "New York",
  "state": "NY",
  "zipCode": "10001",
  "contactName": "John Smith",
  "contactEmail": "john.smith@acme.com",
  "contactPhone": "+1 (555) 123-4567",
  "invoiceNumber": "INV-2024-001",
  "invoiceDate": "2024-01-15",
  "dueDate": "2024-02-15",
  "items": [
    { "description": "Consulting Services", "quantity": 10, "unitPrice": 150.00, "total": 1500.00 },
    { "description": "Software License", "quantity": 1, "unitPrice": 500.00, "total": 500.00 },
    { "description": "Training Session", "quantity": 5, "unitPrice": 100.00, "total": 500.00 }
  ],
  "subtotal": 2500.00,
  "taxRate": 0.08,
  "taxAmount": 200.00,
  "total": 2700.00,
  "notes": "Payment due within 30 days. Thank you for your business!"
}
```
