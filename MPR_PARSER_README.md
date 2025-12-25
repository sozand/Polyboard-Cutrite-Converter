# MPR Format Reference Parser

This directory contains tools for extracting and using MPR (WoodWOP) format command references from the PDF documentation.

## Files Created

1. **`mpr_parser_generator.py`** - Main script that extracts MPR format commands from the PDF and creates a reference dictionary
2. **`mpr_format_reference.json`** - Generated reference dictionary containing all found MPR commands, edge commands, and geometry commands
3. **`mpr_parser_example.py`** - Example script showing how to use the reference dictionary to parse MPR files

## Usage

### Generating the Reference Dictionary

Run the generator script to extract commands from the PDF:

```bash
python mpr_parser_generator.py
```

This will:
- Extract text from `woodwop-mpr4x-format-pdf-free.pdf`
- Scan actual MPR files in `Test_2_3/` directory
- Create `mpr_format_reference.json` with all found commands

### Using the Reference Dictionary

```python
import json

# Load the reference dictionary
with open('mpr_format_reference.json', 'r', encoding='utf-8') as f:
    parser_dict = json.load(f)

# Access command information
command_100 = parser_dict['command_reference']['100']
print(f"Command 100: {command_100['name']}")
print(f"Parameters: {command_100['parameters']}")
print(f"Description: {command_100.get('description', 'No description')}")
```

### Example: Parsing an MPR File

See `mpr_parser_example.py` for a complete example of how to:
- Load the reference dictionary
- Parse MPR files
- Extract commands and parameters
- List all available commands
- Search for commands by name

Run the example:
```bash
python mpr_parser_example.py
```

## Reference Dictionary Structure

The `mpr_format_reference.json` file contains:

```json
{
  "version": "1.0",
  "source_pdf": "woodwop-mpr4x-format-pdf-free.pdf",
  "command_reference": {
    "100": {
      "number": 100,
      "name": "WerkStck",
      "pattern": "<100 \\\\WerkStck\\\\",
      "parameters": ["LA", "BR", "DI", ...],
      "example_parameters": {...},
      "documentation_pages": [...],
      "found_in_files": [...]
    },
    ...
  },
  "edge_reference": {
    "0": {
      "number": 0,
      "pattern": "$E0",
      "documentation_pages": [...]
    },
    ...
  },
  "geometry_reference": {
    "KP": {
      "command": "KP",
      "documentation_pages": [...]
    },
    ...
  },
  "parsing_rules": {
    "command_pattern": "<(\\d+)\\s+\\\\\\\\([A-Za-z_]+)\\\\\\\\",
    "edge_pattern": "\\$E(\\d+)",
    "parameter_pattern": "([A-Z]+)=\"([^\"]+)\"",
    ...
  }
}
```

## Found Commands

The reference dictionary currently contains:
- **64 unique command numbers** extracted from the PDF
- **5 edge commands** ($E0, $E1, $E2, $E3, $E6)
- **9 geometry commands** (KP, KL, KB, KR, KF, KS, KX, KY, KZ)
- **5 commands** found in actual MPR files (100, 103, 139, etc.)

## Common Commands

- **<100 \WerkStck\** - Workpiece definition
- **<103 \BohrHoriz\** - Horizontal drilling
- **<139 \Komponente\** - Component reference

## Requirements

- Python 3.x
- pdfplumber (for PDF text extraction)
  ```bash
  pip install pdfplumber
  ```

## Notes

- The PDF extraction may not find all commands if they're in complex table formats or images
- Commands found in actual MPR files are merged with PDF-extracted commands
- The reference dictionary can be updated by re-running the generator script

