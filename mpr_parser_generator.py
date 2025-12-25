import pdfplumber
import re
from pathlib import Path
from typing import Dict, List, Set
import json

def extract_mpr_commands_from_pdf(pdf_path: str) -> Dict:
    """
    Extract all MPR format commands/processes from the PDF documentation.
    
    Returns a dictionary with all found commands and their details.
    """
    mpr_reference = {
        'commands': {},  # Command number -> details
        'command_names': {},  # Command name -> details
        'edge_commands': {},  # Edge commands ($E0, $E1, etc.)
        'geometry_commands': {},  # Geometry commands (KP, KL, etc.)
        'all_patterns': []  # All found patterns
    }
    
    # Patterns to search for in the PDF
    patterns = {
        'command_block': r'<(\d+)\s+\\\\([A-Za-z_]+)\\\\',  # <100 \WerkStck\
        'edge_command': r'\$E(\d+)',  # $E0, $E1, etc.
        'geometry_command': r'^(KP|KL|KB|KR|KF|KS|KW|KX|KY|KZ)',  # Geometry commands
        'header_section': r'\[H',  # Header section
        'dimension_section': r'\[001',  # Dimension section
    }
    
    print(f"Extracting text from PDF: {pdf_path}")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Total pages: {total_pages}")
            
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                
                if not text:
                    continue
                
                # Extract command blocks: <### \CommandName\
                command_matches = re.finditer(patterns['command_block'], text, re.MULTILINE)
                for match in command_matches:
                    cmd_num = match.group(1)
                    cmd_name = match.group(2)
                    
                    # Get context around the command (next 20 lines)
                    start_pos = match.end()
                    context = text[start_pos:start_pos+500]
                    
                    # Extract parameters from context
                    params = {}
                    param_pattern = r'([A-Z]+)="([^"]+)"'
                    param_matches = re.findall(param_pattern, context)
                    for param_name, param_value in param_matches:
                        if param_name not in params:
                            params[param_name] = []
                        params[param_name].append(param_value)
                    
                    # Store command information
                    if cmd_num not in mpr_reference['commands']:
                        mpr_reference['commands'][cmd_num] = {
                            'number': cmd_num,
                            'name': cmd_name,
                            'full_name': f"<{cmd_num} \\{cmd_name}\\",
                            'parameters': params,
                            'description': '',
                            'pages': []
                        }
                    
                    if page_num not in mpr_reference['commands'][cmd_num]['pages']:
                        mpr_reference['commands'][cmd_num]['pages'].append(page_num)
                    
                    # Also index by name
                    if cmd_name not in mpr_reference['command_names']:
                        mpr_reference['command_names'][cmd_name] = []
                    if cmd_num not in mpr_reference['command_names'][cmd_name]:
                        mpr_reference['command_names'][cmd_name].append(cmd_num)
                
                # Also try to extract from tables
                try:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            if row:
                                row_text = ' '.join([str(cell) if cell else '' for cell in row])
                                # Look for command patterns in table cells
                                table_matches = re.finditer(patterns['command_block'], row_text)
                                for match in table_matches:
                                    cmd_num = match.group(1)
                                    cmd_name = match.group(2)
                                    
                                    if cmd_num not in mpr_reference['commands']:
                                        mpr_reference['commands'][cmd_num] = {
                                            'number': cmd_num,
                                            'name': cmd_name,
                                            'full_name': f"<{cmd_num} \\{cmd_name}\\",
                                            'parameters': {},
                                            'description': row_text[:200] if len(row_text) > 200 else row_text,
                                            'pages': []
                                        }
                                    
                                    if page_num not in mpr_reference['commands'][cmd_num]['pages']:
                                        mpr_reference['commands'][cmd_num]['pages'].append(page_num)
                except:
                    pass
                
                # Also look for command numbers in text (format: "100" or "Command 100" or "<100")
                cmd_num_pattern = r'(?:Command\s+|^|\s)(\d{3,4})(?:\s|$|:)'
                cmd_num_matches = re.finditer(cmd_num_pattern, text)
                for match in cmd_num_matches:
                    cmd_num = match.group(1)
                    # Get surrounding context
                    start = max(0, match.start() - 50)
                    end = min(len(text), match.end() + 200)
                    context = text[start:end]
                    
                    # Try to find command name in context
                    name_match = re.search(r'([A-Z][a-z]+[A-Z]?[a-z]*|\\[A-Za-z_]+\\\\)', context)
                    cmd_name = name_match.group(1).strip('\\') if name_match else f"Command{cmd_num}"
                    
                    if cmd_num not in mpr_reference['commands']:
                        mpr_reference['commands'][cmd_num] = {
                            'number': cmd_num,
                            'name': cmd_name,
                            'full_name': f"<{cmd_num} \\{cmd_name}\\",
                            'parameters': {},
                            'description': context[:300] if len(context) > 300 else context,
                            'pages': []
                        }
                    
                    if page_num not in mpr_reference['commands'][cmd_num]['pages']:
                        mpr_reference['commands'][cmd_num]['pages'].append(page_num)
                
                # Extract edge commands: $E0, $E1, etc.
                edge_matches = re.finditer(patterns['edge_command'], text)
                for match in edge_matches:
                    edge_num = match.group(1)
                    if edge_num not in mpr_reference['edge_commands']:
                        mpr_reference['edge_commands'][edge_num] = {
                            'number': edge_num,
                            'full_name': f"$E{edge_num}",
                            'pages': []
                        }
                    if page_num not in mpr_reference['edge_commands'][edge_num]['pages']:
                        mpr_reference['edge_commands'][edge_num]['pages'].append(page_num)
                
                # Extract geometry commands
                lines = text.split('\n')
                for line in lines:
                    geo_match = re.match(patterns['geometry_command'], line.strip())
                    if geo_match:
                        geo_cmd = geo_match.group(1)
                        if geo_cmd not in mpr_reference['geometry_commands']:
                            mpr_reference['geometry_commands'][geo_cmd] = {
                                'command': geo_cmd,
                                'pages': []
                            }
                        if page_num not in mpr_reference['geometry_commands'][geo_cmd]['pages']:
                            mpr_reference['geometry_commands'][geo_cmd]['pages'].append(page_num)
                
                # Progress indicator
                if page_num % 50 == 0:
                    print(f"Processed {page_num}/{total_pages} pages...")
    
    except Exception as e:
        print(f"Error processing PDF: {e}")
        import traceback
        traceback.print_exc()
        return mpr_reference
    
    print(f"\nExtraction complete!")
    print(f"Found {len(mpr_reference['commands'])} unique command numbers")
    print(f"Found {len(mpr_reference['command_names'])} unique command names")
    print(f"Found {len(mpr_reference['edge_commands'])} edge commands")
    print(f"Found {len(mpr_reference['geometry_commands'])} geometry commands")
    
    return mpr_reference


def scan_mpr_files_for_commands(mpr_directory: str = "Test_2_3") -> Dict:
    """
    Scan actual MPR files to find all commands being used.
    This complements the PDF extraction.
    """
    found_commands = {}
    mpr_dir = Path(mpr_directory)
    
    if not mpr_dir.exists():
        print(f"MPR directory not found: {mpr_directory}")
        return found_commands
    
    print(f"\nScanning MPR files in: {mpr_directory}")
    mpr_files = list(mpr_dir.glob("*.mpr"))
    print(f"Found {len(mpr_files)} MPR files")
    
    # MPR files use single backslash, not double
    command_pattern = r'<(\d+)\s+\\([A-Za-z_]+)\\'
    param_pattern = r'([A-Z]+)="([^"]+)"'
    
    for mpr_file in mpr_files:
        try:
            content = mpr_file.read_text(encoding='utf-8')
            
            for match in re.finditer(command_pattern, content):
                cmd_num = match.group(1)
                cmd_name = match.group(2)
                
                # Get command block
                start_pos = match.end()
                end_pos = content.find('\n<', start_pos)
                if end_pos == -1:
                    end_pos = content.find('\n!', start_pos)
                if end_pos == -1:
                    end_pos = len(content)
                
                cmd_block = content[start_pos:end_pos]
                
                # Extract parameters
                params = {}
                for param_match in re.finditer(param_pattern, cmd_block):
                    param_name = param_match.group(1)
                    param_value = param_match.group(2)
                    if param_name not in params:
                        params[param_name] = []
                    if param_value not in params[param_name]:
                        params[param_name].append(param_value)
                
                # Store command
                if cmd_num not in found_commands:
                    found_commands[cmd_num] = {
                        'number': cmd_num,
                        'name': cmd_name,
                        'full_name': f"<{cmd_num} \\{cmd_name}\\",
                        'parameters': params,
                        'example_parameters': {k: v[0] if v else '' for k, v in params.items()},
                        'found_in_files': []
                    }
                
                if mpr_file.name not in found_commands[cmd_num]['found_in_files']:
                    found_commands[cmd_num]['found_in_files'].append(mpr_file.name)
        
        except Exception as e:
            print(f"Error reading {mpr_file.name}: {e}")
    
    print(f"Found {len(found_commands)} unique commands in MPR files")
    return found_commands


def create_mpr_parser_dictionary(pdf_path: str, output_path: str = None) -> Dict:
    """
    Create a comprehensive MPR format parser dictionary from the PDF.
    
    This dictionary can be used as a reference for parsing MPR files.
    """
    # Extract commands from PDF
    mpr_data = extract_mpr_commands_from_pdf(pdf_path)
    
    # Also scan actual MPR files
    mpr_file_commands = scan_mpr_files_for_commands()
    
    # Merge MPR file commands with PDF commands
    for cmd_num, cmd_info in mpr_file_commands.items():
        if cmd_num in mpr_data['commands']:
            # Merge parameters
            pdf_params = mpr_data['commands'][cmd_num]['parameters']
            file_params = cmd_info['parameters']
            # Combine unique parameters
            for param_name, param_values in file_params.items():
                if param_name not in pdf_params:
                    pdf_params[param_name] = []
                for val in param_values:
                    if val not in pdf_params[param_name]:
                        pdf_params[param_name].append(val)
            mpr_data['commands'][cmd_num]['parameters'] = pdf_params
            mpr_data['commands'][cmd_num]['found_in_files'] = cmd_info['found_in_files']
        else:
            # Add new command from MPR files
            mpr_data['commands'][cmd_num] = {
                'number': cmd_num,
                'name': cmd_info['name'],
                'full_name': cmd_info['full_name'],
                'parameters': cmd_info['parameters'],
                'description': '',
                'pages': [],
                'found_in_files': cmd_info['found_in_files']
            }
    
    # Create a structured parser dictionary
    parser_dict = {
        'version': '1.0',
        'source_pdf': pdf_path,
        'command_reference': {},
        'edge_reference': {},
        'geometry_reference': {},
        'parsing_rules': {
            'command_pattern': r'<(\d+)\s+\\\\([A-Za-z_]+)\\\\',
            'edge_pattern': r'\$E(\d+)',
            'parameter_pattern': r'([A-Z]+)="([^"]+)"',
            'header_pattern': r'\[H',
            'dimension_pattern': r'\[001',
        },
        'known_commands': {}
    }
    
    # Populate command reference
    for cmd_num, cmd_info in mpr_data['commands'].items():
        parser_dict['command_reference'][cmd_num] = {
            'number': int(cmd_num),
            'name': cmd_info['name'],
            'pattern': f"<{cmd_num} \\\\{cmd_info['name']}\\\\",
            'parameters': list(cmd_info['parameters'].keys()),
            'example_parameters': {k: v[0] if v else '' for k, v in cmd_info['parameters'].items()},
            'documentation_pages': cmd_info.get('pages', []),
            'found_in_files': cmd_info.get('found_in_files', [])
        }
    
    # Populate edge reference
    for edge_num, edge_info in mpr_data['edge_commands'].items():
        parser_dict['edge_reference'][edge_num] = {
            'number': int(edge_num),
            'pattern': f"$E{edge_num}",
            'documentation_pages': edge_info['pages']
        }
    
    # Populate geometry reference
    for geo_cmd, geo_info in mpr_data['geometry_commands'].items():
        parser_dict['geometry_reference'][geo_cmd] = {
            'command': geo_cmd,
            'documentation_pages': geo_info['pages']
        }
    
    # Create known commands lookup (from actual MPR files)
    known_commands = {
        '100': {'name': 'WerkStck', 'description': 'Workpiece definition'},
        '103': {'name': 'BohrHoriz', 'description': 'Horizontal drilling'},
        '139': {'name': 'Komponente', 'description': 'Component reference'},
    }
    
    # Merge known commands with extracted ones
    for cmd_num, cmd_info in known_commands.items():
        if cmd_num in parser_dict['command_reference']:
            parser_dict['command_reference'][cmd_num]['description'] = cmd_info['description']
        else:
            parser_dict['command_reference'][cmd_num] = {
                'number': int(cmd_num),
                'name': cmd_info['name'],
                'pattern': f"<{cmd_num} \\\\{cmd_info['name']}\\\\",
                'parameters': [],
                'example_parameters': {},
                'description': cmd_info['description'],
                'documentation_pages': []
            }
    
    parser_dict['known_commands'] = known_commands
    
    # Save to file if output path provided
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(parser_dict, f, indent=2, ensure_ascii=False)
        print(f"\nParser dictionary saved to: {output_path}")
    
    return parser_dict


def parse_mpr_file(mpr_path: str, parser_dict: Dict) -> Dict:
    """
    Parse an MPR file using the reference dictionary.
    
    Returns a structured representation of the MPR file.
    """
    mpr_content = Path(mpr_path).read_text(encoding='utf-8')
    
    parsed = {
        'file': mpr_path,
        'header': {},
        'dimensions': {},
        'edges': [],
        'commands': [],
        'geometry': []
    }
    
    lines = mpr_content.split('\n')
    
    # Parse header
    if '[H' in mpr_content:
        header_section = mpr_content.split('[H')[1].split('[')[0] if '[' in mpr_content.split('[H')[1] else mpr_content.split('[H')[1]
        for line in header_section.split('\n'):
            if '=' in line:
                key, value = line.split('=', 1)
                parsed['header'][key.strip()] = value.strip('"')
    
    # Parse commands
    command_pattern = parser_dict['parsing_rules']['command_pattern']
    for match in re.finditer(command_pattern, mpr_content):
        cmd_num = match.group(1)
        cmd_name = match.group(2)
        
        # Get command block
        start_pos = match.end()
        end_pos = mpr_content.find('\n<', start_pos)
        if end_pos == -1:
            end_pos = mpr_content.find('\n!', start_pos)
        if end_pos == -1:
            end_pos = len(mpr_content)
        
        cmd_block = mpr_content[start_pos:end_pos]
        
        # Extract parameters
        params = {}
        param_pattern = parser_dict['parsing_rules']['parameter_pattern']
        for param_match in re.finditer(param_pattern, cmd_block):
            param_name = param_match.group(1)
            param_value = param_match.group(2)
            params[param_name] = param_value
        
        parsed['commands'].append({
            'number': cmd_num,
            'name': cmd_name,
            'reference': parser_dict['command_reference'].get(cmd_num, {}),
            'parameters': params
        })
    
    # Parse edges
    edge_pattern = parser_dict['parsing_rules']['edge_pattern']
    for match in re.finditer(edge_pattern, mpr_content):
        edge_num = match.group(1)
        parsed['edges'].append({
            'number': edge_num,
            'reference': parser_dict['edge_reference'].get(edge_num, {})
        })
    
    return parsed


# Main execution
if __name__ == "__main__":
    pdf_path = "woodwop-mpr4x-format-pdf-free.pdf"
    output_path = "mpr_format_reference.json"
    
    print("=" * 60)
    print("MPR Format Reference Parser Generator")
    print("=" * 60)
    
    # Create the parser dictionary
    parser_dict = create_mpr_parser_dictionary(pdf_path, output_path)
    
    # Print summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"Total Commands Found: {len(parser_dict['command_reference'])}")
    print(f"Total Edge Commands: {len(parser_dict['edge_reference'])}")
    print(f"Total Geometry Commands: {len(parser_dict['geometry_reference'])}")
    
    # Print first 10 commands as example
    print("\nFirst 10 Commands:")
    for i, (cmd_num, cmd_info) in enumerate(list(parser_dict['command_reference'].items())[:10]):
        print(f"  {cmd_num}: {cmd_info['name']} - {cmd_info.get('description', 'No description')}")
    
    print(f"\nFull reference saved to: {output_path}")
    print("\nYou can now use this dictionary to parse MPR files!")

