"""
Example script showing how to use the MPR format reference dictionary
to parse and analyze MPR files.
"""

import json
import re
from pathlib import Path
from typing import Dict, List

def load_mpr_reference(reference_file: str = "mpr_format_reference.json") -> Dict:
    """Load the MPR format reference dictionary."""
    with open(reference_file, 'r', encoding='utf-8') as f:
        return json.load(f)


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
    
    # Parse header
    if '[H' in mpr_content:
        header_section = mpr_content.split('[H')[1].split('[')[0] if '[' in mpr_content.split('[H')[1] else mpr_content.split('[H')[1]
        for line in header_section.split('\n'):
            if '=' in line and not line.strip().startswith('_'):
                parts = line.split('=', 1)
                if len(parts) == 2:
                    key, value = parts
                    parsed['header'][key.strip()] = value.strip('"')
    
    # Parse commands using the reference dictionary
    command_pattern = parser_dict['parsing_rules']['command_pattern']
    # Adjust pattern for actual MPR files (single backslash)
    command_pattern = r'<(\d+)\s+\\([A-Za-z_]+)\\'
    
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
        
        # Get reference information
        cmd_ref = parser_dict['command_reference'].get(cmd_num, {})
        
        parsed['commands'].append({
            'number': cmd_num,
            'name': cmd_name,
            'reference': {
                'description': cmd_ref.get('description', 'No description'),
                'parameters': cmd_ref.get('parameters', []),
                'documentation_pages': cmd_ref.get('documentation_pages', [])
            },
            'parameters': params
        })
    
    # Parse edges
    edge_pattern = parser_dict['parsing_rules']['edge_pattern']
    for match in re.finditer(edge_pattern, mpr_content):
        edge_num = match.group(1)
        edge_ref = parser_dict['edge_reference'].get(edge_num, {})
        parsed['edges'].append({
            'number': edge_num,
            'reference': {
                'documentation_pages': edge_ref.get('documentation_pages', [])
            }
        })
    
    return parsed


def print_mpr_summary(parsed: Dict):
    """Print a summary of the parsed MPR file."""
    print(f"\n{'='*60}")
    print(f"MPR File: {Path(parsed['file']).name}")
    print(f"{'='*60}")
    
    print(f"\nHeader Information:")
    for key, value in parsed['header'].items():
        print(f"  {key}: {value}")
    
    print(f"\nEdges: {len(parsed['edges'])}")
    for edge in parsed['edges']:
        print(f"  $E{edge['number']}")
    
    print(f"\nCommands: {len(parsed['commands'])}")
    for cmd in parsed['commands']:
        ref = cmd['reference']
        print(f"  <{cmd['number']} \\{cmd['name']}\\")
        if ref.get('description'):
            print(f"    Description: {ref['description']}")
        if cmd['parameters']:
            print(f"    Parameters: {', '.join(cmd['parameters'].keys())}")


def list_all_commands(parser_dict: Dict):
    """List all available commands in the reference dictionary."""
    print(f"\n{'='*60}")
    print("All Available MPR Commands")
    print(f"{'='*60}")
    
    commands = sorted(parser_dict['command_reference'].items(), 
                     key=lambda x: int(x[0]) if x[0].isdigit() else 9999)
    
    print(f"\nTotal Commands: {len(commands)}\n")
    
    for cmd_num, cmd_info in commands:
        print(f"<{cmd_num} \\{cmd_info['name']}\\")
        if cmd_info.get('description'):
            print(f"  Description: {cmd_info['description']}")
        if cmd_info.get('parameters'):
            print(f"  Parameters: {', '.join(cmd_info['parameters'])}")
        if cmd_info.get('found_in_files'):
            print(f"  Found in: {len(cmd_info['found_in_files'])} file(s)")
        print()


def find_command_by_name(parser_dict: Dict, name: str):
    """Find commands by name (partial match)."""
    print(f"\nSearching for commands containing: '{name}'")
    print(f"{'='*60}")
    
    found = []
    for cmd_num, cmd_info in parser_dict['command_reference'].items():
        if name.lower() in cmd_info['name'].lower():
            found.append((cmd_num, cmd_info))
    
    if found:
        for cmd_num, cmd_info in found:
            print(f"<{cmd_num} \\{cmd_info['name']}\\")
            if cmd_info.get('description'):
                print(f"  Description: {cmd_info['description']}")
    else:
        print("No commands found.")


# Example usage
if __name__ == "__main__":
    # Load the reference dictionary
    print("Loading MPR format reference...")
    parser_dict = load_mpr_reference()
    
    # List all commands
    list_all_commands(parser_dict)
    
    # Search for specific commands
    find_command_by_name(parser_dict, "Bohr")
    
    # Parse an example MPR file
    example_file = "Test_2_3/Top_1.mpr"
    if Path(example_file).exists():
        print(f"\n{'='*60}")
        print("Parsing Example MPR File")
        print(f"{'='*60}")
        parsed = parse_mpr_file(example_file, parser_dict)
        print_mpr_summary(parsed)
    else:
        print(f"\nExample file not found: {example_file}")

