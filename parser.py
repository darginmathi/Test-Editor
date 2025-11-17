import re
import os
from pprint import pprint

def parse_webdriver_functions(file_path):
    """
    Parses a WebdriverFunctions.java file to extract public methods and their arguments.
    Returns a dictionary mapping function_name to a list of argument names.
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Regex to find public methods and capture their name and parameters
    # This handles single-line and multi-line function signatures
    function_regex = re.compile(
        r'public\s+(?:void|String|String\[\]|boolean|Double\[\]|Double\[\]\[\]|double\[\])\s+'  # Return type
        r'([a-zA-Z0-9_]+)\s*'  # Function name
        r'\((.*?)\)',  # Parameters
        re.DOTALL  # Allow '.' to match newlines for multi-line parameters
    )
    
    # Regex to extract individual parameters (type and name)
    param_regex = re.compile(r'([a-zA-Z0-9_\\\[\]]+)\s+([a-zA-Z0-9_]+)')
    
    functions = {}
    matches = function_regex.finditer(content)
    
    for match in matches:
        func_name = match.group(1)
        params_str = match.group(2).replace('\n', ' ').strip()
        
        if not params_str:
            functions[func_name] = []
            continue
            
        args = []
        param_matches = param_regex.finditer(params_str)
        for param_match in param_matches:
            args.append(param_match.group(2))
        functions[func_name] = args
        
    return functions

def parse_tc_helper(file_path, func_map):
    """
    Parses a TCHelper.java file to map command strings to their arguments.
    Uses the func_map from parse_webdriver_functions to find the real arg names.
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Regex to find the command string and the block of code associated with it
    command_regex = re.compile(
        r'if\s*\(command\.equalsIgnoreCase\("(.*?)"\)\)\s*{(.*?)}',
        re.DOTALL
    )
    
    # Regex to find the function call within the block
    call_regex = re.compile(r'\w+\.([a-zA-Z0-9_]+)\s*\(')
    
    command_args = {}
    matches = command_regex.finditer(content)
    
    for match in matches:
        command_name = match.group(1)
        block_content = match.group(2)
        
        call_match = call_regex.search(block_content)
        if call_match:
            func_name = call_match.group(1)
            if func_name in func_map:
                command_args[command_name] = func_map[func_name]
            else:
                # Fallback for functions not found (e.g. new Login().enterLogin)
                # Count the number of data[*] usages as a fallback
                data_usages = re.findall(r'data\[(\d+)\]', block_content)
                if data_usages:
                    num_args = max([int(i) for i in data_usages]) + 1
                    command_args[command_name] = [f"arg{i+1}" for i in range(num_args)]
                else:
                    command_args[command_name] = []
        else:
            command_args[command_name] = []
            
    return command_args

def main():
    # Define paths relative to the script's location
    base_path = os.path.join('..', 'synoptionE2E', 'src', 'main', 'java', 'com')
    
    fw_funcs_path = os.path.join(base_path, 'webdriverSpecific', 'FWWebdriverFunctions.java')
    e_funcs_path = os.path.join(base_path, 'projectSpecific', 'EWebdriverFunctions.java')
    
    fw_helper_path = os.path.join(base_path, 'webdriverSpecific', 'FWC_TCHelper.java')
    ec_helper_path = os.path.join(base_path, 'projectSpecific', 'EC_TCHelper.java')

    # Check if files exist
    for path in [fw_funcs_path, e_funcs_path, fw_helper_path, ec_helper_path]:
        if not os.path.exists(path):
            print(f"Error: File not found at {os.path.abspath(path)}")
            return

    # Parse all function definitions
    all_functions = {}
    all_functions.update(parse_webdriver_functions(fw_funcs_path))
    all_functions.update(parse_webdriver_functions(e_funcs_path))
    
    # Parse helpers to get the final command map
    final_command_args = {}
    final_command_args.update(parse_tc_helper(fw_helper_path, all_functions))
    final_command_args.update(parse_tc_helper(ec_helper_path, all_functions))

    # Add the special commands that don't have args and are not in helpers
    special_commands = {
        "StartAppWithLogin": [],
        "StartScenario": [],
        "EndScenario": [],
        "StopApp": [],
    }
    final_command_args.update(special_commands)

    # Sort the final dictionary by command name
    sorted_command_args = dict(sorted(final_command_args.items()))

    # Write to a temporary file
    temp_file_path = os.path.join(os.path.dirname(__file__), 'temp_command_args.txt')
    with open(temp_file_path, 'w', encoding='utf-8') as f:
        f.write("# Auto-generated command arguments\n\n")
        f.write("COMMAND_ARGS = {\n")
        for command, args in sorted_command_args.items():
            f.write(f'    "{command}": {args},\n')
        f.write("}\n")

    print(f"Successfully parsed commands and wrote to {os.path.abspath(temp_file_path)}")

if __name__ == "__main__":
    main()
