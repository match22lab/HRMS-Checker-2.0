import os
import re
import time
import logging
from typing import List
from molmass import Formula
import pandas as pd # also install openpyxl
import fitz  # install PyMuPDF

# Constants for file paths and reporting
source_folder = r"C:\Users\match\Downloads"  # The folder to be searched for PDFs with HRMS data
destination_folder = r"C:\Users\match\Desktop" # The folder for the report is saved if write_report = True
write_report = False # If True, a report is written as Excel file to destination_folder; if False, no report is written


def check_conditions(cleaned_results):
    for row in cleaned_results:
        # Check if the 8th column (index 7) is empty or contains "-0.0001" or "+0.0001"
        if row[7] not in ("", "-0.0001", "+0.0001"):
            return False
        # Check if the 7th column (index 6) as a float is less than 10
        try:
            if float(row[6]) >= 10:
                return False
        except ValueError:
            # If conversion to float fails, return False
            return False
    return True


def fix_floats(text):
    """
    Searches a string for floats in the form "xxxx.xxx" and changes them to "xxxx.xxx0".

    Args:
        text (str): The input text to search and modify.

    Returns:
        str: The modified text with floats in the form "xxxx.xxx0".
    """
    # Define a regular expression pattern to match floats with 3 decimal places
    pattern = r'\b\d+\.\d{3}\b'

    # Use the re.sub() function to replace matches with the modified float
    modified_text = re.sub(pattern, lambda match: match.group() + '0', text)

    return modified_text


def remove_sublists_with_missing_element1_positions_swapped(cleaned_results):
    """
    Removes sublists where element 1 is missing (''), if there exists another sublist
    where elements at positions 2, 3, and 4 are the same (positions 3 and 4 may be swapped)
    and element 1 is present.
    """
    # Create a set to hold indices of sublists to remove
    indices_to_remove = set()
    # Build a dictionary to map keys (elements 2, and positions 3 & 4 as a frozenset) to indices
    element_presence = {}

    # First pass: Collect sublists where element 1 is present
    for idx, sublist in enumerate(cleaned_results):
        if len(sublist) < 4:
            continue  # Skip if sublist doesn't have enough elements
        # Create a frozenset of positions 3 and 4 to handle swapping
        positions_3_4_set = frozenset([sublist[2], sublist[3]])
        key = (sublist[1], positions_3_4_set)  # Element at position 2 and set of positions 3 and 4
        if sublist[0] != '':
            # Element 1 is present, store the index
            if key not in element_presence:
                element_presence[key] = []
            element_presence[key].append(idx)

    # Second pass: Identify sublists to remove
    for idx, sublist in enumerate(cleaned_results):
        if len(sublist) < 4:
            continue  # Skip if sublist doesn't have enough elements
        if sublist[0] == '':
            # Element 1 is missing
            positions_3_4_set = frozenset([sublist[2], sublist[3]])
            key = (sublist[1], positions_3_4_set)
            if key in element_presence:
                # There is at least one sublist where elements 2, 3, 4 (with positions 3 and 4 swapped) are the same and element 1 is present
                indices_to_remove.add(idx)

    # Remove sublists at the collected indices
    cleaned_results = [sublist for idx, sublist in enumerate(cleaned_results) if idx not in indices_to_remove]
    return cleaned_results



def remove_spaces_in_formula(text):
    """
    Removes all spaces within chemical formulas in the input text.

    The function identifies chemical formulas based on sequences of element symbols
    (one or two letters, starting with an uppercase letter), possibly separated by numbers
    and spaces, and removes any spaces within those sequences.

    Args:
      text: The input string containing chemical formulas.

    Returns:
      The processed string with spaces removed from within chemical formulas.
    """

    # Step 1: Protect floats by surrounding them with '#'
    text = re.sub(r'(\d+\.\d+)', r'#\1#', text)

    # Regular expression pattern to match chemical formulas
    element = r'[A-Z][a-z]?'
    number = r'\d+'
    # Pattern matches sequences starting with an element symbol, followed by
    # elements or numbers, possibly with spaces in between
    pattern = r'(' + element + r'(?:\s*(?:' + element + r'|' + number + r'))+)'

    # Function to remove spaces within the matched chemical formula
    def remove_spaces(match):
        return match.group(0).replace(' ', '')

    # Replace matches in the text with spaces removed within chemical formulas
    return re.sub(pattern, remove_spaces, text)


def remove_page_numbers(text):
    """
    Remove lines that appear to be page numbers from a text string.

    Matches:
    - Single integers (e.g., "12")
    - Integers with dashes (e.g., "- 12 -", "-13-")
    - Integers with p/P/s/S prefix (e.g., "P12", "s23")
    - Integers with p/P/s/S prefix and dashes (e.g., "S-12", "p -13")

    Args:
        text (str): Input text containing page numbers

    Returns:
        str: Text with page number lines removed
    """
    # Split text into lines
    lines = text.split('\n')

    # Regular expression patterns for page numbers
    patterns = [
        r'^\s*\d+\s*$',                  # Single integers: "12"
        r'^\s*-\s*\d+\s*-\s*$',          # Dashed integers: "- 12 -"
        r'^\s*-\d+-\s*$',                # Compact dashed integers: "-13-"
        r'^\s*[psPS]\s*-?\s*\d+\s*(?:\n|$)',    # p/P/s/S prefixed: "P12", "s23", "S-12"
    ]

    # Combine patterns
    combined_pattern = '|'.join(f'({pattern})' for pattern in patterns)

    # Filter out lines matching the patterns
    cleaned_lines = [line for line in lines if not re.match(combined_pattern, line)]

    # Rejoin the remaining lines
    return '\n'.join(cleaned_lines)


def is_float(value):
    try:
        float(value)
        return True
    except ValueError:
        return False



def protect_floats(text: str) -> str:

    # Match floats with 3+ digits before decimal and 4+ after
    pattern = r'(\d{3,}\.\d{4,})'

    def add_spaces(match: re.Match) -> str:
        """Add spaces around the matched float if needed."""
        float_num = match.group(1)
        start, end = match.span(1)

        # Get characters before and after the float
        char_before = text[start - 1] if start > 0 else ''
        char_after = text[end] if end < len(text) else ''

        # Only add space if the adjacent characters aren't already spaces
        prefix = '' if char_before.isspace() else ' '
        suffix = '' if char_after.isspace() else ' '

        return f'{prefix}{float_num}{suffix}'

    return re.sub(pattern, add_spaces, text)


from typing import Match


def replace_comma_with_decimal(text: str) -> str:
    # Match numbers with comma decimals that:
    # \b     - Start at a word boundary
    # \d+    - Have one or more digits before the comma
    # ,      - Have a comma
    # \d+    - Have one or more digits after the comma
    # \b     - End at a word boundary
    pattern = r'\b(\d+,\d+)\b'

    def comma_to_decimal(match: Match[str]) -> str:
        """Convert comma to decimal point in matched number."""
        return match.group(0).replace(',', '.')

    return re.sub(pattern, comma_to_decimal, text)


def adjust_space_around_decimal(text):

    if not isinstance(text, str):
        raise TypeError("Input must be a string")

    # Step 1: Remove unwanted spaces around decimal points
    # Handles cases like "23. 4562" → "23.4562"
    text = re.sub(r'(\d+)\s*\.\s*(\d+)', r'\1.\2', text)

    # Step 2: Add space between decimal numbers and following text
    # Handles cases like "2.4beta" → "2.4 beta"
    text = re.sub(r'(\d+\.\d+)([A-Za-z])', r'\1 \2', text)

    # Step 3: Handle special cases where no space is needed
    # For file extensions like ".txt", ".pdf"
    text = re.sub(r'(\s\d+)\s+(\.[A-Za-z]+\b)', r'\1\2', text)

    return text


def decrease_element_count(molecular_formula: str, element_to_decrease: str) -> str:
    """
    Decreases the count of a specific element in a molecular formula by 1.

    Args:
        molecular_formula: The input molecular formula (e.g., 'C6H12O2')
        element_to_decrease: The element whose count should be decreased (e.g., 'C')

    Returns:
        Modified molecular formula with decreased element count

    Example:
        >>> decrease_element_count('C6H12O2', 'C')
        'C5H12O2'
    """
    pattern = fr'({element_to_decrease})(?![a-z])\d*'

    def replace_element(match: re.Match) -> str:
        element_count = match.group()
        element = re.match(r'([A-Z][a-z]*)', element_count).group()

        if count_match := re.search(r'\d+', element_count):
            current_count = int(count_match.group())
            return (f"{element}{current_count - 1}" if current_count > 2
                    else element)  # Remove count when it's 2
        return element

    return re.sub(pattern, replace_element, molecular_formula)


def have_swapped_adjacent_digits(float1: float, float2: float) -> bool:
    # Convert floats to strings
    str1, str2 = str(float1), str(float2)

    # Remove last two digits for comparison
    str1 = str1[:-2]
    str2 = str2[:-2]

    # Remove decimal points for comparison
    str1_no_dot = str1.replace('.', '')
    str2_no_dot = str2.replace('.', '')

    # Check lengths
    if len(str1_no_dot) != len(str2_no_dot) or len(str1_no_dot) < 2:
        return False

    # Find positions that differ
    diff_positions = [i for i in range(len(str1_no_dot))
                      if str1_no_dot[i] != str2_no_dot[i]]

    # Must have exactly 2 differences for a single swap
    if len(diff_positions) != 2:
        return False

    # The positions must be adjacent
    if diff_positions[1] - diff_positions[0] != 1:
        return False

    # Check if it's actually a swap
    pos1, pos2 = diff_positions
    return (str1_no_dot[pos1] == str2_no_dot[pos2] and
            str1_no_dot[pos2] == str2_no_dot[pos1])



def differ_in_single_digit_except_last_two(float1: float, float2: float) -> bool:
    """
    Checks if two floating-point numbers differ by exactly one digit, excluding the last two digits.
    Handles trailing zeros and decimal points in the comparison.

    Args:
        float1: First floating-point number
        float2: Second floating-point number

    Returns:
        True if numbers differ by exactly one digit (excluding last two), False otherwise

    Examples:
        >>> differ_in_single_digit_except_last_two(123.45, 153.45)
        True
        >>> differ_in_single_digit_except_last_two(123.45, 153.46)
        False
        >>> differ_in_single_digit_except_last_two(123.450, 153.45)
        True
    """
    # Convert to strings and normalize by removing trailing zeros and decimal points
    str1 = str(float1).rstrip('0').rstrip('.')
    str2 = str(float2).rstrip('0').rstrip('.')

    # Quick validation checks
    if len(str1) != len(str2) or len(str1) < 3:  # Need at least 3 digits for comparison
        return False

    # Extract main part and last two digits
    main1, last_two1 = str1[:-2], str1[-2:]
    main2, last_two2 = str2[:-2], str2[-2:]

    # Last two digits must match
    if last_two1 != last_two2:
        return False

    # Count differing digits in main part
    return sum(1 for a, b in zip(main1, main2) if a != b) == 1


def calculate_molecular_weight(formula):
    # Dictionary of atomic weights for elements up to Plutonium (94)
    # Values are in atomic mass units (amu) or g/mol
    atomic_weights = {
        "H": 1.008, "D": 2.0141, "He": 4.002602, "Li": 6.94, "Be": 9.0121831, "B": 10.81, "C": 12.011,
        "N": 14.007, "O": 15.999, "F": 18.9984, "Ne": 20.1797, "Na": 22.98977, "Mg": 24.305, "Al": 26.98154,
        "Si": 28.085, "P": 30.97376, "S": 32.06, "Cl": 35.45, "Ar": 39.948, "K": 39.0983, "Ca": 40.078,
        "Sc": 44.955908, "Ti": 47.867, "V": 50.9415, "Cr": 51.9961, "Mn": 54.938044, "Fe": 55.845,
        "Co": 58.933194, "Ni": 58.6934, "Cu": 63.546, "Zn": 65.38, "Ga": 69.723, "Ge": 72.630,
        "As": 74.921595, "Se": 78.971, "Br": 79.904, "Kr": 83.798, "Rb": 85.4678, "Sr": 87.62,
        "Y": 88.90584, "Zr": 91.224, "Nb": 92.90637, "Mo": 95.95, "Tc": 98, "Ru": 101.07,
        "Rh": 102.90550, "Pd": 106.42, "Ag": 107.8682, "Cd": 112.414, "In": 114.818, "Sn": 118.710,
        "Sb": 121.760, "Te": 127.60, "I": 126.90447, "Xe": 131.293, "Cs": 132.90545196, "Ba": 137.327,
        "La": 138.90547, "Ce": 140.116, "Pr": 140.90766, "Nd": 144.242, "Pm": 145, "Sm": 150.36,
        "Eu": 151.964, "Gd": 157.25, "Tb": 158.92535, "Dy": 162.500, "Ho": 164.93033,
        "Er": 167.259, "Tm": 168.93422, "Yb": 173.04, "Lu": 174.9668, "Hf": 178.49,
        "Ta": 180.94788, "W": 183.84, "Re": 186.207, "Os": 190.23, "Ir": 192.217,
        "Pt": 195.084, "Au": 196.96657, "Hg": 200.592, "Tl": 204.38, "Pb": 207.2,
        "Bi": 208.9804, "Po": 209, "At": 210, "Rn": 222, "Fr": 223, "Ra": 226,
        "Ac": 227, "Th": 232.0377, "Pa": 231.03588, "U": 238.02891, "Np": 237, "Pu": 244
    }

    # Parse the molecular formula using regex
    formula_components = re.findall(r"([A-Z][a-z]?)(\d*)", formula)

    # Calculate total molecular weight
    mol_weight = 0.0
    for element, count in formula_components:
        # Get atomic weight from dictionary, default to 0.0 if element not found
        element_weight = atomic_weights.get(element, 0.0)
        # If no count specified, assume 1, otherwise convert string to integer
        mol_weight += element_weight * (int(count) if count else 1)

    return mol_weight

def remove_spaces_within_brackets(s, max_chars=20):
    """
    Removes all spaces within brackets () or [] if the number of non-space characters inside
    is within max_chars. Handles nested brackets appropriately without affecting spaces outside
    the brackets.

    Args:
    - s (str): The input string.
    - max_chars (int): Maximum number of non-space characters between opening and closing brackets.

    Returns:
    - str: The modified string with spaces removed within qualifying brackets.
    """
    stack = []
    # Mapping of opening brackets to their corresponding closing brackets
    opening_to_closing = {'(': ')', '[': ']'}
    # Mapping of closing brackets to their corresponding opening brackets
    closing_to_opening = {')': '(', ']': '['}

    s_list = list(s)  # Convert string to list for mutable operations
    remove_space_ranges = []  # List to hold ranges where spaces need to be removed

    for i, char in enumerate(s_list):
        if char in opening_to_closing:
            # Push opening bracket and its position onto the stack
            stack.append((char, i))
        elif char in closing_to_opening:
            if stack and stack[-1][0] == closing_to_opening[char]:
                # Pop the last opening bracket from the stack
                open_char, open_pos = stack.pop()
                close_pos = i
                # Extract the substring inside the brackets
                content = ''.join(s_list[open_pos + 1:close_pos])
                # Count the number of non-space characters
                non_space_chars = len(content.replace(' ', ''))
                if non_space_chars <= max_chars:
                    # Define the range for space removal (exclusive of brackets)
                    remove_space_ranges.append((open_pos + 1, close_pos))
            else:
                # Unmatched closing bracket; ignore or handle as needed
                pass

    # Sort ranges in descending order of start index to handle inner brackets first
    remove_space_ranges.sort(key=lambda x: x[0], reverse=True)

    for start, end in remove_space_ranges:
        # Extract the substring within the current bracket (excluding brackets)
        substring = ''.join(s_list[start:end])
        # Remove all spaces within this substring
        substring_no_spaces = substring.replace(' ', '')
        # Replace the original substring with the modified one
        s_list[start:end] = list(substring_no_spaces)

    # Join the list back into a string and return
    return ''.join(s_list)


def isotope_correct(text):
    """
    Applies a series of substitutions to a text to correct for isotope labeling and other specific replacements.

    Parameters:
    text (str): The input text to be processed.

    Returns:
    str: The processed text with all substitutions applied.
    """
    # Dictionary of replacements for isotope corrections and other text cleanup
    replacements = {
        "[MALDI]":"","[MALDI-TOF]":"","detected":" ","page": " ", "of": " ",  "𝑀": " ", "EI": " ", " . ": " ", ":": " ", "Δ": " ",
        "𝛼": " ", " a ": " ", " M ": " ", " H ": " ", "ESI": " ", " Na ": " ", " K ": " ",
        " NH4 ": " ", "Obs.": " ", "obs": " ", "78.9183": "", "48Ti": "[48Ti]","54Fe":"[54Fe]",
        "46Ti": "[46Ti]", "47Ti": "[47Ti]", " 2H": "D", " [3H]": "[3H]",
        " 10B": "[10B]", "127I": "[127I]", "120Sn":"[120Sn]", "119Sn":"[119Sn]", "118Sn":"[118Sn]",
        "N23Na": "*N23*Na", "F23Na": "*F23*Na", "H23Na": "*H23*Na", "23Na":"[23Na]","H28Si": "*H28*Si", "H11B": "*H11*B",
        "H13Co": "*H13*Co", "H13Cl": "*H13*Cl", "H18O": "*H18*O", "H218O": "*H218*O", "N18O": "*N18*O",
        "H35Cl": "*H35*Cl", "H37Cl": "*H37*Cl", "H10B":"*H10*B", "H19F": "*H19*F", "H81Br":"*H81*Br","H79Br":"*H79*Br","Br79": "[79Br]",
        " 79Br": "[79Br]", " 81Br": "[81Br]", "18O": "[18O]", "74Ge": "[74Ge]", "65Cu":"[65Cu]",
        "63Cu":"[63Cu]", "Br81": "[81Br]", " 35Cl": "[35Cl]", " 37Cl": "[37Cl]", " 11B": "[11B]",
        " 32S": "S", " 31P": "P", "35Cl":"[35Cl]", "80Se":"[80Se]", "37Cl":"[37Cl]", "28Si":"[28Si]",
        "13C":"[13C]", "[13C]l":"13Cl", "96Ru":"[96Ru]","79Br":"[79Br]", "81Br":"[81Br]", "11B":"[11B]", "10B":"[10B]",
        "[10B]r":"10Br", "[[":"[", "]]":"]", "*H13*Cl": "H13Cl", "*H18*O": "H18O", "*H218*O": "H218O",
        "*N18*O": "N18O", "*H13*Co": "H13Co", "*H37*Cl": "H37Cl", "*H35*Cl": "H35Cl","*H81Br*":"H81Br","*H79Br*":"H79Br",
        "*H28*Si": "H28Si", "*H10*B":"H10B", "*H23*Na": "H23Na", "*F23*Na": "F23Na", "*N23*Na": "N23Na",
        "*H11*B":"H11B", "*H19*F": "H19F", "cacld": "", "calcd.": "calcd ", "calc’d": "calcd ",
        "calcd gcm": " ", " is ": " ", "calcd": "calcd ", "calcd  ": "calcd ","++": "+","(M":"[M", ")+":"]+ ",
        "MALDI":"","Maldi":""," [13C]":"[13C]","  [127I]":"[127I]"," [12C":"C"," [37Cl]":"37Cl"," [35Cl]":"35Cl",
        "C ":"C","H":"H", " N":"N"," O":"O"," Na":"Na", " Br":"Br", "N ":"N"," Cl":"Cl", " F":"F"," S":"S"," P":"P"," B":"B","M]+H+]":"M+H]+","M]-H+]":"M-H]-",
        "MH+":"M+H]+ ","]-(":"]- ","]+)":"]+ ","]-)":"]- ","]2-)":"]2- ","]+C":"]+ C","[MM":"","=":"","[MeOH":" ","[MeCN":" ","m/z":" ","]+2 ":"]2+ ","]+1":"]+","M+ C":"M+C","+]":"]+","+calc":" calc",
        "Na)]":"Na]","+Na)":"+Na]",";":" ","+H)]":"+H]","+K)]":"+K]","+NH4)]":"+NH4]","+H)":"+H]","H+)":"H]+","Na+)":"Na]+","-calcd":"- calcd","[M-H] ":"[M-H]-","--":"-",
        "NH4+)":"MH4]+","M+)":"M]+","M]+)":"M]+","+)":"+","M- ":"M-","+.":"+","[MNa]+":"[M+Na]+","[MH]+":"[M+H]+",
        " M2+ ":" [M]2+ "," M3+ ":" [M]3+ "," M4+ ":" [M]4+ "," M5+ ":" [M]5+ "," M6+ ":" [M]6+ ",
        " M2- ": " [M]2- ", " M3- ": " [M]3- ", " M4- ": " [M]4- ", " M5- ": " [M]5- ", " M6- ": " [M]6- ","[M+H] ":"[M+H]+ ","[M+Na] ":"[M+Na]+ ","[M] ":"[M]+ ","]calcd":"] calcd","-.":"- ","M+1)":"M+1]+ ","+ꞏ":"+","]-calcd":"]- calcd"

    }

    # Apply each replacement in the dictionary to the text
    for original, replacement in replacements.items():
        text = text.replace(original, replacement)

    return text


def transform_expressions_in_text(text):
    """
    Transforms all chemical expressions within a given text into a standardized format.

    Rules for expressions:
    - Starts with M or nM, where n is a single digit integer.
    - Ends with a charge (e.g., +, 2+, -).
    - Can be enclosed in () or [] brackets.
    - May contain spaces which are removed within the expression.
    - Charges can be inside or outside the brackets.

    The transformed expression:
    - Contains no spaces within the expression.
    - Preserves surrounding text intact.

    Args:
    - text (str): The input text containing chemical expressions.

    Returns:
    - str: The text with all expressions transformed accordingly.
    """

    # Step 1: Replace specific symbols with corresponding charges
    symbol_replacements = {

        '⊕': '+',
        '•+': '+',
        '': '+',
        '': "+",
        '+.':'+ ',
        '•': '',
        'ꞏ': '',
        '–': '-',
        '-':'-',
        '−.':'- ',
        '−': '-',  # Minus sign
        '—': '-',  # Em dash
        '―': '-',
        '˗': '-',
        '-.': '- ',
    }

    # Create a regex pattern to match all keys in symbol_replacements
    symbols_pattern = re.compile('|'.join(map(re.escape, symbol_replacements.keys())))
    text = symbols_pattern.sub(lambda match: symbol_replacements[match.group()], text)

    # Step 2: Define regex to find expressions
    # This pattern matches expressions enclosed in [] or () with optional charges outside
    expression_pattern = re.compile(
        r'[\[(]'  # Opening bracket [ or (
        r'(\d*M?\d*[a-zA-Z\d-]*)'  # Capture group (explained above)
        r'[])]'  # Closing bracket ] or )
        r'(\d*\+|-)?'  # Optional charge outside the brackets
        r'[,:]*'  # Optional trailing characters
    )

    def replace_expression(match):
        expression_part = match.group(1)  # The main part of the expression
        charge_outside = match.group(2)  # The charge outside the brackets, if any

        # Step 3: Remove all internal brackets within the main expression
        expression_part = re.sub(r'[\[\]()]', '', expression_part)

        # Step 4: Remove all spaces within the main expression
        expression_part = re.sub(r'\s+', '', expression_part)

        if not charge_outside:
            # Step 5: Extract charge from the main expression if charge_outside is not present
            charge_match = re.search(r'([+-])$', expression_part)
            if charge_match:
                charge = charge_match.group(1)
                expression_part = expression_part[:charge_match.start()]
            else:
                charge = ''
        else:
            charge = charge_outside

        # Step 6: Format the transformed expression
        transformed = f'[{expression_part}]{charge}'

        return transformed

    # Step 7: Substitute all matching expressions in the text
    transformed_text = expression_pattern.sub(replace_expression, text)

    return transformed_text

def transform_molecular_formula(formula):
  """
  Transforms a molecular formula string to a standardized format.

  Args:
    formula: The molecular formula string to transform.

  Returns:
    The transformed molecular formula string.
  """

  # Remove all round brackets and colons
  formula = formula.replace("(", "").replace(")", "").replace(":", "").replace("]+-", "]+")

  # Remove ALL spaces within brackets and move the + or - sign after the bracket (if any)
  formula = re.sub(r'\[(.*?)]', lambda m: '[' + m.group(1).replace(' ', '') + ']' + ('+' if '+' in m.group(1) else '') + ('-' if '-' in m.group(1) else ''), formula)

  # Replace "M-" with "M-"
  formula = re.sub(r'M\s*–', 'M-', formula)

  # Replace "M +" or "M+" with "M+"
  formula = re.sub(r'M\s*\+', 'M+', formula)

  # Ensure standardized ion is surrounded by one space, BUT NOT IF IT IS THE LAST THING
  formula = re.sub(r'([^ ])(\[\w+][+-]?)(?=\S)', r'\1 \2 ', formula)  # Include optional + or - in the ion group

  # Add brackets if "M" is present without brackets
  if "M" in formula and "[" not in formula:
    formula = "[" + formula + "]"

  # Add spaces around "calcd for", "found"
  formula = re.sub(r'(calcd\s*for|found)', r' \1 ', formula)

  # Remove double spaces
  formula = formula.replace("++", "+").replace("++", "+").replace(",", " ")
  formula = re.sub(r'\s+', ' ', formula)
  formula = formula.replace("-+", "+").replace("]+-", "]+").replace("+]+", "]+ ").replace("++", "+").replace("--", "").replace(",", "+")

  return formula


# Configure logging
logging.basicConfig(level=logging.INFO, format='%(message)s')


def generate_error_dictionary(element_list, counts_range, special_cases=None):
    """
    Generates an error dictionary mapping mass differences to element or group descriptions.
    For atoms, includes entries for counts from counts_range.
    For groups, includes entries only for count=1, with descriptions like "1 OH-group".

    Parameters:
    - element_list (list): List of element symbols or groups (e.g., ['H', 'O', 'N', 'OH']).
    - counts_range (range): Range of atom counts for atoms (e.g., range(1, 11) for counts 1-10).
    - special_cases (dict): Optional dictionary for special error cases
                            (e.g., {'0.0005': 'Electron mass error'}).

    Returns:
    - dict: Error dictionary with mass differences as keys and descriptions as values.
    """

    error_dict = {}
    electron_mass = 0.0005486  # Atomic mass units (amu)

    for element in element_list:
        try:
            atomic_mass = Formula(element).monoisotopic_mass
        except Exception as e:
            print(f"Error processing element {element}: {e}")
            continue  # Skip this element if there's an error

        # Determine if the element is a group (more than one capital letter)
        is_group = sum(1 for c in element if c.isupper()) > 1

        if is_group:
            # For groups, create entry only for count=1
            count = 1
            mass_diff_e = atomic_mass * count
            mass_diff_e_rounded = round(mass_diff_e, 4)
            description = f"{count} {element}-group"  # Use the group name with '1' and 'group' with hyphen
            if mass_diff_e_rounded in error_dict:
                if description not in error_dict[mass_diff_e_rounded]:
                    error_dict[mass_diff_e_rounded] += f", {description}"
            else:
                error_dict[mass_diff_e_rounded] = description

            # Positively Charged Ion (E+)
            mass_diff_e_plus = mass_diff_e + (electron_mass * count)
            mass_diff_e_plus_rounded = round(mass_diff_e_plus, 4)
            if mass_diff_e_plus_rounded in error_dict:
                if description not in error_dict[mass_diff_e_plus_rounded]:
                    error_dict[mass_diff_e_plus_rounded] += f", {description}"
            else:
                error_dict[mass_diff_e_plus_rounded] = description

            # Negatively Charged Ion (E-)
            mass_diff_e_minus = mass_diff_e - (electron_mass * count)
            mass_diff_e_minus_rounded = round(mass_diff_e_minus, 4)
            if mass_diff_e_minus_rounded in error_dict:
                if description not in error_dict[mass_diff_e_minus_rounded]:
                    error_dict[mass_diff_e_minus_rounded] += f", {description}"
            else:
                error_dict[mass_diff_e_minus_rounded] = description
        else:
            # For atoms, create entries for counts in counts_range
            for count in counts_range:
                mass_diff_e = atomic_mass * count
                mass_diff_e_rounded = round(mass_diff_e, 4)
                if count == 1:
                    description = f"{count} {element}-atom"
                else:
                    description = f"{count} {element}-atoms"

                if mass_diff_e_rounded in error_dict:
                    if description not in error_dict[mass_diff_e_rounded]:
                        error_dict[mass_diff_e_rounded] += f", {description}"
                else:
                    error_dict[mass_diff_e_rounded] = description

                # Positively Charged Ion (E+)
                mass_diff_e_plus = mass_diff_e + (electron_mass * count)
                mass_diff_e_plus_rounded = round(mass_diff_e_plus, 4)
                if mass_diff_e_plus_rounded in error_dict:
                    if description not in error_dict[mass_diff_e_plus_rounded]:
                        error_dict[mass_diff_e_plus_rounded] += f", {description}"
                else:
                    error_dict[mass_diff_e_plus_rounded] = description

                # Negatively Charged Ion (E-)
                mass_diff_e_minus = mass_diff_e - (electron_mass * count)
                mass_diff_e_minus_rounded = round(mass_diff_e_minus, 4)
                if mass_diff_e_minus_rounded in error_dict:
                    if description not in error_dict[mass_diff_e_minus_rounded]:
                        error_dict[mass_diff_e_minus_rounded] += f", {description}"
                else:
                    error_dict[mass_diff_e_minus_rounded] = description

    # Add Special Cases if Provided
    if special_cases:
        for mass, desc in special_cases.items():
            mass_float = float(mass)
            mass_rounded = round(mass_float, 4)
            if mass_rounded in error_dict:
                if desc not in error_dict[mass_rounded]:
                    error_dict[mass_rounded] += f", {desc}"
            else:
                error_dict[mass_rounded] = desc

    return error_dict


# Define special cases like electron mass error
special_errors = {
    '0.0005': "Electron mass error",
    '0.0006': "Electron mass error",
    '0.0073': "Nominal mass error (H=1.0000)?",
    '0.0072': "Nominal mass error (H=1.0000)?",
    '0.0071': "Nominal mass error (H=1.0000)?",
    '0.0070': "Nominal mass error (H=1.0000)?",
    '1.0005': "Nominal mass error (H=1.0000)?",
    '1.0006': "Nominal mass error (H=1.0000)?",
    '0.0102': "Nominal mass error (Na=23.0000)?",
    '0.0103': "Nominal mass error (Na=23.0000)?",
    '0.0107': "Nominal mass error (Na=23.0000)?",
    '0.0108': "Nominal mass error (Na=23.0000)?",
    '1.0077': '1 H-atom',
    '1.0076': '1 H-atom',
    '1.0075': '1 H-atom',
    '1.0083': '1 H-atom',
    '+22.9897': '1 Na-atom',
    '21.9892':"Nominal mass error [M]+1.0000 (not [M+Na]+)",
    '21.9893':"Nominal mass error [M]+1.0000 (not [M+Na]+)",
    '0.9964': 'Specify measured B-isotope(s)',
    '0.9963': 'Specify measured B-isotope(s)',
    '1.9927': 'Specify measured B-isotopes',
    '1.9928': 'Specify measured B-isotopes',
    '1.9979': 'Specify measured Br-isotope(s)',
    '1.9980': 'Specify measured Br-isotope(s)',
    '+17.9906':"Exchange 1 H- with 1 F-atom",
    '-17.9906':"Exchange 1 F- with 1 H-atom",
    '+14.9871':"Exchange 1 H- with 1 O-atom",
    '-14.9871':"Exchange 1 O- with 1 H-atom",
    '+77.9105':"Exchange 1 H- with 1 Br-atom",
    '-77.9105':"Exchange 1 Br- with 1 H-atom",
    '1.0039': 'Mass calcd for [M+1] (1x 13C)',
    '1.0038': 'Mass calcd for [M+1] (1x 13C)',
    '1.0034': 'Mass calcd for [M+1] (1x 13C)',
    '1.0033': 'Mass calcd for [M+1] (1x 13C)',
    '1.0032': 'Mass calcd for [M+1] (1x 13C)',
    '2.0064': 'Mass calcd for [M+2] (2x 13C)',

}

# Generate the error dictionary
elements = [
    'H',  'He', 'Li', 'Be', 'B',  'C',  'N',  'O',  'F',  'Ne',
    'Na', 'Mg', 'Al', 'Si', 'P',  'S',  'Cl', 'Ar', 'K',  'Ca',
    'Sc', 'Ti', 'V',  'Cr', 'Mn', 'Fe', 'Co', 'Ni', 'Cu', 'Zn',
    'Ga', 'Ge', 'As', 'Se', 'Br', 'Kr', 'Rb', 'Sr', 'Y',  'Zr',
    'Nb', 'Mo', 'Tc', 'Ru', 'Rh', 'Pd', 'Ag', 'Cd', 'In', 'Sn',
    'Sb', 'Te', 'I',  'Xe', 'Cs', 'Ba', 'La', 'Ce', 'Pr', 'Nd',
    'Pm', 'Sm', 'Eu', 'Gd', 'Tb', 'Dy', 'Ho', 'Er', 'Tm', 'Yb',
    'Lu', 'Hf', 'Ta', 'W',  'Re', 'Os', 'Ir', 'Pt', 'Au', 'Hg',
    'Tl', 'Pb', 'Bi','D','CH','CH2','CH3','CH4','NH','NH2','NH3','NH4',
    'OH','H2O','H3O','NO','NO2','OCH3','CF3','C2H5','HF','HCl','HBr','HS','HI'

]

atom_counts = range(1, 11)  # 1 to 10
error_dictionary = generate_error_dictionary(elements, atom_counts, special_errors)

def categorize_error(error_value, known_errors, tolerance=0.0001):
    """
    Categorizes the error based on a given error value and a dictionary of known atomic masses.
    Generates a message indicating whether atoms should be added or removed.

    Parameters:
    error_value (float): The calculated error between the calculated and recalculated mass.
    known_errors (dict): A dictionary where keys are atomic masses and values are the element descriptions.
    tolerance (float): The tolerance range within which the error value should match a known difference.

    Returns:
    str: The dynamically generated error message if a match is found, otherwise returns a blank space for zero difference.
    """
    # Check if the error value is effectively zero within the tolerance range
    if abs(error_value) <= tolerance:
        return ""  # Return a blank space if the difference is zero

    # Special case handling for known mass differences
    for atomic_mass, atom_description in known_errors.items():
        # Check if the error matches the dictionary value or the dictionary value plus 0.0001
        if (abs(abs(error_value) - atomic_mass) <= tolerance or
                abs(abs(error_value) - (atomic_mass + 0.0001)) <= tolerance):

            if len(atom_description) > 13:  # Check if the database entry is longer than six characters
                return atom_description  # Return the database entry directly

            # Extract the count and element from the dictionary entry
            parts = atom_description.split()
            if len(parts) != 2:
                # Handle unexpected format
                return atom_description

            count_str, element = parts
            try:
                count = int(count_str)
            except ValueError:
                # Handle cases where count is not an integer
                return atom_description

            # Generate the correct message based on the sign of the error
            if error_value > 0:
                return f"Add {count} {element} to formula"
            else:
                return f"Remove {count} {element} from formula"

    # If no match found, return the error value as a string with the correct sign
    return f"{error_value:+.4f}"


def hrms_cleanup(result, error_dictionary):
    """
    Processes a list of HRMS data strings and extracts specified components,
    ensuring that the ion notation is correctly captured and then removed from the line.
    Before processing each line, it removes all strings within the line that are shorter than
    5 characters and do not contain a capital 'M'.
    Recalculates the monoisotopic mass using the molmass library and computes error.

    Parameters:
    - result (list of str): The list containing HRMS data strings.
    - error_dictionary (dict): The autogenerated error dictionary with mass differences and descriptions.

    Returns:
    - list of list: A list where each sublist contains extracted data, including error calculations and descriptions.
    """

    # Initialize the parsed_results list
    parsed_results = []

    # Updated ion_pattern to include optional digits before 'M'
    ion_pattern = re.compile(r'\[\d*M[^]]*]\S*')

    # New formula pattern: word starting with 'C', followed by digits, 'H', digits, and possibly other elements
    #formula_pattern = re.compile(r'C\d+H\d+(?:[A-Z][a-z]?\d*|\[\d+[A-Z][a-z]*\d*)*[+-]?')
    #formula_pattern = re.compile(r'C\d+H\d+(?:[A-Z][a-z]?\d*|\[\d+[A-Z][a-z]*\d*\])*[+-]?')
    formula_pattern = re.compile(r'C\d+(?:H\d+|F\d+)(?:[A-Z][a-z]?\d*|\[\d+[A-Z][a-z]*\d*])*[+-]?')


    # Pattern for floats with exactly 4 digits after decimal point
    float_pattern = re.compile(r'\d+\.\d{4}')

    # Process each line in the result list
    for line in result:
        # Remove words shorter than 5 characters that do not contain a capital 'M'
        words = line.split()
        words_filtered = [word for word in words if len(word) >= 5 or ('M' in word)]
        line = ' '.join(words_filtered)

        # Initialize a row with 8 empty elements (added a column for Error)
        row = [''] * 8

        # Extract the ion notation and its charge
        ion_match = ion_pattern.search(line)
        ion_charge = ''
        if ion_match:
            ion = ion_match.group(0)
            row[1] = ion.strip()
            # Extract the charge from the ion notation if present (e.g., ]+, ]-, ]2+)
            ion_charge_match = re.search(r'(\d*[+-])?$', ion)
            if ion_charge_match:
                ion_charge = ion_charge_match.group(1)
            # Remove the ion notation from the line
            line = line.replace(ion, '')
        else:
            row[1] = ''

        # Now proceed to extract the formula, calcd mass, and found mass from the modified line

        # Extract the formula
        formula_match = formula_pattern.search(line)

        if formula_match:
            formula = formula_match.group(0).strip()
            # If the formula ends with ion_charge, remove ion_charge from formula
            if ion_charge and formula.endswith(ion_charge):
                formula = formula[:-len(ion_charge)].strip()
            # Check if there's a charge present in the formula
            charge_match = re.search(r'([+-]\d*)$', formula)
            if charge_match:
                charge = charge_match.group(1)
                formula_no_charge = formula.replace(charge, "")
            else:
                charge = ion_charge if ion_charge else '+'
                formula_no_charge = formula

            # Enclose the formula in square brackets before recalculating the mass

            formula_in_brackets = f'[{formula_no_charge}]{charge}'
            formula_in_brackets = formula_in_brackets.replace("H1HeXe", "[13C]")
            formula_in_brackets = formula_in_brackets.replace("C1F", "CF")
            formula_in_brackets = formula_in_brackets.replace("H1N", "HN")
            row[0] = formula_in_brackets

            # Recalculate the monoisotopic mass using molmass while keeping isotopic notation intact
            try:
                recalculated_mass = Formula(formula_in_brackets).monoisotopic_mass
                if ion_charge:
                    if ion_charge in ("+", "-"):
                        charge_number = 1
                    else:
                        charge_number = int(ion_charge[:-1])  # Extract the numeric part of the charge
                    recalculated_mass /= abs(charge_number)

                row[4] = f'{recalculated_mass:.4f}'  # Store the monoisotopic mass with 4 decimal precision
            except Exception as e:
                row[4] = 'Error'  # Handle the case where the formula is invalid for molmass
        else:
            row[0] = ''
            row[4] = ''

        # Extract all floats with exactly 4 decimal places
        floats_with_4_decimals = float_pattern.findall(line)

        # Extract the calcd mass - first occurring float with 4 decimal places
        if floats_with_4_decimals:
            calcd_mass = floats_with_4_decimals[0]
            row[2] = calcd_mass.strip()
        else:
            row[2] = ''

        # Extract the found mass - second float with 4 decimal places, if it exists
        if len(floats_with_4_decimals) >= 2:
            found_mass = floats_with_4_decimals[1]
            row[3] = found_mass.strip()
        else:
            row[3] = ''

        # Calculate the error between the calculated mass and the recalculated mass
        if row[2] and row[4] and row[2] != 'Error' and row[4] != 'Error':
            try:
                error = float(row[2]) - float(row[4])
                # Categorize the error based on the error value
                error_description = categorize_error(error, error_dictionary)
                # Check for a typo error if no existing error description
                if is_float(error_description) or error==0:

                    if differ_in_single_digit_except_last_two(float(row[2]), float(row[3])):
                        error_description = "Typo (Calcd,Found)"

                    if differ_in_single_digit_except_last_two(float(row[2]), float(row[4])):
                        error_description = "Typo (Calcd,Recalcd)"

                    if have_swapped_adjacent_digits(float(row[2]), float(row[3])):
                        error_description = "Transposed digits (Calcd,Found)"

                    if have_swapped_adjacent_digits(float(row[2]), float(row[4])):
                        error_description = "Transposed digits (Calcd,Recalcd)"

                    if error_description in ("-0.0010", "-0.0011", "-0.0012") and ion_charge == "-":
                        error_description = "Mass was calculated for cation"

                    if error_description in ("-0.0010", "-0.0011", "-0.0012") and "M-" in row[1]:
                        error_description = "Mass was calculated for cation"
                    #print(error_description)

                    mw_plus = round(calculate_molecular_weight(row[0]), 4)
                    if float(row[2]) == mw_plus:
                        error_description = "Molecular weight error"

                    mw_plus_plus1 = round(mw_plus + 1, 4)
                    if float(row[2]) == mw_plus_plus1:
                        error_description = "Molecular weight error"

                    mw_plus_plus23 = round(mw_plus + 23, 4)
                    if float(row[2]) == mw_plus_plus23:
                        error_description = "Molecular weight error"

                    formula_neutral = row[0].replace("+", "")
                    mw_neutral = round(calculate_molecular_weight(formula_neutral), 4)
                    if mw_neutral == float(row[2]):
                        error_description = "Molecular weight error (neutral)"

                    mw_neutral_plus1 = round(mw_neutral + 1, 4)
                    if mw_neutral_plus1 == float(row[2]):
                        error_description = "Molecular weight error (neutral+1)"

                    mw_neutral_plus23 = round(mw_neutral + 23, 4)
                    if mw_neutral_plus23 == float(row[2]):
                        error_description = "Molecular weight error (neutral+23)"

                    if "Na" in row[0]:
                        formula_minus_sodium = row[0].replace("Na", "")
                        mw1 = round(calculate_molecular_weight(formula_minus_sodium), 4) + 23
                        if mw1 == float(row[2]):
                            error_description = "Molecular weight + 23.0000"
                    else:
                        formula_plus_sodium = row[0].replace("[", "").replace("]", "").replace("+", "").replace("-", "")
                        formula_plus_sodium = formula_plus_sodium+"Na"
                        mw_plus_sodium = round(calculate_molecular_weight(formula_plus_sodium), 4)
                        if mw_plus_sodium == float(row[2]):
                            error_description = "Molecular weight error (Formula+Na)"

                    formula_minus_h = row[0].replace("[", "").replace("]", "").replace("+", "").replace("-", "")
                    formula_minus_h = decrease_element_count(formula_minus_h, 'H')
                    mw2 = round(calculate_molecular_weight(formula_minus_h), 4) + 1
                    if mw2 == float(row[2]):
                        error_description = "Molecular weight + 1.0000"

                row[7] = error_description  # Replace the error value with the error description or keep the difference

            except ValueError:
                row[7] = 'Error'
        else:
            row[7] = 'Error'

        if row[1] and row[2] and row[3] and not row[0]:
            row[7] = 'No formula found'

        # Skip the row if both row[0] and row[1] are empty
        if not row[0] and not row[1]:
            continue  # Do not append this row to parsed_results

        # Append the row to the parsed_results list
        parsed_results.append(row)

    return parsed_results


def calc_dev_calcd_and_recalcd(cleaned_results):
    """
    Calculates the absolute deviation between the calculated mass, recalculated mass, and the found mass in ppm,
    and updates the 'Dev (Calcd)' and 'Dev (Recalcd)' columns in the cleaned_results list.

    Parameters:
    cleaned_results (list of list): The list containing extracted data.

    Returns:
    list of list: The updated cleaned_results list with 'Dev (Calcd)' and 'Dev (Recalcd)' columns filled.
    """
    for row in cleaned_results:
        calcd_mass = row[2]
        found_mass = row[3]
        recalcd_mass = row[4]

        # Initialize found_mass_float only if found_mass exists and is valid
        found_mass_float = None
        if found_mass:
            try:
                found_mass_float = float(found_mass)
            except ValueError:
                found_mass_float = None

        # Calculate deviation for the calculated mass
        if calcd_mass and found_mass_float is not None:
            try:
                calcd_mass_float = float(calcd_mass)
                deviation_calcd = abs((found_mass_float - calcd_mass_float) / calcd_mass_float) * 1e6  # ppm
                row[5] = f"{deviation_calcd:.1f}"  # Format to one decimal place
            except ValueError:
                row[5] = ''  # Leave the field empty if conversion fails
        else:
            row[5] = ''

        # Calculate deviation for the recalculated mass
        if recalcd_mass and found_mass_float is not None:
            try:
                recalcd_mass_float = float(recalcd_mass)
                deviation_recalcd = abs((found_mass_float - recalcd_mass_float) / recalcd_mass_float) * 1e6  # ppm
                row[6] = f"{deviation_recalcd:.1f}"  # Format to one decimal place
            except ValueError:
                row[6] = ''  # Leave the field empty if conversion fails
        else:
            row[6] = ''
    return cleaned_results


def print_aligned_table(cleaned_results, pdf_file_path):
    """
    Prints the cleaned_results in an aligned table format,
    highlighting deviations greater than 10 ppm in red and error messages in purple.
    Also exports the table to an Excel file named based on the pdf_file_path.
    """
    headers = ['Formula', 'Ion', 'Calcd Mass', 'Found Mass', 'Recalcd Mass', 'Dev (Calcd)', 'Dev (Recalcd)', 'Error']

    # Calculate column widths for all columns, including the new 'Error' column
    col_widths = [max(len(str(row[i])) for row in [headers] + cleaned_results) for i in range(8)]

    # Print headers
    header_row = '  '.join(f"{headers[i]:<{col_widths[i]}}" for i in range(8))
    print(header_row)
    print('-' * len(header_row))

    # ANSI escape codes for colors
    red = '\033[31m'
    purple = '\033[35m'
    reset = '\033[0m'

    # Prepare data for the Excel file
    excel_data = []

    # Print each row
    for row in cleaned_results:
        # Collect data for Excel
        excel_row = []

        # Ensure that mass values are right-aligned and preserve trailing zeros
        row_output = []
        for i in range(8):  # Loop through all 8 columns
            cell_content = row[i]
            excel_row.append(cell_content)  # Add cell content to excel data

            if i in [2, 3, 4]:  # Calcd Mass, Found Mass, and Recalcd Mass columns
                # Format the mass values to preserve trailing zeros
                formatted_cell = f"{cell_content:>{col_widths[i]}}"
                row_output.append(formatted_cell)
            elif i in [5, 6]:  # Dev (Calcd) and Dev (Recalcd) columns
                # Check if deviation is greater than 10 ppm and highlight in red
                try:
                    deviation = float(cell_content)
                    if deviation > 10:
                        formatted_cell = f"{red}{cell_content:>{col_widths[i]}}{reset}"
                    else:
                        formatted_cell = f"{cell_content:>{col_widths[i]}}"
                except (ValueError, TypeError):
                    formatted_cell = f"{cell_content:>{col_widths[i]}}"
                row_output.append(formatted_cell)
            elif i == 7:  # Error column
                # Highlight error messages in purple only if they are not numeric values
                if isinstance(cell_content, str) and not re.match(r'^[+-]?\d*\.?\d+$', cell_content):
                    formatted_cell = f"{purple}{cell_content:>{col_widths[i]}}{reset}"
                else:
                    formatted_cell = f"{cell_content:>{col_widths[i]}}"
                row_output.append(formatted_cell)
            else:
                row_output.append(f"{str(cell_content):<{col_widths[i]}}")
        print('  '.join(row_output))
        excel_data.append(excel_row)

    # Extract the base filename from pdf_file_path
    base_filename = os.path.basename(pdf_file_path)
    filename_without_ext = os.path.splitext(base_filename)[0]
    # Construct the output filename
    output_filename = f"output {filename_without_ext}.xlsx"
    # Define the output path on the Desktop
    desktop_path = destination_folder
    excel_file_path = os.path.join(desktop_path, output_filename)

    # Create a DataFrame and export to Excel
    if write_report:
        df = pd.DataFrame(excel_data, columns=headers)
        df.to_excel(excel_file_path, index=False)
        print(f"\nData has been exported to Excel file at: {excel_file_path}")

def search_calcd_with_floats(text: str) -> List[str]:
    """
    Search for 'calcd' followed by two floats with four decimal places.
    Extract from up to 25 characters before 'calcd' (if no float present) until the second float.
    Only extract if total length is less than 100 characters.

    Args:
        text (str): Input text to search

    Returns:
        List[str]: List of matching strings
    """
    pattern_float = re.compile(r'\d+\.\d{4}')
    results = []

    # Find all occurrences of 'calcd', case-insensitive
    for calcd_match in re.finditer('calcd', text, re.IGNORECASE):
        calcd_start = calcd_match.start()

        # Look at up to 25 characters before 'calcd'
        pre_calcd_start = max(0, calcd_start - 25)
        pre_calcd_text = text[pre_calcd_start:calcd_start]

        # Check if there's a float in the pre-calcd text
        pre_calcd_floats = list(pattern_float.finditer(pre_calcd_text))

        # Determine the start position based on pre-calcd text
        if not pre_calcd_floats:  # If no floats found before calcd
            extraction_start = pre_calcd_start
        else:
            extraction_start = calcd_start

        # Look ahead for floats after 'calcd'
        post_calcd_text = text[calcd_start:calcd_start + 100]
        post_floats = list(pattern_float.finditer(post_calcd_text))

        if len(post_floats) >= 2:
            # End at the second float
            end_pos = calcd_start + post_floats[1].end()

            # Only extract if total length is less than 100 characters
            if end_pos - extraction_start < 100:
                result = text[extraction_start:end_pos]
                results.append(result)

    return results


def search_hrms_with_floats(text: str) -> List[str]:
    """
    Search for 'HRMS' followed by at least two floats with four decimal places.
    If 'calcd' appears in the 25 characters after the second float, stop at the second float.
    Otherwise, include up to 25 characters after the second float.

    Args:
        text (str): Input text to search

    Returns:
        List[str]: List of matching strings
    """
    pattern_float = re.compile(r'\d+\.\d{4}')
    hrms_positions = [m.start() for m in re.finditer('HRMS', text)]
    results = []

    for hrms_pos in hrms_positions:
        # Extract up to 100 characters from 'HRMS'
        max_length_substring = text[hrms_pos:hrms_pos + 100]
        floats = list(pattern_float.finditer(max_length_substring))

        if len(floats) >= 2:
            second_float_end = floats[1].end()

            # Look at the next 25 characters after the second float
            next_25_chars = max_length_substring[second_float_end:second_float_end + 25]

            # If 'calcd' appears in next 25 chars, stop at second float
            if 'calcd' in next_25_chars.lower():
                end_pos = hrms_pos + second_float_end
            else:
                # If no 'calcd', include up to 25 characters after second float
                end_pos = hrms_pos + second_float_end + 25

            # Ensure end position doesn't exceed text length or 100 characters from 'HRMS'
            end_pos = min(len(text), end_pos, hrms_pos + 100)
            result = text[hrms_pos:end_pos].strip()
            results.append(result)

    return results

def process_replacements(text: str) -> str:
    """
    Perform all necessary string replacements on the text.
    """
    replacements = {
        r'LCMS':'HRMS',
        r'HRESIMS':"HRMS",
        r'HRESI': 'HRMS',
        r'HR-MS': 'HRMS',
        r'ESI-MS': ' HRMS',
        r'‐': '-',
        r'‒':r'-',
        r'MHz':'',
        r'MeOD':'',
        r'Cal':"cal",
        r'calculated': 'calcd ',
        r'calcd.': 'calcd ',
        r'calc. ': 'calcd ',
        r'calc ': 'calcd ',
        r'chemical':'',
        r'formula':'',
        r' ⊕': "+",
        r'•': "",
        r'＋': "+",
        r'Observed':' ',
        r'observed':' ',

    }

    for pattern, replacement in replacements.items():
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    text = ' '.join(text.split()).strip()
    return text

def list_pdfs_in_folder(directory_path: str) -> List[str]:
    """
    List all PDF file paths in the provided directory path.
    """
    try:
        if not os.path.isdir(directory_path):
            logging.error("The provided path is not a valid directory.")
            return []
        return [
            os.path.join(directory_path, filename)
            for filename in os.listdir(directory_path)
            if filename.lower().endswith('.pdf')
        ]
    except Exception as e:
        logging.error(f"An error occurred while listing PDFs: {e}")
        return []

def extract_text_from_pdf(file_path: str) -> str:
    """
    Extract text from a PDF file using PyMuPDF (fitz).
    """
    try:
        with fitz.open(file_path) as pdf_document:
            text_content = ""
            for page_num in range(pdf_document.page_count):
                page = pdf_document.load_page(page_num)
                text_content += page.get_text()
            return text_content
    except Exception as e:
        logging.error(f"Error extracting text from {file_path}: {e}")
        return ""

def main():
    # Source folder
    folder_path = source_folder
    start_time = time.time()
    hrms_total_measurements=0
    pdf_filepaths = list_pdfs_in_folder(folder_path)
    if not pdf_filepaths:
        logging.info("No PDF files found to process.")
        return

    for pdf_file_path in pdf_filepaths:
        text_content = extract_text_from_pdf(pdf_file_path)
        #print(text_content)
        if not text_content:
            logging.warning(f"No text content extracted from {pdf_file_path}")
            continue

        text_content = re.sub(r'\s+', ' ', text_content).strip()  # Replace multiple spaces with a single space
        text_content=process_replacements(text_content)
        text_content=replace_comma_with_decimal(text_content)
        text_content=adjust_space_around_decimal(text_content)
        text_content=fix_floats(text_content)
        text_content = remove_page_numbers(text_content)
        text_content = re.sub(r'\[((C\d+(?:[A-Z][a-z]?\d*)*),\s*([M+][^]]+))',r'\1 [\3]', text_content)
        text_content = re.sub(r'(C)(\d+)(h)(\d+)', lambda m: f'C{m.group(2)}H{m.group(4)}', text_content, flags=re.IGNORECASE)
        text_content = re.sub(r'(c)(\d+)(H)(\d+)', lambda m: f'C{m.group(2)}H{m.group(4)}', text_content, flags=re.IGNORECASE)
        text_content = re.sub(r'\b(C)(\d+)(HD)\b', r'C\2H1D', text_content)
        text_content = re.sub(r'\b(C)\s*(\d*)\s*(H)\s*(\d*)\s*(N)\s*(\d*)\b',
                      lambda
                          m: f"{m.group(1)}{m.group(2) or ''}{m.group(3)}{m.group(4) or ''}{m.group(5)}{m.group(6) or ''}",
                      text_content)

        text_content = re.sub(r'\b(C)\s*(\d*)\s*(H)\s*(\d*)\s*(O)\s*(\d*)\b',
                      lambda
                          m: f"{m.group(1)}{m.group(2) or ''}{m.group(3)}{m.group(4) or ''}{m.group(5)}{m.group(6) or ''}",
                      text_content)
        text_content = text_content.replace("C2o","C20").replace("C1o","C10").replace("Cal","cal")
        text_content = re.sub(r'B(\d+)H(\d+)', r'H\2B\1', text_content)
        text_content = text_content.replace('\n', ' ').replace('+-', '+').replace(':'," ").replace('–','-').replace(','," ")
        text_content = remove_spaces_within_brackets(text_content)
        #Remove nested brackets from [(M+H]]+ etc
        text_content = re.sub(r'\(\[([^]]{1,10})]\+\)', r'[\1]+', text_content)
        text_content = re.sub(r'\[\[([^]]{1,10})]\+]', r'[\1]+', text_content)
        text_content = text_content.replace(' [[', '[').replace(']]', ']')
        replacements = {
            "₁": "1", "₂": "2", "₃": "3", "₄": "4", "₅": "5",
            "₆": "6", "₇": "7", "₈": "8", "₉": "9", "₀": "0", "¹": "1", "²": "2", "³": "3",
            "⁴": "4", "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9", "⁰": "0","С":"C","Н":"H",
            "C ": "C", " H ": "H", " F ":"F", " N ": "N", " Cl ":"Cl", " Br ":"Br", " O ": "O"," I ": "I",
            " P ":"P"," B ":"B", " S ":"S"," NO ":"NO", " Na ": "Na", " SNa ": "SNa"," NNa ":"NNa",
            " + ":"+ ",

        }

        for original, replacement in replacements.items():
            text_content = text_content.replace(original, replacement)
        text_content = remove_spaces_in_formula(text_content)
        text_content = text_content.replace('#', '')
        text_content = re.sub(r'(C\d+)', r' \1', text_content)
        text_content = transform_expressions_in_text(text_content)
        text_content=isotope_correct(text_content)
        text_content=protect_floats(text_content)
        text_content=text_content.replace("[13C]","H1HeXe")
        text_content = text_content.replace("CF", "C1F")
        text_content = text_content.replace("HN", "H1N")
        results1 = search_hrms_with_floats(text_content)
        modified_text = text_content
        for match in results1:
            modified_text = modified_text.replace(match, '')
        # Clean up any double spaces created by the removals
        modified_text = re.sub(r'\s+', ' ', modified_text).strip()
        text_content=modified_text
        results2 = search_calcd_with_floats(text_content)

        results=results1+results2
        cleaned_results = hrms_cleanup(results, error_dictionary)

        cleaned_results = calc_dev_calcd_and_recalcd(cleaned_results)

        cleaned_results = remove_sublists_with_missing_element1_positions_swapped(cleaned_results)

        #cleaned_results = [list(item) for item in set(tuple(sublist) for sublist in cleaned_results)]
        cleaned_results_new = []
        for sublist in cleaned_results:
            if sublist not in cleaned_results_new:
                cleaned_results_new.append(sublist)
        cleaned_results=cleaned_results_new

        # a counter for the total number of measurements
        num_row=len(cleaned_results)
        hrms_total_measurements=hrms_total_measurements+num_row

        if cleaned_results:
            print(" ")
            print(pdf_file_path)
            print_aligned_table(cleaned_results, pdf_file_path)
            if check_conditions(cleaned_results):
                print("\nAwesome! No mistakes!")
            #for result in results:
                #print(result)

        else:
            print(" ")
            print(f"No HRMS matches found in {pdf_file_path}")


    elapsed_time = time.time() - start_time
    minutes, seconds = divmod(elapsed_time, 60)


    print(f"\nTotal number of measurements found: {hrms_total_measurements}")
    print(f"Elapsed time: {int(minutes)} minutes and {int(seconds)} seconds")
if __name__ == '__main__':
    main()