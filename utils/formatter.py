"""
VBA/VB6/VBScript Code Formatter
"""

import re
from typing import List


def format_vba_code(code: str, indent: int = 4) -> str:
    """
    Format VBA/VB6/VBScript code with:
    - Proper indentation
    - Block spacing
    - Procedure spacing
    - Metadata preservation (including indentation within BEGIN/END blocks)
    """
    if not code:
        return code

    lines = code.split("\n")

    processed_lines = []
    current_indent_level = 0
    in_metadata_block = False

    for line in lines:
        original_line = line.rstrip()
        stripped_line = line.strip()

        # Labels must NOT be indented
        if _is_label(stripped_line):
            processed_lines.append(stripped_line)
            continue

        # Preserve metadata formatting (including indentation within BEGIN/END blocks)
        if _is_metadata_line(stripped_line):
            if in_metadata_block:
                # Preserve original indentation inside BEGIN/END metadata blocks
                processed_lines.append(original_line)
            else:
                # For standalone metadata lines, preserve original formatting
                processed_lines.append(original_line)

            if stripped_line.upper().startswith("BEGIN"):
                in_metadata_block = True
            elif stripped_line.upper() == "END" and in_metadata_block:
                in_metadata_block = False

            continue

        # Preserve blank lines for cleanup later
        if not stripped_line:
            processed_lines.append("")
            continue

        # Else/Case handling
        if _is_else_like_statement(stripped_line):
            current_indent_level = max(0, current_indent_level - 1)

            processed_lines.append(
                (" " * (current_indent_level * indent)) + stripped_line
            )

            current_indent_level += 1
            continue

        # Closing blocks
        if _is_closing_statement(stripped_line):
            current_indent_level = max(0, current_indent_level - 1)

            processed_lines.append(
                (" " * (current_indent_level * indent)) + stripped_line
            )

            continue

        # Regular/opening statements
        processed_lines.append(
            (" " * (current_indent_level * indent)) + stripped_line
        )

        if _is_opening_statement(stripped_line):
            current_indent_level += 1

    # Cleanup + spacing passes
    cleaned_lines = _cleanup_blank_lines(processed_lines)
    cleaned_lines = _handle_option_explicit(cleaned_lines)
    cleaned_lines = _add_procedure_spacing(cleaned_lines)
    cleaned_lines = _remove_blank_lines_in_type_enum(cleaned_lines)
    cleaned_lines = _add_blank_lines_around_blocks(cleaned_lines)
    cleaned_lines = _cleanup_blank_lines(cleaned_lines)

    return "\n".join(cleaned_lines)


def _is_metadata_line(line: str) -> bool:
    """
    Detect VBA metadata lines.
    """
    stripped = line.strip()

    metadata_patterns = [
        r"^VERSION\s+",
        r"^BEGIN\s*$",
        r"^END\s*$",
        r"^Attribute\s+VB_",
        r"^Option\s+Explicit\s*$",
        r"^#If\s+",
        r"^#Else\s*$",
        r"^#ElseIf\s+",
        r"^#End\s+If\s*$",
    ]

    for pattern in metadata_patterns:
        if re.match(pattern, stripped, re.IGNORECASE):
            return True

    return False


def _is_opening_statement(line: str) -> bool:
    """
    Detect opening block statements.
    """
    stripped = line.strip()

    # Ignore single-line If statements
    if re.match(
        r"^If\s+.+\s+Then\s+.+",
        stripped,
        re.IGNORECASE,
    ):
        return False

    opening_patterns = [
        r"^(?:Public\s+|Private\s+|Friend\s+|Protected\s+)?Sub\s+",
        r"^(?:Public\s+|Private\s+|Friend\s+|Protected\s+)?Function\s+",
        r"^(?:Public\s+|Private\s+|Friend\s+|Protected\s+)?Property\s+(Get|Let|Set)\s+",
        r"^If\s+.+\s+Then\s*$",
        r"^Select\s+Case\s+",
        r"^For\s+",
        r"^For\s+Each\s+",
        r"^Do\b",
        r"^While\s+",
        r"^With\s+",
        r"^Type\s+",
        r"^Enum\s+",
    ]

    for pattern in opening_patterns:
        if re.match(pattern, stripped, re.IGNORECASE):
            return True

    return False


def _is_else_like_statement(line: str) -> bool:
    """
    Else/Case statements decrease then increase indentation.
    """
    stripped = line.strip()

    patterns = [
        r"^Else\s*$",
        r"^ElseIf\s+",
        r"^Case\s+",
        r"^Case\s+Else\s*$",
    ]

    for pattern in patterns:
        if re.match(pattern, stripped, re.IGNORECASE):
            return True

    return False


def _is_closing_statement(line: str) -> bool:
    """
    Detect closing block statements.
    """
    stripped = line.strip()

    patterns = [
        r"^End\s+Sub\s*$",
        r"^End\s+Function\s*$",
        r"^End\s+Property\s*$",
        r"^End\s+If\s*$",
        r"^End\s+Select\s*$",
        r"^Next\b",
        r"^Loop\b",
        r"^Wend\s*$",
        r"^End\s+With\s*$",
        r"^End\s+Type\s*$",
        r"^End\s+Enum\s*$",
    ]

    for pattern in patterns:
        if re.match(pattern, stripped, re.IGNORECASE):
            return True

    return False


def _cleanup_blank_lines(lines: List[str]) -> List[str]:
    """
    Collapse multiple blank lines into one.
    """
    cleaned = []
    blank_count = 0

    for line in lines:
        stripped = line.rstrip()

        if not stripped:
            blank_count += 1

            if blank_count <= 1:
                cleaned.append("")
        else:
            blank_count = 0
            cleaned.append(stripped)

    while cleaned and not cleaned[-1]:
        cleaned.pop()

    return cleaned

def _is_label(line: str) -> bool:
    """
    Detect real VBA labels (NOT inline statements).
    A label must be the only statement on the line.
    """

    stripped = line.strip()

    # Must end with ":" and nothing else
    if not stripped.endswith(":"):
        return False

    # Must NOT contain inline statements
    if ":" in stripped[:-1]:
        return False

    # Must be a valid identifier
    return bool(
        re.match(r"^[A-Za-z_][A-Za-z0-9_]*:\s*$", stripped)
    )

def _is_option_explicit(line: str) -> bool:
    """
    Detect Option Explicit.
    """
    return re.match(
        r"^Option\s+Explicit\s*$",
        line.strip(),
        re.IGNORECASE,
    ) is not None


def _handle_option_explicit(lines: List[str]) -> List[str]:
    """
    Remove blank lines before Option Explicit.
    """
    result = []

    for line in lines:
        if _is_option_explicit(line):
            while result and not result[-1].strip():
                result.pop()

        result.append(line)

    return result


def _add_procedure_spacing(lines: List[str]) -> List[str]:
    """
    Add spacing inside procedures/functions/properties.

    Result:

    Private Sub Test()

        Debug.Print "Hello"

    End Sub
    """
    result = []

    procedure_start_patterns = [
        r"^(?:Public\s+|Private\s+|Friend\s+|Protected\s+)?Sub\s+",
        r"^(?:Public\s+|Private\s+|Friend\s+|Protected\s+)?Function\s+",
        r"^(?:Public\s+|Private\s+|Friend\s+|Protected\s+)?Property\s+(Get|Let|Set)\s+",
    ]

    procedure_end_patterns = [
        r"^End\s+Sub\s*$",
        r"^End\s+Function\s*$",
        r"^End\s+Property\s*$",
    ]

    def matches_any(line: str, patterns: List[str]) -> bool:
        stripped = line.strip()

        for pattern in patterns:
            if re.match(pattern, stripped, re.IGNORECASE):
                return True

        return False

    for i, line in enumerate(lines):
        stripped = line.strip()

        # Blank line BEFORE End Sub/Function/Property
        if (
            matches_any(stripped, procedure_end_patterns)
            and result
            and result[-1].strip()
        ):
            result.append("")

        result.append(line)

        # Blank line AFTER Sub/Function/Property declaration
        if matches_any(stripped, procedure_start_patterns):
            if i + 1 < len(lines):
                next_line = lines[i + 1]

                if next_line.strip():
                    result.append("")

    return result


def _add_blank_lines_around_blocks(lines: List[str]) -> List[str]:
    """
    Add one blank line BEFORE and AFTER major control blocks.
    """
    result = []

    opening_patterns = [
        r"^If\s+.+\s+Then\s*$",
        r"^Select\s+Case\s+",
        r"^For\s+",
        r"^For\s+Each\s+",
        r"^Do\b",
        r"^While\s+",
    ]

    closing_patterns = [
        r"^End\s+If\s*$",
        r"^End\s+Select\s*$",
        r"^Next\b",
        r"^Loop\b",
        r"^Wend\s*$",
    ]

    def matches_any(line: str, patterns: List[str]) -> bool:
        stripped = line.strip()

        for pattern in patterns:
            if re.match(pattern, stripped, re.IGNORECASE):
                return True

        return False

    for i, line in enumerate(lines):
        stripped = line.strip()

        # Blank line BEFORE opening block
        if (
            matches_any(stripped, opening_patterns)
            and result
            and result[-1].strip()
        ):
            result.append("")

        result.append(line)

        # Blank line AFTER closing block
        if matches_any(stripped, closing_patterns):
            if i + 1 < len(lines):
                next_line = lines[i + 1]

                if next_line.strip():
                    result.append("")

    return result


def _remove_blank_lines_in_type_enum(lines: List[str]) -> List[str]:
    """
    Ensures Type and Enum blocks contain NO blank lines inside them.
    """

    result = []
    inside_block = False

    type_enum_start = re.compile(r"^\s*(Public\s+|Private\s+)?(Type|Enum)\b", re.I)
    type_enum_end = re.compile(r"^\s*End\s+(Type|Enum)\b", re.I)

    for line in lines:
        stripped = line.strip()

        # Enter block
        if type_enum_start.match(line):
            inside_block = True
            result.append(line)
            continue

        # Exit block
        if type_enum_end.match(line):
            inside_block = False
            result.append(line)
            continue

        # Skip blank lines inside Type/Enum
        if inside_block and not stripped:
            continue

        result.append(line)

    return result