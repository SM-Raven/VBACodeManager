"""
VBA/VB6/VBScript Code Formatter
"""

import re
from typing import List


def format_vba_code(code: str, indent: int = 2) -> str:
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
    last_was_select_case = False

    for line in lines:
        original_line = line.rstrip()
        stripped_line = line.strip()

        # Check if we're entering a metadata block
        if _is_begin_statement(stripped_line):
            in_metadata_block = True
            processed_lines.append(original_line)
            continue

        # Check if we're exiting a metadata block
        if _is_end_statement(stripped_line) and in_metadata_block:
            in_metadata_block = False
            processed_lines.append(original_line)
            continue

        # If inside a metadata block, preserve everything as-is (with original indentation)
        if in_metadata_block:
            processed_lines.append(original_line)
            continue

        # Labels must NOT be indented
        if _is_label(stripped_line):
            processed_lines.append(stripped_line)
            continue

        # Preserve other metadata formatting
        if _is_metadata_line(stripped_line):
            processed_lines.append(original_line)
            continue

        # Preserve blank lines for cleanup later
        if not stripped_line:
            processed_lines.append("")
            last_was_select_case = False
            continue

        # Special handling for first Case after Select Case
        # The first Case should be indented normally (not dedented)
        if last_was_select_case and re.match(r"^Case\s+", stripped_line, re.IGNORECASE):
            processed_lines.append(
                (" " * (current_indent_level * indent)) + stripped_line
            )
            current_indent_level += 1
            last_was_select_case = False
            continue

        # Else/Case handling (for subsequent Cases and Else/ElseIf)
        if _is_else_like_statement(stripped_line):
            current_indent_level = max(0, current_indent_level - 1)

            processed_lines.append(
                (" " * (current_indent_level * indent)) + stripped_line
            )

            current_indent_level += 1
            last_was_select_case = False
            continue

        # Closing blocks
        if _is_closing_statement(stripped_line):

            # End Select needs one extra dedent because
            # Case blocks temporarily re-increase indentation
            if re.match(r"^End\s+Select\s*$", stripped_line, re.IGNORECASE):
                current_indent_level = max(0, current_indent_level - 2)
            else:
                current_indent_level = max(0, current_indent_level - 1)

            processed_lines.append(
                (" " * (current_indent_level * indent)) + stripped_line
            )

            last_was_select_case = False
            continue

        # Regular/opening statements
        processed_lines.append(
            (" " * (current_indent_level * indent)) + stripped_line
        )

        if _is_opening_statement(stripped_line):
            current_indent_level += 1
            # Track if this is a Select Case so we handle the first Case specially
            if re.match(r"^Select\s+Case\s+", stripped_line, re.IGNORECASE):
                last_was_select_case = True
        else:
            last_was_select_case = False

    # Cleanup + spacing passes
    cleaned_lines = _cleanup_blank_lines(processed_lines)
    cleaned_lines = _handle_option_explicit(cleaned_lines)
    cleaned_lines = _add_spacing_after_metadata_header(cleaned_lines)
    cleaned_lines = _add_procedure_spacing(cleaned_lines)
    cleaned_lines = _remove_blank_lines_in_type_enum(cleaned_lines)
    cleaned_lines = _add_blank_lines_around_blocks(cleaned_lines)
    cleaned_lines = _remove_blank_lines_in_properties(cleaned_lines)
    cleaned_lines = _cleanup_blank_lines(cleaned_lines)

    return "\n".join(cleaned_lines)


def _is_begin_statement(line: str) -> bool:
    """
    Detect BEGIN statement (start of metadata block).
    Matches: Begin, BEGIN, Begin {...}, Begin <space>, etc.
    Does NOT match: BeginTrans, BeginScope, etc.

    The metadata Begin must be followed by:
    - whitespace
    - opening brace {
    - nothing (end of line)
    """
    stripped = line.strip()
    # Match "Begin" only when followed by whitespace, {, or end of line
    return re.match(r"^Begin(?:\s|{|\s*$)", stripped, re.IGNORECASE) is not None


def _is_end_statement(line: str) -> bool:
    """
    Detect END statement (end of metadata block).
    Matches: End, END, etc. (but not End Sub, End Function, etc.)
    """
    stripped = line.strip()
    # Match standalone "End" but not "End Sub", "End Function", etc.
    return re.match(r"^End\s*$", stripped, re.IGNORECASE) is not None


def _is_metadata_line(line: str) -> bool:
    """
    Detect VBA metadata lines (non-metadata BEGIN/END blocks).
    Note: Metadata BEGIN/END blocks are handled separately by _is_begin_statement/_is_end_statement.
    """
    stripped = line.strip()

    metadata_patterns = [
        r"^VERSION\s+",
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
        r"^(?:Public\s+|Private\s+)?Type\s+",
        r"^(?:Public\s+|Private\s+)?Enum\s+",
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
    Add spacing inside procedures/functions/properties and between them.

    Result:

    Private Sub Test()

        Debug.Print "Hello"

    End Sub

    Private Sub Test2()

        Debug.Print "World"

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

        # Blank line AFTER Sub/Function declaration
        if matches_any(stripped, procedure_start_patterns):

            if i + 1 < len(lines):
                next_line = lines[i + 1]

                if next_line.strip():
                    result.append("")

        # Blank line AFTER End Sub/Function/Property (before next procedure)
        if matches_any(stripped, procedure_end_patterns):
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if next_line and matches_any(next_line, procedure_start_patterns):
                    result.append("")

    return result


def _add_spacing_after_metadata_header(lines: List[str]) -> List[str]:
    """
    Add blank line after metadata header section (VERSION/BEGIN/END/Attributes).

    For class files:
    VERSION 1.0 CLASS
    BEGIN
      MultiUse = -1
    END
    Attribute VB_Name = "casAPI"

    [blank line added here]

    Public Function Test()
    """
    result = []
    last_was_attribute = False
    last_was_metadata_block = False

    for i, line in enumerate(lines):
        stripped = line.strip()
        result.append(line)

        # Track if this is an Attribute line
        if re.match(r"^Attribute\s+VB_", stripped, re.IGNORECASE):
            last_was_attribute = True
            last_was_metadata_block = False
        # Track if we just exited a metadata BEGIN/END block
        elif re.match(r"^End\s*$", stripped, re.IGNORECASE):
            last_was_metadata_block = True
            last_was_attribute = False
        # If we see a non-metadata, non-blank line after attributes/blocks
        elif stripped and not re.match(r"^(?:VERSION|Attribute|Begin|End\s*$)", stripped, re.IGNORECASE):
            if last_was_attribute or last_was_metadata_block:
                # Add blank line if next line isn't already blank
                if i + 1 < len(lines) and lines[i + 1].strip():
                    result.append("")
            last_was_attribute = False
            last_was_metadata_block = False

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


def _remove_blank_lines_in_properties(lines: List[str]) -> List[str]:
    """
    Remove blank lines inside Property Get/Let/Set blocks.
    """

    result = []
    inside_property = False

    property_start = re.compile(
        r"^(?:Public\s+|Private\s+|Friend\s+|Protected\s+)?Property\s+(Get|Let|Set)\b",
        re.I,
    )

    property_end = re.compile(r"^End\s+Property\s*$", re.I)

    for line in lines:
        stripped = line.strip()

        if property_start.match(line):
            inside_property = True
            result.append(line)
            continue

        if property_end.match(line):
            inside_property = False
            result.append(line)
            continue

        if inside_property and not stripped:
            continue

        result.append(line)

    return result