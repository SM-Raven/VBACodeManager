"""
VBA Formatter (VBAPretty-style)

Features:
- Proper VBA indentation
- Smart blank line cleanup
- Preserves module/class metadata
- Keeps attribute blocks together
- Handles If/Else/End If
- Handles Select Case
- Handles With
- Handles Do/Loop
- Handles For/Next
- Handles While/Wend
- Handles Enum/Type
- Handles Property blocks
- Removes excessive blank lines

Usage:
    python formatter.py input.bas
"""

from pathlib import Path
import re
import sys


INDENT = "    "  # 4 spaces


class VBAFormatter:
    def __init__(self):
        self.indent_level = 0

    def format_code(self, code: str) -> str:
        lines = code.splitlines()

        formatted = []

        prev_blank = False
        in_declaration_block = True

        for raw_line in lines:
            original = raw_line.rstrip()
            stripped = original.strip()

            # ---------------------------------------------------------
            # Preserve true empty lines (single only)
            # ---------------------------------------------------------
            if stripped == "":
                if not prev_blank:
                    formatted.append("")
                prev_blank = True
                continue

            prev_blank = False

            # ---------------------------------------------------------
            # Preserve module/class metadata exactly
            # ---------------------------------------------------------
            if self._is_metadata_line(stripped):
                formatted.append(stripped)
                continue

            # ---------------------------------------------------------
            # Add spacing between declarations and procedures
            # ---------------------------------------------------------
            if self._is_procedure_start(stripped):
                if (
                    formatted
                    and formatted[-1] != ""
                ):
                    formatted.append("")

                in_declaration_block = False

            # ---------------------------------------------------------
            # Reduce indent BEFORE writing line
            # ---------------------------------------------------------
            if self._should_dedent(stripped):
                self.indent_level = max(0, self.indent_level - 1)

            # ---------------------------------------------------------
            # Format current line
            # ---------------------------------------------------------
            line = f"{INDENT * self.indent_level}{stripped}"
            formatted.append(line)

            # ---------------------------------------------------------
            # Increase indent AFTER writing line
            # ---------------------------------------------------------
            if self._should_indent(stripped):
                self.indent_level += 1

        # -------------------------------------------------------------
        # Final cleanup
        # -------------------------------------------------------------
        return self._cleanup(formatted)

    # ================================================================
    # RULES
    # ================================================================

    def _is_metadata_line(self, line: str) -> bool:
        upper = line.upper()

        return (
            upper.startswith("VERSION ")
            or upper.startswith("BEGIN ")
            or upper.startswith("END")
            or upper.startswith("ATTRIBUTE ")
        )

    def _is_procedure_start(self, line: str) -> bool:
        upper = line.upper()

        patterns = (
            "SUB ",
            "FUNCTION ",
            "PROPERTY ",
            "PRIVATE SUB ",
            "PUBLIC SUB ",
            "FRIEND SUB ",
            "PRIVATE FUNCTION ",
            "PUBLIC FUNCTION ",
            "FRIEND FUNCTION ",
            "PRIVATE PROPERTY ",
            "PUBLIC PROPERTY ",
            "FRIEND PROPERTY ",
        )

        return upper.startswith(patterns)

    def _should_indent(self, line: str) -> bool:
        upper = line.upper()

        # Single-line If
        if re.search(r"\bIF\b.*\bTHEN\b.+", upper):
            if not upper.endswith("THEN"):
                return False

        indent_keywords = (
            "IF ",
            "FOR ",
            "DO",
            "WITH ",
            "SELECT CASE",
            "WHILE ",
            "ENUM ",
            "TYPE ",
            "SUB ",
            "FUNCTION ",
            "PROPERTY ",
            "PRIVATE SUB ",
            "PUBLIC SUB ",
            "PRIVATE FUNCTION ",
            "PUBLIC FUNCTION ",
            "PRIVATE PROPERTY ",
            "PUBLIC PROPERTY ",
        )

        if upper.startswith(indent_keywords):
            return True

        if upper.startswith("ELSE"):
            return True

        if upper.startswith("CASE "):
            return True

        return False

    def _should_dedent(self, line: str) -> bool:
        upper = line.upper()

        dedent_keywords = (
            "END IF",
            "NEXT",
            "LOOP",
            "WEND",
            "END WITH",
            "END SELECT",
            "END ENUM",
            "END TYPE",
            "END SUB",
            "END FUNCTION",
            "END PROPERTY",
        )

        if upper.startswith(dedent_keywords):
            return True

        if upper.startswith("ELSE"):
            return True

        if upper.startswith("CASE "):
            return True

        return False

    # ================================================================
    # CLEANUP
    # ================================================================

    def _cleanup(self, lines: list[str]) -> str:
        cleaned = []

        prev_blank = False

        for line in lines:
            blank = line.strip() == ""

            # Remove repeated blank lines
            if blank and prev_blank:
                continue

            # Remove blank line after BEGIN
            if (
                cleaned
                and cleaned[-1].strip().upper() == "BEGIN"
                and blank
            ):
                continue

            cleaned.append(line.rstrip())
            prev_blank = blank

        # Trim leading/trailing blank lines
        while cleaned and cleaned[0] == "":
            cleaned.pop(0)

        while cleaned and cleaned[-1] == "":
            cleaned.pop()

        return "\n".join(cleaned) + "\n"


# ====================================================================
# FILE API
# ====================================================================

def format_vba_file(file_path: str):
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(path)

    code = path.read_text(encoding="utf-8")

    formatter = VBAFormatter()
    formatted = formatter.format_code(code)

    path.write_text(formatted, encoding="utf-8")

    print(f"Formatted: {path}")


# ====================================================================
# CLI
# ====================================================================

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python formatter.py <file.bas>")
        sys.exit(1)

    format_vba_file(sys.argv[1])