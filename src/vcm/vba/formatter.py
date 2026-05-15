"""
VBA Formatter - Format VBA code with consistent style
Handles indentation, spacing, and common formatting rules
"""

import re
from typing import List


class VBAFormatter:
    """Formats VBA code"""
    
    # VBA keywords that change indentation
    KEYWORDS_INCREASE_INDENT = {
        'sub', 'function', 'property', 'class', 'if', 'then',
        'for', 'do', 'while', 'with', 'select', 'case', 'else',
        'elseif', 'public', 'private', 'friend'
    }
    
    KEYWORDS_DECREASE_INDENT = {
        'end', 'else', 'elseif', 'loop', 'wend', 'next', 'case'
    }
    
    INDENT_SIZE = 4  # Spaces per indent level
    
    @staticmethod
    def format_code(vba_code: str) -> str:
        """
        Format VBA code with consistent indentation and spacing.
        
        Args:
            vba_code: Raw VBA code
            
        Returns:
            Formatted VBA code
        """
        if not vba_code or not vba_code.strip():
            return vba_code
        
        lines = vba_code.split('\n')
        formatted_lines = []
        indent_level = 0
        
        for line in lines:
            # Strip whitespace
            stripped = line.strip()
            
            # Skip empty lines
            if not stripped:
                formatted_lines.append('')
                continue
            
            # Skip comment lines - preserve at current indent
            if stripped.startswith("'"):
                formatted_lines.append(
                    ' ' * (indent_level * VBAFormatter.INDENT_SIZE) + stripped
                )
                continue
            
            # Check if line decreases indent (End, Else, ElseIf, etc.)
            if VBAFormatter._should_decrease_indent(stripped):
                indent_level = max(0, indent_level - 1)
            
            # Format the line with current indent
            formatted_line = ' ' * (indent_level * VBAFormatter.INDENT_SIZE) + stripped
            formatted_lines.append(formatted_line)
            
            # Check if line increases indent (Sub, If, For, etc.)
            if VBAFormatter._should_increase_indent(stripped):
                indent_level += 1
        
        # Join lines and clean up
        formatted = '\n'.join(formatted_lines)
        formatted = VBAFormatter._clean_spacing(formatted)
        
        return formatted
    
    @staticmethod
    def _should_increase_indent(line: str) -> bool:
        """Check if line should increase indentation for next line"""
        line_lower = line.lower()
        
        # Remove comments from line for checking
        if "'" in line_lower:
            line_lower = line_lower.split("'")[0].lower()
        
        # Check for keywords that increase indent
        for keyword in VBAFormatter.KEYWORDS_INCREASE_INDENT:
            if re.search(r'\b' + keyword + r'\b', line_lower):
                # Exclude "End" which ends blocks
                if keyword != 'end':
                    return True
        
        # Special handling for Case statement
        if re.search(r'\bcase\b', line_lower) and 'case else' not in line_lower.lower():
            return True
        
        return False
    
    @staticmethod
    def _should_decrease_indent(line: str) -> bool:
        """Check if line should decrease indentation"""
        line_lower = line.lower()
        
        # Remove comments
        if "'" in line_lower:
            line_lower = line_lower.split("'")[0].lower()
        
        # Check for keywords that decrease indent
        for keyword in VBAFormatter.KEYWORDS_DECREASE_INDENT:
            if re.search(r'\b' + keyword + r'\b', line_lower):
                # ElseIf should not decrease
                if 'elseif' in line_lower:
                    return False
                # Case Else should not decrease
                if 'case else' in line_lower:
                    return False
                return True
        
        return False
    
    @staticmethod
    def _clean_spacing(code: str) -> str:
        """
        Clean up extra spacing and blank lines.
        
        Args:
            code: Formatted code
            
        Returns:
            Cleaned code
        """
        lines = code.split('\n')
        cleaned = []
        prev_empty = False
        
        for line in lines:
            is_empty = not line.strip()
            
            # Skip multiple consecutive empty lines
            if is_empty:
                if not prev_empty:
                    cleaned.append(line)
                prev_empty = True
            else:
                cleaned.append(line)
                prev_empty = False
        
        # Remove trailing empty lines but keep leading structure
        while cleaned and not cleaned[-1].strip():
            cleaned.pop()
        
        return '\n'.join(cleaned)
    
    @staticmethod
    def format_file(file_path) -> bool:
        """
        Format a VBA file in place.
        
        Args:
            file_path: Path to VBA file
            
        Returns:
            True if file was modified, False otherwise
        """
        from pathlib import Path
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        try:
            # Read original code
            with open(file_path, 'r', encoding='utf-8') as f:
                original = f.read()
            
            # Format code
            formatted = VBAFormatter.format_code(original)
            
            # Check if changed
            if formatted != original:
                # Write back
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(formatted)
                return True
            return False
        
        except Exception as e:
            raise IOError(f"Failed to format file {file_path}: {str(e)}") from e


# Example formatting test
if __name__ == "__main__":
    test_code = """
Sub TestSub()
Dim i As Integer
If True Then
For i = 1 To 10
Debug.Print i
Next i
End If
End Sub

Function GetValue() As String
Dim result As String
Select Case result
Case "A"
result = "Value A"
Case "B"
result = "Value B"
Case Else
result = "Unknown"
End Select
GetValue = result
End Function
"""
    
    print("Original:")
    print(test_code)
    print("\n" + "="*50 + "\n")
    
    formatted = VBAFormatter.format_code(test_code)
    print("Formatted:")
    print(formatted)
