"""
Shared constants for VBA Component Manager
"""

from enum import IntEnum

SUPPORTED_VBA_PROJECTS = {
    ".xlsm",
    ".xls",
    ".xlsb",
    ".xla",
    ".xlam",
}

class ComponentType(IntEnum):
    STANDARD_MODULE = 1
    CLASS_MODULE = 2
    USER_FORM = 3
    DOCUMENT_MODULE = 100


EXPORT_FOLDERS = {
    ComponentType.STANDARD_MODULE: "mod",
    ComponentType.CLASS_MODULE: "cls",
    ComponentType.USER_FORM: "frm",
    ComponentType.DOCUMENT_MODULE: "doc",
}

EXPORT_EXTENSIONS = {
    ComponentType.STANDARD_MODULE: ".bas",
    ComponentType.CLASS_MODULE: ".cls",
    ComponentType.USER_FORM: ".frm",
    ComponentType.DOCUMENT_MODULE: ".cls",
}

SUPPORTED_COMPONENT_TYPES = {
    ComponentType.STANDARD_MODULE,
    ComponentType.CLASS_MODULE,
    ComponentType.USER_FORM,
    ComponentType.DOCUMENT_MODULE,
}