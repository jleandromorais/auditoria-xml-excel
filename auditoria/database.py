"""
Compatibilidade de imports.

O projeto tem o `database.py` no diretório raiz de `auditoria-xml-excel/`,
mas outros módulos (ex.: `auditoria/audit.py`) importam `from .database import AuditDB`.

Este arquivo apenas reexporta `AuditDB` do módulo raiz.
"""

from database import AuditDB  # noqa: F401

