from __future__ import annotations


class formula(str):
    """Excel formula type."""
    pass


# Type alias for supported native cell types
NativeTypes = str | int | float | bool | formula