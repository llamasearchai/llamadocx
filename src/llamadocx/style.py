"""
Classes for working with styles and formatting in Word documents.

This module provides classes for working with document styles, fonts, and
paragraph formatting. These classes wrap python-docx's style and formatting
classes with enhanced functionality.
"""

from typing import Optional, Union, Any

from docx.shared import Pt, RGBColor, Inches
from docx.styles.style import _ParagraphStyle, _CharacterStyle, _Style
from docx.text.parfmt import ParagraphFormat as _ParagraphFormat
from docx.text.font import Font as _Font


class Font:
    """
    Font settings for text.

    This class wraps python-docx's Font class and provides additional
    functionality for font formatting.

    Args:
        font (_Font): The underlying python-docx Font object

    Attributes:
        font (_Font): The underlying python-docx Font object
    """

    def __init__(self, font: _Font) -> None:
        self.font = font

    @property
    def name(self) -> Optional[str]:
        """Get/set the font name."""
        return self.font.name

    @name.setter
    def name(self, value: Optional[str]) -> None:
        self.font.name = value

    @property
    def size(self) -> Optional[float]:
        """Get/set the font size in points."""
        if self.font.size is None:
            return None
        return self.font.size.pt

    @size.setter
    def size(self, value: Optional[Union[int, float]]) -> None:
        if value is None:
            self.font.size = None
        else:
            self.font.size = Pt(value)

    @property
    def color(self) -> Optional[tuple]:
        """Get/set the font color as RGB tuple."""
        if self.font.color.rgb is None:
            return None
        rgb = self.font.color.rgb
        return (rgb.red, rgb.green, rgb.blue)

    @color.setter
    def color(self, value: Optional[Union[str, tuple]]) -> None:
        if value is None:
            self.font.color.rgb = None
        else:
            if isinstance(value, str):
                # Convert hex color to RGB
                value = value.lstrip("#")
                r, g, b = tuple(int(value[i:i+2], 16) for i in (0, 2, 4))
            else:
                r, g, b = value
            self.font.color.rgb = RGBColor(r, g, b)

    @property
    def bold(self) -> Optional[bool]:
        """Get/set whether the font is bold."""
        return self.font.bold

    @bold.setter
    def bold(self, value: Optional[bool]) -> None:
        self.font.bold = value

    @property
    def italic(self) -> Optional[bool]:
        """Get/set whether the font is italic."""
        return self.font.italic

    @italic.setter
    def italic(self, value: Optional[bool]) -> None:
        self.font.italic = value

    @property
    def underline(self) -> Optional[bool]:
        """Get/set whether the font is underlined."""
        return bool(self.font.underline)

    @underline.setter
    def underline(self, value: Optional[bool]) -> None:
        self.font.underline = value

    @property
    def strike(self) -> Optional[bool]:
        """Get/set whether the font is struck through."""
        return self.font.strike

    @strike.setter
    def strike(self, value: Optional[bool]) -> None:
        self.font.strike = value

    @property
    def subscript(self) -> Optional[bool]:
        """Get/set whether the font is subscript."""
        return self.font.subscript

    @subscript.setter
    def subscript(self, value: Optional[bool]) -> None:
        self.font.subscript = value

    @property
    def superscript(self) -> Optional[bool]:
        """Get/set whether the font is superscript."""
        return self.font.superscript

    @superscript.setter
    def superscript(self, value: Optional[bool]) -> None:
        self.font.superscript = value

    @property
    def small_caps(self) -> Optional[bool]:
        """Get/set whether the font uses small caps."""
        return self.font.small_caps

    @small_caps.setter
    def small_caps(self, value: Optional[bool]) -> None:
        self.font.small_caps = value

    @property
    def all_caps(self) -> Optional[bool]:
        """Get/set whether the font uses all caps."""
        return self.font.all_caps

    @all_caps.setter
    def all_caps(self, value: Optional[bool]) -> None:
        self.font.all_caps = value


class ParagraphFormat:
    """
    Paragraph formatting settings.

    This class wraps python-docx's ParagraphFormat class and provides
    additional functionality for paragraph formatting.

    Args:
        format (_ParagraphFormat): The underlying python-docx ParagraphFormat object

    Attributes:
        format (_ParagraphFormat): The underlying python-docx ParagraphFormat object
    """

    def __init__(self, format: _ParagraphFormat) -> None:
        self.format = format

    @property
    def alignment(self) -> Optional[int]:
        """Get/set the paragraph alignment."""
        return self.format.alignment

    @alignment.setter
    def alignment(self, value: Optional[int]) -> None:
        self.format.alignment = value

    @property
    def line_spacing(self) -> Optional[float]:
        """Get/set the line spacing."""
        return self.format.line_spacing

    @line_spacing.setter
    def line_spacing(self, value: Optional[float]) -> None:
        self.format.line_spacing = value

    @property
    def line_spacing_rule(self) -> Optional[int]:
        """Get/set the line spacing rule."""
        return self.format.line_spacing_rule

    @line_spacing_rule.setter
    def line_spacing_rule(self, value: Optional[int]) -> None:
        self.format.line_spacing_rule = value

    @property
    def space_before(self) -> Optional[float]:
        """Get/set the space before the paragraph in points."""
        if self.format.space_before is None:
            return None
        return self.format.space_before.pt

    @space_before.setter
    def space_before(self, value: Optional[Union[int, float]]) -> None:
        if value is None:
            self.format.space_before = None
        else:
            self.format.space_before = Pt(value)

    @property
    def space_after(self) -> Optional[float]:
        """Get/set the space after the paragraph in points."""
        if self.format.space_after is None:
            return None
        return self.format.space_after.pt

    @space_after.setter
    def space_after(self, value: Optional[Union[int, float]]) -> None:
        if value is None:
            self.format.space_after = None
        else:
            self.format.space_after = Pt(value)

    @property
    def left_indent(self) -> Optional[float]:
        """Get/set the left indentation in inches."""
        if self.format.left_indent is None:
            return None
        return self.format.left_indent.inches

    @left_indent.setter
    def left_indent(self, value: Optional[Union[int, float]]) -> None:
        if value is None:
            self.format.left_indent = None
        else:
            self.format.left_indent = Inches(value)

    @property
    def right_indent(self) -> Optional[float]:
        """Get/set the right indentation in inches."""
        if self.format.right_indent is None:
            return None
        return self.format.right_indent.inches

    @right_indent.setter
    def right_indent(self, value: Optional[Union[int, float]]) -> None:
        if value is None:
            self.format.right_indent = None
        else:
            self.format.right_indent = Inches(value)

    @property
    def first_line_indent(self) -> Optional[float]:
        """Get/set the first line indentation in inches."""
        if self.format.first_line_indent is None:
            return None
        return self.format.first_line_indent.inches

    @first_line_indent.setter
    def first_line_indent(self, value: Optional[Union[int, float]]) -> None:
        if value is None:
            self.format.first_line_indent = None
        else:
            self.format.first_line_indent = Inches(value)


class Style:
    """
    A document style.

    This class wraps python-docx's Style class and provides additional
    functionality for working with document styles.

    Args:
        style (_Style): The underlying python-docx Style object

    Attributes:
        style (_Style): The underlying python-docx Style object
    """

    def __init__(self, style: _Style) -> None:
        self.style = style

    @property
    def name(self) -> str:
        """Get the style name."""
        return self.style.name

    @property
    def type(self) -> int:
        """Get the style type."""
        return self.style.type

    @property
    def builtin(self) -> bool:
        """Get whether the style is built-in."""
        return self.style.builtin

    @property
    def base_style(self) -> Optional["Style"]:
        """Get the base style."""
        if self.style.base_style is None:
            return None
        return Style(self.style.base_style)

    @property
    def font(self) -> Optional[Font]:
        """Get the style's font settings."""
        if not hasattr(self.style, "font"):
            return None
        return Font(self.style.font)

    @property
    def paragraph_format(self) -> Optional[ParagraphFormat]:
        """Get the style's paragraph format settings."""
        if not hasattr(self.style, "paragraph_format"):
            return None
        return ParagraphFormat(self.style.paragraph_format)

    def __eq__(self, other: Any) -> bool:
        if isinstance(other, Style):
            return self.style == other.style
        return self.style == other

    def __str__(self) -> str:
        return f"Style(name='{self.name}')"

    def __repr__(self) -> str:
        return str(self) 