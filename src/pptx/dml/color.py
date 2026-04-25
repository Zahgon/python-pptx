"""DrawingML objects related to color, ColorFormat being the most prominent."""

from __future__ import annotations

from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR
from pptx.oxml.dml.color import (
    CT_HslColor,
    CT_PresetColor,
    CT_SchemeColor,
    CT_ScRgbColor,
    CT_SRgbColor,
    CT_SystemColor,
)


class ColorFormat(object):
    """
    Provides access to color settings such as RGB color, theme color, and
    luminance adjustments.
    """

    def __init__(self, eg_colorChoice_parent, color):
        super(ColorFormat, self).__init__()
        self._xFill = eg_colorChoice_parent
        self._color = color

    @property
    def brightness(self):
        """
        Read/write float value between -1.0 and 1.0 indicating the brightness
        adjustment for this color, e.g. -0.25 is 25% darker and 0.4 is 40%
        lighter. 0 means no brightness adjustment.
        """
        pass

    @brightness.setter
    def brightness(self, value):
        pass

    @classmethod
    def from_colorchoice_parent(cls, eg_colorChoice_parent):
        pass

    @property
    def rgb(self):
        """
        |RGBColor| value of this color, or None if no RGB color is explicitly
        defined for this font. Setting this value to an |RGBColor| instance
        causes its type to change to MSO_COLOR_TYPE.RGB. If the color was a
        theme color with a brightness adjustment, the brightness adjustment
        is removed when changing it to an RGB color.
        """
        pass

    @rgb.setter
    def rgb(self, rgb):
        pass

    @property
    def theme_color(self):
        """Theme color value of this color.

        Value is a member of :ref:`MsoThemeColorIndex`, e.g.
        ``MSO_THEME_COLOR.ACCENT_1``. Raises AttributeError on access if the
        color is not type ``MSO_COLOR_TYPE.SCHEME``. Assigning a member of
        :ref:`MsoThemeColorIndex` causes the color's type to change to
        ``MSO_COLOR_TYPE.SCHEME``.
        """
        pass

    @theme_color.setter
    def theme_color(self, mso_theme_color_idx):
        # change to theme color format if not already
        pass

    @property
    def type(self):
        """
        Read-only. A value from :ref:`MsoColorType`, either RGB or SCHEME,
        corresponding to the way this color is defined, or None if no color
        is defined at the level of this font.
        """
        pass

    def _validate_brightness_value(self, value):
        pass


class _Color(object):
    """
    Object factory for color object of the appropriate type, also the base
    class for all color type classes such as SRgbColor.
    """

    def __new__(cls, xClr):
        color_cls = {
            type(None): _NoneColor,
            CT_HslColor: _HslColor,
            CT_PresetColor: _PrstColor,
            CT_SchemeColor: _SchemeColor,
            CT_ScRgbColor: _ScRgbColor,
            CT_SRgbColor: _SRgbColor,
            CT_SystemColor: _SysColor,
        }[type(xClr)]
        return super(_Color, cls).__new__(color_cls)

    def __init__(self, xClr):
        super(_Color, self).__init__()
        self._xClr = xClr

    @property
    def brightness(self):
        pass

    @brightness.setter
    def brightness(self, value):
        pass

    @property
    def color_type(self):  # pragma: no cover
        pass

    @property
    def rgb(self):
        """
        Raises TypeError on access unless overridden by subclass.
        """
        pass

    @property
    def theme_color(self):
        """
        Raises TypeError on access unless overridden by subclass.
        """
        pass

    def _shade(self, value):
        pass

    def _tint(self, value):
        pass


class _HslColor(_Color):
    @property
    def color_type(self):
        pass


class _NoneColor(_Color):
    @property
    def color_type(self):
        pass

    @property
    def theme_color(self):
        """
        Raise TypeError on attempt to access .theme_color when no color
        choice is present.
        """
        pass


class _PrstColor(_Color):
    @property
    def color_type(self):
        pass


class _SchemeColor(_Color):
    def __init__(self, schemeClr):
        super(_SchemeColor, self).__init__(schemeClr)
        self._schemeClr = schemeClr

    @property
    def color_type(self):
        pass

    @property
    def theme_color(self):
        """
        Theme color value of this color, one of those defined in the
        MSO_THEME_COLOR enumeration, e.g. MSO_THEME_COLOR.ACCENT_1. None if
        no theme color is explicitly defined for this font. Setting this to a
        value in MSO_THEME_COLOR causes the color's type to change to
        ``MSO_COLOR_TYPE.SCHEME``.
        """
        pass

    @theme_color.setter
    def theme_color(self, mso_theme_color_idx):
        pass


class _ScRgbColor(_Color):
    @property
    def color_type(self):
        pass


class _SRgbColor(_Color):
    def __init__(self, srgbClr):
        super(_SRgbColor, self).__init__(srgbClr)
        self._srgbClr = srgbClr

    @property
    def color_type(self):
        pass

    @property
    def rgb(self):
        """
        |RGBColor| value of this color, corresponding to the value in the
        required ``val`` attribute of the ``<a:srgbColr>`` element.
        """
        pass

    @rgb.setter
    def rgb(self, rgb):
        pass


class _SysColor(_Color):
    @property
    def color_type(self):
        pass


class RGBColor(tuple):
    """
    Immutable value object defining a particular RGB color.
    """

    def __new__(cls, r, g, b):
        msg = "RGBColor() takes three integer values 0-255"
        for val in (r, g, b):
            if not isinstance(val, int) or val < 0 or val > 255:
                raise ValueError(msg)
        return super(RGBColor, cls).__new__(cls, (r, g, b))

    def __str__(self):
        """
        Return a hex string rgb value, like '3C2F80'
        """
        return "%02X%02X%02X" % self

    @classmethod
    def from_string(cls, rgb_hex_str):
        """
        Return a new instance from an RGB color hex string like ``'3C2F80'``.
        """
        pass
