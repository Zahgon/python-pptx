"""DrawingML objects related to fill."""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING

from pptx.dml.color import ColorFormat
from pptx.enum.dml import MSO_FILL
from pptx.oxml.dml.fill import (
    CT_BlipFillProperties,
    CT_GradientFillProperties,
    CT_GroupFillProperties,
    CT_NoFillProperties,
    CT_PatternFillProperties,
    CT_SolidColorFillProperties,
)
from pptx.oxml.xmlchemy import BaseOxmlElement
from pptx.shared import ElementProxy
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.enum.dml import MSO_FILL_TYPE
    from pptx.oxml.xmlchemy import BaseOxmlElement


class FillFormat(object):
    """Provides access to the current fill properties.

    Also provides methods to change the fill type.
    """

    def __init__(self, eg_fill_properties_parent: BaseOxmlElement, fill_obj: _Fill):
        super(FillFormat, self).__init__()
        self._xPr = eg_fill_properties_parent
        self._fill = fill_obj

    @classmethod
    def from_fill_parent(cls, eg_fillProperties_parent: BaseOxmlElement) -> FillFormat:
        """
        Return a |FillFormat| instance initialized to the settings contained
        in *eg_fillProperties_parent*, which must be an element having
        EG_FillProperties in its child element sequence in the XML schema.
        """
        pass

    @property
    def back_color(self):
        """Return a |ColorFormat| object representing background color.

        This property is only applicable to pattern fills and lines.
        """
        pass

    def background(self):
        """
        Sets the fill type to noFill, i.e. transparent.
        """
        pass

    @property
    def fore_color(self):
        """
        Return a |ColorFormat| instance representing the foreground color of
        this fill.
        """
        pass

    def gradient(self):
        """Sets the fill type to gradient.

        If the fill is not already a gradient, a default gradient is added.
        The default gradient corresponds to the default in the built-in
        PowerPoint "White" template. This gradient is linear at angle
        90-degrees (upward), with two stops. The first stop is Accent-1 with
        tint 100%, shade 100%, and satMod 130%. The second stop is Accent-1
        with tint 50%, shade 100%, and satMod 350%.
        """
        pass

    @property
    def gradient_angle(self):
        """Angle in float degrees of line of a linear gradient.

        Read/Write. May be |None|, indicating the angle should be inherited
        from the style hierarchy. An angle of 0.0 corresponds to
        a left-to-right gradient. Increasing angles represent
        counter-clockwise rotation of the line, for example 90.0 represents
        a bottom-to-top gradient. Raises |TypeError| when the fill type is
        not MSO_FILL_TYPE.GRADIENT. Raises |ValueError| for a non-linear
        gradient (e.g. a radial gradient).
        """
        pass

    @gradient_angle.setter
    def gradient_angle(self, value):
        pass

    @property
    def gradient_stops(self):
        """|GradientStops| object providing access to stops of this gradient.

        Raises |TypeError| when fill is not gradient (call `fill.gradient()`
        first). Each stop represents a color between which the gradient
        smoothly transitions.
        """
        pass

    @property
    def pattern(self):
        """Return member of :ref:`MsoPatternType` indicating fill pattern.

        Raises |TypeError| when fill is not patterned (call
        `fill.patterned()` first). Returns |None| if no pattern has been set;
        PowerPoint may display the default `PERCENT_5` pattern in this case.
        Assigning |None| will remove any explicit pattern setting, although
        relying on the default behavior is discouraged and may produce
        rendering differences across client applications.
        """
        pass

    @pattern.setter
    def pattern(self, pattern_type):
        pass

    def patterned(self):
        """Selects the pattern fill type.

        Note that calling this method does not by itself set a foreground or
        background color of the pattern. Rather it enables subsequent
        assignments to properties like fore_color to set the pattern and
        colors.
        """
        pass

    def solid(self):
        """
        Sets the fill type to solid, i.e. a solid color. Note that calling
        this method does not set a color or by itself cause the shape to
        appear with a solid color fill; rather it enables subsequent
        assignments to properties like fore_color to set the color.
        """
        pass

    @property
    def type(self) -> MSO_FILL_TYPE:
        """The type of this fill, e.g. `MSO_FILL_TYPE.SOLID`."""
        pass


class _Fill(object):
    """
    Object factory for fill object of class matching fill element, such as
    _SolidFill for ``<a:solidFill>``; also serves as the base class for all
    fill classes
    """

    def __new__(cls, xFill):
        if xFill is None:
            fill_cls = _NoneFill
        elif isinstance(xFill, CT_BlipFillProperties):
            fill_cls = _BlipFill
        elif isinstance(xFill, CT_GradientFillProperties):
            fill_cls = _GradFill
        elif isinstance(xFill, CT_GroupFillProperties):
            fill_cls = _GrpFill
        elif isinstance(xFill, CT_NoFillProperties):
            fill_cls = _NoFill
        elif isinstance(xFill, CT_PatternFillProperties):
            fill_cls = _PattFill
        elif isinstance(xFill, CT_SolidColorFillProperties):
            fill_cls = _SolidFill
        else:
            fill_cls = _Fill
        return super(_Fill, cls).__new__(fill_cls)

    @property
    def back_color(self):
        """Raise TypeError for types that do not override this property."""
        pass

    @property
    def fore_color(self):
        """Raise TypeError for types that do not override this property."""
        pass

    @property
    def pattern(self):
        """Raise TypeError for fills that do not override this property."""
        pass

    @property
    def type(self) -> MSO_FILL_TYPE:  # pragma: no cover
        raise NotImplementedError(
            f".type property must be implemented on {self.__class__.__name__}"
        )


class _BlipFill(_Fill):
    @property
    def type(self):
        pass


class _GradFill(_Fill):
    """Proxies an `a:gradFill` element."""

    def __init__(self, gradFill):
        self._element = self._gradFill = gradFill

    @property
    def gradient_angle(self):
        """Angle in float degrees of line of a linear gradient.

        Read/Write. May be |None|, indicating the angle is inherited from the
        style hierarchy. An angle of 0.0 corresponds to a left-to-right
        gradient. Increasing angles represent clockwise rotation of the line,
        for example 90.0 represents a top-to-bottom gradient. Raises
        |TypeError| when the fill type is not MSO_FILL_TYPE.GRADIENT. Raises
        |ValueError| for a non-linear gradient (e.g. a radial gradient).
        """
        pass

    @gradient_angle.setter
    def gradient_angle(self, value):
        pass

    @lazyproperty
    def gradient_stops(self):
        """|_GradientStops| object providing access to gradient colors.

        Each stop represents a color between which the gradient smoothly
        transitions.
        """
        pass

    @property
    def type(self):
        pass


class _GrpFill(_Fill):
    @property
    def type(self):
        pass


class _NoFill(_Fill):
    @property
    def type(self):
        pass


class _NoneFill(_Fill):
    @property
    def type(self):
        pass


class _PattFill(_Fill):
    """Provides access to patterned fill properties."""

    def __init__(self, pattFill):
        super(_PattFill, self).__init__()
        self._element = self._pattFill = pattFill

    @lazyproperty
    def back_color(self):
        """Return |ColorFormat| object that controls background color."""
        pass

    @lazyproperty
    def fore_color(self):
        """Return |ColorFormat| object that controls foreground color."""
        pass

    @property
    def pattern(self):
        """Return member of :ref:`MsoPatternType` indicating fill pattern.

        Returns |None| if no pattern has been set; PowerPoint may display the
        default `PERCENT_5` pattern in this case. Assigning |None| will
        remove any explicit pattern setting.
        """
        pass

    @pattern.setter
    def pattern(self, pattern_type):
        pass

    @property
    def type(self):
        pass


class _SolidFill(_Fill):
    """Provides access to fill properties such as color for solid fills."""

    def __init__(self, solidFill):
        super(_SolidFill, self).__init__()
        self._solidFill = solidFill

    @lazyproperty
    def fore_color(self):
        """Return |ColorFormat| object controlling fill color."""
        pass

    @property
    def type(self):
        pass


class _GradientStops(Sequence):
    """Collection of |GradientStop| objects defining gradient colors.

    A gradient must have a minimum of two stops, but can have as many more
    than that as required to achieve the desired effect (three is perhaps
    most common). Stops are sequenced in the order they are transitioned
    through.
    """

    def __init__(self, gsLst):
        self._gsLst = gsLst

    def __getitem__(self, idx):
        return _GradientStop(self._gsLst[idx])

    def __len__(self):
        return len(self._gsLst)


class _GradientStop(ElementProxy):
    """A single gradient stop.

    A gradient stop defines a color and a position.
    """

    def __init__(self, gs):
        super(_GradientStop, self).__init__(gs)
        self._gs = gs

    @lazyproperty
    def color(self):
        """Return |ColorFormat| object controlling stop color."""
        pass

    @property
    def position(self):
        """Location of stop in gradient path as float between 0.0 and 1.0.

        The value represents a percentage, where 0.0 (0%) represents the
        start of the path and 1.0 (100%) represents the end of the path. For
        a linear gradient, these would represent opposing extents of the
        filled area.
        """
        pass

    @position.setter
    def position(self, value):
        pass
