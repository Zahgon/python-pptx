"""Connector (line) shape and related objects.

A connector is a line shape having end-points that can be connected to other
objects (but not to other connectors). A connector can be straight, have
elbows, or can be curved.
"""

from __future__ import annotations

from pptx.dml.line import LineFormat
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.base import BaseShape
from pptx.util import Emu, lazyproperty


class Connector(BaseShape):
    """Connector (line) shape.

    A connector is a linear shape having end-points that can be connected to
    other objects (but not to other connectors). A connector can be straight,
    have elbows, or can be curved.
    """

    def begin_connect(self, shape, cxn_pt_idx):
        """
        **EXPERIMENTAL** - *The current implementation only works properly
        with rectangular shapes, such as pictures and rectangles. Use with
        other shape types may cause unexpected visual alignment of the
        connected end-point and could lead to a load error if cxn_pt_idx
        exceeds the connection point count available on the connected shape.
        That said, a quick test should reveal what to expect when using this
        method with other shape types.*

        Connect the beginning of this connector to *shape* at the connection
        point specified by *cxn_pt_idx*. Each shape has zero or more
        connection points and they are identified by index, starting with 0.
        Generally, the first connection point of a shape is at the top center
        of its bounding box and numbering proceeds counter-clockwise from
        there. However this is only a convention and may vary, especially
        with non built-in shapes.
        """
        pass

    @property
    def begin_x(self):
        """
        Return the X-position of the begin point of this connector, in
        English Metric Units (as a |Length| object).
        """
        pass

    @begin_x.setter
    def begin_x(self, value):
        pass

    @property
    def begin_y(self):
        """
        Return the Y-position of the begin point of this connector, in
        English Metric Units (as a |Length| object).
        """
        pass

    @begin_y.setter
    def begin_y(self, value):
        pass

    def end_connect(self, shape, cxn_pt_idx):
        """
        **EXPERIMENTAL** - *The current implementation only works properly
        with rectangular shapes, such as pictures and rectangles. Use with
        other shape types may cause unexpected visual alignment of the
        connected end-point and could lead to a load error if cxn_pt_idx
        exceeds the connection point count available on the connected shape.
        That said, a quick test should reveal what to expect when using this
        method with other shape types.*

        Connect the ending of this connector to *shape* at the connection
        point specified by *cxn_pt_idx*.
        """
        pass

    @property
    def end_x(self):
        """
        Return the X-position of the end point of this connector, in English
        Metric Units (as a |Length| object).
        """
        pass

    @end_x.setter
    def end_x(self, value):
        pass

    @property
    def end_y(self):
        """
        Return the Y-position of the end point of this connector, in English
        Metric Units (as a |Length| object).
        """
        pass

    @end_y.setter
    def end_y(self, value):
        pass

    def get_or_add_ln(self):
        """Helper method required by |LineFormat|."""
        pass

    @lazyproperty
    def line(self):
        """|LineFormat| instance for this connector.

        Provides access to line properties such as line color, width, and
        line style.
        """
        pass

    @property
    def ln(self):
        """Helper method required by |LineFormat|.

        The ``<a:ln>`` element containing the line format properties such as
        line color and width. |None| if no `<a:ln>` element is present.
        """
        pass

    @property
    def shape_type(self):
        """Member of `MSO_SHAPE_TYPE` identifying the type of this shape.

        Unconditionally `MSO_SHAPE_TYPE.LINE` for a `Connector` object.
        """
        pass

    def _connect_begin_to(self, shape, cxn_pt_idx):
        """
        Add or update a stCxn element for this connector that connects its
        begin point to the connection point of *shape* specified by
        *cxn_pt_idx*.
        """
        pass

    def _connect_end_to(self, shape, cxn_pt_idx):
        """
        Add or update an endCxn element for this connector that connects its
        end point to the connection point of *shape* specified by
        *cxn_pt_idx*.
        """
        pass

    def _move_begin_to_cxn(self, shape, cxn_pt_idx):
        """
        Move the begin point of this connector to coordinates of the
        connection point of *shape* specified by *cxn_pt_idx*.
        """
        pass

    def _move_end_to_cxn(self, shape, cxn_pt_idx):
        """
        Move the end point of this connector to the coordinates of the
        connection point of *shape* specified by *cxn_pt_idx*.
        """
        pass
