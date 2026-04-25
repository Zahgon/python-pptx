"""Objects related to construction of freeform shapes."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterable, Iterator, Sequence

from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from typing_extensions import TypeAlias

    from pptx.oxml.shapes.autoshape import (
        CT_Path2D,
        CT_Path2DClose,
        CT_Path2DLineTo,
        CT_Path2DMoveTo,
        CT_Shape,
    )
    from pptx.shapes.shapetree import _BaseGroupShapes  # pyright: ignore[reportPrivateUsage]
    from pptx.util import Length

CT_DrawingOperation: TypeAlias = "CT_Path2DClose | CT_Path2DLineTo | CT_Path2DMoveTo"
DrawingOperation: TypeAlias = "_LineSegment | _MoveTo | _Close"


class FreeformBuilder(Sequence[DrawingOperation]):
    """Allows a freeform shape to be specified and created.

    The initial pen position is provided on construction. From there, drawing proceeds using
    successive calls to draw line segments. The freeform shape may be closed by calling the
    :meth:`close` method.

    A shape may have more than one contour, in which case overlapping areas are "subtracted". A
    contour is a sequence of line segments beginning with a "move-to" operation. A move-to
    operation is automatically inserted in each new freeform; additional move-to ops can be
    inserted with the `.move_to()` method.
    """

    def __init__(
        self,
        shapes: _BaseGroupShapes,
        start_x: Length,
        start_y: Length,
        x_scale: float,
        y_scale: float,
    ):
        super(FreeformBuilder, self).__init__()
        self._shapes = shapes
        self._start_x = start_x
        self._start_y = start_y
        self._x_scale = x_scale
        self._y_scale = y_scale

    def __getitem__(  # pyright: ignore[reportIncompatibleMethodOverride]
        self, idx: int
    ) -> DrawingOperation:
        return self._drawing_operations.__getitem__(idx)

    def __iter__(self) -> Iterator[DrawingOperation]:
        return self._drawing_operations.__iter__()

    def __len__(self):
        return self._drawing_operations.__len__()

    @classmethod
    def new(
        cls,
        shapes: _BaseGroupShapes,
        start_x: float,
        start_y: float,
        x_scale: float,
        y_scale: float,
    ):
        """Return a new |FreeformBuilder| object.

        The initial pen location is specified (in local coordinates) by
        (`start_x`, `start_y`).
        """
        return cls(shapes, Emu(int(round(start_x))), Emu(int(round(start_y))), x_scale, y_scale)

    def add_line_segments(self, vertices: Iterable[tuple[float, float]], close: bool = True):
        """Add a straight line segment to each point in `vertices`.

        `vertices` must be an iterable of (x, y) pairs (2-tuples). Each x and y value is rounded
        to the nearest integer before use. The optional `close` parameter determines whether the
        resulting contour is `closed` or left `open`.

        Returns this |FreeformBuilder| object so it can be used in chained calls.
        """
        pass

    def convert_to_shape(self, origin_x: Length = Emu(0), origin_y: Length = Emu(0)):
        """Return new freeform shape positioned relative to specified offset.

        `origin_x` and `origin_y` locate the origin of the local coordinate system in slide
        coordinates (EMU), perhaps most conveniently by use of a |Length| object.

        Note that this method may be called more than once to add multiple shapes of the same
        geometry in different locations on the slide.
        """
        pass

    def move_to(self, x: float, y: float):
        """Move pen to (x, y) (local coordinates) without drawing line.

        Returns this |FreeformBuilder| object so it can be used in chained calls.
        """
        pass

    @property
    def shape_offset_x(self) -> Length:
        """Return x distance of shape origin from local coordinate origin.

        The returned integer represents the leftmost extent of the freeform shape, in local
        coordinates. Note that the bounding box of the shape need not start at the local origin.
        """
        pass

    @property
    def shape_offset_y(self) -> Length:
        """Return y distance of shape origin from local coordinate origin.

        The returned integer represents the topmost extent of the freeform shape, in local
        coordinates. Note that the bounding box of the shape need not start at the local origin.
        """
        pass

    def _add_close(self):
        """Add a close |_Close| operation to the drawing sequence."""
        pass

    def _add_freeform_sp(self, origin_x: Length, origin_y: Length):
        """Add a freeform `p:sp` element having no drawing elements.

        `origin_x` and `origin_y` are specified in slide coordinates, and represent the location
        of the local coordinates origin on the slide.
        """
        pass

    def _add_line_segment(self, x: float, y: float) -> None:
        """Add a |_LineSegment| operation to the drawing sequence."""
        pass

    @lazyproperty
    def _drawing_operations(self) -> list[DrawingOperation]:
        """Return the sequence of drawing operation objects for freeform."""
        pass

    @property
    def _dx(self) -> Length:
        """Return width of this shape's path in local units."""
        pass

    @property
    def _dy(self) -> Length:
        """Return integer height of this shape's path in local units."""
        pass

    @property
    def _height(self):
        """Return vertical size of this shape's path in slide coordinates.

        This value is based on the actual extents of the shape and does not include any
        positioning offset.
        """
        pass

    @property
    def _left(self):
        """Return leftmost extent of this shape's path in slide coordinates.

        Note that this value does not include any positioning offset; it assumes the drawing
        (local) coordinate origin is at (0, 0) on the slide.
        """
        pass

    def _local_to_shape(self, local_x: Length, local_y: Length) -> tuple[Length, Length]:
        """Translate local coordinates point to shape coordinates.

        Shape coordinates have the same unit as local coordinates, but are offset such that the
        origin of the shape coordinate system (0, 0) is located at the top-left corner of the
        shape bounding box.
        """
        pass

    def _start_path(self, sp: CT_Shape) -> CT_Path2D:
        """Return a newly created `a:path` element added to `sp`.

        The returned `a:path` element has an `a:moveTo` element representing the shape starting
        point as its only child.
        """
        pass

    @property
    def _top(self):
        """Return topmost extent of this shape's path in slide coordinates.

        Note that this value does not include any positioning offset; it assumes the drawing
        (local) coordinate origin is located at slide coordinates (0, 0) (top-left corner of
        slide).
        """
        pass

    @property
    def _width(self):
        """Return width of this shape's path in slide coordinates.

        This value is based on the actual extents of the shape path and does not include any
        positioning offset.
        """
        pass


class _BaseDrawingOperation(object):
    """Base class for freeform drawing operations.

    A drawing operation has at least one location (x, y) in local coordinates.
    """

    def __init__(self, freeform_builder: FreeformBuilder, x: Length, y: Length):
        super(_BaseDrawingOperation, self).__init__()
        self._freeform_builder = freeform_builder
        self._x = x
        self._y = y

    def apply_operation_to(self, path: CT_Path2D) -> CT_DrawingOperation:
        """Add the XML element(s) implementing this operation to `path`.

        Must be implemented by each subclass.
        """
        raise NotImplementedError("must be implemented by each subclass")

    @property
    def x(self) -> Length:
        """Return the horizontal (x) target location of this operation.

        The returned value is an integer in local coordinates.
        """
        pass

    @property
    def y(self) -> Length:
        """Return the vertical (y) target location of this operation.

        The returned value is an integer in local coordinates.
        """
        pass


class _Close(object):
    """Specifies adding a `<a:close/>` element to the current contour."""

    @classmethod
    def new(cls) -> _Close:
        """Return a new _Close object."""
        return cls()

    def apply_operation_to(self, path: CT_Path2D) -> CT_Path2DClose:
        """Add `a:close` element to `path`."""
        pass


class _LineSegment(_BaseDrawingOperation):
    """Specifies a straight line segment ending at the specified point."""

    @classmethod
    def new(cls, freeform_builder: FreeformBuilder, x: float, y: float) -> _LineSegment:
        """Return a new _LineSegment object ending at point *(x, y)*.

        Both `x` and `y` are rounded to the nearest integer before use.
        """
        return cls(freeform_builder, Emu(int(round(x))), Emu(int(round(y))))

    def apply_operation_to(self, path: CT_Path2D) -> CT_Path2DLineTo:
        """Add `a:lnTo` element to `path` for this line segment.

        Returns the `a:lnTo` element newly added to the path.
        """
        pass


class _MoveTo(_BaseDrawingOperation):
    """Specifies a new pen position."""

    @classmethod
    def new(cls, freeform_builder: FreeformBuilder, x: float, y: float) -> _MoveTo:
        """Return a new _MoveTo object for move to point `(x, y)`.

        Both `x` and `y` are rounded to the nearest integer before use.
        """
        return cls(freeform_builder, Emu(int(round(x))), Emu(int(round(y))))

    def apply_operation_to(self, path: CT_Path2D) -> CT_Path2DMoveTo:
        """Add `a:moveTo` element to `path` for this line segment."""
        pass
