"""Common shape-related oxml objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

from pptx.dml.fill import CT_GradientFillProperties
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import (
    ST_Angle,
    ST_Coordinate,
    ST_Direction,
    ST_DrawingElementId,
    ST_LineWidth,
    ST_PlaceholderSize,
    ST_PositiveCoordinate,
    XsdBoolean,
    XsdString,
    XsdUnsignedInt,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OptionalAttribute,
    OxmlElement,
    RequiredAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
)
from pptx.util import Emu

if TYPE_CHECKING:
    from pptx.oxml.action import CT_Hyperlink
    from pptx.oxml.shapes.autoshape import CT_CustomGeometry2D, CT_PresetGeometry2D
    from pptx.util import Length


class BaseShapeElement(BaseOxmlElement):
    """Provides common behavior for shape element classes like CT_Shape, CT_Picture, etc."""

    spPr: CT_ShapeProperties

    @property
    def cx(self) -> Length:
        pass

    @cx.setter
    def cx(self, value):
        pass

    @property
    def cy(self) -> Length:
        pass

    @cy.setter
    def cy(self, value):
        pass

    @property
    def flipH(self):
        pass

    @flipH.setter
    def flipH(self, value):
        pass

    @property
    def flipV(self):
        pass

    @flipV.setter
    def flipV(self, value):
        pass

    def get_or_add_xfrm(self):
        """Return the `a:xfrm` grandchild element, newly-added if not present.

        This version works for `p:sp`, `p:cxnSp`, and `p:pic` elements, others will need to
        override.
        """
        pass

    @property
    def has_ph_elm(self):
        """
        True if this shape element has a `p:ph` descendant, indicating it
        is a placeholder shape. False otherwise.
        """
        pass

    @property
    def ph(self) -> CT_Placeholder | None:
        """The `p:ph` descendant element if there is one, None otherwise."""
        pass

    @property
    def ph_idx(self) -> int:
        """Integer value of placeholder idx attribute.

        Raises |ValueError| if shape is not a placeholder.
        """
        pass

    @property
    def ph_orient(self) -> str:
        """Placeholder orientation, e.g. 'vert'.

        Raises |ValueError| if shape is not a placeholder.
        """
        pass

    @property
    def ph_sz(self) -> str:
        """Placeholder size, e.g. ST_PlaceholderSize.HALF.

        Raises `ValueError` if shape is not a placeholder.
        """
        pass

    @property
    def ph_type(self):
        """Placeholder type, e.g. ST_PlaceholderType.TITLE ('title').

        Raises `ValueError` if shape is not a placeholder.
        """
        pass

    @property
    def rot(self) -> float:
        """Float representing degrees this shape is rotated clockwise."""
        pass

    @rot.setter
    def rot(self, value: float):
        pass

    @property
    def shape_id(self):
        """
        Integer id of this shape
        """
        pass

    @property
    def shape_name(self):
        """
        Name of this shape
        """
        pass

    @property
    def txBody(self):
        """Child `p:txBody` element, None if not present."""
        pass

    @property
    def x(self) -> Length:
        pass

    @x.setter
    def x(self, value):
        pass

    @property
    def xfrm(self):
        """The `a:xfrm` grandchild element or |None| if not found.

        This version works for `p:sp`, `p:cxnSp`, and `p:pic` elements, others will need to
        override.
        """
        pass

    @property
    def y(self) -> Length:
        pass

    @y.setter
    def y(self, value):
        pass

    @property
    def _nvXxPr(self):
        """
        Required non-visual shape properties element for this shape. Actual
        name depends on the shape type, e.g. `p:nvPicPr` for picture
        shape.
        """
        pass

    def _get_xfrm_attr(self, name: str) -> Length | None:
        pass

    def _set_xfrm_attr(self, name, value):
        pass


class CT_ApplicationNonVisualDrawingProps(BaseOxmlElement):
    """`p:nvPr` element."""

    get_or_add_ph: Callable[[], CT_Placeholder]

    ph = ZeroOrOne(
        "p:ph",
        successors=(
            "a:audioCd",
            "a:wavAudioFile",
            "a:audioFile",
            "a:videoFile",
            "a:quickTimeFile",
            "p:custDataLst",
            "p:extLst",
        ),
    )


class CT_LineProperties(BaseOxmlElement):
    """Custom element class for <a:ln> element"""

    _tag_seq = (
        "a:noFill",
        "a:solidFill",
        "a:gradFill",
        "a:pattFill",
        "a:prstDash",
        "a:custDash",
        "a:round",
        "a:bevel",
        "a:miter",
        "a:headEnd",
        "a:tailEnd",
        "a:extLst",
    )
    eg_lineFillProperties = ZeroOrOneChoice(
        (
            Choice("a:noFill"),
            Choice("a:solidFill"),
            Choice("a:gradFill"),
            Choice("a:pattFill"),
        ),
        successors=_tag_seq[4:],
    )
    prstDash = ZeroOrOne("a:prstDash", successors=_tag_seq[5:])
    custDash = ZeroOrOne("a:custDash", successors=_tag_seq[6:])
    del _tag_seq
    w = OptionalAttribute("w", ST_LineWidth, default=Emu(0))

    @property
    def eg_fillProperties(self):
        """
        Required to fulfill the interface used by dml.fill.
        """
        pass

    @property
    def prstDash_val(self):
        """Return value of `val` attribute of `a:prstDash` child.

        Return |None| if not present.
        """
        pass

    @prstDash_val.setter
    def prstDash_val(self, val):
        pass


class CT_NonVisualDrawingProps(BaseOxmlElement):
    """`p:cNvPr` custom element class."""

    get_or_add_hlinkClick: Callable[[], CT_Hyperlink]
    get_or_add_hlinkHover: Callable[[], CT_Hyperlink]

    _tag_seq = ("a:hlinkClick", "a:hlinkHover", "a:extLst")
    hlinkClick: CT_Hyperlink | None = ZeroOrOne("a:hlinkClick", successors=_tag_seq[1:])
    hlinkHover: CT_Hyperlink | None = ZeroOrOne("a:hlinkHover", successors=_tag_seq[2:])
    id = RequiredAttribute("id", ST_DrawingElementId)
    name = RequiredAttribute("name", XsdString)
    del _tag_seq


class CT_Placeholder(BaseOxmlElement):
    """`p:ph` custom element class."""

    type: PP_PLACEHOLDER = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "type", PP_PLACEHOLDER, default=PP_PLACEHOLDER.OBJECT
    )
    orient: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "orient", ST_Direction, default=ST_Direction.HORZ
    )
    sz: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "sz", ST_PlaceholderSize, default=ST_PlaceholderSize.FULL
    )
    idx: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "idx", XsdUnsignedInt, default=0
    )


class CT_Point2D(BaseOxmlElement):
    """
    Custom element class for <a:off> element.
    """

    x: Length = RequiredAttribute("x", ST_Coordinate)  # pyright: ignore[reportAssignmentType]
    y: Length = RequiredAttribute("y", ST_Coordinate)  # pyright: ignore[reportAssignmentType]


class CT_PositiveSize2D(BaseOxmlElement):
    """
    Custom element class for <a:ext> element.
    """

    cx = RequiredAttribute("cx", ST_PositiveCoordinate)
    cy = RequiredAttribute("cy", ST_PositiveCoordinate)


class CT_ShapeProperties(BaseOxmlElement):
    """Custom element class for `p:spPr` element.

    Shared by `p:sp`, `p:cxnSp`,  and `p:pic` elements as well as a few more obscure ones.
    """

    get_or_add_xfrm: Callable[[], CT_Transform2D]
    get_or_add_ln: Callable[[], CT_LineProperties]
    _add_prstGeom: Callable[[], CT_PresetGeometry2D]
    _remove_custGeom: Callable[[], None]

    _tag_seq = (
        "a:xfrm",
        "a:custGeom",
        "a:prstGeom",
        "a:noFill",
        "a:solidFill",
        "a:gradFill",
        "a:blipFill",
        "a:pattFill",
        "a:grpFill",
        "a:ln",
        "a:effectLst",
        "a:effectDag",
        "a:scene3d",
        "a:sp3d",
        "a:extLst",
    )
    xfrm: CT_Transform2D | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:xfrm", successors=_tag_seq[1:]
    )
    custGeom: CT_CustomGeometry2D | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:custGeom", successors=_tag_seq[2:]
    )
    prstGeom: CT_PresetGeometry2D | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:prstGeom", successors=_tag_seq[3:]
    )
    eg_fillProperties = ZeroOrOneChoice(
        (
            Choice("a:noFill"),
            Choice("a:solidFill"),
            Choice("a:gradFill"),
            Choice("a:blipFill"),
            Choice("a:pattFill"),
            Choice("a:grpFill"),
        ),
        successors=_tag_seq[9:],
    )
    ln: CT_LineProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:ln", successors=_tag_seq[10:]
    )
    effectLst = ZeroOrOne("a:effectLst", successors=_tag_seq[11:])
    del _tag_seq

    @property
    def cx(self):
        """
        Shape width as an instance of Emu, or None if not present.
        """
        pass

    @property
    def cy(self):
        """
        Shape height as an instance of Emu, or None if not present.
        """
        pass

    @property
    def x(self) -> Length | None:
        """Distance between the left edge of the slide and left edge of the shape.

        0 if not present.
        """
        pass

    @property
    def y(self):
        """
        The offset of the top of the shape from the top of the slide, as an
        instance of Emu. None if not present.
        """
        pass

    def _new_gradFill(self):
        pass


class CT_Transform2D(BaseOxmlElement):
    """`a:xfrm` custom element class.

    NOTE: this is a composite including CT_GroupTransform2D, which appears
    with the `a:xfrm` tag in a group shape (including a slide `p:spTree`).
    """

    _tag_seq = ("a:off", "a:ext", "a:chOff", "a:chExt")
    off: CT_Point2D | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:off", successors=_tag_seq[1:]
    )
    ext = ZeroOrOne("a:ext", successors=_tag_seq[2:])
    chOff = ZeroOrOne("a:chOff", successors=_tag_seq[3:])
    chExt = ZeroOrOne("a:chExt", successors=_tag_seq[4:])
    del _tag_seq
    rot: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rot", ST_Angle, default=0.0
    )
    flipH = OptionalAttribute("flipH", XsdBoolean, default=False)
    flipV = OptionalAttribute("flipV", XsdBoolean, default=False)

    @property
    def x(self):
        pass

    @x.setter
    def x(self, value):
        pass

    @property
    def y(self):
        pass

    @y.setter
    def y(self, value):
        pass

    @property
    def cx(self):
        pass

    @cx.setter
    def cx(self, value):
        pass

    @property
    def cy(self):
        pass

    @cy.setter
    def cy(self, value):
        pass

    def _new_ext(self):
        pass

    def _new_off(self):
        pass
