# pyright: reportPrivateUsage=false

"""lxml custom element classes for shape-related XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.shapes.shared import BaseShapeElement
from pptx.oxml.simpletypes import (
    ST_Coordinate,
    ST_PositiveCoordinate,
    XsdBoolean,
    XsdString,
)
from pptx.oxml.text import CT_TextBody
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from pptx.oxml.shapes.shared import (
        CT_ApplicationNonVisualDrawingProps,
        CT_NonVisualDrawingProps,
        CT_ShapeProperties,
    )
    from pptx.util import Length


class CT_AdjPoint2D(BaseOxmlElement):
    """`a:pt` custom element class."""

    x: Length = RequiredAttribute("x", ST_Coordinate)  # pyright: ignore[reportAssignmentType]
    y: Length = RequiredAttribute("y", ST_Coordinate)  # pyright: ignore[reportAssignmentType]


class CT_CustomGeometry2D(BaseOxmlElement):
    """`a:custGeom` custom element class."""

    get_or_add_pathLst: Callable[[], CT_Path2DList]

    _tag_seq = ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst", "a:rect", "a:pathLst")
    pathLst: CT_Path2DList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:pathLst", successors=_tag_seq[6:]
    )


class CT_GeomGuide(BaseOxmlElement):
    """`a:gd` custom element class.

    Defines a "guide", corresponding to a yellow diamond-shaped handle on an autoshape.
    """

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    fmla: str = RequiredAttribute("fmla", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_GeomGuideList(BaseOxmlElement):
    """`a:avLst` custom element class."""

    _add_gd: Callable[[], CT_GeomGuide]

    gd_lst: list[CT_GeomGuide]

    gd = ZeroOrMore("a:gd")


class CT_NonVisualDrawingShapeProps(BaseShapeElement):
    """`p:cNvSpPr` custom element class."""

    spLocks = ZeroOrOne("a:spLocks")
    txBox: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "txBox", XsdBoolean
    )


class CT_Path2D(BaseOxmlElement):
    """`a:path` custom element class."""

    _add_close: Callable[[], CT_Path2DClose]
    _add_lnTo: Callable[[], CT_Path2DLineTo]
    _add_moveTo: Callable[[], CT_Path2DMoveTo]

    close = ZeroOrMore("a:close", successors=())
    lnTo = ZeroOrMore("a:lnTo", successors=())
    moveTo = ZeroOrMore("a:moveTo", successors=())
    w: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w", ST_PositiveCoordinate
    )
    h: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "h", ST_PositiveCoordinate
    )

    def add_close(self) -> CT_Path2DClose:
        """Return a newly created `a:close` element.

        The new `a:close` element is appended to this `a:path` element.
        """
        pass

    def add_lnTo(self, x: Length, y: Length) -> CT_Path2DLineTo:
        """Return a newly created `a:lnTo` subtree with end point *(x, y)*.

        The new `a:lnTo` element is appended to this `a:path` element.
        """
        pass

    def add_moveTo(self, x: Length, y: Length):
        """Return a newly created `a:moveTo` subtree with point `(x, y)`.

        The new `a:moveTo` element is appended to this `a:path` element.
        """
        pass


class CT_Path2DClose(BaseOxmlElement):
    """`a:close` custom element class."""


class CT_Path2DLineTo(BaseOxmlElement):
    """`a:lnTo` custom element class."""

    _add_pt: Callable[[], CT_AdjPoint2D]

    pt = ZeroOrOne("a:pt", successors=())


class CT_Path2DList(BaseOxmlElement):
    """`a:pathLst` custom element class."""

    _add_path: Callable[[], CT_Path2D]

    path = ZeroOrMore("a:path", successors=())

    def add_path(self, w: Length, h: Length):
        """Return a newly created `a:path` child element."""
        pass


class CT_Path2DMoveTo(BaseOxmlElement):
    """`a:moveTo` custom element class."""

    _add_pt: Callable[[], CT_AdjPoint2D]

    pt = ZeroOrOne("a:pt", successors=())


class CT_PresetGeometry2D(BaseOxmlElement):
    """`a:prstGeom` custom element class."""

    _add_avLst: Callable[[], CT_GeomGuideList]
    _remove_avLst: Callable[[], None]

    avLst: CT_GeomGuideList | None = ZeroOrOne("a:avLst")  # pyright: ignore[reportAssignmentType]
    prst: MSO_AUTO_SHAPE_TYPE = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "prst", MSO_AUTO_SHAPE_TYPE
    )

    @property
    def gd_lst(self) -> list[CT_GeomGuide]:
        """Sequence of `a:gd` element children of `a:avLst`. Empty if none are present."""
        pass

    def rewrite_guides(self, guides: list[tuple[str, int]]):
        """Replace any `a:gd` element children of `a:avLst` with ones forme from `guides`."""
        pass


class CT_Shape(BaseShapeElement):
    """`p:sp` custom element class."""

    get_or_add_txBody: Callable[[], CT_TextBody]

    nvSpPr: CT_ShapeNonVisual = OneAndOnlyOne("p:nvSpPr")  # pyright: ignore[reportAssignmentType]
    spPr: CT_ShapeProperties = OneAndOnlyOne("p:spPr")  # pyright: ignore[reportAssignmentType]
    txBody: CT_TextBody | None = ZeroOrOne("p:txBody", successors=("p:extLst",))  # pyright: ignore

    def add_path(self, w: Length, h: Length) -> CT_Path2D:
        pass

    def get_or_add_ln(self):
        """Return the `a:ln` grandchild element, newly added if not present."""
        pass

    @property
    def has_custom_geometry(self):
        """True if this shape has custom geometry, i.e. is a freeform shape.

        A shape has custom geometry if it has a `p:spPr/a:custGeom`
        descendant (instead of `p:spPr/a:prstGeom`).
        """
        pass

    @property
    def is_autoshape(self):
        """True if this shape is an auto shape.

        A shape is an auto shape if it has a `a:prstGeom` element and does not have a txBox="1"
        attribute on cNvSpPr.
        """
        pass

    @property
    def is_textbox(self):
        """True if this shape is a text box.

        A shape is a text box if it has a `txBox` attribute on cNvSpPr that resolves to |True|.
        The default when the txBox attribute is missing is |False|.
        """
        pass

    @property
    def ln(self):
        """`a:ln` grand-child element or |None| if not present."""
        pass

    @staticmethod
    def new_autoshape_sp(
        id_: int, name: str, prst: str, left: int, top: int, width: int, height: int
    ) -> CT_Shape:
        """Return a new `p:sp` element tree configured as a base auto shape."""
        pass

    @staticmethod
    def new_freeform_sp(shape_id: int, name: str, x: int, y: int, cx: int, cy: int):
        """Return new `p:sp` element tree configured as freeform shape.

        The returned shape has a `a:custGeom` subtree but no paths in its
        path list.
        """
        pass

    @staticmethod
    def new_placeholder_sp(
        id_: int, name: str, ph_type: PP_PLACEHOLDER, orient: str, sz, idx
    ) -> CT_Shape:
        """Return a new `p:sp` element tree configured as a placeholder shape."""
        sp = cast(
            CT_Shape,
            parse_xml(
                f"<p:sp {nsdecls('a', 'p')}>\n"
                f"  <p:nvSpPr>\n"
                f'    <p:cNvPr id="{id_}" name="{name}"/>\n'
                f"    <p:cNvSpPr>\n"
                f'      <a:spLocks noGrp="1"/>\n'
                f"    </p:cNvSpPr>\n"
                f"    <p:nvPr/>\n"
                f"  </p:nvSpPr>\n"
                f"  <p:spPr/>\n"
                f"</p:sp>"
            ),
        )

        ph = sp.nvSpPr.nvPr.get_or_add_ph()
        ph.type = ph_type
        ph.idx = idx
        ph.orient = orient
        ph.sz = sz

        placeholder_types_that_have_a_text_frame = (
            PP_PLACEHOLDER.TITLE,
            PP_PLACEHOLDER.CENTER_TITLE,
            PP_PLACEHOLDER.SUBTITLE,
            PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.OBJECT,
        )

        if ph_type in placeholder_types_that_have_a_text_frame:
            sp.append(CT_TextBody.new())

        return sp

    @staticmethod
    def new_textbox_sp(id_, name, left, top, width, height):
        """Return a new `p:sp` element tree configured as a base textbox shape."""
        pass

    @property
    def prst(self):
        """Value of `prst` attribute of `a:prstGeom` element or |None| if not present."""
        pass

    @property
    def prstGeom(self) -> CT_PresetGeometry2D:
        """Reference to `a:prstGeom` child element.

        |None| if this shape doesn't have one, for example, if it's a placeholder shape.
        """
        pass

    def _new_txBody(self):
        pass

    @staticmethod
    def _textbox_sp_tmpl():
        pass


class CT_ShapeNonVisual(BaseShapeElement):
    """`p:nvSpPr` custom element class."""

    cNvPr: CT_NonVisualDrawingProps = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cNvPr"
    )
    cNvSpPr: CT_NonVisualDrawingShapeProps = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cNvSpPr"
    )
    nvPr: CT_ApplicationNonVisualDrawingProps = (  # pyright: ignore[reportAssignmentType]
        OneAndOnlyOne("p:nvPr")
    )
