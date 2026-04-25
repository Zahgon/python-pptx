"""lxml custom element classes for shape-tree-related XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator

from pptx.enum.shapes import MSO_CONNECTOR_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.shapes.autoshape import CT_Shape
from pptx.oxml.shapes.connector import CT_Connector
from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectFrame
from pptx.oxml.shapes.picture import CT_Picture
from pptx.oxml.shapes.shared import BaseShapeElement
from pptx.oxml.xmlchemy import BaseOxmlElement, OneAndOnlyOne, ZeroOrOne
from pptx.util import Emu

if TYPE_CHECKING:
    from pptx.enum.shapes import PP_PLACEHOLDER
    from pptx.oxml.shapes import ShapeElement
    from pptx.oxml.shapes.shared import CT_Transform2D


class CT_GroupShape(BaseShapeElement):
    """Used for shape tree (`p:spTree`) as well as the group shape (`p:grpSp`) elements."""

    nvGrpSpPr: CT_GroupShapeNonVisual = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:nvGrpSpPr"
    )
    grpSpPr: CT_GroupShapeProperties = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:grpSpPr"
    )

    _shape_tags = (
        qn("p:sp"),
        qn("p:grpSp"),
        qn("p:graphicFrame"),
        qn("p:cxnSp"),
        qn("p:pic"),
        qn("p:contentPart"),
    )

    def add_autoshape(
        self, id_: int, name: str, prst: str, x: int, y: int, cx: int, cy: int
    ) -> CT_Shape:
        """Return new `p:sp` appended to the group/shapetree with specified attributes."""
        pass

    def add_cxnSp(
        self,
        id_: int,
        name: str,
        type_member: MSO_CONNECTOR_TYPE,
        x: int,
        y: int,
        cx: int,
        cy: int,
        flipH: bool,
        flipV: bool,
    ) -> CT_Connector:
        """Return new `p:cxnSp` appended to the group/shapetree with the specified attribues."""
        pass

    def add_freeform_sp(self, x: int, y: int, cx: int, cy: int) -> CT_Shape:
        """Append a new freeform `p:sp` with specified position and size."""
        pass

    def add_grpSp(self) -> CT_GroupShape:
        """Return `p:grpSp` element newly appended to this shape tree.

        The element contains no sub-shapes, is positioned at (0, 0), and has
        width and height of zero.
        """
        pass

    def add_pic(
        self, id_: int, name: str, desc: str, rId: str, x: int, y: int, cx: int, cy: int
    ) -> CT_Picture:
        """Append a `p:pic` shape to the group/shapetree having properties as specified in call."""
        pass

    def add_placeholder(
        self, id_: int, name: str, ph_type: PP_PLACEHOLDER, orient: str, sz: str, idx: int
    ) -> CT_Shape:
        """Append a newly-created placeholder `p:sp` shape having the specified properties."""
        sp = CT_Shape.new_placeholder_sp(id_, name, ph_type, orient, sz, idx)
        self.insert_element_before(sp, "p:extLst")
        return sp

    def add_table(
        self, id_: int, name: str, rows: int, cols: int, x: int, y: int, cx: int, cy: int
    ) -> CT_GraphicalObjectFrame:
        """Append a `p:graphicFrame` shape containing a table as specified in call."""
        pass

    def add_textbox(self, id_: int, name: str, x: int, y: int, cx: int, cy: int) -> CT_Shape:
        """Append a newly-created textbox `p:sp` shape having the specified position and size."""
        pass

    @property
    def chExt(self):
        """Descendent `p:grpSpPr/a:xfrm/a:chExt` element."""
        pass

    @property
    def chOff(self):
        """Descendent `p:grpSpPr/a:xfrm/a:chOff` element."""
        pass

    def get_or_add_xfrm(self) -> CT_Transform2D:
        """Return the `a:xfrm` grandchild element, newly-added if not present."""
        pass

    def iter_ph_elms(self):
        """Generate each placeholder shape child element in document order."""
        pass

    def iter_shape_elms(self) -> Iterator[ShapeElement]:
        """Generate each child of this `p:spTree` element that corresponds to a shape.

        Items appear in XML document order.
        """
        for elm in self.iterchildren():
            if elm.tag in self._shape_tags:
                yield elm

    @property
    def max_shape_id(self) -> int:
        """Maximum int value assigned as @id in this slide.

        This is generally a shape-id, but ids can be assigned to other
        objects so we just check all @id values anywhere in the document
        (XML id-values have document scope).

        In practice, its minimum value is 1 because the spTree element itself
        is always assigned id="1".
        """
        pass

    @classmethod
    def new_grpSp(cls, id_: int, name: str) -> CT_GroupShape:
        """Return new "loose" `p:grpSp` element having `id_` and `name`."""
        pass

    def recalculate_extents(self) -> None:
        """Adjust x, y, cx, and cy to incorporate all contained shapes.

        This would typically be called when a contained shape is added,
        removed, or its position or size updated.

        This method is recursive "upwards" since a change in a group shape
        can change the position and size of its containing group.
        """
        pass

    @property
    def xfrm(self) -> CT_Transform2D | None:
        """The `a:xfrm` grandchild element or |None| if not found."""
        pass

    @property
    def _child_extents(self) -> tuple[int, int, int, int]:
        """(x, y, cx, cy) tuple representing net position and size.

        The values are formed as a composite of the contained child shapes.
        """
        pass

    @property
    def _next_shape_id(self) -> int:
        """Return unique shape id suitable for use with a new shape element.

        The returned id is the next available positive integer drawing object
        id in shape tree, starting from 1 and making use of any gaps in
        numbering. In practice, the minimum id is 2 because the spTree
        element itself is always assigned id="1".
        """
        pass


class CT_GroupShapeNonVisual(BaseShapeElement):
    """`p:nvGrpSpPr` element."""

    cNvPr = OneAndOnlyOne("p:cNvPr")


class CT_GroupShapeProperties(BaseOxmlElement):
    """p:grpSpPr element"""

    get_or_add_xfrm: Callable[[], CT_Transform2D]

    _tag_seq = (
        "a:xfrm",
        "a:noFill",
        "a:solidFill",
        "a:gradFill",
        "a:blipFill",
        "a:pattFill",
        "a:grpFill",
        "a:effectLst",
        "a:effectDag",
        "a:scene3d",
        "a:extLst",
    )
    xfrm: CT_Transform2D | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:xfrm", successors=_tag_seq[1:]
    )
    effectLst = ZeroOrOne("a:effectLst", successors=_tag_seq[8:])
    del _tag_seq
