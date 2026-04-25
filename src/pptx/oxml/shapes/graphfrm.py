"""lxml custom element class for CT_GraphicalObjectFrame XML element."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.oxml import parse_xml
from pptx.oxml.chart.chart import CT_Chart
from pptx.oxml.ns import nsdecls
from pptx.oxml.shapes.shared import BaseShapeElement
from pptx.oxml.simpletypes import XsdBoolean, XsdString
from pptx.oxml.table import CT_Table
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)
from pptx.spec import (
    GRAPHIC_DATA_URI_CHART,
    GRAPHIC_DATA_URI_OLEOBJ,
    GRAPHIC_DATA_URI_TABLE,
)

if TYPE_CHECKING:
    from pptx.oxml.shapes.shared import (
        CT_ApplicationNonVisualDrawingProps,
        CT_NonVisualDrawingProps,
        CT_Transform2D,
    )


class CT_GraphicalObject(BaseOxmlElement):
    """`a:graphic` element.

    The container for the reference to or definition of the framed graphical object (table, chart,
    etc.).
    """

    graphicData: CT_GraphicalObjectData = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:graphicData"
    )

    @property
    def chart(self) -> CT_Chart | None:
        """The `c:chart` grandchild element, or |None| if not present."""
        pass


class CT_GraphicalObjectData(BaseShapeElement):
    """`p:graphicData` element.

    The direct container for a table, a chart, or another graphical object.
    """

    chart: CT_Chart | None = ZeroOrOne("c:chart")  # pyright: ignore[reportAssignmentType]
    tbl: CT_Table | None = ZeroOrOne("a:tbl")  # pyright: ignore[reportAssignmentType]
    uri: str = RequiredAttribute("uri", XsdString)  # pyright: ignore[reportAssignmentType]

    @property
    def blob_rId(self) -> str | None:
        """Optional `r:id` attribute value of `p:oleObj` descendent element.

        This value is `None` when this `p:graphicData` element does not enclose an OLE object.
        This value could also be `None` if an enclosed OLE object does not specify this attribute
        (it is specified optional in the schema) but so far, all OLE objects we've encountered
        specify this value.
        """
        pass

    @property
    def is_embedded_ole_obj(self) -> bool | None:
        """Optional boolean indicating an embedded OLE object.

        Returns `None` when this `p:graphicData` element does not enclose an OLE object. `True`
        indicates an embedded OLE object and `False` indicates a linked OLE object.
        """
        pass

    @property
    def progId(self) -> str | None:
        """Optional str value of "progId" attribute of `p:oleObj` descendent.

        This value identifies the "type" of the embedded object in terms of the application used
        to open it.

        This value is `None` when this `p:graphicData` element does not enclose an OLE object.
        This could also be `None` if an enclosed OLE object does not specify this attribute (it is
        specified optional in the schema) but so far, all OLE objects we've encountered specify
        this value.
        """
        pass

    @property
    def showAsIcon(self) -> bool | None:
        """Optional value of "showAsIcon" attribute value of `p:oleObj` descendent.

        This value is `None` when this `p:graphicData` element does not enclose an OLE object. It
        is False when the `showAsIcon` attribute is omitted on the `p:oleObj` element.
        """
        pass

    @property
    def _oleObj(self) -> CT_OleObject | None:
        """Optional `p:oleObj` element contained in this `p:graphicData' element.

        Returns `None` when this graphic-data element does not enclose an OLE object. Note that
        this returns the last `p:oleObj` element found. There can be more than one `p:oleObj`
        element because an `mc.AlternateContent` element may appear as the child of
        `p:graphicData` and that alternate-content subtree can contain multiple compatibility
        choices. The last one should suit best for reading purposes because it contains the lowest
        common denominator.
        """
        pass


class CT_GraphicalObjectFrame(BaseShapeElement):
    """`p:graphicFrame` element.

    A container for a table, a chart, or another graphical object.
    """

    nvGraphicFramePr: CT_GraphicalObjectFrameNonVisual = (  # pyright: ignore[reportAssignmentType]
        OneAndOnlyOne("p:nvGraphicFramePr")
    )
    xfrm: CT_Transform2D = OneAndOnlyOne("p:xfrm")  # pyright: ignore
    graphic: CT_GraphicalObject = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:graphic"
    )

    @property
    def chart(self) -> CT_Chart | None:
        """The `c:chart` great-grandchild element, or |None| if not present."""
        pass

    @property
    def chart_rId(self) -> str | None:
        """The `rId` attribute of the `c:chart` great-grandchild element.

        |None| if not present.
        """
        pass

    def get_or_add_xfrm(self) -> CT_Transform2D:
        """Return the required `p:xfrm` child element.

        Overrides version on BaseShapeElement.
        """
        pass

    @property
    def graphicData(self) -> CT_GraphicalObjectData:
        """`a:graphicData` grandchild of this graphic-frame element."""
        pass

    @property
    def graphicData_uri(self) -> str:
        """str value of `uri` attribute of `a:graphicData` grandchild."""
        pass

    @property
    def has_oleobj(self) -> bool:
        """`True` for graphicFrame containing an OLE object, `False` otherwise."""
        pass

    @property
    def is_embedded_ole_obj(self) -> bool | None:
        """Optional boolean indicating an embedded OLE object.

        Returns `None` when this `p:graphicFrame` element does not enclose an OLE object. `True`
        indicates an embedded OLE object and `False` indicates a linked OLE object.
        """
        pass

    @classmethod
    def new_chart_graphicFrame(
        cls, id_: int, name: str, rId: str, x: int, y: int, cx: int, cy: int
    ) -> CT_GraphicalObjectFrame:
        """Return a `p:graphicFrame` element tree populated with a chart element."""
        pass

    @classmethod
    def new_graphicFrame(
        cls, id_: int, name: str, x: int, y: int, cx: int, cy: int
    ) -> CT_GraphicalObjectFrame:
        """Return a new `p:graphicFrame` element tree suitable for containing a table or chart.

        Note that a graphicFrame element is not a valid shape until it contains a graphical object
        such as a table.
        """
        pass

    @classmethod
    def new_ole_object_graphicFrame(
        cls,
        id_: int,
        name: str,
        ole_object_rId: str,
        progId: str,
        icon_rId: str,
        x: int,
        y: int,
        cx: int,
        cy: int,
        imgW: int,
        imgH: int,
    ) -> CT_GraphicalObjectFrame:
        """Return newly-created `p:graphicFrame` for embedded OLE-object.

        `ole_object_rId` identifies the relationship to the OLE-object part.

        `progId` is a str identifying the object-type in terms of the application (program) used
        to open it. This becomes an attribute of the same name in the `p:oleObj` element.

        `icon_rId` identifies the relationship to an image part used to display the OLE-object as
        an icon (vs. a preview).
        """
        pass

    @classmethod
    def new_table_graphicFrame(
        cls, id_: int, name: str, rows: int, cols: int, x: int, y: int, cx: int, cy: int
    ) -> CT_GraphicalObjectFrame:
        """Return a `p:graphicFrame` element tree populated with a table element."""
        pass


class CT_GraphicalObjectFrameNonVisual(BaseOxmlElement):
    """`p:nvGraphicFramePr` element.

    This contains the non-visual properties of a graphic frame, such as name, id, etc.
    """

    cNvPr: CT_NonVisualDrawingProps = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cNvPr"
    )
    nvPr: CT_ApplicationNonVisualDrawingProps = (  # pyright: ignore[reportAssignmentType]
        OneAndOnlyOne("p:nvPr")
    )


class CT_OleObject(BaseOxmlElement):
    """`p:oleObj` element, container for an OLE object (e.g. Excel file).

    An OLE object can be either linked or embedded (hence the name).
    """

    progId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "progId", XsdString
    )
    rId: str | None = OptionalAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]
    showAsIcon: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "showAsIcon", XsdBoolean, default=False
    )

    @property
    def is_embedded(self) -> bool:
        """True when this OLE object is embedded, False when it is linked."""
        pass
