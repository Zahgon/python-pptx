"""Custom element classes for top-level chart-related XML elements."""

from __future__ import annotations

from typing import cast

from pptx.oxml import parse_xml
from pptx.oxml.chart.shared import CT_Title
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import ST_Style, XsdString
from pptx.oxml.text import CT_TextBody
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_Chart(BaseOxmlElement):
    """`c:chart` custom element class."""

    _tag_seq = (
        "c:title",
        "c:autoTitleDeleted",
        "c:pivotFmts",
        "c:view3D",
        "c:floor",
        "c:sideWall",
        "c:backWall",
        "c:plotArea",
        "c:legend",
        "c:plotVisOnly",
        "c:dispBlanksAs",
        "c:showDLblsOverMax",
        "c:extLst",
    )
    title = ZeroOrOne("c:title", successors=_tag_seq[1:])
    autoTitleDeleted = ZeroOrOne("c:autoTitleDeleted", successors=_tag_seq[2:])
    plotArea = OneAndOnlyOne("c:plotArea")
    legend = ZeroOrOne("c:legend", successors=_tag_seq[9:])
    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]

    @property
    def has_legend(self):
        """
        True if this chart has a legend defined, False otherwise.
        """
        pass

    @has_legend.setter
    def has_legend(self, bool_value):
        """
        Add, remove, or leave alone the ``<c:legend>`` child element depending
        on current state and *bool_value*. If *bool_value* is |True| and no
        ``<c:legend>`` element is present, a new default element is added.
        When |False|, any existing legend element is removed.
        """
        pass

    @staticmethod
    def new_chart(rId: str) -> CT_Chart:
        """Return a new `c:chart` element."""
        pass

    def _new_title(self):
        pass


class CT_ChartSpace(BaseOxmlElement):
    """`c:chartSpace` root element of a chart part."""

    _tag_seq = (
        "c:date1904",
        "c:lang",
        "c:roundedCorners",
        "c:style",
        "c:clrMapOvr",
        "c:pivotSource",
        "c:protection",
        "c:chart",
        "c:spPr",
        "c:txPr",
        "c:externalData",
        "c:printSettings",
        "c:userShapes",
        "c:extLst",
    )
    date1904 = ZeroOrOne("c:date1904", successors=_tag_seq[1:])
    style = ZeroOrOne("c:style", successors=_tag_seq[4:])
    chart = OneAndOnlyOne("c:chart")
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[10:])
    externalData = ZeroOrOne("c:externalData", successors=_tag_seq[11:])
    del _tag_seq

    @property
    def catAx_lst(self):
        pass

    @property
    def date_1904(self):
        """
        Return |True| if the `c:date1904` child element resolves truthy,
        |False| otherwise. This value indicates whether date number values
        are based on the 1900 or 1904 epoch.
        """
        pass

    @property
    def dateAx_lst(self):
        pass

    def get_or_add_title(self):
        """Return the `c:title` grandchild, newly created if not present."""
        pass

    @property
    def plotArea(self):
        """
        Return the required `c:chartSpace/c:chart/c:plotArea` grandchild
        element.
        """
        pass

    @property
    def valAx_lst(self):
        pass

    @property
    def xlsx_part_rId(self):
        """
        The string in the required ``r:id`` attribute of the
        `<c:externalData>` child, or |None| if no externalData element is
        present.
        """
        pass

    def _add_externalData(self):
        """
        Always add a ``<c:autoUpdate val="0"/>`` child so auto-updating
        behavior is off by default.
        """
        pass

    def _new_txPr(self):
        pass


class CT_ExternalData(BaseOxmlElement):
    """
    `<c:externalData>` element, defining link to embedded Excel package part
    containing the chart data.
    """

    autoUpdate = ZeroOrOne("c:autoUpdate")
    rId = RequiredAttribute("r:id", XsdString)


class CT_PlotArea(BaseOxmlElement):
    """
    ``<c:plotArea>`` element.
    """

    catAx = ZeroOrMore("c:catAx")
    valAx = ZeroOrMore("c:valAx")

    def iter_sers(self):
        """
        Generate each of the `c:ser` elements in this chart, ordered first by
        the document order of the containing xChart element, then by their
        ordering within the xChart element (not necessarily document order).
        """
        pass

    def iter_xCharts(self):
        """
        Generate each xChart child element in document.
        """
        pass

    @property
    def last_ser(self):
        """
        Return the last `<c:ser>` element in the last xChart element, based
        on series order (not necessarily the same element as document order).
        """
        pass

    @property
    def next_idx(self):
        """
        Return the next available `c:ser/c:idx` value within the scope of
        this chart, the maximum idx value found on existing series,
        incremented by one.
        """
        pass

    @property
    def next_order(self):
        """
        Return the next available `c:ser/c:order` value within the scope of
        this chart, the maximum order value found on existing series,
        incremented by one.
        """
        pass

    @property
    def sers(self):
        """
        Return a sequence containing all the `c:ser` elements in this chart,
        ordered first by the document order of the containing xChart element,
        then by their ordering within the xChart element (not necessarily
        document order).
        """
        pass

    @property
    def xCharts(self):
        """
        Return a sequence containing all the `c:{x}Chart` elements in this
        chart, in document order.
        """
        pass


class CT_Style(BaseOxmlElement):
    """
    ``<c:style>`` element; defines the chart style.
    """

    val = RequiredAttribute("val", ST_Style)
