"""Chart data-label related oxml objects."""

from __future__ import annotations

from pptx.enum.chart import XL_DATA_LABEL_POSITION
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.text import CT_TextBody
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_DLbl(BaseOxmlElement):
    """
    ``<c:dLbl>`` element specifying the properties of the data label for an
    individual data point.
    """

    _tag_seq = (
        "c:idx",
        "c:layout",
        "c:tx",
        "c:numFmt",
        "c:spPr",
        "c:txPr",
        "c:dLblPos",
        "c:showLegendKey",
        "c:showVal",
        "c:showCatName",
        "c:showSerName",
        "c:showPercent",
        "c:showBubbleSize",
        "c:separator",
        "c:extLst",
    )
    idx = OneAndOnlyOne("c:idx")
    tx = ZeroOrOne("c:tx", successors=_tag_seq[3:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[5:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[6:])
    dLblPos = ZeroOrOne("c:dLblPos", successors=_tag_seq[7:])
    del _tag_seq

    def get_or_add_rich(self):
        """
        Return the `c:rich` descendant representing the text frame of the
        data label, newly created if not present. Any existing `c:strRef`
        element is removed along with its contents.
        """
        pass

    def get_or_add_tx_rich(self):
        """
        Return the `c:tx[c:rich]` subtree, newly created if not present.
        """
        pass

    @property
    def idx_val(self):
        """
        The integer value of the `val` attribute on the required `c:idx`
        child.
        """
        pass

    @classmethod
    def new_dLbl(cls):
        """Return a newly created "loose" `c:dLbl` element.

        The `c:dLbl` element contains the same (fairly extensive) default
        subtree added by PowerPoint when an individual data label is
        customized in the UI. Note that the idx value must be set by the
        client. Failure to set the idx value will likely result in any
        changes not being visible and may result in a repair error on open.
        """
        pass

    def remove_tx_rich(self):
        """
        Remove any `c:tx[c:rich]` child, or do nothing if not present.
        """
        pass

    def _new_txPr(self):
        pass


class CT_DLblPos(BaseOxmlElement):
    """
    ``<c:dLblPos>`` element specifying the positioning of a data label with
    respect to its data point.
    """

    val = RequiredAttribute("val", XL_DATA_LABEL_POSITION)


class CT_DLbls(BaseOxmlElement):
    """`c:dLbls` element specifying properties for a set of data labels."""

    _tag_seq = (
        "c:dLbl",
        "c:numFmt",
        "c:spPr",
        "c:txPr",
        "c:dLblPos",
        "c:showLegendKey",
        "c:showVal",
        "c:showCatName",
        "c:showSerName",
        "c:showPercent",
        "c:showBubbleSize",
        "c:separator",
        "c:showLeaderLines",
        "c:leaderLines",
        "c:extLst",
    )
    dLbl = ZeroOrMore("c:dLbl", successors=_tag_seq[1:])
    numFmt = ZeroOrOne("c:numFmt", successors=_tag_seq[2:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[4:])
    dLblPos = ZeroOrOne("c:dLblPos", successors=_tag_seq[5:])
    showLegendKey = ZeroOrOne("c:showLegendKey", successors=_tag_seq[6:])
    showVal = ZeroOrOne("c:showVal", successors=_tag_seq[7:])
    showCatName = ZeroOrOne("c:showCatName", successors=_tag_seq[8:])
    showSerName = ZeroOrOne("c:showSerName", successors=_tag_seq[9:])
    showPercent = ZeroOrOne("c:showPercent", successors=_tag_seq[10:])
    del _tag_seq

    @property
    def defRPr(self):
        """
        ``<a:defRPr>`` great-great-grandchild element, added with its
        ancestors if not present.
        """
        pass

    def get_dLbl_for_point(self, idx):
        """
        Return the `c:dLbl` child representing the label for the data point
        at index *idx*.
        """
        pass

    def get_or_add_dLbl_for_point(self, idx):
        """
        Return the `c:dLbl` element representing the label of the point at
        index *idx*.
        """
        pass

    @classmethod
    def new_dLbls(cls):
        """Return a newly created "loose" `c:dLbls` element."""
        pass

    def _insert_dLbl_in_sequence(self, idx):
        """
        Return a newly created `c:dLbl` element having `c:idx` child of *idx*
        and inserted in numeric sequence among the `c:dLbl` children of this
        element.
        """
        pass

    def _new_dLbl(self):
        pass

    def _new_showCatName(self):
        """Return a new `c:showCatName` with value initialized.

        This method is called by the metaclass-generated code whenever a new
        `c:showCatName` element is required. In this case, it defaults to
        `val=true`, which is not what we need so we override to make val
        explicitly False.
        """
        pass

    def _new_showLegendKey(self):
        pass

    def _new_showPercent(self):
        pass

    def _new_showSerName(self):
        pass

    def _new_showVal(self):
        pass

    def _new_txPr(self):
        pass
