"""lxml custom element classes for core properties-related XML elements."""

from __future__ import annotations

import datetime as dt
import re
from typing import Callable, cast

from lxml.etree import _Element  # pyright: ignore[reportPrivateUsage]

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.xmlchemy import BaseOxmlElement, ZeroOrOne


class CT_CoreProperties(BaseOxmlElement):
    """`cp:coreProperties` element.

    The root element of the Core Properties part stored as `/docProps/core.xml`. Implements many
    of the Dublin Core document metadata elements. String elements resolve to an empty string ('')
    if the element is not present in the XML. String elements are limited in length to 255 unicode
    characters.
    """

    get_or_add_revision: Callable[[], _Element]

    category = ZeroOrOne("cp:category", successors=())
    contentStatus = ZeroOrOne("cp:contentStatus", successors=())
    created = ZeroOrOne("dcterms:created", successors=())
    creator = ZeroOrOne("dc:creator", successors=())
    description = ZeroOrOne("dc:description", successors=())
    identifier = ZeroOrOne("dc:identifier", successors=())
    keywords = ZeroOrOne("cp:keywords", successors=())
    language = ZeroOrOne("dc:language", successors=())
    lastModifiedBy = ZeroOrOne("cp:lastModifiedBy", successors=())
    lastPrinted = ZeroOrOne("cp:lastPrinted", successors=())
    modified = ZeroOrOne("dcterms:modified", successors=())
    revision: _Element | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cp:revision", successors=()
    )
    subject = ZeroOrOne("dc:subject", successors=())
    title = ZeroOrOne("dc:title", successors=())
    version = ZeroOrOne("cp:version", successors=())

    _coreProperties_tmpl = "<cp:coreProperties %s/>\n" % nsdecls("cp", "dc", "dcterms")

    @staticmethod
    def new_coreProperties() -> CT_CoreProperties:
        """Return a new `cp:coreProperties` element"""
        pass

    @property
    def author_text(self) -> str:
        pass

    @author_text.setter
    def author_text(self, value: str):
        pass

    @property
    def category_text(self) -> str:
        pass

    @category_text.setter
    def category_text(self, value: str):
        pass

    @property
    def comments_text(self) -> str:
        pass

    @comments_text.setter
    def comments_text(self, value: str):
        pass

    @property
    def contentStatus_text(self) -> str:
        pass

    @contentStatus_text.setter
    def contentStatus_text(self, value: str):
        pass

    @property
    def created_datetime(self):
        pass

    @created_datetime.setter
    def created_datetime(self, value: dt.datetime):
        pass

    @property
    def identifier_text(self) -> str:
        pass

    @identifier_text.setter
    def identifier_text(self, value: str):
        pass

    @property
    def keywords_text(self) -> str:
        pass

    @keywords_text.setter
    def keywords_text(self, value: str):
        pass

    @property
    def language_text(self) -> str:
        pass

    @language_text.setter
    def language_text(self, value: str):
        pass

    @property
    def lastModifiedBy_text(self) -> str:
        pass

    @lastModifiedBy_text.setter
    def lastModifiedBy_text(self, value: str):
        pass

    @property
    def lastPrinted_datetime(self):
        pass

    @lastPrinted_datetime.setter
    def lastPrinted_datetime(self, value: dt.datetime):
        pass

    @property
    def modified_datetime(self):
        pass

    @modified_datetime.setter
    def modified_datetime(self, value: dt.datetime):
        pass

    @property
    def revision_number(self) -> int:
        """Integer value of revision property."""
        pass

    @revision_number.setter
    def revision_number(self, value: int):
        """Set revision property to string value of integer `value`."""
        pass

    @property
    def subject_text(self) -> str:
        pass

    @subject_text.setter
    def subject_text(self, value: str):
        pass

    @property
    def title_text(self) -> str:
        pass

    @title_text.setter
    def title_text(self, value: str):
        pass

    @property
    def version_text(self) -> str:
        pass

    @version_text.setter
    def version_text(self, value: str):
        pass

    def _datetime_of_element(self, property_name: str) -> dt.datetime | None:
        pass

    def _get_or_add(self, prop_name: str):
        """Return element returned by 'get_or_add_' method for `prop_name`."""
        pass

    @classmethod
    def _offset_dt(cls, datetime: dt.datetime, offset_str: str):
        """Return |datetime| instance offset from `datetime` by offset specified in `offset_str`.

        `offset_str` is a string like `'-07:00'`.
        """
        pass

    _offset_pattern = re.compile(r"([+-])(\d\d):(\d\d)")

    @classmethod
    def _parse_W3CDTF_to_datetime(cls, w3cdtf_str: str) -> dt.datetime:
        # valid W3CDTF date cases:
        # yyyy e.g. '2003'
        # yyyy-mm e.g. '2003-12'
        # yyyy-mm-dd e.g. '2003-12-31'
        # UTC timezone e.g. '2003-12-31T10:14:55Z'
        # numeric timezone e.g. '2003-12-31T10:14:55-08:00'
        pass

    def _set_element_datetime(self, prop_name: str, value: dt.datetime) -> None:
        """Set date/time value of child element having `prop_name` to `value`."""
        pass

    def _set_element_text(self, prop_name: str, value: str) -> None:
        """Set string value of `name` property to `value`."""
        pass

    def _text_of_element(self, property_name: str) -> str:
        pass
