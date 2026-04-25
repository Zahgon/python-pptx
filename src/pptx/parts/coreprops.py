"""Core properties part, corresponds to ``/docProps/core.xml`` part in package."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.coreprops import CT_CoreProperties

if TYPE_CHECKING:
    from pptx.package import Package


class CorePropertiesPart(XmlPart):
    """Corresponds to part named `/docProps/core.xml`.

    Contains the core document properties for this document package.
    """

    _element: CT_CoreProperties

    @classmethod
    def default(cls, package: Package):
        """Return default new |CorePropertiesPart| instance suitable as starting point.

        This provides a base for adding core-properties to a package that doesn't yet
        have any.
        """
        pass

    @property
    def author(self) -> str:
        pass

    @author.setter
    def author(self, value: str):
        pass

    @property
    def category(self) -> str:
        pass

    @category.setter
    def category(self, value: str):
        pass

    @property
    def comments(self) -> str:
        pass

    @comments.setter
    def comments(self, value: str):
        pass

    @property
    def content_status(self) -> str:
        pass

    @content_status.setter
    def content_status(self, value: str):
        pass

    @property
    def created(self):
        pass

    @created.setter
    def created(self, value: dt.datetime):
        pass

    @property
    def identifier(self) -> str:
        pass

    @identifier.setter
    def identifier(self, value: str):
        pass

    @property
    def keywords(self) -> str:
        pass

    @keywords.setter
    def keywords(self, value: str):
        pass

    @property
    def language(self) -> str:
        pass

    @language.setter
    def language(self, value: str):
        pass

    @property
    def last_modified_by(self) -> str:
        pass

    @last_modified_by.setter
    def last_modified_by(self, value: str):
        pass

    @property
    def last_printed(self):
        pass

    @last_printed.setter
    def last_printed(self, value: dt.datetime):
        pass

    @property
    def modified(self):
        pass

    @modified.setter
    def modified(self, value: dt.datetime):
        pass

    @property
    def revision(self):
        pass

    @revision.setter
    def revision(self, value: int):
        pass

    @property
    def subject(self) -> str:
        pass

    @subject.setter
    def subject(self, value: str):
        pass

    @property
    def title(self) -> str:
        pass

    @title.setter
    def title(self, value: str):
        pass

    @property
    def version(self) -> str:
        pass

    @version.setter
    def version(self, value: str):
        pass

    @classmethod
    def _new(cls, package: Package) -> CorePropertiesPart:
        """Return new empty |CorePropertiesPart| instance."""
        pass
