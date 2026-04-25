"""Base classes and other objects used by enumerations."""

from __future__ import annotations

import enum
import textwrap
from typing import TYPE_CHECKING, Any, Type, TypeVar

if TYPE_CHECKING:
    from typing_extensions import Self

_T = TypeVar("_T", bound="BaseXmlEnum")


class BaseEnum(int, enum.Enum):
    """Base class for Enums that do not map XML attr values.

    The enum's value will be an integer, corresponding to the integer assigned the
    corresponding member in the MS API enum of the same name.
    """

    def __new__(cls, ms_api_value: int, docstr: str):
        self = int.__new__(cls, ms_api_value)
        self._value_ = ms_api_value
        self.__doc__ = docstr.strip()
        return self

    def __str__(self):
        """The symbolic name and string value of this member, e.g. 'MIDDLE (3)'."""
        return f"{self.name} ({self.value})"


class BaseXmlEnum(int, enum.Enum):
    """Base class for Enums that also map XML attr values.

    The enum's value will be an integer, corresponding to the integer assigned the
    corresponding member in the MS API enum of the same name.
    """

    xml_value: str | None

    def __new__(cls, ms_api_value: int, xml_value: str | None, docstr: str):
        self = int.__new__(cls, ms_api_value)
        self._value_ = ms_api_value
        self.xml_value = xml_value
        self.__doc__ = docstr.strip()
        return self

    def __str__(self):
        """The symbolic name and string value of this member, e.g. 'MIDDLE (3)'."""
        return f"{self.name} ({self.value})"

    @classmethod
    def from_xml(cls, xml_value: str) -> Self:
        """Enumeration member corresponding to XML attribute value `xml_value`.

        Raises `ValueError` if `xml_value` is the empty string ("") or is not an XML attribute
        value registered on the enumeration. Note that enum members that do not correspond to one
        of the defined values for an XML attribute have `xml_value == ""`. These
        "return-value only" members cannot be automatically mapped from an XML attribute value and
        must be selected explicitly by code, based on the appropriate conditions.

        Example::

            >>> WD_PARAGRAPH_ALIGNMENT.from_xml("center")
            WD_PARAGRAPH_ALIGNMENT.CENTER

        """
        # -- the empty string never maps to a member --
        member = (
            next((member for member in cls if member.xml_value == xml_value), None)
            if xml_value
            else None
        )

        if member is None:
            raise ValueError(f"{cls.__name__} has no XML mapping for {repr(xml_value)}")

        return member

    @classmethod
    def to_xml(cls: Type[_T], value: int | _T) -> str:
        """XML value of this enum member, generally an XML attribute value."""
        pass

    @classmethod
    def validate(cls: Type[_T], value: _T):
        """Raise |ValueError| if `value` is not an assignable value."""
        pass


class DocsPageFormatter(object):
    """Formats a reStructuredText documention page (string) for an enumeration."""

    def __init__(self, clsname: str, clsdict: dict[str, Any]):
        self._clsname = clsname
        self._clsdict = clsdict

    @property
    def page_str(self):
        """
        The RestructuredText documentation page for the enumeration. This is
        the only API member for the class.
        """
        pass

    @property
    def _intro_text(self):
        """
        The docstring of the enumeration, formatted for use at the top of the
        documentation page
        """
        pass

    def _member_def(self, member: BaseEnum | BaseXmlEnum):
        """Return an individual member definition formatted as an RST glossary entry.

        Output is wrapped to fit within 78 columns.
        """
        pass

    @property
    def _member_defs(self):
        """
        A single string containing the aggregated member definitions section
        of the documentation page
        """
        pass

    @property
    def _ms_name(self):
        """
        The Microsoft API name for this enumeration
        """
        pass

    @property
    def _page_title(self):
        """
        The title for the documentation page, formatted as code (surrounded
        in double-backtics) and underlined with '=' characters
        """
        pass
