"""Base and meta classes enabling declarative definition of custom element classes."""

from __future__ import annotations

import re
from typing import Any, Callable, Iterable, Protocol, Sequence, Type, cast

from lxml import etree
from lxml.etree import ElementBase, _Element  # pyright: ignore[reportPrivateUsage]

from pptx.exc import InvalidXmlError
from pptx.oxml import oxml_parser
from pptx.oxml.ns import NamespacePrefixedTag, _nsmap, qn  # pyright: ignore[reportPrivateUsage]
from pptx.util import lazyproperty


class AttributeType(Protocol):
    """Interface for an object that can act as an attribute type.

    An attribute-type specifies how values are transformed to and from the XML "string" value of the
    attribute.
    """

    @classmethod
    def from_xml(cls, xml_value: str) -> Any:
        """Transform an attribute value to a Python value."""
        ...

    @classmethod
    def to_xml(cls, value: Any) -> str:
        """Transform a Python value to a str value suitable to this XML attribute."""
        ...


def OxmlElement(nsptag_str: str, nsmap: dict[str, str] | None = None) -> BaseOxmlElement:
    """Return a "loose" lxml element having the tag specified by `nsptag_str`.

    `nsptag_str` must contain the standard namespace prefix, e.g. 'a:tbl'. The resulting element is
    an instance of the custom element class for this tag name if one is defined.
    """
    pass


def serialize_for_reading(element: ElementBase):
    """
    Serialize *element* to human-readable XML suitable for tests. No XML
    declaration.
    """
    xml = etree.tostring(element, encoding="unicode", pretty_print=True)
    return XmlString(xml)


class XmlString(str):
    """Provides string comparison override suitable for serialized XML; useful for tests."""

    # '    <w:xyz xmlns:a="http://ns/decl/a" attr_name="val">text</w:xyz>'
    # |          |                                          ||           |
    # +----------+------------------------------------------++-----------+
    #  front      attrs                                     | text
    #                                                     close

    _xml_elm_line_patt = re.compile(r"( *</?[\w:]+)(.*?)(/?>)([^<]*</[\w:]+>)?")

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, str):
            return False
        lines = self.splitlines()
        lines_other = other.splitlines()
        if len(lines) != len(lines_other):
            return False
        for line, line_other in zip(lines, lines_other):
            if not self._eq_elm_strs(line, line_other):
                return False
        return True

    def __ne__(self, other: object) -> bool:
        return not self.__eq__(other)

    def _attr_seq(self, attrs: str) -> list[str]:
        """Return a sequence of attribute strings parsed from *attrs*.

        Each attribute string is stripped of whitespace on both ends.
        """
        pass

    def _eq_elm_strs(self, line: str, line_2: str) -> bool:
        """True if the element in `line_2` is XML-equivalent to the element in `line`.

        In particular, the order of attributes in XML is not significant.
        """
        pass

    def _parse_line(self, line: str):
        """Return front, attrs, close, text 4-tuple result of parsing XML element string `line`."""
        pass


class MetaOxmlElement(type):
    """Metaclass for BaseOxmlElement."""

    def __init__(cls, clsname: str, bases: tuple[type, ...], clsdict: dict[str, Any]):
        dispatchable = (
            OneAndOnlyOne,
            OneOrMore,
            OptionalAttribute,
            RequiredAttribute,
            ZeroOrMore,
            ZeroOrOne,
            ZeroOrOneChoice,
        )
        for key, value in clsdict.items():
            if isinstance(value, dispatchable):
                value.populate_class_members(cls, key)


class BaseAttribute:
    """Base class for OptionalAttribute and RequiredAttribute, providing common methods."""

    def __init__(self, attr_name: str, simple_type: type[AttributeType]):
        self._attr_name = attr_name
        self._simple_type = simple_type

    def populate_class_members(self, element_cls: Type[BaseOxmlElement], prop_name: str):
        """
        Add the appropriate methods to *element_cls*.
        """
        pass

    def _add_attr_property(self):
        """Add a read/write `{prop_name}` property to the element class.

        The property returns the interpreted value of this attribute on access and changes the
        attribute value to its ST_* counterpart on assignment.
        """
        pass

    @property
    def _clark_name(self):
        pass

    @property
    def _getter(self) -> Callable[[BaseOxmlElement], Any]:
        """Callable suitable for the "get" side of the attribute property descriptor."""
        raise NotImplementedError("must be implemented by each subclass")

    @property
    def _setter(self) -> Callable[[BaseOxmlElement, Any], None]:
        """Callable suitable for the "set" side of the attribute property descriptor."""
        raise NotImplementedError("must be implemented by each subclass")


class OptionalAttribute(BaseAttribute):
    """Defines an optional attribute on a custom element class.

    An optional attribute returns a default value when not present for reading. When assigned
    |None|, the attribute is removed.
    """

    def __init__(self, attr_name: str, simple_type: type[AttributeType], default: Any = None):
        super(OptionalAttribute, self).__init__(attr_name, simple_type)
        self._default = default

    @property
    def _docstring(self):
        """
        Return the string to use as the ``__doc__`` attribute of the property
        for this attribute.
        """
        pass

    @property
    def _getter(self) -> Callable[[BaseOxmlElement], Any]:
        """Callable suitable for the "get" side of the attribute property descriptor."""
        pass

    @property
    def _setter(self) -> Callable[[BaseOxmlElement, Any], None]:
        """Callable suitable for the "set" side of the attribute property descriptor."""
        pass


class RequiredAttribute(BaseAttribute):
    """Defines a required attribute on a custom element class.

    A required attribute is assumed to be present for reading, so does not have a default value;
    its actual value is always used. If missing on read, an |InvalidXmlError| is raised. It also
    does not remove the attribute if |None| is assigned. Assigning |None| raises |TypeError| or
    |ValueError|, depending on the simple type of the attribute.
    """

    @property
    def _getter(self) -> Callable[[BaseOxmlElement], Any]:
        """Callable suitable for the "get" side of the attribute property descriptor."""
        pass

    @property
    def _docstring(self):
        """
        Return the string to use as the ``__doc__`` attribute of the property
        for this attribute.
        """
        pass

    @property
    def _setter(self) -> Callable[[BaseOxmlElement, Any], None]:
        """Callable suitable for the "set" side of the attribute property descriptor."""
        pass


class _BaseChildElement:
    """Base class for the child element classes corresponding to varying cardinalities.

    Subclasses include ZeroOrOne and ZeroOrMore.
    """

    def __init__(self, nsptagname: str, successors: Sequence[str] = ()):
        super(_BaseChildElement, self).__init__()
        self._nsptagname = nsptagname
        self._successors = successors

    def populate_class_members(self, element_cls: Type[BaseOxmlElement], prop_name: str):
        """Baseline behavior for adding the appropriate methods to `element_cls`."""
        pass

    def _add_adder(self):
        """Add an ``_add_x()`` method to the element class for this child element."""
        pass

    def _add_creator(self):
        """Add a `_new_{prop_name}()` method to the element class.

        This method creates a new, empty element of the correct type, having no attributes.
        """
        pass

    def _add_getter(self):
        """Add a read-only `{prop_name}` property to the parent element class.

        The property locates and returns this child element or `None` if not present.
        """
        pass

    def _add_inserter(self):
        """Add an ``_insert_x()`` method to the element class for this child element."""
        pass

    def _add_list_getter(self):
        """
        Add a read-only ``{prop_name}_lst`` property to the element class to
        retrieve a list of child elements matching this type.
        """
        pass

    @lazyproperty
    def _add_method_name(self):
        pass

    def _add_to_class(self, name: str, method: Callable[..., Any]):
        """Add `method` to the target class as `name`, unless `name` is already defined there."""
        pass

    @property
    def _creator(self) -> Callable[[BaseOxmlElement], BaseOxmlElement]:
        """Callable that creates a new, empty element of the child type, having no attributes."""
        pass

    @property
    def _getter(self) -> Callable[[BaseOxmlElement], BaseOxmlElement | None]:
        """Callable suitable for the "get" side of the property descriptor.

        This default getter returns the child element with matching tag name or |None| if not
        present.
        """
        pass

    @lazyproperty
    def _insert_method_name(self):
        pass

    @property
    def _list_getter(self) -> Callable[[BaseOxmlElement], list[BaseOxmlElement]]:
        """Callable suitable for the "get" side of a list property descriptor."""
        pass

    @lazyproperty
    def _remove_method_name(self):
        pass

    @lazyproperty
    def _new_method_name(self):
        pass


class Choice(_BaseChildElement):
    """Defines a child element belonging to a group, only one of which may appear as a child."""

    @property
    def nsptagname(self):
        pass

    def populate_class_members(  # pyright: ignore[reportIncompatibleMethodOverride]
        self, element_cls: Type[BaseOxmlElement], group_prop_name: str, successors: Sequence[str]
    ):
        """Add the appropriate methods to `element_cls`."""
        pass

    def _add_get_or_change_to_method(self) -> None:
        """Add a `get_or_change_to_x()` method to the element class for this child element."""
        pass

    @property
    def _prop_name(self):
        """
        Calculate property name from tag name, e.g. a:schemeClr -> schemeClr.
        """
        pass

    @lazyproperty
    def _get_or_change_to_method_name(self):
        pass

    @lazyproperty
    def _remove_group_method_name(self):
        pass


class OneAndOnlyOne(_BaseChildElement):
    """Defines a required child element for MetaOxmlElement."""

    def __init__(self, nsptagname: str):
        super(OneAndOnlyOne, self).__init__(nsptagname, ())

    def populate_class_members(self, element_cls: Type[BaseOxmlElement], prop_name: str):
        """
        Add the appropriate methods to *element_cls*.
        """
        pass

    @property
    def _getter(self) -> Callable[[BaseOxmlElement], BaseOxmlElement]:
        """Callable suitable for the "get" side of the property descriptor."""
        pass


class OneOrMore(_BaseChildElement):
    """Defines a repeating child element for MetaOxmlElement that must appear at least once."""

    def populate_class_members(self, element_cls: Type[BaseOxmlElement], prop_name: str):
        """Add the appropriate methods to *element_cls*."""
        pass

    def _add_public_adder(self) -> None:
        """Add a public `.add_x()` method to the parent element class."""
        pass

    @lazyproperty
    def _public_add_method_name(self):
        """
        add_childElement() is public API for a repeating element, allowing
        new elements to be added to the sequence. May be overridden to
        provide a friendlier API to clients having domain appropriate
        parameter names for required attributes.
        """
        pass


class ZeroOrMore(_BaseChildElement):
    """
    Defines an optional repeating child element for MetaOxmlElement.
    """

    def populate_class_members(self, element_cls: Type[BaseOxmlElement], prop_name: str):
        """
        Add the appropriate methods to *element_cls*.
        """
        pass


class ZeroOrOne(_BaseChildElement):
    """Defines an optional child element for MetaOxmlElement."""

    def populate_class_members(self, element_cls: Type[BaseOxmlElement], prop_name: str):
        """Add the appropriate methods to `element_cls`."""
        pass

    def _add_get_or_adder(self):
        """Add a `.get_or_add_x()` method to the element class for this child element."""
        pass

    def _add_remover(self):
        """Add a `._remove_x()` method to the element class for this child element."""
        pass

    @lazyproperty
    def _get_or_add_method_name(self):
        pass


class ZeroOrOneChoice(_BaseChildElement):
    """An `EG_*` element group where at most one of its members may appear as a child."""

    def __init__(self, choices: Iterable[Choice], successors: Iterable[str] = ()):
        self._choices = tuple(choices)
        self._successors = tuple(successors)

    def populate_class_members(self, element_cls: Type[BaseOxmlElement], prop_name: str):
        """Add the appropriate methods to `element_cls`."""
        pass

    def _add_choice_getter(self):
        """Add a read-only `.{prop_name}` property to the element class.

        The property returns the present member of this group, or |None| if none are present.
        """
        pass

    def _add_group_remover(self):
        """Add a `._remove_eg_x()` method to the element class for this choice group."""
        pass

    @property
    def _choice_getter(self):
        """
        Return a function object suitable for the "get" side of the property
        descriptor.
        """
        pass

    @lazyproperty
    def _member_nsptagnames(self) -> list[str]:
        """Sequence of namespace-prefixed tagnames, one for each member element of choice group."""
        pass

    @lazyproperty
    def _remove_choice_group_method_name(self):
        """Function-name for choice remover."""
        pass


# -- lxml typing isn't quite right here, just ignore this error on _Element --
class BaseOxmlElement(etree.ElementBase, metaclass=MetaOxmlElement):
    """Effective base class for all custom element classes.

    Adds standardized behavior to all classes in one place.
    """

    def __repr__(self):
        return "<%s '<%s>' at 0x%0x>" % (
            self.__class__.__name__,
            self._nsptag,
            id(self),
        )

    def first_child_found_in(self, *tagnames: str) -> _Element | None:
        """First child with tag in `tagnames`, or None if not found."""
        for tagname in tagnames:
            child = self.find(qn(tagname))
            if child is not None:
                return child
        return None

    def insert_element_before(self, elm: ElementBase, *tagnames: str):
        successor = self.first_child_found_in(*tagnames)
        if successor is not None:
            successor.addprevious(elm)
        else:
            self.append(elm)
        return elm

    def remove_all(self, *tagnames: str) -> None:
        """Remove child elements with tagname (e.g. "a:p") in `tagnames`."""
        pass

    @property
    def xml(self) -> str:
        """XML string for this element, suitable for testing purposes.

        Pretty printed for readability and without an XML declaration at the top.
        """
        return serialize_for_reading(self)

    def xpath(self, xpath_str: str) -> Any:  # pyright: ignore[reportIncompatibleMethodOverride]
        """Override of `lxml` _Element.xpath() method.

        Provides standard Open XML namespace mapping (`nsmap`) in centralized location.
        """
        return super().xpath(xpath_str, namespaces=_nsmap)

    @property
    def _nsptag(self) -> str:
        pass
