"""Simple-type classes.

A "simple-type" is a scalar type, generally serving as an XML attribute. This is in contrast to a
"complex-type" which would specify an XML element.

These objects providing validation and format translation for values stored in XML element
attributes. Naming generally corresponds to the simple type in the associated XML schema.
"""

from __future__ import annotations

import numbers
from typing import Any

from pptx.exc import InvalidXmlError
from pptx.util import Centipoints, Emu


class BaseSimpleType:
    @classmethod
    def from_xml(cls, xml_value: str) -> Any:
        return cls.convert_from_xml(xml_value)

    @classmethod
    def to_xml(cls, value: Any) -> str:
        pass

    @classmethod
    def validate_float(cls, value: Any):
        """Note that int values are accepted."""
        pass

    @classmethod
    def validate_int(cls, value):
        pass

    @classmethod
    def validate_float_in_range(cls, value, min_inclusive, max_inclusive):
        pass

    @classmethod
    def validate_int_in_range(cls, value, min_inclusive, max_inclusive):
        pass

    @classmethod
    def validate_string(cls, value):
        pass


class BaseFloatType(BaseSimpleType):
    @classmethod
    def convert_from_xml(cls, str_value):
        return float(str_value)

    @classmethod
    def convert_to_xml(cls, value):
        pass

    @classmethod
    def validate(cls, value):
        pass


class BaseIntType(BaseSimpleType):
    @classmethod
    def convert_from_percent_literal(cls, str_value):
        int_str = str_value.replace("%", "")
        return int(int_str)

    @classmethod
    def convert_from_xml(cls, str_value):
        return int(str_value)

    @classmethod
    def convert_to_xml(cls, value):
        pass

    @classmethod
    def validate(cls, value):
        pass


class BaseStringType(BaseSimpleType):
    @classmethod
    def convert_from_xml(cls, str_value):
        return str_value

    @classmethod
    def convert_to_xml(cls, value):
        pass

    @classmethod
    def validate(cls, value):
        pass


class BaseStringEnumerationType(BaseStringType):
    @classmethod
    def validate(cls, value):
        pass


class XsdAnyUri(BaseStringType):
    """
    There's a regular expression this is supposed to meet but so far thinking
    spending cycles on validating wouldn't be worth it for the number of
    programming errors it would catch.
    """


class XsdBoolean(BaseSimpleType):
    @classmethod
    def convert_from_xml(cls, str_value):
        if str_value not in ("1", "0", "true", "false"):
            raise InvalidXmlError(
                "value must be one of '1', '0', 'true' or 'false', got '%s'" % str_value
            )
        return str_value in ("1", "true")

    @classmethod
    def convert_to_xml(cls, value):
        pass

    @classmethod
    def validate(cls, value):
        pass


class XsdDouble(BaseFloatType):
    pass


class XsdId(BaseStringType):
    """
    String that must begin with a letter or underscore and cannot contain any
    colons. Not fully validated because not used in external API.
    """


class XsdInt(BaseIntType):
    @classmethod
    def validate(cls, value):
        pass


class XsdLong(BaseIntType):
    @classmethod
    def validate(cls, value):
        pass


class XsdString(BaseStringType):
    pass


class XsdStringEnumeration(BaseStringEnumerationType):
    """
    Set of enumerated xsd:string values.
    """


class XsdToken(BaseStringType):
    """
    xsd:string with whitespace collapsing, e.g. multiple spaces reduced to
    one, leading and trailing space stripped.
    """


class XsdTokenEnumeration(BaseStringEnumerationType):
    """
    xsd:string with whitespace collapsing, e.g. multiple spaces reduced to
    one, leading and trailing space stripped.
    """


class XsdUnsignedByte(BaseIntType):
    @classmethod
    def validate(cls, value):
        pass


class XsdUnsignedInt(BaseIntType):
    @classmethod
    def validate(cls, value):
        pass


class XsdUnsignedShort(BaseIntType):
    @classmethod
    def validate(cls, value):
        pass


class ST_Angle(XsdInt):
    """
    Valid values for `rot` attribute on `<a:xfrm>` element. 60000ths of
    a degree rotation.
    """

    DEGREE_INCREMENTS = 60000
    THREE_SIXTY = 360 * DEGREE_INCREMENTS

    @classmethod
    def convert_from_xml(cls, str_value: str) -> float:
        rot = int(str_value) % cls.THREE_SIXTY
        return float(rot) / cls.DEGREE_INCREMENTS

    @classmethod
    def convert_to_xml(cls, value):
        """
        Convert signed angle float like -42.42 to int 60000 per degree,
        normalized to positive value.
        """
        pass

    @classmethod
    def validate(cls, value):
        pass


class ST_AxisUnit(XsdDouble):
    """
    Valid values for val attribute on c:majorUnit and others.
    """

    @classmethod
    def validate(cls, value):
        pass


class ST_BarDir(XsdStringEnumeration):
    """
    Valid values for <c:barDir val="?"> attribute
    """

    BAR = "bar"
    COL = "col"

    _members = (BAR, COL)


class ST_BubbleScale(BaseIntType):
    """
    String value is an integer in range 0-300, representing a percent,
    optionally including a '%' suffix.
    """

    @classmethod
    def convert_from_xml(cls, str_value):
        if "%" in str_value:
            return cls.convert_from_percent_literal(str_value)
        return super(ST_BubbleScale, cls).convert_from_xml(str_value)

    @classmethod
    def validate(cls, value):
        pass


class ST_ContentType(XsdString):
    """
    Has a pretty wicked regular expression it needs to match in the schema,
    but figuring it's not worth the trouble or run time to identify
    a programming error (as opposed to a user/runtime error).
    """

    pass


class ST_Coordinate(BaseSimpleType):
    @classmethod
    def convert_from_xml(cls, str_value):
        if "i" in str_value or "m" in str_value or "p" in str_value:
            return ST_UniversalMeasure.convert_from_xml(str_value)
        return Emu(int(str_value))

    @classmethod
    def convert_to_xml(cls, value):
        pass

    @classmethod
    def validate(cls, value):
        pass


class ST_Coordinate32(BaseSimpleType):
    """
    xsd:union of ST_Coordinate32Unqualified, ST_UniversalMeasure
    """

    @classmethod
    def convert_from_xml(cls, str_value):
        if "i" in str_value or "m" in str_value or "p" in str_value:
            return ST_UniversalMeasure.convert_from_xml(str_value)
        return ST_Coordinate32Unqualified.convert_from_xml(str_value)

    @classmethod
    def convert_to_xml(cls, value):
        pass

    @classmethod
    def validate(cls, value):
        pass


class ST_Coordinate32Unqualified(XsdInt):
    @classmethod
    def convert_from_xml(cls, str_value):
        return Emu(int(str_value))


class ST_CoordinateUnqualified(XsdLong):
    @classmethod
    def validate(cls, value):
        pass


class ST_Direction(XsdTokenEnumeration):
    """Valid values for `<p:ph orient="...">` attribute."""

    HORZ = "horz"
    VERT = "vert"

    _members = (HORZ, VERT)


class ST_DrawingElementId(XsdUnsignedInt):
    pass


class ST_Extension(XsdString):
    """
    Has a regular expression it needs to match in the schema, but figuring
    it's not worth the trouble or run time to identify a programming error
    (as opposed to a user/runtime error).
    """

    pass


class ST_GapAmount(BaseIntType):
    """
    String value is an integer in range 0-500, representing a percent,
    optionally including a '%' suffix.
    """

    @classmethod
    def convert_from_xml(cls, str_value):
        if "%" in str_value:
            return cls.convert_from_percent_literal(str_value)
        return super(ST_GapAmount, cls).convert_from_xml(str_value)

    @classmethod
    def validate(cls, value):
        pass


class ST_Grouping(XsdStringEnumeration):
    """
    Valid values for <c:grouping val=""> attribute. Overloaded for use as
    ST_BarGrouping using same tag name.
    """

    CLUSTERED = "clustered"
    PERCENT_STACKED = "percentStacked"
    STACKED = "stacked"
    STANDARD = "standard"

    _members = (CLUSTERED, PERCENT_STACKED, STACKED, STANDARD)


class ST_HexColorRGB(BaseStringType):
    @classmethod
    def convert_to_xml(cls, value):
        """
        Keep alpha characters all uppercase just for consistency.
        """
        pass

    @classmethod
    def validate(cls, value):
        # must be string ---------------
        pass


class ST_LayoutMode(XsdStringEnumeration):
    """
    Valid values for `val` attribute on c:xMode and other elements of type
    CT_LayoutMode.
    """

    EDGE = "edge"
    FACTOR = "factor"

    _members = (EDGE, FACTOR)


class ST_LblOffset(XsdUnsignedShort):
    """
    Unsigned integer value between 0 and 1000 inclusive, with optional
    percent character ('%') suffix.
    """

    @classmethod
    def convert_from_xml(cls, str_value):
        if str_value.endswith("%"):
            return cls.convert_from_percent_literal(str_value)
        return int(str_value)

    @classmethod
    def validate(cls, value):
        pass


class ST_LineWidth(XsdInt):
    @classmethod
    def convert_from_xml(cls, str_value):
        return Emu(int(str_value))

    @classmethod
    def validate(cls, value):
        pass


class ST_MarkerSize(XsdUnsignedByte):
    @classmethod
    def validate(cls, value):
        pass


class ST_Orientation(XsdStringEnumeration):
    """Valid values for `val` attribute on c:orientation (CT_Orientation)."""

    MAX_MIN = "maxMin"
    MIN_MAX = "minMax"

    _members = (MAX_MIN, MIN_MAX)


class ST_Overlap(BaseIntType):
    """
    String value is an integer in range -100..100, representing a percent,
    optionally including a '%' suffix.
    """

    @classmethod
    def convert_from_xml(cls, str_value):
        if "%" in str_value:
            return cls.convert_from_percent_literal(str_value)
        return super(ST_Overlap, cls).convert_from_xml(str_value)

    @classmethod
    def validate(cls, value):
        pass


class ST_Percentage(BaseIntType):
    """Percentage value like 42000 or '42.0%'

    Either an integer literal representing 1000ths of a percent
    (e.g. "42000"), or a floating point literal with a '%' suffix
    (e.g. "42.0%).
    """

    @classmethod
    def convert_from_xml(cls, str_value):
        if "%" in str_value:
            return cls._convert_from_percent_literal(str_value)
        return int(str_value) / 100000.0

    @classmethod
    def convert_to_xml(cls, value):
        pass

    @classmethod
    def validate(cls, value):
        pass

    @classmethod
    def _convert_from_percent_literal(cls, str_value):
        float_part = str_value[:-1]  # trim off '%' character
        return float(float_part) / 100.0


class ST_PlaceholderSize(XsdTokenEnumeration):
    """
    Valid values for <p:ph> sz (size) attribute
    """

    FULL = "full"
    HALF = "half"
    QUARTER = "quarter"

    _members = (FULL, HALF, QUARTER)


class ST_PositiveCoordinate(XsdLong):
    @classmethod
    def convert_from_xml(cls, str_value):
        int_value = super(ST_PositiveCoordinate, cls).convert_from_xml(str_value)
        return Emu(int_value)

    @classmethod
    def validate(cls, value):
        pass


class ST_PositiveFixedAngle(ST_Angle):
    """Valid values for `a:lin@ang`.

    60000ths of a degree rotation, constained to positive angles less than
    360 degrees.
    """

    @classmethod
    def convert_to_xml(cls, degrees):
        """Convert signed angle float like -427.42 to int 60000 per degree.

        Value is normalized to a positive value less than 360 degrees.
        """
        pass


class ST_PositiveFixedPercentage(ST_Percentage):
    """Percentage value between 0 and 100% like 42000 or '42.0%'

    Either an integer literal representing 1000ths of a percent
    (e.g. "42000"), or a floating point literal with a '%' suffix
    (e.g. "42.0%). Value is constrained to range of 0% to 100%. The source
    value is a float between 0.0 and 1.0.
    """

    @classmethod
    def validate(cls, value):
        pass


class ST_RelationshipId(XsdString):
    pass


class ST_SlideId(XsdUnsignedInt):
    @classmethod
    def validate(cls, value):
        pass


class ST_SlideSizeCoordinate(BaseIntType):
    @classmethod
    def convert_from_xml(cls, str_value):
        return Emu(str_value)

    @classmethod
    def validate(cls, value):
        pass


class ST_Style(XsdUnsignedByte):
    @classmethod
    def validate(cls, value):
        pass


class ST_TargetMode(XsdString):
    """
    The valid values for the ``TargetMode`` attribute in a Relationship
    element, either 'External' or 'Internal'.
    """

    @classmethod
    def validate(cls, value):
        pass


class ST_TextFontScalePercentOrPercentString(BaseFloatType):
    """
    Valid values for the `fontScale` attribute of ``<a:normAutofit>``.
    Translates to a float value.
    """

    @classmethod
    def convert_from_xml(cls, str_value):
        if str_value.endswith("%"):
            return float(str_value[:-1])  # trim off '%' character
        return int(str_value) / 1000.0

    @classmethod
    def convert_to_xml(cls, value):
        pass

    @classmethod
    def validate(cls, value):
        pass


class ST_TextFontSize(BaseIntType):
    @classmethod
    def validate(cls, value):
        pass


class ST_TextIndentLevelType(BaseIntType):
    @classmethod
    def validate(cls, value):
        pass


class ST_TextSpacingPercentOrPercentString(BaseFloatType):
    @classmethod
    def convert_from_xml(cls, str_value):
        if str_value.endswith("%"):
            return cls._convert_from_percent_literal(str_value)
        return int(str_value) / 100000.0

    @classmethod
    def _convert_from_percent_literal(cls, str_value):
        float_part = str_value[:-1]  # trim off '%' character
        percent_value = float(float_part)
        lines_value = percent_value / 100.0
        return lines_value

    @classmethod
    def convert_to_xml(cls, value):
        """
        1.75 -> '175000'
        """
        pass

    @classmethod
    def validate(cls, value):
        pass


class ST_TextSpacingPoint(BaseIntType):
    @classmethod
    def convert_from_xml(cls, str_value):
        """
        Reads string integer centipoints, returns |Length| value.
        """
        return Centipoints(int(str_value))

    @classmethod
    def convert_to_xml(cls, value):
        pass

    @classmethod
    def validate(cls, value):
        pass


class ST_TextTypeface(XsdString):
    pass


class ST_TextWrappingType(XsdTokenEnumeration):
    """
    Valid values for <a:bodyPr wrap=""> attribute
    """

    NONE = "none"
    SQUARE = "square"

    _members = (NONE, SQUARE)


class ST_UniversalMeasure(BaseSimpleType):
    @classmethod
    def convert_from_xml(cls, str_value):
        float_part, units_part = str_value[:-2], str_value[-2:]
        quantity = float(float_part)
        multiplier = {
            "mm": 36000,
            "cm": 360000,
            "in": 914400,
            "pt": 12700,
            "pc": 152400,
            "pi": 152400,
        }[units_part]
        emu_value = Emu(int(round(quantity * multiplier)))
        return emu_value
