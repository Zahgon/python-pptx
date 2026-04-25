"""lxml custom element classes for picture-related XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast
from xml.sax.saxutils import escape

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.shapes.shared import BaseShapeElement
from pptx.oxml.xmlchemy import BaseOxmlElement, OneAndOnlyOne

if TYPE_CHECKING:
    from pptx.oxml.shapes.shared import CT_ShapeProperties
    from pptx.util import Length


class CT_Picture(BaseShapeElement):
    """`p:pic` element.

    Represents a picture shape (an image placement on a slide).
    """

    nvPicPr = OneAndOnlyOne("p:nvPicPr")
    blipFill = OneAndOnlyOne("p:blipFill")
    spPr: CT_ShapeProperties = OneAndOnlyOne("p:spPr")  # pyright: ignore[reportAssignmentType]

    @property
    def blip_rId(self) -> str | None:
        """Value of `p:blipFill/a:blip/@r:embed`.

        Returns |None| if not present.
        """
        pass

    def crop_to_fit(self, image_size, view_size):
        """
        Set cropping values in `p:blipFill/a:srcRect` such that an image of
        *image_size* will stretch to exactly fit *view_size* when its aspect
        ratio is preserved.
        """
        pass

    def get_or_add_ln(self):
        """
        Return the <a:ln> grandchild element, newly added if not present.
        """
        pass

    @property
    def ln(self):
        """
        ``<a:ln>`` grand-child element or |None| if not present
        """
        pass

    @classmethod
    def new_ph_pic(cls, id_, name, desc, rId):
        """
        Return a new `p:pic` placeholder element populated with the supplied
        parameters.
        """
        pass

    @classmethod
    def new_pic(cls, shape_id, name, desc, rId, x, y, cx, cy):
        """Return new `<p:pic>` element tree configured with supplied parameters."""
        pass

    @classmethod
    def new_video_pic(
        cls,
        shape_id: int,
        shape_name: str,
        video_rId: str,
        media_rId: str,
        poster_frame_rId: str,
        x: Length,
        y: Length,
        cx: Length,
        cy: Length,
    ) -> CT_Picture:
        """Return a new `p:pic` populated with the specified video."""
        pass

    @property
    def srcRect_b(self):
        """Value of `p:blipFill/a:srcRect/@b` or 0.0 if not present."""
        pass

    @srcRect_b.setter
    def srcRect_b(self, value):
        pass

    @property
    def srcRect_l(self):
        """Value of `p:blipFill/a:srcRect/@l` or 0.0 if not present."""
        pass

    @srcRect_l.setter
    def srcRect_l(self, value):
        pass

    @property
    def srcRect_r(self):
        """Value of `p:blipFill/a:srcRect/@r` or 0.0 if not present."""
        pass

    @srcRect_r.setter
    def srcRect_r(self, value):
        pass

    @property
    def srcRect_t(self):
        """Value of `p:blipFill/a:srcRect/@t` or 0.0 if not present."""
        pass

    @srcRect_t.setter
    def srcRect_t(self, value):
        pass

    def _fill_cropping(self, image_size, view_size):
        """
        Return a (left, top, right, bottom) 4-tuple containing the cropping
        values required to display an image of *image_size* in *view_size*
        when stretched proportionately. Each value is a percentage expressed
        as a fraction of 1.0, e.g. 0.425 represents 42.5%. *image_size* and
        *view_size* are each (width, height) pairs.
        """
        pass

    @classmethod
    def _pic_ph_tmpl(cls):
        pass

    @classmethod
    def _pic_tmpl(cls):
        pass

    @classmethod
    def _pic_video_tmpl(cls):
        pass

    def _srcRect_x(self, attr_name):
        """
        Value of `p:blipFill/a:srcRect/@{attr_name}` or 0.0 if not present.
        """
        pass


class CT_PictureNonVisual(BaseOxmlElement):
    """
    ``<p:nvPicPr>`` element, containing non-visual properties for a picture
    shape.
    """

    cNvPr = OneAndOnlyOne("p:cNvPr")
    nvPr = OneAndOnlyOne("p:nvPr")
