"""Objects related to mouse click and hover actions on a shape or text."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.enum.action import PP_ACTION
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.shapes import Subshape
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.action import CT_Hyperlink
    from pptx.oxml.shapes.shared import CT_NonVisualDrawingProps
    from pptx.oxml.text import CT_TextCharacterProperties
    from pptx.parts.slide import SlidePart
    from pptx.shapes.base import BaseShape
    from pptx.slide import Slide, Slides


class ActionSetting(Subshape):
    """Properties specifying how a shape or run reacts to mouse actions."""

    # -- The Subshape base class provides access to the Slide Part, which is needed to access
    # -- relationships, which is where hyperlinks live.

    def __init__(
        self,
        xPr: CT_NonVisualDrawingProps | CT_TextCharacterProperties,
        parent: BaseShape,
        hover: bool = False,
    ):
        super(ActionSetting, self).__init__(parent)
        # xPr is either a cNvPr or rPr element
        self._element = xPr
        # _hover determines use of `a:hlinkClick` or `a:hlinkHover`
        self._hover = hover

    @property
    def action(self):
        """Member of :ref:`PpActionType` enumeration, such as `PP_ACTION.HYPERLINK`.

        The returned member indicates the type of action that will result when the
        specified shape or text is clicked or the mouse pointer is positioned over the
        shape during a slide show.

        If there is no click-action or the click-action value is not recognized (is not
        one of the official `MsoPpAction` values) then `PP_ACTION.NONE` is returned.
        """
        pass

    @lazyproperty
    def hyperlink(self) -> Hyperlink:
        """
        A |Hyperlink| object representing the hyperlink action defined on
        this click or hover mouse event. A |Hyperlink| object is always
        returned, even if no hyperlink or other click action is defined.
        """
        pass

    @property
    def target_slide(self) -> Slide | None:
        """
        A reference to the slide in this presentation that is the target of
        the slide jump action in this shape. Slide jump actions include
        `PP_ACTION.FIRST_SLIDE`, `LAST_SLIDE`, `NEXT_SLIDE`,
        `PREVIOUS_SLIDE`, and `NAMED_SLIDE`. Returns |None| for all other
        actions. In particular, the `LAST_SLIDE_VIEWED` action and the `PLAY`
        (start other presentation) actions are not supported.

        A slide object may be assigned to this property, which makes the
        shape an "internal hyperlink" to the assigened slide::

            slide, target_slide = prs.slides[0], prs.slides[1]
            shape = slide.shapes[0]
            shape.target_slide = target_slide

        Assigning |None| removes any slide jump action. Note that this is
        accomplished by removing any action present (such as a hyperlink),
        without first checking that it is a slide jump action.
        """
        pass

    @target_slide.setter
    def target_slide(self, slide: Slide | None):
        pass

    def _clear_click_action(self):
        """Remove any existing click action."""
        pass

    @property
    def _hlink(self) -> CT_Hyperlink | None:
        """
        Reference to the `a:hlinkClick` or `a:hlinkHover` element for this
        click action. Returns |None| if the element is not present.
        """
        pass

    @lazyproperty
    def _slide(self):
        """
        Reference to the slide containing the shape having this click action.
        """
        pass

    @lazyproperty
    def _slide_index(self):
        """
        Position in the slide collection of the slide containing the shape
        having this click action.
        """
        pass

    @lazyproperty
    def _slides(self) -> Slides:
        """
        Reference to the slide collection for this presentation.
        """
        pass


class Hyperlink(Subshape):
    """Represents a hyperlink action on a shape or text run."""

    def __init__(
        self,
        xPr: CT_NonVisualDrawingProps | CT_TextCharacterProperties,
        parent: BaseShape,
        hover: bool = False,
    ):
        super(Hyperlink, self).__init__(parent)
        # xPr is either a cNvPr or rPr element
        self._element = xPr
        # _hover determines use of `a:hlinkClick` or `a:hlinkHover`
        self._hover = hover

    @property
    def address(self) -> str | None:
        """Read/write. The URL of the hyperlink.

        URL can be on http, https, mailto, or file scheme; others may work. Returns |None| if no
        hyperlink is defined, including when another action such as `RUN_MACRO` is defined on the
        object. Assigning |None| removes any action defined on the object, whether it is a hyperlink
        action or not.
        """
        pass

    @address.setter
    def address(self, url: str | None):
        # implements all three of add, change, and remove hyperlink
        pass

    def _get_or_add_hlink(self) -> CT_Hyperlink:
        """Get the `a:hlinkClick` or `a:hlinkHover` element for the Hyperlink object.

        The actual element depends on the value of `self._hover`. Create the element if not present.
        """
        pass

    @property
    def _hlink(self) -> CT_Hyperlink | None:
        """Reference to the `a:hlinkClick` or `h:hlinkHover` element for this click action.

        Returns |None| if the element is not present.
        """
        pass

    def _remove_hlink(self):
        """Remove the a:hlinkClick or a:hlinkHover element.

        Also drops any relationship it might have.
        """
        pass
