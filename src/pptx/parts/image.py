"""ImagePart and related objects."""

from __future__ import annotations

import hashlib
import io
import os
from typing import IO, TYPE_CHECKING, Any, cast

from PIL import Image as PIL_Image

from pptx.opc.package import Part
from pptx.opc.spec import image_content_types
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.opc.packuri import PackURI
    from pptx.package import Package
    from pptx.util import Length


class ImagePart(Part):
    """An image part.

    An image part generally has a partname matching the regex `ppt/media/image[1-9][0-9]*.*`.
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        package: Package,
        blob: bytes,
        filename: str | None = None,
    ):
        super(ImagePart, self).__init__(partname, content_type, package, blob)
        self._blob = blob
        self._filename = filename

    @classmethod
    def new(cls, package: Package, image: Image) -> ImagePart:
        """Return new |ImagePart| instance containing `image`.

        `image` is an |Image| object.
        """
        return cls(
            package.next_image_partname(image.ext),
            image.content_type,
            package,
            image.blob,
            image.filename,
        )

    @property
    def desc(self) -> str:
        """The filename associated with this image.

        Either the filename of the original image or a generic name of the form `image.ext` where
        `ext` is appropriate to the image file format, e.g. `'jpg'`. An image created using a path
        will have that filename; one created with a file-like object will have a generic name.
        """
        pass

    @property
    def ext(self) -> str:
        """File-name extension for this image e.g. `'png'`."""
        pass

    @property
    def image(self) -> Image:
        """An |Image| object containing the image in this image part.

        Note this is a `pptx.image.Image` object, not a PIL Image.
        """
        pass

    def scale(self, scaled_cx: int | None, scaled_cy: int | None) -> tuple[int, int]:
        """Return scaled image dimensions in EMU based on the combination of parameters supplied.

        If `scaled_cx` and `scaled_cy` are both |None|, the native image size is returned. If
        neither `scaled_cx` nor `scaled_cy` is |None|, their values are returned unchanged. If a
        value is provided for either `scaled_cx` or `scaled_cy` and the other is |None|, the
        missing value is calculated such that the image's aspect ratio is preserved.
        """
        pass

    @lazyproperty
    def sha1(self) -> str:
        """The 40-character SHA1 hash digest for the image binary of this image part.

        like: `"1be010ea47803b00e140b852765cdf84f491da47"`.
        """
        pass

    @property
    def _dpi(self) -> tuple[int, int]:
        """(horz_dpi, vert_dpi) pair representing the dots-per-inch resolution of this image."""
        pass

    @property
    def _native_size(self) -> tuple[Length, Length]:
        """A (width, height) 2-tuple representing the native dimensions of the image in EMU.

        Calculated based on the image DPI value, if present, assuming 72 dpi as a default.
        """
        pass

    @property
    def _px_size(self) -> tuple[int, int]:
        """A (width, height) 2-tuple representing the dimensions of this image in pixels."""
        pass


class Image(object):
    """Immutable value object representing an image such as a JPEG, PNG, or GIF."""

    def __init__(self, blob: bytes, filename: str | None):
        super(Image, self).__init__()
        self._blob = blob
        self._filename = filename

    @classmethod
    def from_blob(cls, blob: bytes, filename: str | None = None) -> Image:
        """Return a new |Image| object loaded from the image binary in `blob`."""
        pass

    @classmethod
    def from_file(cls, image_file: str | IO[bytes]) -> Image:
        """Return a new |Image| object loaded from `image_file`.

        `image_file` can be either a path (str) or a file-like object.
        """
        pass

    @property
    def blob(self) -> bytes:
        """The binary image bytestream of this image."""
        pass

    @lazyproperty
    def content_type(self) -> str:
        """MIME-type of this image, e.g. `"image/jpeg"`."""
        pass

    @lazyproperty
    def dpi(self) -> tuple[int, int]:
        """A (horz_dpi, vert_dpi) 2-tuple specifying the dots-per-inch resolution of this image.

        A default value of (72, 72) is used if the dpi is not specified in the image file.
        """
        pass

    @lazyproperty
    def ext(self) -> str:
        """Canonical file extension for this image e.g. `'png'`.

        The returned extension is all lowercase and is the canonical extension for the content type
        of this image, regardless of what extension may have been used in its filename, if any.
        """
        pass

    @property
    def filename(self) -> str | None:
        """Filename from path used to load this image, if loaded from the filesystem.

        |None| if no filename was used in loading, such as when loaded from an in-memory stream.
        """
        pass

    @lazyproperty
    def sha1(self) -> str:
        """SHA1 hash digest of the image blob."""
        pass

    @lazyproperty
    def size(self) -> tuple[int, int]:
        """A (width, height) 2-tuple specifying the dimensions of this image in pixels."""
        pass

    @property
    def _format(self) -> str | None:
        """The PIL Image format of this image, e.g. 'PNG'."""
        pass

    @lazyproperty
    def _pil_props(self) -> tuple[str | None, tuple[int, int], tuple[int, int] | None]:
        """tuple of image properties extracted from this image using Pillow."""
        pass
