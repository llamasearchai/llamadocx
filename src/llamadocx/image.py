"""
Classes for working with images in Word documents.

This module provides classes for working with images in Word documents,
including basic image manipulation and processing capabilities.
"""

import os
from pathlib import Path
from typing import Optional, Union, Tuple, List

from docx.shape import InlineShape
from docx.shared import Inches, Pt
from PIL import Image as PILImage
import numpy as np

try:
    import cv2
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False


class Image:
    """
    An image in a Word document.

    This class wraps python-docx's InlineShape class and provides additional
    functionality for image manipulation.

    Args:
        shape (InlineShape): The underlying python-docx InlineShape object

    Attributes:
        shape (InlineShape): The underlying python-docx InlineShape object
    """

    def __init__(self, shape: InlineShape) -> None:
        self.shape = shape

    @property
    def width(self) -> Optional[float]:
        """Get/set the image width in inches."""
        if self.shape.width is None:
            return None
        return self.shape.width.inches

    @width.setter
    def width(self, value: Optional[Union[int, float]]) -> None:
        if value is None:
            self.shape.width = None
        else:
            self.shape.width = Inches(value)

    @property
    def height(self) -> Optional[float]:
        """Get/set the image height in inches."""
        if self.shape.height is None:
            return None
        return self.shape.height.inches

    @height.setter
    def height(self, value: Optional[Union[int, float]]) -> None:
        if value is None:
            self.shape.height = None
        else:
            self.shape.height = Inches(value)

    def resize(
        self,
        width: Optional[Union[int, float]] = None,
        height: Optional[Union[int, float]] = None,
        keep_aspect_ratio: bool = True
    ) -> None:
        """
        Resize the image.

        Args:
            width (Union[int, float], optional): New width in inches
            height (Union[int, float], optional): New height in inches
            keep_aspect_ratio (bool): Whether to maintain aspect ratio
        """
        if width is None and height is None:
            return

        current_width = self.width
        current_height = self.height

        if keep_aspect_ratio:
            if width is not None and height is None:
                ratio = width / current_width
                height = current_height * ratio
            elif height is not None and width is None:
                ratio = height / current_height
                width = current_width * ratio
            elif width is not None and height is not None:
                width_ratio = width / current_width
                height_ratio = height / current_height
                ratio = min(width_ratio, height_ratio)
                width = current_width * ratio
                height = current_height * ratio

        if width is not None:
            self.width = width
        if height is not None:
            self.height = height


class ImageProcessor:
    """
    Class for processing images before adding them to Word documents.

    This class provides methods for basic image processing operations using
    PIL and optionally OpenCV.

    Args:
        path (Union[str, Path]): Path to the image file

    Attributes:
        path (Path): Path to the image file
        image (PILImage.Image): The PIL Image object
    """

    def __init__(self, path: Union[str, Path]) -> None:
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(f"Image file not found: {self.path}")
        self.image = PILImage.open(str(self.path))

    def resize(
        self,
        width: Optional[int] = None,
        height: Optional[int] = None,
        keep_aspect_ratio: bool = True
    ) -> "ImageProcessor":
        """
        Resize the image.

        Args:
            width (int, optional): New width in pixels
            height (int, optional): New height in pixels
            keep_aspect_ratio (bool): Whether to maintain aspect ratio

        Returns:
            ImageProcessor: self for method chaining
        """
        if width is None and height is None:
            return self

        current_width, current_height = self.image.size

        if keep_aspect_ratio:
            if width is not None and height is None:
                ratio = width / current_width
                height = int(current_height * ratio)
            elif height is not None and width is None:
                ratio = height / current_height
                width = int(current_width * ratio)
            elif width is not None and height is not None:
                width_ratio = width / current_width
                height_ratio = height / current_height
                ratio = min(width_ratio, height_ratio)
                width = int(current_width * ratio)
                height = int(current_height * ratio)

        if width is not None and height is not None:
            self.image = self.image.resize((width, height), PILImage.LANCZOS)

        return self

    def rotate(self, angle: float) -> "ImageProcessor":
        """
        Rotate the image.

        Args:
            angle (float): Rotation angle in degrees

        Returns:
            ImageProcessor: self for method chaining
        """
        self.image = self.image.rotate(angle, expand=True)
        return self

    def crop(
        self,
        left: int,
        top: int,
        right: int,
        bottom: int
    ) -> "ImageProcessor":
        """
        Crop the image.

        Args:
            left (int): Left coordinate
            top (int): Top coordinate
            right (int): Right coordinate
            bottom (int): Bottom coordinate

        Returns:
            ImageProcessor: self for method chaining
        """
        self.image = self.image.crop((left, top, right, bottom))
        return self

    def adjust_brightness(self, factor: float) -> "ImageProcessor":
        """
        Adjust image brightness.

        Args:
            factor (float): Brightness adjustment factor (1.0 = original)

        Returns:
            ImageProcessor: self for method chaining
        """
        enhancer = PILImage.ImageEnhance.Brightness(self.image)
        self.image = enhancer.enhance(factor)
        return self

    def adjust_contrast(self, factor: float) -> "ImageProcessor":
        """
        Adjust image contrast.

        Args:
            factor (float): Contrast adjustment factor (1.0 = original)

        Returns:
            ImageProcessor: self for method chaining
        """
        enhancer = PILImage.ImageEnhance.Contrast(self.image)
        self.image = enhancer.enhance(factor)
        return self

    def convert_to_grayscale(self) -> "ImageProcessor":
        """
        Convert the image to grayscale.

        Returns:
            ImageProcessor: self for method chaining
        """
        self.image = self.image.convert("L")
        return self

    def apply_filter(self, filter_name: str) -> "ImageProcessor":
        """
        Apply a predefined filter to the image.

        Args:
            filter_name (str): Name of the filter to apply
                ("blur", "sharpen", "edge_enhance", "emboss")

        Returns:
            ImageProcessor: self for method chaining
        """
        filter_map = {
            "blur": PILImage.BLUR,
            "sharpen": PILImage.SHARPEN,
            "edge_enhance": PILImage.EDGE_ENHANCE,
            "emboss": PILImage.EMBOSS,
        }
        if filter_name.lower() not in filter_map:
            raise ValueError(f"Invalid filter name: {filter_name}")
        self.image = self.image.filter(filter_map[filter_name.lower()])
        return self

    def detect_faces(self) -> List[Tuple[int, int, int, int]]:
        """
        Detect faces in the image using OpenCV.

        Returns:
            List[Tuple[int, int, int, int]]: List of face bounding boxes
                (x, y, width, height)

        Raises:
            ImportError: If OpenCV is not installed
        """
        if not HAS_CV2:
            raise ImportError(
                "OpenCV (cv2) is required for face detection. "
                "Install it with: pip install opencv-python"
            )

        # Convert PIL Image to OpenCV format
        cv_image = cv2.cvtColor(np.array(self.image), cv2.COLOR_RGB2BGR)

        # Load face cascade classifier
        cascade_path = cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        face_cascade = cv2.CascadeClassifier(cascade_path)

        # Detect faces
        gray = cv2.cvtColor(cv_image, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(
            gray,
            scaleFactor=1.1,
            minNeighbors=5,
            minSize=(30, 30)
        )

        return [(x, y, w, h) for (x, y, w, h) in faces]

    def save(
        self,
        path: Optional[Union[str, Path]] = None,
        format: Optional[str] = None,
        quality: int = 95
    ) -> None:
        """
        Save the processed image.

        Args:
            path (Union[str, Path, None]): Path to save the image to
                (if None, overwrites original)
            format (str, optional): Output format (e.g., "JPEG", "PNG")
            quality (int): Output quality for JPEG (1-100)
        """
        save_path = Path(path) if path is not None else self.path
        self.image.save(str(save_path), format=format, quality=quality)

    def __enter__(self) -> "ImageProcessor":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.image.close() 