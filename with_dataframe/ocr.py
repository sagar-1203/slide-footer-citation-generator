from aspose.pydrawing import imaging
from aspose.pydrawing import Color
import aspose.slides as slides
import easyocr
import numpy as np
from PIL import Image
import io

dim = 72


def ppt_to_images_in_memory(ppt_path, scale=2):
    """Returns list of (PIL.Image, dpi_x, dpi_y) for each slide."""
    with slides.Presentation(ppt_path) as pres:
        slide_width_pt = pres.slide_size.size.width
        slide_height_pt = pres.slide_size.size.height

        # Render first slide to get pixel dimensions
        first_image = pres.slides[0].get_thumbnail(scale, scale)
        img_width_px = first_image.width
        img_height_px = first_image.height

        # Calculate DPI
        dpi_x = img_width_px / (slide_width_pt / dim)
        dpi_y = img_height_px / (slide_height_pt / dim)

        images = []
        for slide in pres.slides:
            image = slide.get_thumbnail(scale, scale)
            img_stream = io.BytesIO()
            image.save(img_stream, imaging.ImageFormat.png)  # ✅ FIXED
            img_stream.seek(0)
            pil_img = Image.open(img_stream).convert("RGB")
            images.append(pil_img)

    return images, dpi_x, dpi_y


def ocr_on_ppt_in_memory(ppt_path, work_area):
    # pres = slides.Presentation(ppt_path)
    # slide = pres.slides[0]  # pick the slide you want

    # Step 1: Get images and DPI
    images, dpi_x, dpi_y = ppt_to_images_in_memory(ppt_path, scale=2)

    # Step 2: Convert work_area bottom (points → pixels)
    work_area_bottom_px = work_area["bottom"] * (dpi_y / dim)
    print(f"Work area bottom in pixels: {work_area_bottom_px}")

    # Step 3: Run OCR directly on images
    reader = easyocr.Reader(['en'])
    for i, img in enumerate(images):
        np_img = np.array(img)  # EasyOCR expects numpy array
        results = reader.readtext(np_img)

        filtered_results = []
        for (bbox, text, prob) in results:
            top_y = min(bbox[0][1], bbox[1][1])
            if top_y > work_area_bottom_px:
                filtered_results.append((bbox, text, prob))

        # print(f"--- Slide {i+1} ---")
        # for bbox, text, prob in filtered_results:
        #     print(f"Text: {text}, TopY: {min(bbox[0][1], bbox[1][1])}, BBox: {bbox}")
        # for bbox, text, prob in filtered_results:
        #     top_y = min(bbox[0][1], bbox[1][1])
        #     bottom_y = max(bbox[2][1], bbox[3][1])
        #     left_x = min(bbox[0][0], bbox[3][0])
        #     right_x = max(bbox[1][0], bbox[2][0])

        #     width = right_x - left_x
        #     height = bottom_y - top_y

        #     print(f"Text: {text}, TopY: {top_y}, Width: {width}, Height: {height}, BBox: {bbox}")
        shape_list = []
        print(f"--- Slide {i + 1} ---")
        for bbox, text, prob in filtered_results:
            # Pixel coordinates
            top_y_px = min(bbox[0][1], bbox[1][1])
            bottom_y_px = max(bbox[2][1], bbox[3][1])
            left_x_px = min(bbox[0][0], bbox[3][0])
            right_x_px = max(bbox[1][0], bbox[2][0])

            width_px = right_x_px - left_x_px
            height_px = bottom_y_px - top_y_px

            # Convert to Aspose points
            x_pt = (left_x_px / dpi_x) * dim
            y_pt = (top_y_px / dpi_y) * dim
            width_pt = (width_px / dpi_x) * dim
            height_pt = (height_px / dpi_y) * dim

            print(f"Text: {text}")
            print(f" Pixel coords -> X: {left_x_px:.2f}, Y: {top_y_px:.2f}, W: {width_px:.2f}, H: {height_px:.2f}")
            print(
                f" Aspose coords -> X: {x_pt:.2f} pt, Y: {y_pt:.2f} pt, W: {width_pt:.2f} pt, H: {height_pt:.2f} pt\n")
            shape_list.append((x_pt, y_pt, width_pt, height_pt, text))
            # # Add a rectangle
            # shape = slide.shapes.add_auto_shape(
            #     slides.ShapeType.RECTANGLE,
            #     x_pt,
            #     y_pt,
            #     width_pt,
            #     height_pt
            # )
            # # Optional: set some style so you can see it
            # shape.fill_format.fill_type = slides.FillType.SOLID
            # shape.fill_format.solid_fill_color.color = Color.from_argb(255, 200, 200, 255)
            # # Add OCR text to the shape
            # shape.text_frame.text = text

            # # Set custom properties
            # # shape.custom_data.set_custom_property("content", text)
            # # shape.custom_data.set_custom_property("image", True)
            # # shape.tags.add("content", ocr_text)
            # # shape.tags.add("image", "true")  # tags store strings
            # pres.save("output.pptx", slides.export.SaveFormat.PPTX)
        return shape_list


# Example usage
# ppt_path = file_path
# work_area = {"bottom": 461.36952209472656}  # from Aspose (points)
# ocr_on_ppt_in_memory(ppt_path, work_area)


import aspose.slides as slides
import aspose.pydrawing as draw


def add_ocr_shape(slide, x_pt, y_pt, width_pt, height_pt, ocr_text):
    # Add shape
    shape = slide.shapes.add_auto_shape(
        slides.ShapeType.RECTANGLE,
        x_pt,
        y_pt,
        width_pt,
        height_pt
    )

    # Fill color
    para = shape.text_frame.paragraphs[0]
    para.portions[0].portion_format.font_height = 3
    # para.portions[0].portion_format.font_bold = True
    para.portions[0].portion_format.fill_format.fill_type = slides.FillType.SOLID
    para.portions[0].portion_format.fill_format.solid_fill_color.color = draw.Color.white

    # Add OCR text
    shape.text_frame.text = ocr_text

    return shape


def merge_overlapping_boxes(boxes, overlap_threshold=0.1):
    """
    Merge overlapping bounding boxes.
    boxes: list of (x, y, w, h, text)
    overlap_threshold: fraction of area overlap to consider merge-worthy
    """
    merged = []

    def boxes_overlap(b1, b2):
        x1, y1, w1, h1, _ = b1
        x2, y2, w2, h2, _ = b2

        # Coordinates of intersection
        ix1 = max(x1, x2)
        iy1 = max(y1, y2)
        ix2 = min(x1 + w1, x2 + w2)
        iy2 = min(y1 + h1, y2 + h2)

        if ix1 < ix2 and iy1 < iy2:
            intersection_area = (ix2 - ix1) * (iy2 - iy1)
            area1 = w1 * h1
            area2 = w2 * h2
            # Compare overlap relative to smaller box
            return intersection_area / min(area1, area2) > overlap_threshold
        return False

    used = [False] * len(boxes)

    for i in range(len(boxes)):
        if used[i]:
            continue

        # Start with current box
        x, y, w, h, text = boxes[i]
        merged_text = [text]

        for j in range(i + 1, len(boxes)):
            if not used[j] and boxes_overlap((x, y, w, h, text), boxes[j]):
                # Expand box to include the other
                x2, y2, w2, h2, t2 = boxes[j]
                new_x1 = min(x, x2)
                new_y1 = min(y, y2)
                new_x2 = max(x + w, x2 + w2)
                new_y2 = max(y + h, y2 + h2)

                x = new_x1
                y = new_y1
                w = new_x2 - new_x1
                h = new_y2 - new_y1

                merged_text.append(t2)
                used[j] = True

        merged.append((x, y, w, h, " ".join(merged_text)))
        used[i] = True

    return merged


def remove_shapes_by_info(slide, shapes_info):
    shape_list = list(slide.shapes)
    for shp in shape_list:
        # Check coordinates and text match
        try:
            if (abs(shp.x - shapes_info[0]) < 0.1 and
                    abs(shp.y - shapes_info[1]) < 0.1 and
                    abs(shp.width - shapes_info[2]) < 0.1 and
                    abs(shp.height - shapes_info[3]) < 0.1 and
                    getattr(shp.text_frame, "text", "") == shapes_info[4]):
                slide.shapes.remove(shp)
        except AttributeError:
            continue


