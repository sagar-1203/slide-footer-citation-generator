import aspose.slides as slides
import aspose.pydrawing as drawing
import math

def get_the_max_font_in_column_or_row_wise(min_font, max_font, footer_shape, text_list):
    """
    Find the maximum fitting font size for footer text in row-wise or column-wise layout.

    Args:
        min_font (int): Minimum font size to test.
        max_font (int): Maximum font size to test.
        footer_shape (dict): Dictionary containing footer shape dimensions (x, y, width, height).
        text_list (list[str]): List of footer text entries.

    Returns:
        tuple: (final_font_size (int), add_footer_row_wise (bool))
            - final_font_size: The maximum font size that fits, or -1 if none fit.
            - add_footer_row_wise: True if row-wise layout chosen, False if column-wise.
    """
    # Prepare text for row-wise arrangement (all sources joined in one line)
    text = "Sources: " + "; ".join(text_list)

    # Check row-wise fitting
    status_row, font_size_row = find_largest_fitting_font(
        min_font, max_font, footer_shape["width"], footer_shape["height"], [text]
    )
    print("Row-wise fitting -> status:", status_row, "font size:", font_size_row)

    # Check column-wise fitting (split text across equal width boxes)
    equal_width = (footer_shape['width'] - 5 * len(text_list)) / len(text_list)
    max_len_text = max(text_list, key=len)  # Longest text will determine column fitting
    status_column, font_size_column = find_largest_fitting_font(
        min_font, max_font, equal_width, footer_shape["height"], [max_len_text]
    )
    print("Column-wise fitting -> status:", status_column, "font size:", font_size_column)
    print("Equal width for columns:", equal_width)

    final_font_size = -1
    add_footer_row_wise = False

    # Case 1: Both row and column work → choose larger font
    if status_row and status_column:
        if font_size_row >= font_size_column:
            print("✅ Row-wise chosen (better or equal font size)")
            final_font_size = font_size_row
            add_footer_row_wise = True
        else:
            print("✅ Column-wise chosen (better font size)")
            final_font_size = font_size_column

    # Case 2: Only row-wise works
    elif status_row:
        print("✅ Only row-wise fits")
        final_font_size = font_size_row
        add_footer_row_wise = True

    # Case 3: Only column-wise works
    elif status_column:
        print("✅ Only column-wise fits")
        final_font_size = font_size_column

    # Case 4: Neither works
    else:
        print("❌ Cannot add footer (font size too small)")
        final_font_size = -1

    return final_font_size, add_footer_row_wise


def can_text_list_fit_in_area(font_size, box_width, box_height, lines_of_text, fill_ratio=0.95):
    """
    Check if a list of text lines can fit inside a given rectangular area at a given font size.

    Args:
        font_size (int or float): Font size in points.
        box_width (float): Width of the bounding box in points.
        box_height (float): Height of the bounding box in points.
        lines_of_text (list[str]): List of text lines to fit.
        fill_ratio (float, optional): Adjustment factor (0–1) for usable width. Default is 0.95.

    Returns:
        bool: True if all lines fit inside the area, False otherwise.
    """
    try:
        # Approximate character width and line height based on font size
        avg_character_width = 0.6 * font_size
        line_height = 1.3 * font_size

        print(f"\n🔎 Checking fit for font_size={font_size}, box=({box_width}x{box_height})")

        # Calculate available capacity
        max_chars_per_line = int((box_width // avg_character_width) * fill_ratio)
        max_lines_in_box = round(box_height / line_height)

        print("➡ Capacity: max_chars_per_line =", max_chars_per_line, 
              ", max_lines_in_box =", max_lines_in_box)

        # Check if box has enough lines for total text
        if max_lines_in_box >= len(lines_of_text):
            remaining_lines = max_lines_in_box
            print("Starting with remaining_lines =", remaining_lines)

            for line in lines_of_text:
                required_lines_for_text = math.ceil(len(line) / max_chars_per_line)
                print(f"  Line length={len(line)} → requires {required_lines_for_text} line(s)")

                if remaining_lines >= required_lines_for_text:
                    remaining_lines -= required_lines_for_text
                else:
                    print("❌ Not enough space for this line")
                    return False

            print("✅ Text fits. Remaining lines =", remaining_lines)
            return True
        else:
            print("❌ Box cannot hold even the number of lines")
            return False

    except Exception as e:
        print("==" * 25)
        print("⚠ Exception in can_text_list_fit_in_area:", str(e))
        print("==" * 25)
        return False



def find_largest_fitting_font(min_font_size, max_font_size, box_width, box_height, lines_of_text):
    """
    Find the largest font size that allows a list of text lines 
    to fit within a given rectangular area.

    Args:
        min_font_size (float): Minimum allowed font size.
        max_font_size (float): Maximum allowed font size.
        box_width (float): Width of the bounding box in points.
        box_height (float): Height of the bounding box in points.
        lines_of_text (list[str]): List of text lines to test.

    Returns:
        tuple:
            bool: True if text fits at some font size, False otherwise.
            float: Best fitting font size found (rounded to 2 decimals).
    """
    # Start from the smaller of "height-per-line" or max_font_size
    font_size = min((box_height / len(lines_of_text)), max_font_size)

    print(f"\n🔎 Finding largest fitting font in box=({box_width}x{box_height}), "
          f"lines={len(lines_of_text)}, "
          f"range=({min_font_size}–{max_font_size})")
    print("Initial font_size candidate:", round(font_size, 2))
    print("lines_of_text:", lines_of_text)

    # Step down gradually until text fits or below min font size
    while font_size >= min_font_size:
        fits = can_text_list_fit_in_area(font_size, box_width, box_height, lines_of_text)
        if fits:
            print(f"✅ Fits at font_size={round(font_size, 2)}")
            return True, round(font_size, 2)
        font_size -= 0.1  # step down

    print(f"❌ No fit found down to min_font_size={min_font_size}")
    return False, round(font_size, 2)



print(slides.SchemeColor)
SCHEME_COLOR_MAP = {
    "CH_TEXT1": slides.SchemeColor.TEXT1,
    "CH_TEXT2": slides.SchemeColor.TEXT2,
    "CH_BACKGROUND1": slides.SchemeColor.BACKGROUND1,
    "CH_BACKGROUND2": slides.SchemeColor.BACKGROUND2,
    "CH_ACCENT1": slides.SchemeColor.ACCENT1,
    "CH_ACCENT2": slides.SchemeColor.ACCENT2,
    "CH_ACCENT3": slides.SchemeColor.ACCENT3,
    "CH_ACCENT4": slides.SchemeColor.ACCENT4,
    "CH_ACCENT5": slides.SchemeColor.ACCENT5,
    "CH_ACCENT6": slides.SchemeColor.ACCENT6,
    "CH_HYPERLINK": slides.SchemeColor.HYPERLINK,
    "CH_FOLLOWED_HYPERLINK": slides.SchemeColor.FOLLOWED_HYPERLINK,
}

def apply_font_color_from_config(portion_format, footer_config):
    """
    Apply font color to a text portion based on footer configuration.

    Args:
        portion_format (slides.PortionFormat): Portion format object to update.
        footer_config (dict): Configuration containing font_body_color settings.

    Returns:
        None
    """
    color_cfg = footer_config.get("font_body_color", {}) or {}
    color_type = color_cfg.get("color_type")
    print(f"\n🎨 Applying font color, color_type={color_type}")

    portion_format.fill_format.fill_type = slides.FillType.SOLID

    if color_type == "CT_SCHEME":
        # Use scheme color mapping (default: TEXT1)
        scheme_key = color_cfg.get("scheme_color", "CH_TEXT1")
        scheme_color = SCHEME_COLOR_MAP.get(scheme_key, slides.SchemeColor.TEXT1)
        print(f"→ Using scheme color: {scheme_key} ({scheme_color})")

        portion_format.fill_format.solid_fill_color.color_type = slides.ColorType.SCHEME
        portion_format.fill_format.solid_fill_color.scheme_color = scheme_color

    elif color_type == "CT_RGB":
        # Parse RGB hex string (default: black)
        hex_color = color_cfg.get("color_name", "000000").lstrip("#")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        print(f"→ Using RGB color: #{hex_color} -> ({r},{g},{b})")

        portion_format.fill_format.solid_fill_color.color_type = slides.ColorType.RGB
        portion_format.fill_format.solid_fill_color.color = drawing.Color.from_argb(r, g, b)

    elif color_type == "CT_TEXT":
        # Use inherited default text color
        print("→ Using CT_TEXT (inheriting default text color).")
        portion_format.fill_format.solid_fill_color.color_type = slides.ColorType.NOT_DEFINED

    elif color_type == "CT_NONE":
        # Transparent text
        print("→ Using CT_NONE (transparent text).")
        portion_format.fill_format.fill_type = slides.FillType.NO_FILL

    elif color_type == "CT_PRESET":
        # Named preset color like "Red", "Blue"
        preset_name = color_cfg.get("color_name", "Black")
        print(f"→ Using CT_PRESET: {preset_name}")
        portion_format.fill_format.solid_fill_color.color_type = slides.ColorType.PRESET
        portion_format.fill_format.solid_fill_color.preset_color = getattr(
            slides.PresetColor, preset_name.upper(), slides.PresetColor.BLACK
        )

    else:
        # Fallback to black
        print("⚠️ Unknown color_type. Using default black.")
        portion_format.fill_format.solid_fill_color.color_type = slides.ColorType.RGB
        portion_format.fill_format.solid_fill_color.color = drawing.Color.black

    # (Optional) if you want to use luminance_score later for adjustments
    if "luminance_score" in color_cfg:
        lum = color_cfg["luminance_score"]
        print(f"ℹ️ Luminance score available: {lum}")


def apply_font_color_from_config1(portion_format, footer_config):
    """
    Apply font color to a text portion based on footer configuration.

    Args:
        portion_format (slides.PortionFormat): Portion format object to update.
        footer_config (dict): Configuration containing color type and values.

    Returns:
        None
    """
    color_type = footer_config.get("color_type")
    print(f"\n🎨 Applying font color, color_type={color_type}")

    if color_type == "CT_SCHEME":
        # Use scheme color mapping (default: TEXT1)
        scheme_key = footer_config.get("scheme_color", "CH_TEXT1")
        scheme_color = SCHEME_COLOR_MAP.get(scheme_key, slides.SchemeColor.TEXT1)
        print(f"→ Using scheme color: {scheme_key} ({scheme_color})")

        portion_format.fill_format.fill_type = slides.FillType.SOLID
        portion_format.fill_format.solid_fill_color.color_type = slides.ColorType.SCHEME
        portion_format.fill_format.solid_fill_color.scheme_color = scheme_color

    elif color_type == "CT_RGB":
        # Parse RGB hex string (default: black)
        hex_color = footer_config.get("color_name", "000000").lstrip("#")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        print(f"→ Using RGB color: #{hex_color} -> ({r},{g},{b})")

        portion_format.fill_format.fill_type = slides.FillType.SOLID
        portion_format.fill_format.solid_fill_color.color_type = slides.ColorType.RGB
        portion_format.fill_format.solid_fill_color.color = drawing.Color.from_argb(r, g, b)

    else:
        # Fallback to default black
        print("⚠️ Unknown color_type. Using default black.")
        portion_format.fill_format.fill_type = slides.FillType.SOLID
        portion_format.fill_format.solid_fill_color.color = drawing.Color.black


from aspose.slides import TextAutofitType
import aspose.pydrawing as draw


# from aspose.slides import ShapeType, TextAutofitType
def add_rectangle_box(
    slide, footer_x, footer_y, footer_width, footer_height,
    footer_config, text, num_columns=3, font_size=None
):
    """
    Add a rectangle text box to the slide for footer content.

    Args:
        slide (slides.Slide): The slide where the box is added.
        footer_x (float): X position of the footer box.
        footer_y (float): Y position of the footer box.
        footer_width (float): Width of the footer box.
        footer_height (float): Height of the footer box.
        footer_config (dict): Footer configuration (colors, style).
        text (str): Footer text to display.
        num_columns (int, optional): Number of text columns. Default is 3.
        font_size (float, optional): Explicit font size (if set).

    Returns:
        footer_shape (slides.Shape): The created rectangle shape.
    """
    # 1. Create rectangle shape
    footer_shape = slide.shapes.add_auto_shape(
        slides.ShapeType.RECTANGLE,
        footer_x, footer_y,
        footer_width, footer_height
    )

    # 2. Set text box margins & wrapping
    tf = footer_shape.text_frame
    tf.text_frame_format.margin_left = 2.5
    tf.text_frame_format.margin_right = 2.5
    tf.text_frame_format.margin_top = 2.5
    tf.text_frame_format.margin_bottom = 2.5
    tf.text_frame_format.autofit_type = TextAutofitType.NORMAL
    tf.text_frame_format.wrap_text = slides.NullableBool.TRUE

    # 3. Set column layout
    tf.text_frame_format.column_count = num_columns
    tf.text_frame_format.column_spacing = 7  # spacing in points

    # 4. Insert footer text
    tf.text = text

    # 5. Remove box background/line
    footer_shape.fill_format.fill_type = slides.FillType.NO_FILL
    footer_shape.line_format.fill_format.fill_type = slides.FillType.NO_FILL

    # 6. Text formatting (left align, color, font size)
    para = tf.paragraphs[0]
    para.paragraph_format.alignment = slides.TextAlignment.LEFT

    portion_fmt = para.portions[0].portion_format
    portion_fmt.fill_format.fill_type = slides.FillType.SOLID
    portion_fmt.fill_format.solid_fill_color.color = draw.Color.black

    # Apply color from config (overwrites default black if configured)
    apply_font_color_from_config(portion_fmt, footer_config)

    # Apply font size if provided
    if font_size:
        tf.text_frame_format.autofit_type = TextAutofitType.SHAPE
        portion_fmt.font_height = font_size

    print(f"📦 Added footer box at ({footer_x}, {footer_y}), size=({footer_width}x{footer_height}), font={font_size}")
    return footer_shape
