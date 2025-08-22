import aspose.slides as slides
from aspose.slides import TextAutofitType
import aspose.pydrawing as draw
from aspose.slides import AutoShape
from utils import remove_shapes_outside_slide_dicts,is_fillable

def add_footer_shape(slide, work_area, footer_height=30, padding=0, collision_threshold=2):
    footer_left = work_area["left"]
    footer_y = work_area["bottom"]
    slide_width = slide.presentation.slide_size.size.width
    slide_height = slide.presentation.slide_size.size.height
    shapes_to_remove = []
    is_footer_present = False
    keep_footer_shape = None
    bottom_shapes = []
    threshold_ratio = 0.9
    seen_shapes = set()
    disqualifying_keywords = ["internal use", "copyright", "copy right", "external use"]
    qualifying_keywords = ["footer", "source goes here", "goes here", "disclaimer place holder", "edit source",
                           "footnotes", "foot"]
    print("footer_y ", footer_y)
    slide_shapes = list(slide.shapes)
    #     count = 0
    for shape in slide_shapes + list(slide.presentation.masters[0].shapes) + list(
            slide.presentation.layout_slides[0].shapes):  # use list() to avoid iterator issues while removing
        if shape.hidden:
            continue
        if shape.y + shape.height < footer_y:
            continue
        # Skip large shapes likely to be background or title
        if shape.height >= slide_height * 0.8:
            continue

        shape_id = getattr(shape, "id", None)
        shape_key = (shape_id if shape_id else (shape.x, shape.y, shape.width, shape.height))
        print("shape_key ", shape_key)
        # Duplicate shape: check if already seen
        if shape_key in seen_shapes:
            shape_text = shape.text_frame.text.strip().lower() if hasattr(shape,
                                                                          "text_frame") and shape.text_frame else ""
            name = shape.name.lower() if shape.name else ""

            # Check if duplicate is footer-like
            if any(keyword in shape_text for keyword in qualifying_keywords):
                print(f"Duplicate footer-like shape: '{shape_text}' at ({shape.x}, {shape.y})")
                print("shape type ", str(shape.get_type()))

                # Try to find and remove one of the duplicates from the slide
                for s in slide.shapes:
                    if hasattr(s, "id") and s.id == shape_id:
                        # slide.shapes.remove(s)
                        is_footer_present = True
                        keep_footer_shape = s
                        print("Removed duplicate shape by ID from slide.")
                        break
                    elif (s.x, s.y, s.width, s.height) == (shape.x, shape.y, shape.width, shape.height):
                        # slide.shapes.remove(s)
                        is_footer_present = True
                        keep_footer_shape = s
                        print("Removed duplicate shape by coordinates from slide.")
                        break

                # is_footer_present = True
                # keep_footer_shape = shape
            continue  # Skip processing duplicate again

        seen_shapes.add(shape_key)
        # Placeholder and name check for footer
        placeholder_type = shape.placeholder.type if hasattr(shape, "placeholder") and shape.placeholder else None
        name = shape.name.lower() if shape.name else ""

        if shape in slide_shapes and placeholder_type == slides.PlaceholderType.FOOTER or "footer" in name:
            if hasattr(shape, "text_frame") and shape.text_frame:
                shape_text = shape.text_frame.text.strip().lower()
                if len(shape_text) == 0 or any(keyword in shape_text for keyword in qualifying_keywords):
                    print("shape type ", str(shape.get_type()))
                    print(f"Footer placeholder found: '{shape_text}' at ({shape.x}, {shape.y})")
                    is_footer_present = True
                    keep_footer_shape = shape
        #             continue

        # Final fallback heuristic
        if shape.y > footer_y or ((shape.y + shape.height) > footer_y and shape.width > 1 and not is_fillable(shape)):
            if not is_footer_present and shape in slide_shapes and isinstance(shape, AutoShape) and hasattr(shape,
                                                                                                            "text_frame") and shape.text_frame:
                shape_text = shape.text_frame.text.strip().lower()
                print("shape name ", name)
                if len(shape_text) == 0 or any(keyword in shape_text for keyword in qualifying_keywords):
                    print(f"Detected footer-like shape: '{shape_text}' at ({shape.x}, {shape.y})")
                    is_footer_present = True
                    keep_footer_shape = shape
            bottom_shapes.append(shape)

    print("Estimated footer position:", footer_y)
    print("is_footer_present ", is_footer_present)
    print("bottom_shpes positions: 0")
    for shape in bottom_shapes:
        print(f"Shape at x={shape.x}, y={shape.y}, width={shape.width}, height={shape.height}")

    bottom_shapes = remove_shapes_outside_slide_dicts(bottom_shapes, slide_width, slide_height)
    # Step 2: Sort shapes from left to right
    bottom_shapes.sort(key=lambda s: s.x)
    if len(bottom_shapes) >= 2:
        prev_shape = bottom_shapes[0]  ## treat this shape as a q main shape
        for i in range(1, len(bottom_shapes)):
            curr_shape = bottom_shapes[i]
            collision_info = get_collision_info_2d_(prev_shape, curr_shape)
            print("collision_info ", collision_info)
            if collision_info['inter_vertical'] in ["G_ENCLOSED_BY_Q_V", "Q_ENCLOSED_BY_G_V"] or collision_info[
                'inter_horizontal'] in ["G_ENCLOSED_BY_Q_H", "Q_ENCLOSED_BY_G_H"]:
                print("shape prev_shape ", prev_shape.x, prev_shape.y)
                print("shape curr_shape ", curr_shape.x, curr_shape.y)
                if collision_info['inter_vertical'] == "Q_ENCLOSED_BY_G_V" or collision_info[
                    'inter_horizontal'] == "Q_ENCLOSED_BY_G_H":
                    prev_shape = curr_shape
            else:
                prev_shape = curr_shape

    # print(f"Found {len(bottom_shapes)} shapes near the footer region.")
    print("bottom_shpes positions:")
    for shape in bottom_shapes:
        print(f"Shape at x={shape.x}, y={shape.y}, width={shape.width}, height={shape.height}")

    center_x = slide_width // 2
    #     available_left = 0
    available_right = slide_width
    #     min_y = min(sh.y for sh in bottom_shapes)
    #     avg_y = sum(sh.y for sh in bottom_shapes) / len(bottom_shapes) if bottom_shapes else footer_y
    ys = [sh.y for sh in bottom_shapes if sh.width >= threshold_ratio * slide_width]
    avg_y = sum(ys) / len(ys) if ys else 0
    print("avg_y ", avg_y)
    if avg_y == 0:
        avg_y = (
            sum(sh.y for sh in sorted(bottom_shapes, key=lambda sh: sh.y)[:2]) / 2
            if len(bottom_shapes) >= 2 else
            (bottom_shapes[0].y if bottom_shapes else footer_y)
        )
        print("NEW avg_y ", avg_y)
        avg_y = sum(sh.y for sh in bottom_shapes) / len(bottom_shapes) if bottom_shapes else footer_y
        print("OLD avg_y ", avg_y)
    #     print("split position ", min_y)
    print("slide height ", slide_height)
    avg_footer_height = slide_height - avg_y
    #     max_footer_height = max(sh.height for sh in bottom_shapes) if bottom_shape else 0
    #     print("max_footer_height ",max_footer_height)
    print("avg_footer_height ", avg_footer_height)
    #     tmp_footer_height = max(max_footer_height,tmp_footer_height)
    footer_shape_list = list()
    #     Shape at x=0.0, y=448.3125305175781, width=162.9999237060547, height=91.6874771118164
    #     Shape at x=20.05072021484375, y=469.3670959472656, width=910.6329345703125, height=49.47149658203125
    #     Shape at x=691.4732055664062, y=481.5976257324219, width=194.7489776611328, height=28.75
    #     Shape at x=833.3802490234375, y=460.5278625488281, width=236.27685546875, height=21.56590461730957
    #     Shape at x=886.2222290039062, y=481.5975646972656, width=33.09008026123047, height=28.75
    if not is_footer_present:
        if len(bottom_shapes) > 1:
            bottom_shapes = [shape for shape in bottom_shapes if shape.width < slide_width * threshold_ratio]
        shape_bounds = sorted([(sh.x, sh.x + sh.width) for sh in bottom_shapes], key=lambda b: b[0])
        print("length of shape_bounds ", len(shape_bounds))
        # Optional: Merge overlapping/adjacent shapes
        # Add virtual bounds for left and right slide edges
        #         resulted_shape = get_available_gaps_from_shapes(shape_bounds,slide_width)
        #         print("resulted_shape ",resulted_shape)
        shape_bounds.insert(0, (0, 0))
        shape_bounds.append((slide_width, slide_width))
        # Step 3: Find widest horizontal gap
        max_gap = 20
        best_position = 0
        print("shape_bounds ", shape_bounds)
        for i in range(len(shape_bounds) - 1):
            gap_start = shape_bounds[i][1]
            gap_end = shape_bounds[i + 1][0]
            gap_width = gap_end - gap_start

            if gap_width > max_gap:  # and gap_width > already_present_foooter_width:
                #                 max_gap = gap_width
                best_position = gap_start  # center of the gap
                print(f"max_gap {max_gap} , best_position {best_position}, gap_width {gap_width}")
                #                 if max_gap < 20:
                #                     print("⚠️ Not enough space for footer on this slide. ",max_gap)
                #                     continue
                # Step 4: Add the footer shape
                footer_width = gap_width - 2 * collision_threshold  # min(max_gap * 0.9, 200)  # Fit 90% of gap or max 200
                #                 footer_width = gap_width
                if footer_width < 30:
                    print("⚠️ Not enough width for footer on this slide. ", footer_width)
                    continue
                # footer_height = 20
                footer_x = best_position + collision_threshold
                if footer_x < footer_left and footer_width < footer_left:
                    print("⚠️ outside to working area on left ", footer_x, footer_width)
                    continue
                elif footer_x < footer_left:
                    print("taking work area x as footer_x ")
                    footer_width -= (footer_left - footer_x)
                    footer_x = footer_left
                # footer_y = slide_height - footer_height - 10  # 10 units from bottom
                #                 for shape in bottom_shapes:
                #                     print(f"Shape at x={shape.x}, y={shape.y}, width={shape.width}, height={shape.height}")
                #                 footer_y = sum(sh.y for sh in bottom_shapes) / len(bottom_shapes) if bottom_shapes else footer_y
                footer_y_shapes = [sh.y for sh in bottom_shapes if sh.y > avg_y]
                print("footer_y_shapes ", footer_y_shapes)
                footer_y = sum(footer_y_shapes) / len(footer_y_shapes) if footer_y_shapes else footer_y
                print(
                    f"footer_width {footer_width},footer_height {avg_footer_height}, footer_x {footer_x} , footer_y {footer_y} ")
                #                 text = "Environmental Protection Agency. (2023). Air Quality Index Report 2023. EPA Publications." + "\n" + "Green Research Institute. (2023). Environmental Impact Assessment of Electric Vehicles. Journal of Environmental Studies, 15(3), 45-62." + "\n" + "Urban Planning Department. (2023). Urban Noise Study: Impact of Electric Vehicles. City Planning Review, 28(4), 112-128."
                #                 add_rectangle_box(slide,footer_x,avg_y+1,footer_width,avg_footer_height,text,3)
                footer_shape = {"width": footer_width, "height": footer_height, "x": footer_x, "y": footer_y}
                footer_shape_list.append(footer_shape)
    else:
        expand_footer_shape(keep_footer_shape, bottom_shapes, work_area, slide_width, slide_width)
        print("footer shape AFter expansion ", keep_footer_shape.x, keep_footer_shape.y, keep_footer_shape.width,
              keep_footer_shape.height)
        footer_shape = {"width": keep_footer_shape.width, "height": keep_footer_shape.height, "x": keep_footer_shape.x,
                        "y": keep_footer_shape.y}
        footer_shape_list.append(footer_shape)
        slide.shapes.remove(keep_footer_shape)

    return avg_y, bottom_shapes, footer_shape_list

def add_rectangle_box(slide, footer_x, footer_y, footer_width, footer_height, footer_config, text, num_columns=3,
                      font_size=None):
    footer_shape = slide.shapes.add_auto_shape(
        slides.ShapeType.RECTANGLE,
        footer_x,
        footer_y,
        footer_width,
        footer_height
    )
    footer_shape.text_frame.text_frame_format.margin_left = 2.5
    footer_shape.text_frame.text_frame_format.margin_right = 2.5
    footer_shape.text_frame.text_frame_format.margin_top = 2.5
    footer_shape.text_frame.text_frame_format.margin_bottom = 2.5
    footer_shape.text_frame.text_frame_format.autofit_type = TextAutofitType.NORMAL
    footer_shape.text_frame.text_frame_format.wrap_text = slides.NullableBool.TRUE
    footer_shape.text_frame.text_frame_format.column_count = num_columns
    footer_shape.text_frame.text_frame_format.column_spacing = 7
    footer_shape.text_frame.text = text
    footer_shape.fill_format.fill_type = slides.FillType.NO_FILL
    footer_shape.line_format.fill_format.fill_type = slides.FillType.NO_FILL
    footer_shape.text_frame.paragraphs[0].paragraph_format.alignment = slides.TextAlignment.LEFT
    footer_shape.text_frame.paragraphs[0].portions[0].portion_format.fill_format.fill_type = slides.FillType.SOLID
    footer_shape.text_frame.paragraphs[0].portions[
        0].portion_format.fill_format.solid_fill_color.color = draw.Color.black
    apply_font_color_from_config(footer_shape.text_frame.paragraphs[0].portions[0].portion_format, footer_config)
    if font_size:
        footer_shape.text_frame.text_frame_format.autofit_type = TextAutofitType.SHAPE
        footer_shape.text_frame.paragraphs[0].portions[0].portion_format.font_height = font_size

def expand_footer_shape(footer_shape, other_shapes, work_area, slide_width, slide_height, padding=1,
                        max_footer_height=30):
    # Footer original bounds
    x, y, width, height = footer_shape.x, footer_shape.y, footer_shape.width, footer_shape.height
    footer_bottom = y + height

    min_distance_to_shape_below = float("inf")
    min_distance_to_shape_above = float("inf")

    footer_top = footer_shape.y
    footer_bottom = footer_shape.y + footer_shape.height

    for shape in other_shapes:
        if shape == footer_shape:
            continue

        shape_top = shape.y
        shape_bottom = shape.y + shape.height

        # Footer is fully inside another shape
        if shape_top <= footer_top and shape_bottom >= footer_bottom:
            top_gap = footer_top - shape_top
            bottom_gap = shape_bottom - footer_bottom

            if top_gap < min_distance_to_shape_above:
                min_distance_to_shape_above = top_gap
            if bottom_gap < min_distance_to_shape_below:
                min_distance_to_shape_below = bottom_gap

        else:
            # Shape is fully above
            if shape_bottom <= footer_top:
                distance = footer_top - shape_bottom
                if distance < min_distance_to_shape_above:
                    min_distance_to_shape_above = distance

            # Shape is fully below
            elif shape_top >= footer_bottom:
                distance = shape_top - footer_bottom
                if distance < min_distance_to_shape_below:
                    min_distance_to_shape_below = distance

    # Default to 0 if no restriction was found
    expand_up = min_distance_to_shape_above if min_distance_to_shape_above != float("inf") else 0
    expand_down = min_distance_to_shape_below if min_distance_to_shape_below != float("inf") else 0
    print("expand up ", expand_up)
    # Total possible expansion
    total_possible_expansion = expand_up + expand_down
    current_height = footer_shape.height
    max_expandable = max_footer_height - current_height

    # Restrict expansion to not exceed max_footer_height
    if total_possible_expansion > 0 and total_possible_expansion > max_expandable:
        # Proportionally scale expansion if needed
        scale_factor = max_expandable / total_possible_expansion
        expand_up = int(expand_up * scale_factor)
        expand_down = int(expand_down * scale_factor)
    print("expand up ", expand_up)
    # Update footer shape's Y-position and height
    new_y = footer_shape.y - expand_up
    new_height = footer_shape.height + expand_up + expand_down

    # Clamp to max height just in case
    new_height = min(new_height, max_footer_height)
    footer_shape.y = new_y
    if hasattr(footer_shape, "height"):
        try:
            footer_shape.height = new_height
        except Exception as e:
            print("Failed to set footer shape height/y:", e)

    # Apply changes
    #     footer_shape.y = new_y
    #     footer_shape.height = new_height

    print("Expanded footer shape:")
    print("  New Y:", footer_shape.y)
    print("  New height:", footer_shape.height)

    # Expand upward
    #     new_y = max_y_above_footer + padding
    #     print("expandable_height ", expandable_height)
    #     footer_bottom += max(expandable_height - 2,0)

    # Expand upward
    #     new_y = max_y_above_footer + padding
    #     footer_bottom = max_y_above_footer
    #     height = footer_bottom - y
    #     footer_shape.height = height
    #     print("height fo ",height)
    # Horizontal expansion logic
    # Try expanding left
    # Work area boundaries
    work_area_left = work_area["left"]
    work_area_right = work_area["right"]

    # Try expanding left
    new_x = x
    while new_x - padding >= work_area_left:
        collided = False
        for shape in other_shapes:
            if shape == footer_shape:
                continue
            if (shape.y < footer_shape.y + footer_shape.height and
                    shape.y + shape.height > footer_shape.y):
                # Vertically overlapping
                if shape.x + shape.width > new_x - padding > shape.x:
                    collided = True
                    break
        if collided:
            break
        new_x -= padding
    new_width_left = x - new_x

    # Try expanding right
    new_right = x + width
    while new_right + padding <= work_area_right:
        collided = False
        for shape in other_shapes:
            if shape == footer_shape:
                continue
            if (shape.y < footer_shape.y + footer_shape.height and
                    shape.y + shape.height > footer_shape.y):
                if shape.x < new_right + padding < shape.x + shape.width:
                    collided = True
                    break
        if collided:
            break
        new_right += padding
    new_width_right = new_right - (x + width)

    # Apply new width and x position
    footer_shape.x = new_x + padding
    footer_shape.width = width + new_width_left + new_width_right

    print("footer shape expand", footer_shape.x, footer_shape.y, footer_shape.width, footer_shape.height)


#     return footer_shape

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
    color_type = footer_config.get("color_type")

    if color_type == "CT_SCHEME":
        scheme_key = footer_config.get("scheme_color", "CH_TEXT1")
        scheme_color = SCHEME_COLOR_MAP.get(scheme_key, slides.SchemeColor.TEXT1)

        portion_format.fill_format.fill_type = slides.FillType.SOLID
        portion_format.fill_format.solid_fill_color.color_type = slides.ColorType.SCHEME
        portion_format.fill_format.solid_fill_color.scheme_color = scheme_color

    elif color_type == "CT_RGB":
        hex_color = footer_config.get("color_name", "000000").lstrip("#")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)

        portion_format.fill_format.fill_type = slides.FillType.SOLID
        portion_format.fill_format.solid_fill_color.color_type = slides.ColorType.RGB
        portion_format.fill_format.solid_fill_color.color = draw.Color.from_argb(r, g, b)

    else:
        # fallback default (black)
        portion_format.fill_format.fill_type = slides.FillType.SOLID
        portion_format.fill_format.solid_fill_color.color = draw.Color.black


def get_collision_info_2d_(q, g):
    """
    Compute vertical and horizontal collision metadata between two shapes `q` and `g`.

    Args:
        q, g: Shape-like objects with attributes x, y, width, height.

    Returns:
        dict with intersection types, crop values, and aspect ratio.
    """
    top_crop = bottom_crop = left_crop = right_crop = 0
    vertical_intersect = horizontal_intersect = 0

    # Vertical intersection check
    if g.y <= q.y <= g.y + g.height <= q.y + q.height:
        inter_vertical = 'TOP_INTERSECT_V'
        top_crop = (g.y + g.height) - q.y
        bottom_crop = (q.y + q.height) - (g.y + g.height)
        vertical_intersect = min(top_crop, bottom_crop)
    elif g.y + g.height > q.y + q.height and q.y <= g.y <= q.y + q.height:
        inter_vertical = 'BOTTOM_INTERSECT_V'
        bottom_crop = (q.y + q.height) - g.y
        top_crop = g.y - q.y
        vertical_intersect = min(top_crop, bottom_crop)
    elif q.y <= g.y <= q.y + q.height and q.y <= g.y + g.height <= q.y + q.height:
        inter_vertical = 'G_ENCLOSED_BY_Q_V'
        top_crop = (g.y + g.height) - q.y
        bottom_crop = (q.y + q.height) - g.y
        vertical_intersect = min(top_crop, bottom_crop)
    elif g.y <= q.y <= g.y + g.height and g.y <= q.y + q.height <= g.y + g.height:
        inter_vertical = 'Q_ENCLOSED_BY_G_V'
    else:
        inter_vertical = 'NONE_V'

    # Horizontal intersection check
    if g.x <= q.x <= g.x + g.width <= q.x + q.width:
        inter_horizontal = 'LEFT_INTERSECT_H'
        left_crop = (g.x + g.width) - q.x
        right_crop = (q.x + q.width) - (g.x + g.width)
        horizontal_intersect = min(left_crop, right_crop)
    elif g.x + g.width > q.x + q.width and q.x <= g.x <= q.x + q.width:
        inter_horizontal = 'RIGHT_INTERSECT_H'
        right_crop = (q.x + q.width) - g.x
        left_crop = g.x - q.x
        horizontal_intersect = min(left_crop, right_crop)
    elif q.x <= g.x <= q.x + q.width and q.x <= g.x + g.width <= q.x + q.width:
        inter_horizontal = 'G_ENCLOSED_BY_Q_H'
        left_crop = (g.x + g.width) - q.x
        right_crop = (q.x + q.width) - g.x
        horizontal_intersect = min(left_crop, right_crop)
    elif g.x <= q.x <= g.x + g.width and g.x <= q.x + q.width <= g.x + g.width:
        inter_horizontal = 'Q_ENCLOSED_BY_G_H'
    else:
        inter_horizontal = 'NONE_H'

    return {
        'inter_vertical': inter_vertical,
        'inter_horizontal': inter_horizontal,
        'left_crop': left_crop,
        'right_crop': right_crop,
        'top_crop': top_crop,
        'bottom_crop': bottom_crop,
        'g_wh_ratio': g.width / g.height if g.height else 0,
    }
