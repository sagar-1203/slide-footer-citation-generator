import aspose.slides as slides
import aspose.pydrawing as drawing
import math


def get_the_max_font_in_column_or_row_wise(new_slide, min_font, max_font, footer_shape, text_list, footer_config):
    text = "\n".join(text_list)
    status_row, font_size_row = find_largest_fitting_font(min_font, max_font, footer_shape["width"],
                                                          footer_shape["height"], text_list)
    print("status_row,font_size_row ", status_row, font_size_row)
    equal_width = (footer_shape['width'] - 5 * (len(text_list))) / len(text_list)
    print("equal_width ", equal_width)
    max_len_text = max(text_list, key=len)
    status_column, font_size_column = find_largest_fitting_font(min_font, max_font, equal_width, footer_shape["height"],
                                                                [max_len_text])
    print("status_column,font_size_column ", status_column, font_size_column)
    final_font_size = -1
    add_footer_row_wise = False
    if status_row and status_column:
        if font_size_row >= font_size_column:
            print("Row wise solution as added ")
            ## Creating new rectangle box with row wise data
            ## add_rectangle_box(new_slide,footer_shape['x'],footer_shape['y'],footer_shape['width'],footer_shape['height'],footer_config,text,1,font_size_row)
            final_font_size = font_size_row
            add_footer_row_wise = True
        else:
            ## Creating new rectangle box with column wise data
            print("Column wise solution as added ")
            #             spacing = 0
            #             for text in text_list:
            #                 add_rectangle_box(new_slide,footer_shape['x'] + spacing,footer_shape['y'],equal_width,footer_shape['height'],footer_config,text,1,font_size_column)
            #                 spacing += equal_width + 5
            final_font_size = font_size_column
    elif status_row:
        print("Row wise solution as added ")
        add_footer_row_wise = True
        # add_rectangle_box(new_slide,footer_shape['x'],footer_shape['y'],footer_shape['width'],footer_shape['height'],footer_config,text,1,font_size_row)
        final_font_size = font_size_row
    elif status_column:
        print("Column wise solution as added ")
        spacing = 0
        #         for text in text_list:
        #             add_rectangle_box(new_slide,footer_shape['x'] + spacing,footer_shape['y'],equal_width,footer_shape['height'],footer_config,text,1,font_size_column)
        #             spacing += equal_width + 5
        final_font_size = font_size_column
    else:
        #         slide_width = new_slide.presentation.slide_size.size.width
        #         if footer_shape["width"] >= slide_width // 2:
        #             add_rectangle_box(new_slide,footer_shape['x'],footer_shape['y'],footer_shape["width"],footer_shape['height'],footer_config,text,len(text_list))
        #         else:
        #             add_rectangle_box(new_slide,footer_shape['x'], footer_shape['y'],footer_shape["width"],footer_shape['height'],footer_config,text,1)
        final_font_size = -1
        #         final_font_size = font_size_column
        print("Cannot add footer box due to font size is getting very low ", )

    return final_font_size, add_footer_row_wise


def can_text_list_fit_in_area(font_size, box_width, box_height, lines_of_text, fill_ratio=0.95):
    try:
        avg_character_width = 0.6 * font_size  # Average width of a character in points
        line_height = 1.3 * font_size  # Estimated line height with spacing
        print("font_size ", font_size)
        # Calculate capacity
        max_chars_per_line = int(box_width // avg_character_width) * fill_ratio
        max_lines_in_box = round(box_height / line_height)
        #     total_chars = chars_per_line * num_lines
        print("max_chars_per_line , max_lines_in_box ", max_chars_per_line, max_lines_in_box)
        #     print("total_chars ",total_chars)
        # Check if enough lines are available for the text list
        if max_lines_in_box >= len(lines_of_text):
            remaining_lines = max_lines_in_box
            print("Before remaining_lines ", remaining_lines)
            for line in lines_of_text:
                print("characters ", len(line))
                required_lines_for_text = math.ceil(len(line) / max_chars_per_line)
                print("required_lines_for_text ", required_lines_for_text)
                #             if len(line) > max_chars_per_line:
                if remaining_lines >= required_lines_for_text:
                    remaining_lines -= required_lines_for_text
                else:
                    return False
            print("After remaining_lines ", remaining_lines)
            return True
        else:
            return False
    except Exception as e:
        print("==" * 25)
        print("exception in can_text_list_fit_in_area ", str(e))
        print("==" * 25)
        return False


def find_largest_fitting_font(min_font_size, max_font_size, box_width, box_height, lines_of_text):
    font_size = min((box_height / len(lines_of_text)), max_font_size)
    print("lines_of_text ", lines_of_text)
    while font_size >= min_font_size:
        if can_text_list_fit_in_area(font_size, box_width, box_height, lines_of_text):
            return True, round(font_size, 2)
        font_size -= 0.1

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
        portion_format.fill_format.solid_fill_color.color = drawing.Color.from_argb(r, g, b)

    else:
        # fallback default (black)
        portion_format.fill_format.fill_type = slides.FillType.SOLID
        portion_format.fill_format.solid_fill_color.color = drawing.Color.black


from aspose.slides import Portion, PortionFormat, TextAutofitType
import aspose.pydrawing as draw


# from aspose.slides import ShapeType, TextAutofitType
def add_rectangle_box(slide, footer_x, footer_y, footer_width, footer_height, footer_config, text, num_columns=3,
                      font_size=None):
    footer_shape = slide.shapes.add_auto_shape(
        slides.ShapeType.RECTANGLE,
        footer_x,
        footer_y,
        footer_width,
        footer_height
    )
    #     text = "Environmental Protection Agency. (2023). Air Quality Index Report 2023. EPA Publications." + "\n" + "Green Research Institute. (2023). Environmental Impact Assessment of Electric Vehicles. Journal of Environmental Studies, 15(3), 45-62." + "\n" + "Urban Planning Department. (2023). Urban Noise Study: Impact of Electric Vehicles. City Planning Review, 28(4), 112-128."
    #     text = "Environmental Protection Agency. (2023). Air Quality Index Report 2023. EPA Publications." + "Green Research Institute. (2023). Environmental Impact Assessment of Electric Vehicles. Journal of Environmental Studies, 15(3), 45-62." + "Urban Planning Department. (2023). Urban Noise Study: Impact of Electric Vehicles. City Planning Review, 28(4), 112-128."

    #     text = "Footer"
    footer_shape.text_frame.text_frame_format.margin_left = 2.5
    footer_shape.text_frame.text_frame_format.margin_right = 2.5
    footer_shape.text_frame.text_frame_format.margin_top = 2.5
    footer_shape.text_frame.text_frame_format.margin_bottom = 2.5
    footer_shape.text_frame.text_frame_format.autofit_type = TextAutofitType.NORMAL
    footer_shape.text_frame.text_frame_format.wrap_text = slides.NullableBool.TRUE
    # 3. Set column properties
    footer_shape.text_frame.text_frame_format.column_count = num_columns

    # Aspose.Slides' column_spacing is typically specified in inches, convert from inches to points.
    # 1 inch = 72 points. If spacing is 0.25 inches, that's 18 points.
    footer_shape.text_frame.text_frame_format.column_spacing = 7
    #     footer_shape.text_frame.paragraphs[0].paragraph_format.indent = 7.0
    #     footer_shape.text_frame.paragraphs[0].paragraph_format.bullet.type = slides.BulletType.SYMBOL
    #     footer_shape.text_frame.paragraphs[0].paragraph_format.bullet.char = chr(8226)
    footer_shape.text_frame.text = text

    #     print("font ",footer_shape.text_frame.paragraphs[0].portions[0].portion_format.font_height)
    #     if footer_shape.text_frame.paragraphs and footer_shape.text_frame.paragraphs[0].portions:
    #         # Correct: Call get_effective() on the 'portion_format' object
    #         effective_portion_format = footer_shape.text_frame.paragraphs[0].portions[0].portion_format.get_effective()
    #         calculated_font_height = effective_portion_format.font_height
    #         print(f"Calculated (Effective) font height: {calculated_font_height}")
    #     else:
    #         print("No paragraphs or portions found in text frame.")
    #     footer_shape.text_frame.paragraphs[0].portions[0].portion_format.font_height= font_size
    #     new_ratio = refit_fonts(slide,footer_shape,font_size)
    #     print("current font size ,new ration, new_font ",font_size, new_ratio, new_ratio*font_size)
    #     footer_shape.text_frame.paragraphs[0].portions[0].portion_format.font_height= font_size * new_ratio
    footer_shape.fill_format.fill_type = slides.FillType.NO_FILL
    footer_shape.line_format.fill_format.fill_type = slides.FillType.NO_FILL
    footer_shape.text_frame.paragraphs[0].paragraph_format.alignment = slides.TextAlignment.LEFT
    footer_shape.text_frame.paragraphs[0].portions[0].portion_format.fill_format.fill_type = slides.FillType.SOLID
    footer_shape.text_frame.paragraphs[0].portions[
        0].portion_format.fill_format.solid_fill_color.color = draw.Color.black
    #     footer_shape.text_frame.paragraphs[0].portions[0].portion_format.fill_format.solid_fill_color.color_type = slides.ColorType.SCHEME
    #     footer_shape.text_frame.paragraphs[0].portions[0].portion_format.fill_format.solid_fill_color.scheme_color = slides.SchemeColor.TEXT_1  # maps to "CH_TEXT1"

    #     # Optional: font name, size, bold, etc.
    #     footer_shape.text_frame.paragraphs[0].portions[0].portion_format.latin_font = slides.PortionFormat.create_font("Arial")
    apply_font_color_from_config(footer_shape.text_frame.paragraphs[0].portions[0].portion_format, footer_config)
    #     footer_shape.text_frame.text_frame_format.autofit_type =  TextAutofitType.NORMAL
    if font_size:
        footer_shape.text_frame.text_frame_format.autofit_type = TextAutofitType.SHAPE
        footer_shape.text_frame.paragraphs[0].portions[0].portion_format.font_height = font_size
#     footer_shape.text_frame.text_frame_format.column_count = num_columns
#     print("footer_shape.txt ", footer_shape.text_frame.text)
