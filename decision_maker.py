import math
from find_footer_at_bottom_area import *


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
    final_font_size = 10
    if status_row and status_column:
        if font_size_row >= font_size_column:
            print("Row wise solution as added ")
            ## Creating new rectangle box with row wise data
            add_rectangle_box(new_slide, footer_shape['x'], footer_shape['y'], footer_shape['width'],
                              footer_shape['height'], footer_config, text, 1, font_size_row)
            final_font_size = font_size_row
        else:
            ## Creating new rectangle box with column wise data
            print("Column wise solution as added ")
            spacing = 0
            for text in text_list:
                add_rectangle_box(new_slide, footer_shape['x'] + spacing, footer_shape['y'], equal_width,
                                  footer_shape['height'], footer_config, text, 1, font_size_column)
                spacing += equal_width + 5
            final_font_size = font_size_column
    elif status_row:
        print("Row wise solution as added ")
        add_rectangle_box(new_slide, footer_shape['x'], footer_shape['y'], footer_shape['width'],
                          footer_shape['height'], footer_config, text, 1, font_size_row)
        final_font_size = font_size_row
    elif status_column:
        print("Column wise solution as added ")
        spacing = 0
        for text in text_list:
            add_rectangle_box(new_slide, footer_shape['x'] + spacing, footer_shape['y'], equal_width,
                              footer_shape['height'], footer_config, text, 1, font_size_column)
            spacing += equal_width + 5
        final_font_size = font_size_column
    else:
        slide_width = new_slide.presentation.slide_size.size.width
        # if footer_shape["width"] >= slide_width // 2:
        #     add_rectangle_box(new_slide,footer_shape['x'],footer_shape['y'],footer_shape["width"],footer_shape['height'],footer_config,text,len(text_list))
        # else:
        #     add_rectangle_box(new_slide,footer_shape['x'], footer_shape['y'],footer_shape["width"],footer_shape['height'],footer_config,text,1)
        #         final_font_size = font_size_column
        print("Cannot add footer box due to font size is getting very low ", )

    return final_font_size


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
        print(f"Error in can_text_list_fit_in_area: {e}")
        return False


def find_largest_fitting_font(min_font_size, max_font_size, box_width, box_height, lines_of_text):
    font_size = min((box_height / len(lines_of_text)), max_font_size)
    print("lines_of_text ", lines_of_text)
    while font_size >= min_font_size:
        if can_text_list_fit_in_area(font_size, box_width, box_height, lines_of_text):
            return True, round(font_size, 2)
        font_size -= 0.1

    return False, round(font_size, 2)





