import aspose.slides as slides
import pandas as pd
from footer_locator_bottom import *
from footer_space_detector import *
from footer_orientation_manager import *
from utils import *
from aspose.slides import Portion


## Sample input text
input_text = {
    'title': ["Ferrari's Grand Entrance: Roaring into the Indian Luxury Market "],
    'subtitle': [],  # "How the F-Series drives America's economy engine"
    'content': [
        # "Conduct comprehensive workforce analysis identifying representation gaps and cultural barriers to inclusion" + addintional_data,
        "Established 3 state-of-the-art showrooms in key Indian cities",
        "Introduced personalized Tailor Made program for Indian customers",
        "Hosted Ferrari Challenge racing events to drive brand enthusiasm",
        "Announced plans to double the dealer network within 3 years",
        "Announced plans to double the dealer network within 3 years"
    ],
    'section_header': [
        # 'Inaugural Launch',
        'Exclusive Dealerships',
        'Customization Offerings',
        'Motorsport Engagement',
        'Expanding Footprint',
        'Expanding Footprint'
    ],
    # 'hub_title': '',
    # 'graph_data': {},
    # 'lang': 'english',
}


def add_citations_as_superscript(slide, text_matches, default_font_size=6, min_font_size=4):
    """
    Add superscript citation indices to matched text in slide shapes.

    Args:
        slide (Slide): PowerPoint slide to process.
        text_matches (list[str]): List of citation texts to match.
        default_font_size (int, optional): Default font size for superscripts. Defaults to 6.
        min_font_size (int, optional): Minimum font size for superscripts. Defaults to 4.

    Returns:
        Slide: Slide with superscript citations added.
    """
    lower_matches = [s.lower() for s in text_matches]

    # Loop over matches FIRST (this controls numbering)
    for count, match in enumerate(lower_matches, start=1):
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame") or shape.text_frame is None:
                continue

            shape_text = shape.text_frame.text.lower()
            if match not in shape_text:
                continue  # skip shapes that don't contain the match at all

            paras = shape.text_frame.paragraphs
            target_para = None

            # First try to find a paragraph containing the match
            for para in paras:
                if match in para.text.lower():
                    target_para = para
                    break

            # If no specific paragraph matched, fallback to last paragraph
            if target_para is None and paras.count > 0:
                target_para = paras[paras.count - 1]

            if target_para is None:
                continue

            portions = target_para.portions
            if portions.count == 0:
                continue

            # Get last portion
            last_portion = portions[portions.count - 1]

            # Create superscript
            superscript = Portion()
            superscript.text = f"{count}"
            superscript_format = superscript.portion_format
            superscript_format.escapement = 30

            # Inherit style from previous portion
            prev_format = last_portion.portion_format
            superscript_format.font_height = prev_format.font_height or default_font_size
            superscript_format.latin_font = prev_format.latin_font
            superscript.portion_format.fill_format.fill_type = prev_format.fill_format.fill_type
            superscript.portion_format.fill_format.solid_fill_color.color = prev_format.fill_format.solid_fill_color.color
            superscript.portion_format.font_bold = prev_format.font_bold
            superscript.portion_format.font_italic = prev_format.font_italic
            superscript.portion_format.font_underline = prev_format.font_underline
            superscript.portion_format.east_asian_font = prev_format.east_asian_font
            superscript.portion_format.complex_script_font = prev_format.complex_script_font

            target_para.portions.add(superscript)

            shape.text_frame.text_frame_format.autofit_type = slides.TextAutofitType.SHAPE
            shape.text_frame.text_frame_format.wrap_text = slides.NullableBool.TRUE

            break  # stop after first shape containing this match
    return slide

def add_superscript_references(slide, text_matches, default_font_size=6, min_font_size=4):
    """
    For each shape, check if any string in text_matches is present in its text.
    If found, append superscript [i] where i is the index of the matched string (1-based).
    Handles portion/paragraph indexing correctly and adjusts layout if overflow occurs.

    Args:
        slide (slides.Slide): The slide to process.
        text_matches (list[str]): List of strings to match in shapes.
        default_font_size (int): Default font size for newly created portion.
        min_font_size (int): Minimum font size if shape overflows.
    """
    lower_matches = [s.lower() for s in text_matches]
    count = 1
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or shape.text_frame is None:
            continue

        shape_text = shape.text_frame.text.lower()
        for i, match in enumerate(lower_matches, start=1):
            if match in shape_text:
                print("match index ", i, match)
                paras = shape.text_frame.paragraphs
                for para_index in range(paras.count):
                    para = paras[para_index]
                    portions = para.portions
                    # print("port count ",portions.count)
                    if portions.count == 0:
                        continue
                        # new_portion = Portion()
                        # new_portion.text = " "
                        # new_portion.portion_format.font_height = default_font_size
                        # para.portions.add(new_portion)

                    # Get last portion
                    last_portion = portions[portions.count - 1]

                    # Ensure space before superscript
                    # if last_portion.text and not last_portion.text.endswith(" "):
                    #     last_portion.text += " "
                    # Create superscript
                    superscript = Portion()
                    superscript.text = f"{count}"
                    count += 1
                    superscript_format = superscript.portion_format
                    superscript_format.escapement = 30
                    # Inherit style from previous portion
                    prev_format = last_portion.portion_format
                    superscript_format.font_height = prev_format.font_height or default_font_size
                    superscript_format.latin_font = prev_format.latin_font
                    superscript.portion_format.fill_format.fill_type = prev_format.fill_format.fill_type
                    superscript.portion_format.fill_format.solid_fill_color.color = prev_format.fill_format.solid_fill_color.color
                    superscript.portion_format.font_bold = prev_format.font_bold
                    superscript.portion_format.font_italic = prev_format.font_italic
                    superscript.portion_format.font_underline = prev_format.font_underline
                    superscript.portion_format.east_asian_font = prev_format.east_asian_font
                    superscript.portion_format.complex_script_font = prev_format.complex_script_font
                    para.portions.add(superscript)
                    shape.text_frame.text_frame_format.autofit_type = slides.TextAutofitType.SHAPE
                    shape.text_frame.text_frame_format.wrap_text = slides.NullableBool.TRUE
                    break  # Only first match per shape
                break

#  ============= Add final Footer shape ===============

def add_final_footer_shape(slide, footer_shape, footer_config, font_size, text_list, add_footer_row_wise):
    """
    Add footer text to a slide either as a single row or split across columns.

    Args:
        slide (Slide): PowerPoint slide to add footer to.
        footer_shape (dict): Dictionary with footer box position and size (x, y, width, height).
        footer_config (dict): Footer styling configuration.
        font_size (int): Font size for footer text.
        text_list (list[str]): List of footer text entries.
        add_footer_row_wise (bool): If True, add as single row; otherwise split into columns.

    Returns:
        None
    """
    text = "Sources: " + "; ".join(text_list)
    if add_footer_row_wise:
        add_rectangle_box(slide, footer_shape['x'], footer_shape['y'], footer_shape['width'], footer_shape['height'],
                          footer_config, text, 1, font_size)
    else:
        equal_width = (footer_shape['width'] - 5 * (len(text_list))) / len(text_list)
        spacing = 0
        for text in text_list:
            add_rectangle_box(slide, footer_shape['x'] + spacing, footer_shape['y'], equal_width,
                              footer_shape['height'], footer_config, text, 1, font_size)
            spacing += equal_width + 5

def format_input_footer_text(footer_text_list):
    formatted_list = []
    for sublist in footer_text_list:
        for item in sublist:
            formatted_list.append(f"{item['id']} - {item['citation']}")
    return formatted_list

def add_footer(slide,work_area, footer_text_list, footer_config, MIN_FONT_SIZE=4, MAX_FONT_SIZE=7):
    """
    Determine and add the best-fitting footer text to a slide.

    Args:
        slide (Slide): PowerPoint slide to add the footer.
        work_area (dict): Dictionary defining the usable slide area (left, bottom, etc.).
        footer_text_list (list[list[dict]]): Nested list of dictionaries containing "id" and "citation".
        footer_config (dict): Styling configuration for the footer.
        MIN_FONT_SIZE (int, optional): Minimum font size allowed for footer text. Defaults to 4.
        MAX_FONT_SIZE (int, optional): Maximum font size allowed for footer text. Defaults to 7.

    Returns:
        Slide: The slide with the footer added or updated.
    """
    decision_message = None
    new_footer_text = format_input_footer_text(footer_text_list)
    original_slides = [slide]
    if work_area:
        for slide in original_slides:
            all_shapes_df, layout_df_master, layout_df = find_all_the_shapes(slide)
            slide_width = slide.presentation.slide_size.size.width
            slide_height = slide.presentation.slide_size.size.height
            print("slide width and height ", slide_width, slide_height)
            #         find_all_the_shapes(slide)
            print("initial rendering of slide")
            # render_ppt(pres,slide)
            split_y, bottom_shapes, footer_shape_list = add_footer_shape_df(slide, work_area, padding=0,
                                                                            collision_threshold=2)
            print("split_y ", split_y)
            print("footer shape ", footer_shape_list)
            max_footer_font_1 = -1
            max_footer_font_2 = -1
            resulted_footer_shape = None
            add_footer_row_wise1 = None
            add_footer_row_wise2 = None
            results = None
            if footer_shape_list:
                for footer_shape in footer_shape_list:
                    if footer_shape['x'] >= work_area["left"]:
                        # max_footer_font_1 = min(max_footer_font_1,get_the_max_font_in_column_or_row_wise(new_slide,min_font,max_font,footer_shape,text_list,footer_config))
                        resulted_font_size1, add_footer_row_wise1 = get_the_max_font_in_column_or_row_wise(
                                                                                                           MIN_FONT_SIZE,
                                                                                                           MAX_FONT_SIZE,
                                                                                                           footer_shape,
                                                                                                           new_footer_text,
                                                                                                           )
                        #                     max_footer_font_1 = min(max_footer_font_1,resulted_font_size1)
                        if resulted_font_size1 > max_footer_font_1:
                            max_footer_font_1 = resulted_font_size1
                            resulted_footer_shape = footer_shape
                        # add_rectangle_box(new_slide,footer_shape['x'],footer_shape['y'],footer_shape['width'],footer_shape['height'],footer_config,text,1,font_size)
            if split_y > work_area["bottom"] + 5:
                results = find_max_footer_area_df(slide, work_area, split_y, all_shapes_df, layout_df_master,
                                                  layout_df, min_width=20, min_height=10, max_height=20, )
                #             results = find_max_footer_area(new_slide_2, work_area,split_y)
                #         results = find_best_footer_area(slide,work_area["abbvie_aquipta"]["work_area"]["left"], split_y)
                print("results ", results)
                if results:
                    #                 remove_present_footer(slide)
                    resulted_font_size2, add_footer_row_wise2 = get_the_max_font_in_column_or_row_wise(
                                                                                                       MIN_FONT_SIZE,
                                                                                                       MAX_FONT_SIZE,
                                                                                                       results,
                                                                                                       new_footer_text,
                                                                                                       )
                    if resulted_font_size2 > max_footer_font_2:
                        max_footer_font_2 = resulted_font_size2
                print("max_footer_font_1, max_footer_font_2 ", max_footer_font_1, max_footer_font_2)
            if max_footer_font_1 == -1 and max_footer_font_2 == -1:
                print("No footer found in both slides")
                decision_message = "No footer found in both slides"
            elif max_footer_font_1 >= max_footer_font_2:
                print("font. Use slide 2 As final footer solution")
                add_final_footer_shape(slide, resulted_footer_shape, footer_config, max_footer_font_1, new_footer_text,
                                       add_footer_row_wise1)
                decision_message = "Use slide 2 As final footer solution"
            elif max_footer_font_2 > max_footer_font_1:
                add_final_footer_shape(slide, results, footer_config, max_footer_font_2, new_footer_text,
                                       add_footer_row_wise2)
                print("font. Use slide 1 As final footer solution")
                decision_message = "Use slide 1 As final footer solution"
            else:
                print("No solution fount")
                decision_message = "No solution found"
            print("final rendering of slide")
            # for shape_info in merged_results:
            #     remove_shapes_by_info(slide,shape_info)
    else:
        decision_message = "Work Area not found in template configuration."
    print("decision_message ", decision_message)
    return slide

def add_footer_and_citation(slide,work_area, footer_text, footer_config, MIN_FONT_SIZE=4, MAX_FONT_SIZE=7):
    slide = add_footer(slide, work_area, footer_text, footer_config, MIN_FONT_SIZE, MAX_FONT_SIZE)
    slide = add_citations_as_superscript(slide,input_text["content"])
    return slide