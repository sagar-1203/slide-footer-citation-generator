import aspose.slides as slides
from aspose.slides import TextAutofitType
from aspose.slides import AutoShape
from app.utils.add_footer_and_citation.utils import find_all_the_shapes
import pandas as pd
from utils import remove_shapes_outside_slide, get_collision_info_2d
# from app.utils.generate_text_fill_utils.text_fill.format_slides.collision_utils import get_collision_info_2d
# from app.utils.generate_text_fill_utils.text_fill.format_slides.format_slide import remove_shapes_outside_slide

def add_footer_shape_df_old(slide, work_area, footer_height=30, padding=0, collision_threshold=2):
    """
    DataFrame-native version of add_footer_shape. If df_shapes is None, builds it via unpack_shapes.
    Returns: avg_y, bottom_shapes_df, footer_shape_list (same dicts as before)
    """
    slide_width = slide.presentation.slide_size.size.width
    slide_height = slide.presentation.slide_size.size.height
    footer_left = work_area["left"]
    work_area_right = work_area["right"]
    footer_y = work_area["bottom"]
    threshold_height = 0.8
    threshold_width = 0.85
    all_shapes_df, layout_df_master, layout_df = find_all_the_shapes(slide)

    # Combine the DataFrames (rows are stacked vertically)
    df_shapes = pd.concat([all_shapes_df, layout_df_master, layout_df], ignore_index=True)
    # Normalize missing columns if necessary
    df = df_shapes.copy()
    print("combined df")
    # display(df)
    # Filter hidden shapes
    df = df[~df.apply(lambda row: getattr(getattr(row["shape"], "shape", None), "hidden", False), axis=1)]
    print("removing hidden shape df")
    # display(df)
    #     df = df[~df.get("hidden", False)]

    # Filter shapes entirely above footer_y
    df = df[~((df["top"] + df["height"]) < footer_y)]
    print("removing shapes bottom < footer_y shape df")
    # display(df)
    # Skip huge shapes (likely background/title)
    # df = df[df["height"] < slide_height * threshold_height]
    df = df[(df["is_image"]) | (df["height"] < slide_height * threshold_height)]
    df = df[
        ~(
                (df["height"] >= 0.99 * slide_height) &
                (df["width"] >= 0.99 * slide_width)
        )
    ]
    print("removing shapes height > slide_height*threshold shape df")
    # display(df)
    # df = df[~df.apply(lambda row: (row["isfillable"] == False) and (row["is_image"] == False and (row["parent"] == [-1])), axis=1)]
    # print("After removing isfillable and is_image = false shape")
    # display(df)
    # Dedup logic
    seen_shapes = set()
    is_footer_present = False
    keep_footer_shape_obj = None
    bottom_shapes = []
    footer_shape_list = []

    disqualifying_keywords = ["internal use", "copyright", "copy right", "external use"]
    qualifying_keywords = [
        "footer", "source goes here", "goes here", "disclaimer place holder", "edit source",
        "footnote", "foot", "reference"
    ]
    print("initial dataframe after all the conditions")
    # display(df)
    # We'll iterate row-wise because of the complex logic
    for _, row in df.iterrows():
        shape_obj = row['shape'].shape
        top = row["top"]
        height = row["height"]
        width = row["width"]
        left = row["left"]
        bottom = top + height
        shape_key = (left, top, width, height)

        # Duplicate detection
        if shape_key in seen_shapes:
            lower_text = (row.get("lower_case_content") or "")
            name = ""
            try:
                name = getattr(shape_obj, "name", "") or ""
                name = name.lower()
            except Exception as e:
                print("Exception ", str(e))
                pass

            if any(keyword in lower_text for keyword in qualifying_keywords):  # or "footer" in name:
                # Duplicate footer-like shape: prefer one in slide.shapes if possible
                is_footer_present = True
                keep_footer_shape_obj = shape_obj
                for s in slide.shapes:
                    if (hasattr(s, "x") and hasattr(s, "y") and hasattr(s, "width") and hasattr(s, "height") and (s.x,
                                                                                                                  s.y,
                                                                                                                  s.width,
                                                                                                                  s.height) == shape_key):
                        is_footer_present = True
                        keep_footer_shape_obj = s
                        print("for shape key ", shape_key)
                        print("Removed duplicate shape by coordinates from slide.")
                        break
            continue  # skip further processing of this duplicate

        seen_shapes.add(shape_key)

        # Placeholder/name-based detection (if placeholder is accessible)
        placeholder_type = None
        name = ""
        try:
            placeholder_type = shape_obj.placeholder.type if hasattr(shape_obj,
                                                                     "placeholder") and shape_obj.placeholder else None
            name = getattr(shape_obj, "name", "") or ""
            name = name.lower()
        except Exception as e:
            print("exception ", str(e))

        lower_text = (row.get("lower_case_content") or "")

        # Footer placeholder or explicit "footer" in name
        # (shape_obj in slide.shapes) and
        if ((placeholder_type == slides.PlaceholderType.FOOTER or "footer" in name)):
            if len(lower_text) == 0 or any(keyword in lower_text for keyword in qualifying_keywords):
                is_footer_present = True
                keep_footer_shape_obj = shape_obj

        # Final fallback heuristic: bottom region and footer-like
        if (top > footer_y) or ((bottom) > footer_y + 1 and width >= 2 and not row["isfillable"]):
            if (not is_footer_present) and (shape_obj in slide.shapes) and isinstance(shape_obj, AutoShape):
                if len(lower_text) == 0 or any(keyword in lower_text for keyword in qualifying_keywords):
                    is_footer_present = True
                    keep_footer_shape_obj = shape_obj
                    print("footer shape found")
            bottom_shapes.append(row)
    print("is footer present ", is_footer_present)
    print("keep footer shape ", keep_footer_shape_obj)
    # Convert bottom_shapes list of rows to DataFrame
    bottom_df = pd.DataFrame(bottom_shapes)
    print("initial bottom shapes")
    # display(bottom_df)
    # Remove shapes outside slide bounds (you have a helper; apply it)
    bottom_df = remove_shapes_outside_slide(bottom_df, slide_width, slide_height)  # assume you wrap original function

    # Sort left-to-right
    bottom_df = bottom_df.sort_values(by="left").reset_index(drop=True)

    # Collision-based selection (mirrors original logic)
    # if len(bottom_df) >= 2:
    #     prev_row = bottom_df.loc[0]
    #     for i in range(1, len(bottom_df)):
    #         curr_row = bottom_df.loc[i]
    #         prev_shape = prev_row['shape']
    #         curr_shape = curr_row['shape']
    #         collision_info = get_collision_info_2d(prev_shape, curr_shape)
    #         if (collision_info['inter_vertical'] in ["G_ENCLOSED_BY_Q_V", "Q_ENCLOSED_BY_G_V"] or
    #                 collision_info['inter_horizontal'] in ["G_ENCLOSED_BY_Q_H", "Q_ENCLOSED_BY_G_H"]):
    #             print("shape prev_shape ",prev_shape.left,prev_shape.top)
    #             print("shape curr_shape ",curr_shape.left,curr_shape.top)
    #             if (collision_info['inter_vertical'] == "Q_ENCLOSED_BY_G_V" or
    #                     collision_info['inter_horizontal'] == "Q_ENCLOSED_BY_G_H"):
    #                 prev_row = curr_row
    #         else:
    #             prev_row = curr_row
    if len(bottom_df) >= 2:
        to_drop = set()

        for i in range(len(bottom_df)):
            for j in range(i + 1, len(bottom_df)):
                row_i = bottom_df.loc[i]
                row_j = bottom_df.loc[j]

                shape_i = row_i['shape']
                shape_j = row_j['shape']

                collision_info = get_collision_info_2d(shape_i, shape_j)
                vert = collision_info['inter_vertical']
                horiz = collision_info['inter_horizontal']

                print(f"Comparing i={i}, j={j}")
                print("  shape_i:", shape_i.left, shape_i.top)
                print("  shape_j:", shape_j.left, shape_j.top)
                print("  collision_info:", collision_info)

                # --- Drop Q (row_i) if horizontally enclosed by G and vertically enclosed or intersecting ---
                if horiz == "Q_ENCLOSED_BY_G_H" and vert not in ["NONE_V"]:
                    to_drop.add(row_i.name)

                # --- Drop G (row_j) if horizontally enclosed by Q and vertically enclosed or intersecting ---
                elif horiz == "G_ENCLOSED_BY_Q_H" and vert not in ["NONE_V"]:
                    to_drop.add(row_j.name)

                # --- Strict full enclosures (both axes say enclosed) ---
                elif vert == "Q_ENCLOSED_BY_G_V" and horiz not in ["NONE_H"]:
                    to_drop.add(row_i.name)
                elif vert == "G_ENCLOSED_BY_Q_V" and horiz not in ["NONE_H"]:
                    to_drop.add(row_j.name)
                ## above or below if below then correct some edge cases
                # if horiz == "Q_ENCLOSED_BY_G_H" and vert == "Q_ENCLOSED_BY_G_V":
                #     to_drop.add(row_i.name)
                # elif horiz == "G_ENCLOSED_BY_Q_H" and vert == "G_ENCLOSED_BY_Q_V":
                #     to_drop.add(row_j.name)

        # Remove at once
        if to_drop:
            bottom_df = bottom_df.drop(list(to_drop)).reset_index(drop=True)

    # bottom_df = bottom_df.reset_index(drop=True)

    print("bottom df")
    # display(bottom_df)
    height_thresh = 7
    # Compute avg_y logic

    ys = [
        row["top"] for _, row in bottom_df.iterrows()
        if row["width"] >= threshold_width * slide_width
           and row["height"] > height_thresh
           and row["top"] > footer_y
    ]

    avg_y = 0  # default

    thin_shapes = bottom_df[bottom_df["height"] <= height_thresh]
    other_shapes = bottom_df[bottom_df["height"] > height_thresh]

    if not thin_shapes.empty and not other_shapes.empty:
        for _, thin_row in thin_shapes.iterrows():
            thin_top = thin_row["top"]

            # Compare thin shape with others
            all_above = all(other_shapes["top"] + other_shapes["height"] <= thin_top)
            all_below = all(other_shapes["top"] >= thin_top)

            if all_above:
                # Remove this thin shape
                bottom_df = bottom_df.drop(thin_row.name)
            elif all_below:
                # Set avg_y to this shape's top
                avg_y = thin_top
            else:
                # Mixed case → fallback to average of ys
                avg_y = sum(ys) / len(ys) if ys else 0
    else:
        # Fallback to average of ys
        avg_y = sum(ys) / len(ys) if ys else 0
    # ys = [
    #     row["top"] for _, row in bottom_df.iterrows()
    #     if row["width"] >= threshold_width * slide_width and row["height"] > height_thresh and row["top"] > footer_y
    # ]
    print(ys)
    # avg_y = sum(ys) / len(ys) if ys else 0
    print("avg_y ", avg_y)
    if avg_y <= 0:
        #         if len(bottom_df) >= 2:
        #             sorted_by_y = bottom_df.sort_values(by="y").iloc[:2]
        #             avg_y = sum(sorted_by_y["y"]) / 2
        #         elif len(bottom_df) == 1:
        #             avg_y = bottom_df.iloc[0]["y"]

        # fallback to full average
        filtered_df = bottom_df[bottom_df["height"] > height_thresh]
        filtered_df = filtered_df[filtered_df["top"] > footer_y]
        if len(filtered_df):
            old_avg = filtered_df["top"].mean()
            avg_y = old_avg
        else:
            avg_y = footer_y
    print("avg_y ", avg_y)
    avg_footer_height = slide_height - avg_y

    # Build footer shapes if none present yet
    if not is_footer_present:
        # Filter wide shapes if more than one
        if len(bottom_df) > 1:
            bottom_df = bottom_df[bottom_df["width"] < slide_width * threshold_width]

        # Compute bounds
        shape_bounds = sorted(
            [(row["left"], row["left"] + row["width"]) for _, row in bottom_df.iterrows()],
            key=lambda b: b[0]
        )
        # Add virtual edges
        shape_bounds.insert(0, (0, 0))
        shape_bounds.append((slide_width, slide_width))

        max_gap = 20
        best_position = 0
        for i in range(len(shape_bounds) - 1):
            gap_start = shape_bounds[i][1]
            gap_end = shape_bounds[i + 1][0]
            gap_width = gap_end - gap_start
            if gap_width > max_gap:
                best_position = gap_start
                footer_width = gap_width - 2 * collision_threshold
                if footer_width < 30:
                    print("⚠️ Not enough width for footer on this slide. ", footer_width)
                    continue
                footer_x = best_position + collision_threshold
                if footer_x < footer_left and footer_width < footer_left:
                    print("⚠️ outside to working area on left ", footer_x, footer_width)
                    continue
                elif footer_x < footer_left:
                    print("taking work area x as footer_x ")
                    footer_width -= (footer_left - footer_x)
                    footer_x = footer_left
                elif footer_x > work_area_right:
                    print("work area right ", work_area_right)
                    print("⚠️ outside to working area on right ", footer_x, footer_width)
                    continue

                footer_y_shapes = [
                    row["top"] for _, row in bottom_df.iterrows()
                    if row["top"] > avg_y
                ]
                print("footer_y_shapes ", footer_y_shapes)
                computed_footer_y = (sum(footer_y_shapes) / len(footer_y_shapes)) if footer_y_shapes else footer_y
                print(
                    f"footer_width {footer_width},footer_height {avg_footer_height}, footer_x {footer_x} , footer_y {footer_y} ")
                footer_shape = {
                    "width": footer_width,
                    "height": footer_height,
                    "x": footer_x,
                    "y": computed_footer_y
                }
                footer_shape_list.append(footer_shape)
    else:
        # Expand existing footer-like shape
        expand_footer_shape(keep_footer_shape_obj, bottom_df, work_area, slide_width, slide_width)
        footer_shape = {
            "width": keep_footer_shape_obj.width,
            "height": keep_footer_shape_obj.height,
            "x": keep_footer_shape_obj.x,
            "y": keep_footer_shape_obj.y
        }
        footer_shape_list.append(footer_shape)
        if keep_footer_shape_obj in slide.shapes:
            slide.shapes.remove(keep_footer_shape_obj)

    return avg_y, bottom_df, footer_shape_list



def add_footer_shape_df(
    slide,
    work_area,
    footer_height: int = 30,
    padding: int = 0,
    collision_threshold: int = 2
):
    """
    DataFrame-native version of `add_footer_shape`.

    Args:
        slide (aspose.slides.Slide):
            The PowerPoint slide object to analyze and possibly add a footer to.
        work_area (dict):
            A dictionary describing the usable area of the slide. Expected keys:
            - "left": left boundary (float or int)
            - "right": right boundary
            - "top": top boundary
            - "bottom": bottom boundary
        footer_height (int, optional):
            Default footer shape height. Defaults to 30.
        padding (int, optional):
            Padding to apply around the footer shape. Currently unused. Defaults to 0.
        collision_threshold (int, optional):
            Buffer space used when resolving overlaps with other shapes. Defaults to 2.

    Returns:
        tuple:
            avg_y (float):
                Estimated Y-coordinate (top position) where the footer region starts.
            bottom_shapes_df (pandas.DataFrame):
                DataFrame of shapes near the bottom of the slide that influenced footer placement.
            footer_shape_list (list of dict):
                Candidate footer shapes (dicts with "x", "y", "width", "height").
                Usually contains 1 entry if a footer was created or adjusted.
    """
    slide_width = slide.presentation.slide_size.size.width
    slide_height = slide.presentation.slide_size.size.height
    footer_left = work_area["left"]
    work_area_right = work_area["right"]
    footer_y = work_area["bottom"]
    threshold_height = 0.8
    threshold_width = 0.85

    all_shapes_df, layout_df_master, layout_df = find_all_the_shapes(slide)

    # Combine shape dataframes
    df_shapes = pd.concat([all_shapes_df, layout_df_master, layout_df], ignore_index=True)
    df = df_shapes.copy()
    print("combined df")

    # Remove hidden shapes
    df = df[~df.apply(lambda row: getattr(getattr(row["shape"], "shape", None), "hidden", False), axis=1)]
    print("removing hidden shape df")

    # Filter shapes entirely above footer_y
    df = df[~((df["top"] + df["height"]) < footer_y)]
    print("removing shapes bottom < footer_y shape df")

    # Skip very tall shapes (likely background/title), except images
    df = df[(df["is_image"]) | (df["height"] < slide_height * threshold_height)]
    df = df[
        ~(
            (df["height"] >= 0.99 * slide_height) &
            (df["width"] >= 0.99 * slide_width)
        )
    ]
    print("removing shapes height > slide_height*threshold shape df")

    # Deduplication + footer detection
    seen_shapes = set()
    is_footer_present = False
    keep_footer_shape_obj = None
    bottom_shapes = []
    footer_shape_list = []

    disqualifying_keywords = ["internal use", "copyright", "copy right", "external use"]
    qualifying_keywords = [
        "footer", "source goes here", "goes here", "disclaimer place holder", "edit source",
        "footnote", "foot", "reference"
    ]
    print("initial dataframe after all the conditions")

    for _, row in df.iterrows():
        shape_obj = row['shape'].shape
        top, height, width, left = row["top"], row["height"], row["width"], row["left"]
        bottom = top + height
        shape_key = (left, top, width, height)

        # --- Dedup detection
        if shape_key in seen_shapes:
            lower_text = (row.get("lower_case_content") or "")
            name = ""
            try:
                name = getattr(shape_obj, "name", "") or ""
                name = name.lower()
            except Exception as e:
                print("Exception ", str(e))

            if any(keyword in lower_text for keyword in qualifying_keywords):
                is_footer_present = True
                keep_footer_shape_obj = shape_obj
                for s in slide.shapes:
                    if (
                        hasattr(s, "x") and hasattr(s, "y") and
                        hasattr(s, "width") and hasattr(s, "height") and
                        (s.x, s.y, s.width, s.height) == shape_key
                    ):
                        keep_footer_shape_obj = s
                        print("Removed duplicate shape by coordinates from slide.")
                        break
            continue

        seen_shapes.add(shape_key)

        # --- Placeholder/name-based detection
        placeholder_type, name = None, ""
        try:
            placeholder_type = (
                shape_obj.placeholder.type
                if hasattr(shape_obj, "placeholder") and shape_obj.placeholder
                else None
            )
            name = getattr(shape_obj, "name", "") or ""
            name = name.lower()
        except Exception as e:
            print("exception ", str(e))

        lower_text = (row.get("lower_case_content") or "")

        if ((placeholder_type == slides.PlaceholderType.FOOTER) or ("footer" in name)):
            if len(lower_text) == 0 or any(keyword in lower_text for keyword in qualifying_keywords):
                is_footer_present = True
                keep_footer_shape_obj = shape_obj

        # --- Fallback heuristic: bottom region
        if (top > footer_y) or ((bottom > footer_y + 1) and width >= 2 and not row["isfillable"]):
            if (not is_footer_present) and (shape_obj in slide.shapes) and isinstance(shape_obj, AutoShape):
                if len(lower_text) == 0 or any(keyword in lower_text for keyword in qualifying_keywords):
                    is_footer_present = True
                    keep_footer_shape_obj = shape_obj
                    print("footer shape found")
            bottom_shapes.append(row)

    print("is footer present ", is_footer_present)

    # Build bottom_shapes dataframe
    bottom_df = pd.DataFrame(bottom_shapes)
    bottom_df = remove_shapes_outside_slide(bottom_df, slide_width, slide_height)
    bottom_df = bottom_df.sort_values(by="left").reset_index(drop=True)

    # --- Collision filtering
    if len(bottom_df) >= 2:
        to_drop = set()
        for i in range(len(bottom_df)):
            for j in range(i + 1, len(bottom_df)):
                row_i = bottom_df.loc[i]
                row_j = bottom_df.loc[j]
                collision_info = get_collision_info_2d(row_i['shape'], row_j['shape'])
                vert, horiz = collision_info['inter_vertical'], collision_info['inter_horizontal']

                if horiz == "Q_ENCLOSED_BY_G_H" and vert != "NONE_V":
                    to_drop.add(row_i.name)
                elif horiz == "G_ENCLOSED_BY_Q_H" and vert != "NONE_V":
                    to_drop.add(row_j.name)
                elif vert == "Q_ENCLOSED_BY_G_V" and horiz != "NONE_H":
                    to_drop.add(row_i.name)
                elif vert == "G_ENCLOSED_BY_Q_V" and horiz != "NONE_H":
                    to_drop.add(row_j.name)

        if to_drop:
            bottom_df = bottom_df.drop(list(to_drop)).reset_index(drop=True)

    # --- Average footer Y calculation
    height_thresh = 7
    ys = [
        row["top"] for _, row in bottom_df.iterrows()
        if row["width"] >= threshold_width * slide_width
        and row["height"] > height_thresh
        and row["top"] > footer_y
    ]

    avg_y = sum(ys) / len(ys) if ys else 0

    # Fallback
    if avg_y <= 0:
        filtered_df = bottom_df[(bottom_df["height"] > height_thresh) & (bottom_df["top"] > footer_y)]
        avg_y = filtered_df["top"].mean() if len(filtered_df) else footer_y

    avg_footer_height = slide_height - avg_y

    # --- Build footer shapes
    if not is_footer_present:
        if len(bottom_df) > 1:
            bottom_df = bottom_df[bottom_df["width"] < slide_width * threshold_width]

        shape_bounds = sorted(
            [(row["left"], row["left"] + row["width"]) for _, row in bottom_df.iterrows()],
            key=lambda b: b[0]
        )
        shape_bounds.insert(0, (0, 0))
        shape_bounds.append((slide_width, slide_width))

        max_gap = 20
        for i in range(len(shape_bounds) - 1):
            gap_start, gap_end = shape_bounds[i][1], shape_bounds[i + 1][0]
            gap_width = gap_end - gap_start
            if gap_width > max_gap:
                footer_x = gap_start + collision_threshold
                footer_width = gap_width - 2 * collision_threshold
                if footer_width >= 30:
                    computed_footer_y = (
                        bottom_df[bottom_df["top"] > avg_y]["top"].mean()
                        if len(bottom_df[bottom_df["top"] > avg_y]) else footer_y
                    )
                    footer_shape_list.append({
                        "width": footer_width,
                        "height": footer_height,
                        "x": footer_x,
                        "y": computed_footer_y
                    })
    else:
        expand_footer_shape(keep_footer_shape_obj, bottom_df, work_area, slide_width, slide_width)
        footer_shape_list.append({
            "width": keep_footer_shape_obj.width,
            "height": keep_footer_shape_obj.height,
            "x": keep_footer_shape_obj.x,
            "y": keep_footer_shape_obj.y
        })
        if keep_footer_shape_obj in slide.shapes:
            slide.shapes.remove(keep_footer_shape_obj)

    return avg_y, bottom_df, footer_shape_list


def expand_footer_shape(footer_shape, other_rows_df, work_area, slide_width, slide_height,
                        padding=1, max_footer_height=30):
    """
    Expand footer shape vertically and horizontally within available space, 
    avoiding collisions with other shapes and staying inside work area.

    Args:
        footer_shape: The Aspose.Slides shape object representing the footer.
        other_rows_df (pd.DataFrame): DataFrame of other shape wrappers; each row has 'shape', 'top', 'bottom', 'left', 'right'.
        work_area (dict): Dictionary with 'left' and 'right' boundaries of usable slide area.
        slide_width (float): Width of the slide.
        slide_height (float): Height of the slide.
        padding (int, optional): Extra spacing buffer around the footer. Defaults to 1.
        max_footer_height (float, optional): Maximum allowed height of the footer. Defaults to 30.

    Returns:
        None: Mutates `footer_shape` in place (adjusts x, y, width, height).
    """
    print("=== expand_footer_shape called ===")
    print("Initial footer wrapper:", footer_shape)
    print("Underlying footer shape initial geometry:",
          f"x={footer_shape.x}, y={footer_shape.y}, width={footer_shape.width}, height={footer_shape.height}")

    x = footer_shape.x
    y = footer_shape.y
    width = footer_shape.width
    height = footer_shape.height
    footer_top = y
    footer_bottom = y + height

    min_distance_to_shape_below = float("inf")
    min_distance_to_shape_above = float("inf")

    # Vertical space analysis
    for _, other_row in other_rows_df.iterrows():
        other_wrapper = other_row["shape"]
        other_shape = other_wrapper.shape
        if other_shape == footer_shape:
            continue
        shape_top = other_shape.y
        shape_bottom = other_shape.y + other_shape.height

        # Footer fully inside another shape
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

    expand_up = min_distance_to_shape_above if min_distance_to_shape_above != float("inf") else 0
    expand_down = min_distance_to_shape_below if min_distance_to_shape_below != float("inf") else 0
    print(f"Computed vertical gaps: expand_up={expand_up}, expand_down={expand_down}")

    total_possible_expansion = expand_up + expand_down
    current_height = footer_shape.height
    max_expandable = max_footer_height - current_height
    print(f"Current height: {current_height}, max_footer_height: {max_footer_height}, max_expandable: {max_expandable}")

    if total_possible_expansion > 0 and total_possible_expansion > max_expandable:
        scale_factor = max_expandable / total_possible_expansion
        expand_up = int(expand_up * scale_factor)
        expand_down = int(expand_down * scale_factor)
        print(
            f"Scaling expansions with factor {scale_factor:.3f}: new expand_up={expand_up}, expand_down={expand_down}")

    new_y = footer_shape.y - expand_up
    new_height = footer_shape.height + expand_up + expand_down
    new_height = min(new_height, max_footer_height)
    print(
        f"Vertical expansion applied: new_y={new_y}, new_height(before clamp)={footer_shape.height + expand_up + expand_down}, after clamp={new_height}")

    footer_shape.y = new_y
    try:
        footer_shape.height = new_height
    except Exception as e:
        print("Warning: failed to set footer shape height:", e)

    # Horizontal expansion
    work_area_left = work_area["left"]
    work_area_right = work_area["right"]
    print(f"Work area bounds: left={work_area_left}, right={work_area_right}")
    print("before expansion ")
    # display(other_rows_df)
    # ---------------- Left Expansion ----------------
    # Find all shapes to the left of footer that vertically overlap
    left_candidates = []
    for _, other_row in other_rows_df.iterrows():
        other = other_row["shape"].shape
        if other == footer_shape:
            continue
        # vertical overlap check
        if other_row["top"] < y + height and other_row["bottom"] > y:
            right_edge = other_row["right"]
            if right_edge <= x:  # valid left candidate
                left_candidates.append(right_edge)

    if left_candidates:
        print("left_candidates ", left_candidates)
        nearest_left = max(left_candidates)
        new_x = max(work_area_left, nearest_left)  # don't go beyond work area
    elif width < slide_width * 0.70:
        new_x = work_area_left  ## previously it was x
    else:
        new_x = max(x, work_area_left)

    new_width_left = x - new_x
    print(f"Expanded left by {new_width_left} (from x={x} to new_x={new_x})")

    # ---------------- Right Expansion ----------------
    right_candidates = []
    for _, other_row in other_rows_df.iterrows():
        other = other_row["shape"].shape
        if other == footer_shape:
            continue
        # vertical overlap check
        if other_row["top"] < y + height and other_row["bottom"] > y:
            left_edge = other_row["left"]
            print("left_edge ", left_edge)
            if left_edge >= x + width:  # valid right candidate
                right_candidates.append(left_edge)

    if right_candidates:
        print("right_candidates ", right_candidates)
        nearest_right = min(right_candidates)
        new_right = min(work_area_right, nearest_right)  # stop at nearest shape or work area
    else:
        new_right = x + width

    new_width_right = new_right - (x + width)
    print(f"Expanded right by {new_width_right} (from right={x + width} to new_right={new_right})")

    # ---------------- Apply Expansion ----------------
    footer_shape.x = new_x
    footer_shape.width = width + new_width_left + new_width_right

    print("Final expanded footer shape geometry:")
    print(f"  x={footer_shape.x}, y={footer_shape.y}, width={footer_shape.width}, height={footer_shape.height}")
    print("=== expand_footer_shape done ===")

    # # Expand left
    # new_x = x
    # while new_x - padding >= work_area_left:
    #     collided = False
    #     for _, other_row in other_rows_df.iterrows():
    #         other_wrapper = other_row["shape"]
    #         other_shape = other_wrapper.shape
    #         if other_shape == footer_shape:
    #             continue
    #         # check vertical overlap with updated footer
    #         if (other_shape.y < footer_shape.y + footer_shape.height and
    #                 other_shape.y + other_shape.height > footer_shape.y):
    #             if other_shape.x + other_shape.width > new_x - padding > other_shape.x:
    #                 collided = True
    #                 break
    #     if collided:
    #         print(f"Left expansion collision at x={new_x - padding}")
    #         break
    #     new_x -= padding
    # new_width_left = x - new_x
    # print(f"Expanded left by {new_width_left} (from x={x} to new_x={new_x})")
    # print("before right expansion")
    # display(other_rows_df)
    # # Expand right
    # new_right = x + width
    # while new_right + padding <= work_area_right:
    #     collided = False
    #     for _, other_shape_row in other_rows_df.iterrows():
    #         other_wrapper = other_shape_row["shape"]
    #         other_shape = other_wrapper.shape
    #         if other_shape == footer_shape:
    #             continue
    #         if (other_shape.y < footer_shape.y + footer_shape.height and
    #                 other_shape.y + other_shape.height > footer_shape.y):
    #             if other_shape.x < new_right + padding < other_shape.x + other_shape.width:
    #                 collided = True
    #                 break
    #     if collided:
    #         print(f"Right expansion collision at right={new_right + padding}")
    #         break
    #     new_right += padding
    # new_width_right = new_right - (x + width)
    # print(f"Expanded right by {new_width_right} (from right={x + width} to new_right={new_right})")

    # # Apply horizontal changes
    # footer_shape.x = new_x + padding
    # footer_shape.width = width + new_width_left + new_width_right

    # print("Final expanded footer shape geometry:")
    # print(f"  x={footer_shape.x}, y={footer_shape.y}, width={footer_shape.width}, height={footer_shape.height}")
    # print("=== expand_footer_shape done ===")
