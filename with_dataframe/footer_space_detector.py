import pandas as pd
from utils import remove_shapes_outside_slide_with_threshold
# from app.utils.generate_text_fill_utils.text_fill.format_slides.collision_utils import get_collision_info_2d
from utils import remove_shapes_outside_slide, get_collision_info_2d

def _unique_shapes_from_dfs(df1, df2, df3):
    """
    Iterate df1, df2, df3 in order and return a list of unique underlying shape objects.
    Dedup key is shape.id if present else (left, top, width, height).
    Assumes each row has a 'shape' wrapper and real shape is at row['shape'].shape.
    """
    seen = set()
    unique_shapes = []
    for df in (df1, df2, df3):
        if df is None or df.empty:
            continue
        for _, row in df.iterrows():
            wrapper = row.get("shape")
            if wrapper is None:
                continue
            shape = getattr(wrapper, "shape", None)
            if shape is None:
                continue
            key = (row.get("left"), row.get("top"), row.get("width"), row.get("height"))
            if key in seen:
                continue
            seen.add(key)
            unique_shapes.append(row)
    return unique_shapes


class PseudoShape:
    def __init__(self, x, y, width, height):
        self.x = x
        self.y = y
        self.width = width
        self.height = height

        # Derived properties
        self.left = x
        self.top = y
        self.right = x + width
        self.bottom = y + height
        self.start = -1

    def true_z_position(self):
        return -1

    def __repr__(self):
        return (f"PseudoShape(x={self.x}, y={self.y}, width={self.width}, height={self.height}, "
                f"left={self.left}, top={self.top}, right={self.right}, bottom={self.bottom})")


def find_max_footer_area_df_2(
        slide,
        work_area,
        bottom,
        all_shapes_df,
        layout_df_master,
        layout_df,
        min_width=20,
        min_height=10,
        max_height=20,
):
    left = work_area["left"]
    right = work_area["right"]
    work_area_top = work_area["top"]
    work_area_bottom = work_area["bottom"]

    # Normalize bottom (can be list)
    bottom_val = bottom[0] if isinstance(bottom, (list, tuple)) else bottom
    top = max(bottom_val - max_height, work_area_bottom)
    top = top[0] if isinstance(top, (list, tuple)) else top

    q = {"x": left, "y": top, "width": right - left, "height": bottom_val - top}
    print("Initial q ", q)

    all_shapes = _unique_shapes_from_dfs(all_shapes_df, layout_df_master, layout_df)
    print("all unique shapes (after dedupe):")
    for shape in all_shapes:
        print(" x,y,w,h ", shape.x, shape.y, shape.width, shape.height)

    # Filter shapes similar to original conditions
    shapes = []
    for sh in all_shapes:
        shape_top = sh.y
        shape_bottom = sh.y + sh.height

        condition1 = (
                shape_bottom > work_area_bottom
                and shape_top < work_area_bottom
                and not is_fillable(sh)
        )
        condition2 = (
                shape_top > work_area_bottom
                and shape_top < bottom_val
        )

        if condition1 or condition2:
            shapes.append(sh)

    slide_width = slide.presentation.slide_size.size.width
    slide_height = slide.presentation.slide_size.size.height

    shapes = remove_shapes_outside_slide_dicts(shapes, slide_width, slide_height, threshold=0)
    shapes.sort(key=lambda sh: sh.y + sh.height, reverse=True)
    print("Filtered shapes for footer proximity:")
    for shape in shapes:
        print(" x,y,w,h ", shape.x, shape.y, shape.width, shape.height)

    ### VERTICAL COLLISION PHASE
    vertical_data = []
    for sh in shapes:
        col = get_collision_info_2d(q, sh)
        if col["inter_horizontal"] != "NONE_H" and col["inter_vertical"] != "NONE_V":
            col.update({
                "shape_x": sh.x,
                "shape_y": sh.y,
                "shape_w": sh.width,
                "shape_h": sh.height
            })
            vertical_data.append(col)

    df_vert = pd.DataFrame(vertical_data)
    print("vertical collision data")
    # print(df_vert)  # optionally display

    tc = 0
    bc = 0
    crop_thresh = 0.01
    for _, row in df_vert.iterrows():
        top_crop = row.get("top_crop", 0)
        bottom_crop = row.get("bottom_crop", 0)
        if top_crop or bottom_crop:
            # original logic uses always-first branch
            if top_crop < bottom_crop:
                tc = max(tc, top_crop) + crop_thresh
            else:
                bc = max(bc, bottom_crop) + crop_thresh
    print("tc, bc ", tc, bc)

    if tc > bc:
        if bc != 0:
            q["height"] -= bc
        else:
            q["y"] += tc
            q["height"] -= tc
    elif bc > tc:
        if tc != 0:
            q["y"] += tc
            q["height"] -= tc
        else:
            q["height"] -= bc
    print(f'After vertical collision: x={q["x"]}, y={q["y"]}, w={q["width"]}, h={q["height"]}')

    ### HORIZONTAL COLLISION PHASE
    horizontal_data = []
    for sh in shapes:
        col = get_collision_info_2d(q, sh)
        if col["inter_horizontal"] != "NONE_H" and col["inter_vertical"] != "NONE_V":
            col.update({
                "shape_x": sh.x,
                "shape_y": sh.y,
                "shape_w": sh.width,
                "shape_h": sh.height
            })
            horizontal_data.append(col)

    df_horiz = pd.DataFrame(horizontal_data)
    print("horizontal collision data")
    # print(df_horiz)  # optionally display

    lc = 0
    rc = 0
    for _, row in df_horiz.iterrows():
        left_crop = row.get("left_crop", 0)
        right_crop = row.get("right_crop", 0)
        if left_crop or right_crop:
            # original branch always taken
            if row.get("inter_horizontal") == "LEFT_INTERSECT_H":
                lc = max(lc, left_crop)
            elif row.get("inter_horizontal") == "RIGHT_INTERSECT_H":
                rc = max(rc, right_crop)
            else:
                if row.get("inter_vertical") in ["Q_ENCLOSED_BY_G_V", "TOP_INTERSECT_V"]:
                    rc = max(rc, right_crop)
    print("after lc, rc ", lc, rc)

    if lc + rc < right:
        q["x"] += lc
        q["width"] -= (lc + rc + 1)
    else:
        if lc > rc:
            q["x"] += lc
            q["width"] -= (lc + 1)
        else:
            q["width"] -= (rc + 1)
    print(f'After horizontal collision: x={q["x"]}, y={q["y"]}, w={q["width"]}, h={q["height"]}')

    if q["width"] >= min_width and q["height"] >= min_height:
        return {
            "x": q["x"],
            "y": q["y"],
            "width": q["width"],
            "height": q["height"]
        }
    return None


def find_max_footer_area_df(
        slide,
        work_area,
        bottom,
        all_shapes_df,
        layout_df_master,
        layout_df,
        min_width=20,
        min_height=10,
        max_height=20,
):
    """
    Find the largest available footer area on a slide by analyzing shape collisions 
    and adjusting a candidate rectangle.

    Args:
        slide (slides.Slide): Current slide object.
        work_area (dict): Work area boundaries with 'left', 'right', 'top', 'bottom'.
        bottom (float | list | tuple): Y-coordinate (or list) marking the footer baseline.
        all_shapes_df (pd.DataFrame): DataFrame of all shapes in the slide.
        layout_df_master (pd.DataFrame): Master layout shapes.
        layout_df (pd.DataFrame): Layout-specific shapes.
        min_width (float, optional): Minimum footer width. Defaults to 20.
        min_height (float, optional): Minimum footer height. Defaults to 10.
        max_height (float, optional): Maximum footer height. Defaults to 20.

    Returns:
        dict | None: Footer box geometry {"x", "y", "width", "height"} if valid, 
                     otherwise None.
    """
    left = work_area["left"]
    right = work_area["right"]
    work_area_top = work_area["top"]
    work_area_bottom = work_area["bottom"]

    # Normalize bottom (can be list)
    bottom_val = bottom[0] if isinstance(bottom, (list, tuple)) else bottom
    top = max(bottom_val - max_height, work_area_bottom)
    top = top[0] if isinstance(top, (list, tuple)) else top

    # Initial footer candidate box
    pseudo_q = PseudoShape(x=left, y=top, width=right - left, height=bottom_val - top)
    print("Initial pseudo_q:", pseudo_q.__dict__)

    # Combine dataframes and deduplicate
    all_shapes_rows = _unique_shapes_from_dfs(all_shapes_df, layout_df_master, layout_df)
    shapes_df = pd.DataFrame(all_shapes_rows)
    print("Unique shapes (after dedupe):", shapes_df.shape)
    # display(shapes_df)
    slide_width = slide.presentation.slide_size.size.width
    slide_height = slide.presentation.slide_size.size.height

    # Filter shapes near footer area
    def is_relevant(row):
        top, bottom = row["top"], row["bottom"]
        return (
                (bottom > work_area_bottom and top < work_area_bottom and not row["isfillable"]) or
                (top > work_area_bottom and top < bottom_val)
        )

    shapes_df = shapes_df[shapes_df.apply(is_relevant, axis=1)]
    print("removing shapes after relevancy check")
    # display(shapes_df)
    shapes_df = remove_shapes_outside_slide_with_threshold(shapes_df, slide_width, slide_height, threshold=1)
    print("removing shapes after outside_slide_with_threshold")
    # display(shapes_df)
    shapes_df = shapes_df.sort_values(by=["bottom"], ascending=False)

    print("Filtered shapes for footer proximity:", shapes_df.shape)
    # display(shapes_df)
    ### VERTICAL COLLISION PHASE
    vertical_data = []
    for _, row in shapes_df.iterrows():
        sh = row["shape"]  # use real shape from row
        col = get_collision_info_2d(pseudo_q, sh)
        if col["inter_horizontal"] != "NONE_H" and col["inter_vertical"] != "NONE_V":
            col.update({
                "shape_x": row.left,
                "shape_y": row.top,
                "shape_w": row.width,
                "shape_h": row.height
            })
            vertical_data.append(col)

    df_vert = pd.DataFrame(vertical_data)
    print("Vertical collision data:")
    # display(df_vert)

    tc, bc = 0, 0
    crop_thresh = 0.01
    for _, row in df_vert.iterrows():
        top_crop = row.get("top_crop", 0)
        bottom_crop = row.get("bottom_crop", 0)
        if top_crop or bottom_crop:
            if top_crop < bottom_crop:
                tc = max(tc, top_crop) + crop_thresh
            else:
                bc = max(bc, bottom_crop) + crop_thresh
    print("tc, bc:", tc, bc)

    if tc > bc:
        if bc != 0:
            pseudo_q.height -= bc
        else:
            pseudo_q.y += tc
            pseudo_q.height -= tc
    elif bc > tc:
        if tc != 0:
            pseudo_q.y += tc
            pseudo_q.height -= tc
        else:
            pseudo_q.height -= bc
    print("After vertical collision:", pseudo_q.__dict__)

    ### HORIZONTAL COLLISION PHASE
    horizontal_data = []
    for _, row in shapes_df.iterrows():
        sh = row["shape"]
        col = get_collision_info_2d(pseudo_q, sh)
        if col["inter_horizontal"] != "NONE_H" and col["inter_vertical"] != "NONE_V":
            col.update({
                "shape_x": row.left,
                "shape_y": row.top,
                "shape_w": row.width,
                "shape_h": row.height
            })
            horizontal_data.append(col)

    df_horiz = pd.DataFrame(horizontal_data)
    print("Horizontal collision data:")
    # display(df_horiz)

    lc, rc = 0, 0
    for _, row in df_horiz.iterrows():
        left_crop = row.get("left_crop", 0)
        right_crop = row.get("right_crop", 0)
        if left_crop or right_crop:
            if row.get("inter_horizontal") == "LEFT_INTERSECT_H":
                lc = max(lc, left_crop)
            elif row.get("inter_horizontal") == "RIGHT_INTERSECT_H":
                rc = max(rc, right_crop)
    #             elif row.get("inter_vertical") in ["Q_ENCLOSED_BY_G_V", "TOP_INTERSECT_V"]:
    #                 rc = max(rc, right_crop)
    print("after lc, rc:", lc, rc)

    if lc + rc < right:
        pseudo_q.x += lc
        pseudo_q.width -= (lc + rc + 1)
    else:
        if lc > rc:
            pseudo_q.x += lc
            pseudo_q.width -= (lc + 1)
        else:
            pseudo_q.width -= (rc + 1)
    print("After horizontal collision:", pseudo_q.__dict__)

    if pseudo_q.width >= min_width and pseudo_q.height >= min_height:
        return {
            "x": pseudo_q.x,
            "y": pseudo_q.y,
            "width": pseudo_q.width,
            "height": pseudo_q.height
        }
    return None

