from utils import remove_shapes_outside_slide_dicts, is_fillable
import pandas as pd
def find_max_footer_area(slide, work_area, bottom, min_width=20, min_height=10, max_height=20):
    left, right, work_area_top, work_area_bottom = work_area["left"], work_area["right"], work_area["top"], work_area[
        "bottom"]
    def get_unique_shapes(slide):
        seen = set()
        unique_shapes = []
        for shape in list(slide.shapes) + list(slide.presentation.masters[0].shapes) + list(
                slide.presentation.layout_slides[0].shapes):  # list(slide.presentation.masters[0].shapes) +
            if shape.hidden:
                continue
            key = (shape.x, shape.y, shape.width, shape.height)
            if key not in seen:
                seen.add(key)
                unique_shapes.append(shape)
        return unique_shapes

    def get_collision_info_2d(q, g):
        top_crop = 0;
        bottom_crop = 0
        left_crop = 0;
        right_crop = 0
        vertical_intersect = 0;
        horizontal_intersect = 0
        if g.y <= q['y'] and q['y'] <= g.y + g.height <= q['y'] + q['height']:
            inter_vertical = 'TOP_INTERSECT_V'
            top_crop = (g.y + g.height) - q['y']
            bottom_crop = (q['y'] + q['height']) - (g.y + g.height)
            vertical_intersect = min(top_crop, bottom_crop)
        elif g.y + g.height > q['y'] + q['height'] and q['y'] <= g.y <= q['y'] + q['height']:
            inter_vertical = 'BOTTOM_INTERSECT_V'
            bottom_crop = (q['y'] + q['height']) - g.y
            top_crop = g.y - q['y']
            vertical_intersect = min(top_crop, bottom_crop)
        elif q['y'] <= g.y <= q['y'] + q['height'] and q['y'] <= g.y + g.height <= q['y'] + q['height']:
            inter_vertical = 'G_ENCLOSED_BY_Q_V'
            top_crop = (g.y + g.height) - q['y']
            bottom_crop = (q['y'] + q['height']) - g.y
            vertical_intersect = min(top_crop, bottom_crop)
        elif g.y <= q['y'] <= g.y + g.height and g.y <= q['y'] + q['height'] <= g.y + g.height:
            inter_vertical = 'Q_ENCLOSED_BY_G_V'
        else:
            inter_vertical = 'NONE_V'

        if g.x <= q['x'] and q['x'] <= g.x + g.width <= q['x'] + q['width']:
            inter_horizontal = 'LEFT_INTERSECT_H'
            left_crop = (g.x + g.width) - q['x']
            right_crop = (q['x'] + q['width']) - (g.x + g.width)
            horizontal_intersect = min(left_crop, right_crop)
        elif g.x + g.width > q['x'] + q['width'] and q['x'] <= g.x <= q['x'] + q['width']:
            inter_horizontal = 'RIGHT_INTERSECT_H'
            right_crop = (q['x'] + q['width']) - g.x
            left_crop = g.x - q['x']
            horizontal_intersect = min(left_crop, right_crop)
        elif q['x'] <= g.x <= q['x'] + q['width'] and q['x'] <= g.x + g.width <= q['x'] + q['width']:
            inter_horizontal = 'G_ENCLOSED_BY_Q_H'
            left_crop = (g.x + g.width) - q['x']
            right_crop = (q['x'] + q['width']) - g.x
            horizontal_intersect = min(left_crop, right_crop)
        elif g.x <= q['x'] <= g.x + g.width and g.x <= q['x'] + q['width'] <= g.x + g.width:
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

    # Define initial q
    top = max(bottom - max_height, work_area_bottom)
    bottom = bottom[0] if isinstance(bottom, list) else bottom
    top = top[0] if isinstance(top, list) else top
    q = {'x': left, 'y': top, 'width': right - left, 'height': bottom - top}
    print("q ", q)
    all_shapes = get_unique_shapes(slide)
    print("all shapes")
    for shape in all_shapes:
        print("x,y,w,h ", shape.x, shape.y, shape.width, shape.height)
    #     bottom_shapes_df_for_between_footer = filtered_df[((filtered_df['bottom'] > working_area_bottom) &
    #                                                            (filtered_df['top'] < working_area_bottom) &
    #                                                            (~filtered_df['isfillable'])) |
    #                                                          ((filtered_df['top'] > working_area_bottom) &
    #                                                           (filtered_df['top'] < footer_split_y))].copy()
    #     shapes = [sh for sh in all_shapes if ((sh.y + sh.height > work_area_bottom)]
    shapes = []
    for sh in all_shapes:
        shape_top = sh.y
        shape_bottom = sh.y + sh.height

        condition1 = (
                shape_bottom > work_area_bottom and
                shape_top < work_area_bottom and
                not is_fillable(sh)
        )

        condition2 = (
                shape_top > work_area_bottom and
                shape_top < bottom
        )

        if condition1 or condition2:
            shapes.append(sh)
    slide_width = slide.presentation.slide_size.size.width
    slide_height = slide.presentation.slide_size.size.height
    shapes = remove_shapes_outside_slide_dicts(shapes, slide_width, slide_height)
    shapes.sort(key=lambda sh: sh.y + sh.height, reverse=True)
    print("shapes ")
    for shape in shapes:
        print("x,y,w,h ", shape.x, shape.y, shape.width, shape.height)
    ### HORIZONTAL COLLISION PHASE
    #     horizontal_data = []
    #     for sh in shapes:
    #         col = get_collision_info_2d(q, sh)
    #         if col['inter_horizontal'] != 'NONE_H' and col['inter_vertical'] != 'NONE_V':
    #             col.update({
    #                 'shape_x': sh.x,
    #                 'shape_y': sh.y,
    #                 'shape_w': sh.width,
    #                 'shape_h': sh.height
    #             })
    #             horizontal_data.append(col)

    #     df_horiz = pd.DataFrame(horizontal_data)
    #     print("horizontal")
    #     display(df_horiz)
    #     lc = 0; rc = 0
    #     for _, row in df_horiz.iterrows():
    #         if row.left_crop is not None or row.right_crop is not None:
    #             if row.g_wh_ratio < 3:
    #                 if row.inter_horizontal == 'LEFT_INTERSECT_H':
    #                     lc = max(lc, row.left_crop)
    #                 elif row.inter_horizontal == 'RIGHT_INTERSECT_H':
    #                     rc = max(rc, row.right_crop)
    #             else:
    #                 if row.inter_vertical in ['Q_ENCLOSED_BY_G_V', 'TOP_INTERSECT_V']:
    #                     rc = max(rc, row.right_crop)
    # #     print("before lc, rc", lc,rc)
    # #     if lc >= rc and rc != 0:
    # #         lc = 0
    # #     elif lc < rc and lc != 0:
    # #         rc = 0
    #     print("After lc , rc ", lc,rc)
    #     if lc + rc < right:
    #         q['x'] += lc
    #         q['width'] -= (lc + rc+ 1)
    #     else:
    #         if lc > rc:
    #             q['x'] += lc
    #             q['width'] -= (lc+ 1)
    #         else:
    #             q['width'] -= (rc+ 1)
    #     print(f'After horizontal collision: x={q["x"]}, y={q["y"]}, w={q["width"]}, h={q["height"]}')

    ### VERTICAL COLLISION PHASE
    vertical_data = []
    for sh in shapes:
        col = get_collision_info_2d(q, sh)
        #         if col['inter_vertical'] != 'NONE_V':
        if col['inter_horizontal'] != 'NONE_H' and col['inter_vertical'] != 'NONE_V':
            col.update({
                'shape_x': sh.x,
                'shape_y': sh.y,
                'shape_w': sh.width,
                'shape_h': sh.height
            })
            vertical_data.append(col)

    df_vert = pd.DataFrame(vertical_data)
    print("vertical")
    # display(df_vert)
    tc = 0;
    bc = 0
    crop_thresh = 0.01
    for _, row in df_vert.iterrows():
        if row.top_crop or row.bottom_crop:
            if True or row.g_wh_ratio < 4:
                if row.top_crop < row.bottom_crop:
                    tc = max(tc, row.top_crop) + crop_thresh
                else:
                    bc = max(bc, row.bottom_crop) + crop_thresh
            else:
                if row.top_crop > row.bottom_crop:
                    tc = max(tc, row.top_crop) + crop_thresh
                else:
                    bc = max(bc, row.bottom_crop) + crop_thresh
    print("tc,bc ", tc, bc)
    if tc > bc:
        if bc != 0:
            q['height'] -= bc
        else:
            q['y'] += tc
            q['height'] -= tc
    elif bc > tc:
        if tc != 0:
            q['y'] += tc
            q['height'] -= tc
        else:
            q['height'] -= bc
    #     q['y'] += tc
    #     q['height'] -= (tc + bc)
    print(f'After vertical collision: x={q["x"]}, y={q["y"]}, w={q["width"]}, h={q["height"]}')

    ### HORIZONTAL COLLISION PHASE
    horizontal_data = []
    for sh in shapes:
        col = get_collision_info_2d(q, sh)
        if col['inter_horizontal'] != 'NONE_H' and col['inter_vertical'] != 'NONE_V':
            col.update({
                'shape_x': sh.x,
                'shape_y': sh.y,
                'shape_w': sh.width,
                'shape_h': sh.height
            })
            horizontal_data.append(col)

    df_horiz = pd.DataFrame(horizontal_data)
    print("horizontal")
    # display(df_horiz)
    lc = 0;
    rc = 0
    for _, row in df_horiz.iterrows():
        if row.left_crop or row.right_crop:
            if True or row.g_wh_ratio < 4:
                if row.inter_horizontal == 'LEFT_INTERSECT_H':
                    lc = max(lc, row.left_crop)
                elif row.inter_horizontal == 'RIGHT_INTERSECT_H':
                    rc = max(rc, row.right_crop)
            else:
                if row.inter_vertical in ['Q_ENCLOSED_BY_G_V', 'TOP_INTERSECT_V']:
                    rc = max(rc, row.right_crop)

    #     if lc >= rc and rc != 0:
    #         lc = 0
    #     elif lc < rc and lc != 0:
    #         rc = 0

    #     q['x'] += lc
    #     q['width'] -= (lc + rc)
    print("after lc,rc ", lc, rc)
    # if lc > rc:
    #     if rc != 0:
    #         q['width'] -= (rc+ 1)
    #     else:
    #         q['x'] += (lc+1)
    #         q['width'] -= (lc+ 1)
    # elif rc > lc:
    #     if lc!=0:
    #         q['x'] += (lc+1)
    #         q['width'] -= (lc+ 1)
    #     else:
    #         q['width'] -= (rc+ 1)
    if lc + rc < right:
        q['x'] += lc
        q['width'] -= (lc + rc + 1)
    else:
        if lc > rc:
            q['x'] += lc
            q['width'] -= (lc + 1)
        else:
            q['width'] -= (rc + 1)
    print(f'After horizontal collision: x={q["x"]}, y={q["y"]}, w={q["width"]}, h={q["height"]}')

    if q['width'] >= min_width and q['height'] >= min_height:
        return {
            "x": q['x'],
            "y": q['y'],
            "width": q['width'],
            "height": q['height']
        }
    return None
