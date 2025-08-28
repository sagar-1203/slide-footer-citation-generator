import pandas as pd
from app.utils.generate_text_fill_utils.text_fill.format_slides.unpack_slide import unpack_shapes

def find_all_the_shapes(master):
    """
    Collect all shapes from a slide master, its parent presentation master, 
    and its layout slide into DataFrames.

    Args:
        master (slides.MasterSlide): Slide master object.

    Returns:
        tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
            - all_shapes_df: Shapes directly from the given master.
            - layout_df_master: Shapes from the presentation master (excluding placeholders).
            - layout_df: Shapes from the first layout slide.
    """
    fin = []
    key = 0

    for sh in master.shapes:
        out, key = unpack_shapes(sh, [0], key)
        fin += out
    all_shapes_df = pd.DataFrame(fin)
    fin = []
    ##Adding Shapes from Master and adding to layout
    for sh in master.presentation.masters[0].shapes:
        if sh.hidden:
            continue
        out, key = unpack_shapes(sh, [-1], key)
        fin += out
    layout_df_master=pd.DataFrame(fin)
    layout_df_master.to_csv("layout_df_master.csv")
    for index,row in layout_df_master.iterrows():
        if row['shape'].shape.placeholder is not None:
            layout_df_master.drop(index, inplace=True)
    fin=[]
    for sh in master.presentation.layout_slides[0].shapes:
        out, key = unpack_shapes(sh, [-1], key)
        fin += out
    layout_df = pd.DataFrame(fin)
    return all_shapes_df, layout_df_master,layout_df

def remove_shapes_outside_slide_with_threshold(df, slide_width, slide_height, threshold=5):
    """
    Remove shapes that are completely outside slide boundaries unless 'isfillable' is True.

    Parameters:
    df (pd.DataFrame): DataFrame with 'top', 'bottom', 'left', 'right', 'isfillable' columns.
    slide_width (float): Width of the slide.
    slide_height (float): Height of the slide.
    threshold (float): Margin threshold to allow slight overflow.

    Returns:
    pd.DataFrame: Filtered DataFrame.
    """
    # print("This one is called ")
    required_cols = ['top', 'bottom', 'left', 'right', 'isfillable']
    for col in required_cols:
        if col not in df.columns:
            return df

    # Check whether each shape is outside bounds
    outside_horizontally = (df['left'] < 0) | (df['right'] > slide_width + threshold)
    outside_vertically = (df['top'] < 0) | (df['bottom'] > slide_height + threshold)

    # Filter condition: shape is outside AND is not fillable
    filter_mask = ~(outside_horizontally | outside_vertically) | (df['isfillable'] == True)

    return df[filter_mask].copy()

def remove_shapes_outside_slide(df, slide_width, slide_height):
    """
    Remove shapes that are completely outside slide boundaries unless 'isfillable' is True.

    Parameters:
    df (pd.DataFrame): DataFrame with 'top', 'bottom', 'left', 'right', 'isfillable' columns.
    slide_width (float): Width of the slide.
    slide_height (float): Height of the slide.

    Returns:
    pd.DataFrame: Filtered DataFrame.
    """
    required_cols = ['top', 'bottom', 'left', 'right', 'isfillable']
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"DataFrame must contain '{col}' column")

    # Boolean mask for shapes completely outside vertical bounds
    outside_vertically = ((df['top'] < 0) & (df['bottom'] < 0)) | \
                         ((df['top'] >= slide_height) & (df['bottom'] >= slide_height))

    # Boolean mask for shapes completely outside horizontal bounds
    outside_horizontally = ((df['left'] < 0) & (df['right'] <= 0)) | \
                           ((df['left'] >= slide_width) & (df['right'] > slide_width))

    # Final mask: shape is outside AND isfillable is False
    mask_outside_and_not_fillable = (outside_vertically | outside_horizontally) # & (df['isfillable'] == False)

    # Filter out those rows
    filtered_df = df[~mask_outside_and_not_fillable].copy()
    return filtered_df.copy()

def get_collision_info_2d(q,g):
    top_crop = None; bottom_crop = None
    left_crop = None; right_crop = None
    intersect = 0
    if g.top <= q.top and q.top<=g.bottom<=q.bottom:
        inter_vertical = 'TOP_INTERSECT_V'
        top_crop = g.bottom - q.top
        bottom_crop = q.bottom - g.bottom
        intersect = min(top_crop,bottom_crop)
    elif g.bottom > q.bottom and q.top<=g.top<=q.bottom:
        inter_vertical = 'BOTTOM_INTERSECT_V'
        bottom_crop = q.bottom - g.top
        top_crop = g.top - q.top
        intersect = min(top_crop,bottom_crop)
    elif q.top<= g.top<=q.bottom and q.top<=g.bottom<=q.bottom:
        inter_vertical = 'G_ENCLOSED_BY_Q_V'
        top_crop = g.bottom - q.top
        bottom_crop = q.bottom - g.top
        intersect = min(top_crop,bottom_crop)
    elif g.top<= q.top<=g.bottom and g.top<=q.bottom<=g.bottom:
        inter_vertical = 'Q_ENCLOSED_BY_G_V'
    else:
        inter_vertical = 'NONE_V'

    if g.left <= q.left and q.left<=g.right<=q.right:
        inter_horizontal = 'LEFT_INTERSECT_H'
        left_crop = g.right - q.left
        right_crop = q.right - g.right
        intersect = min(left_crop,right_crop)
    elif g.right > q.right and q.left<=g.left<=q.right:
        inter_horizontal = 'RIGHT_INTERSECT_H'
        right_crop = q.right - g.left
        left_crop = g.left - q.left
        intersect = min(left_crop,right_crop)
    elif q.left<= g.left<=q.right and q.left<=g.right<=q.right:
        inter_horizontal = 'G_ENCLOSED_BY_Q_H'
        left_crop = g.right - q.left
        right_crop = q.right - g.left
        intersect = min(left_crop,right_crop)

    elif g.left<= q.left<=g.right and g.left<=q.right<=g.right:
        inter_horizontal = 'Q_ENCLOSED_BY_G_H'
    else:
        inter_horizontal = 'NONE_H'

    return {'inter_vertical':inter_vertical, 'inter_horizontal':inter_horizontal,'left_crop':left_crop,
             'right_crop':right_crop, 'top_crop':top_crop, 'bottom_crop':bottom_crop, 'intersect':intersect,
             'z_coor_g':g.true_z_position(), 'z_coor_q':q.true_z_position(), 'g_wh_ratio':g.width/g.height, 'g':g.item_key,
             'q_corr': q.start,'g_corr':g.start}
