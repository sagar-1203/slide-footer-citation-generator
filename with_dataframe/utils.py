import pandas as pd
from app.utils.generate_text_fill_utils.text_fill.format_slides.unpack_slide import unpack_shapes
# from app.utils.add_footer_and_citation.main import add_footer, add_citations_as_superscript
from main import add_footer, add_citations_as_superscript
import math
import aspose.slides as slides

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

def extract_titles_and_sources(data, title_key, data_for_citation, sources_list):
    """
    Extract titles and title_sources from either a dict or a list of dicts.

    Args:
        data: dict or list of dicts
        title_key: str, the key used for the title ("title" or "table_title")
        data_for_citation: list, where titles will be appended
        sources_list: list, where title_sources will be extended
    """
    if not data:
        return

    # Normalize to list of dicts
    if isinstance(data, dict):
        data = [data]

    for item in data:
        title = item.get(title_key)
        title_sources = item.get("title_sources", [])
        if title_sources:  # only include if title_sources is non-empty
            sources_list.extend(title_sources)
            if title:
                data_for_citation.append(title)

def reorder_and_deduplicate_content_sources(slide, content_data, content_sources, start_id=1):
    """
    Handle content sources for content_data (list of list of dicts or strings).
    Deduplicate citations globally, assign unique IDs, and map IDs per content item.
    Combined sources are ordered according to slide coordinates.

    Args:
        slide: Aspose.Slides Slide object
        content_data: list of content items (each item can be string or dict)
        content_sources: list of list of dicts/strs aligned to content_data
        start_id: starting ID for combined sources

    Returns:
        sorted_data: [{"id": [ids], "data": content_item_text}]
        combined_sources: [{"id": id, "citation": text}] (unique globally, sorted by slide order)
    """
    import math

    # Normalize content sources to align with content_data
    srcs = list(content_sources) if content_sources else []
    if len(srcs) < len(content_data):
        srcs.extend([[]] * (len(content_data) - len(srcs)))

    # Step 1: Prepare per-item normalized sources
    items = []
    for item, src_list in zip(content_data, srcs):
        # extract text
        if isinstance(item, str):
            text = item
        elif isinstance(item, dict):
            text = item.get("data") or item.get("title") or item.get("text") or ""
        else:
            text = str(item)

        # normalize sources into list of dicts with 'citation'
        normalized_sources = []
        if src_list is None:
            normalized_sources = []
        elif isinstance(src_list, (list, tuple)):
            for s in src_list:
                if isinstance(s, str):
                    normalized_sources.append({"citation": s})
                elif isinstance(s, dict):
                    normalized_sources.append({"citation": s.get("citation", "")})
                else:
                    normalized_sources.append({"citation": str(s)})
        elif isinstance(src_list, dict):
            normalized_sources.append({"citation": src_list.get("citation", "")})
        elif isinstance(src_list, str):
            normalized_sources.append({"citation": src_list})
        else:
            normalized_sources.append({"citation": str(src_list)})

        items.append({
            "text": text,
            "text_norm": text.strip().lower(),
            "sources": normalized_sources
        })

    # Step 2: Map content items to slide coordinates
    coords = {}
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or not shape.text_frame:
            continue
        shape_text = shape.text_frame.text
        if not shape_text:
            continue
        shape_norm = shape_text.strip().lower()
        for itm in items:
            if not itm["text_norm"]:
                continue
            if shape_norm == itm["text_norm"] or itm["text_norm"] in shape_norm or shape_norm in itm["text_norm"]:
                coords[itm["text_norm"]] = (float(getattr(shape, "y", 0.0)), float(getattr(shape, "x", 0.0)))
                break

    # Step 3: Sort items by slide coordinates
    sorted_items = sorted(items, key=lambda itm: coords.get(itm["text_norm"], (math.inf, math.inf)))

    # Step 4: Deduplicate citations globally in slide order
    seen_citations = {}    # citation_text -> id
    combined_sources = []
    current_id = start_id
    sorted_data = []

    for itm in sorted_items:
        text = itm["text"]
        ids_for_item = []

        for s in itm["sources"]:
            citation_text = str(s.get("citation", "")).strip()
            if not citation_text:
                continue
            if citation_text not in seen_citations:
                seen_citations[citation_text] = current_id
                combined_sources.append({"id": current_id, "citation": citation_text})
                ids_for_item.append(current_id)
                current_id += 1
            else:
                ids_for_item.append(seen_citations[citation_text])

        # fallback if no valid citation
        if not ids_for_item:
            seen_citations[""] = current_id
            combined_sources.append({"id": current_id, "citation": ""})
            ids_for_item.append(current_id)
            current_id += 1

        sorted_data.append({
            "id": ids_for_item,
            "data": text
        })

    return sorted_data, combined_sources


def reorder_and_deduplicate_sources(slide, data_for_citation, graph_title_sources, table_title_sources, start_id=1):
    # Step 1: Assign sources per title
    sorted_graph_sources = graph_title_sources or []
    sorted_table_sources = table_title_sources or []

    title_to_sources = {}
    for idx, title in enumerate(data_for_citation):
        if idx < len(sorted_graph_sources):
            src_list = sorted_graph_sources[idx] if isinstance(sorted_graph_sources[idx], list) else [sorted_graph_sources[idx]]
        else:
            t_idx = idx - len(sorted_graph_sources)
            src_list = sorted_table_sources[t_idx] if isinstance(sorted_table_sources[t_idx], list) else [sorted_table_sources[t_idx]]
        title_to_sources[title] = [s if isinstance(s, dict) else {"citation": s} for s in src_list]

    # Step 2: Map titles to slide coordinates
    coords = {}
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or not shape.text_frame:
            continue
        text = shape.text_frame.text
        if not text:
            continue
        norm = text.strip().lower()
        for t in data_for_citation:
            if t.strip().lower() == norm:
                coords[t] = (float(getattr(shape, "y", 0.0)), float(getattr(shape, "x", 0.0)))
                break

    # Step 3: Sort titles by coordinates
    sorted_titles = sorted(data_for_citation, key=lambda t: coords.get(t, (math.inf, math.inf)))

    # Step 4: Deduplicate globally and assign IDs according to sorted_titles
    seen_citations = {}
    combined_sources = []
    current_id = start_id
    sorted_data = []

    for t in sorted_titles:
        ids_for_title = []
        for s in title_to_sources.get(t, []):
            citation_text = s.get("citation", "").strip()
            if not citation_text:
                continue
            if citation_text not in seen_citations:
                seen_citations[citation_text] = current_id
                combined_sources.append({"id": current_id, "citation": citation_text})
                current_id += 1
            ids_for_title.append(seen_citations[citation_text])
        sorted_data.append({"id": ids_for_title, "data": t})

    return sorted_data, combined_sources


def add_footer_and_citation(slide, input_body, presentation, file_path):
    try:
        is_footer_added = False
        is_footer_needed = input_body.get('is_footer_needed', False)
        if is_footer_needed:
            print("inside footer")
            data_for_citation = []
            graph_title_sources, table_title_sources = [], []

            graph_data = input_body.get("graph_data", [])
            print("length of graph data", len(graph_data))
            extract_titles_and_sources(graph_data, "title", data_for_citation,
                                       graph_title_sources)
            table_data = input_body.get("table_data", {})
            print("length of table_data", len(table_data))
            extract_titles_and_sources(table_data, "table_title", data_for_citation,
                                       table_title_sources)
            length_for_citation_data = 0
            if not (graph_title_sources or table_title_sources):
                content_sources = input_body.get("input_text", {}).get(
                    "content_sources") or []
                # title_sources = input_body.get("input_text", {}).get("title_sources") or []
                input_text_title = input_body.get("input_text", {}).get("title") or ""
                # if input_text_title:
                #     data_for_citation += input_text_title
                input_text_content = input_body.get("input_text", {}).get("content", [])
                length_for_citation_data = len(input_text_content)
                if input_text_content and content_sources:
                    data_for_citation += input_text_content
                sorted_content, combined_sources = reorder_and_deduplicate_content_sources(
                    slide,
                    data_for_citation,
                    content_sources,
                    start_id=1
                )

                print("Sorted content data with IDs:")
                for d in sorted_content:
                    print(d)

                print("\nDeduplicated combined content sources:")
                for src in combined_sources:
                    print(src)
            else:
                print("inside else part")
                length_for_citation_data = len(graph_data) + len(table_data)
                sorted_content, combined_sources = reorder_and_deduplicate_sources(
                    slide,
                    data_for_citation,
                    graph_title_sources,
                    table_title_sources,
                    start_id=1
                )

                print("Sorted Data with IDs:")
                for d in sorted_content:
                    print(d)

                print("\nCombined Sources (deduplicated):")
                for src in combined_sources:
                    print(src)

            # print("graph_title_sources, table_title_sources, content_sources",
            #     graph_title_sources, table_title_sources, content_sources)
            if combined_sources and len(combined_sources) > 0:
                work_area = None
                footer_config = None
                template_config = input_body.get("config", {}).get("template_config", {})

                if isinstance(template_config, dict):
                    footer_config = template_config.get('font_body_color')
                    footer_font = template_config.get('footer_font')

                    if footer_config and footer_font:
                        footer_config = {"font_body_color": footer_config,
                                         "footer_font": footer_font}
                    elif footer_config or footer_font:
                        footer_config = {"font_body_color": footer_config,
                                         "footer_font": footer_font}

                    work_area = template_config.get("work_area")
                is_citation_needed = True
                if len(combined_sources) == 1 and len(sorted_content) == length_for_citation_data:
                    is_citation_needed = False
                    combined_sources[0]["id"] = -1
                is_footer_needed = True
                for src in combined_sources:
                    citation = src.get("citation")
                    if citation is None or len(citation) == 0:
                        is_footer_needed = False
                        break
                if is_footer_needed:
                    slide, is_footer_added = add_footer(slide, work_area, combined_sources,
                                                        footer_config, MIN_FONT_SIZE=4,
                                                        MAX_FONT_SIZE=7)
                # print("finieshed the add_footer", is_footer_added)
                # print("data for citation", data_for_citation)

                if is_footer_added and is_citation_needed:
                    # presentation.save(
                    #     "/home/sagar/project/graph_table_chart_slide_generator/new_one/slide-generator/build/after_add_footer.pptx",
                    #     slides.export.SaveFormat.PPTX)
                    slide = add_citations_as_superscript(slide, sorted_content, wrap_text=True)
                if is_footer_added:
                    presentation.save(file_path,
                                      slides.export.SaveFormat.PPTX)
        return slide, is_footer_added
    except Exception as e:
        print("Exception while adding footer", e)
        return slide,False
