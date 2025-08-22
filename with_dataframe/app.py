import streamlit as st
import os
import io
import logging
import aspose.slides as slides
import pandas as pd
import ast
import json
from io import BytesIO
from with_dataframe.footer_locator_bottom import *
from with_dataframe.footer_space_detector import *
from with_dataframe.footer_orientation_manager import *
import requests
from with_dataframe.s3_utils import *
from with_dataframe.utils import *
from botocore.exceptions import ClientError

# ========== Logger Configuration ==========

LOG_FILENAME = "ppt_processing.log"

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

if not logger.hasHandlers():
    fh = logging.FileHandler(LOG_FILENAME)
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] - %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)


# # ========== Aspose License Setup ==========
# license = slides.License()
# license.set_license("/home/sagar/project/slide-generator/Aspose.Slides.Python.NET.lic")
# folders = ["/home/sagar/project/slide-generator/01-Prezent-Fonts"]
# slides.FontsLoader.load_external_fonts(folders)
# logger.info("Aspose license set and fonts loaded.")

# # ========== Load Config ==========
# config_csv = pd.read_csv("/home/sagar/project/rl-templates-090625.csv")
# logger.info("Template configuration CSV loaded.")


@st.cache_resource
def setup_aspose_license_and_fonts():
    license = slides.License()
    license.set_license("/home/sagar/project/slide-generator/Aspose.Slides.Python.NET.lic")
    folders = ["/home/sagar/project/slide-generator/01-Prezent-Fonts"]
    slides.FontsLoader.load_external_fonts(folders)
    logger.info("Aspose license set and fonts loaded.")
    return True


@st.cache_data
def load_template_config():
    config = pd.read_csv("/home/sagar/project/rl-data-design-data.csv")
    logger.info("Template configuration CSV loaded.")
    return config


# Initialize once
setup_aspose_license_and_fonts()
config_csv = load_template_config()
# ========== User Auth ==========
USER_CREDENTIALS = {"admin": "admin123", "tester": "testpass"}

# ========== Sample Footer Text ==========
text = ("Environmental Protection Agency. (2023). Air Quality Index Report 2023. EPA Publications.\n"
        "Green Research Institute. (2023). Environmental Impact Assessment of Electric Vehicles. Journal of Environmental Studies, 15(3), 45-62.\n"
        "Urban Planning Department. (2023). Urban Noise Study: Impact of Electric Vehicles. City Planning Review, 28(4), 112-128.")

text_list = text.split('\n')


# ========== Authentication ==========
def authenticate():
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if USER_CREDENTIALS.get(username) == password:
            st.session_state.authenticated = True
            logger.info(f"User '{username}' authenticated successfully.")
            st.success("Authenticated successfully!")
        else:
            logger.warning(f"Authentication failed for username: {username}")
            st.error("Invalid credentials")


# ========== Template Metadata Fetch ==========
def get_work_area_from_file_name(file_name: str):
    matched_template = None
    max_match_length = 0

    file_name_lower = file_name.lower()

    # Find the best match by longest matching substring
    for template_name in config_csv['template_internal_name']:
        template_name_lower = template_name.lower()
        if template_name_lower in file_name_lower:
            if len(template_name_lower) > max_match_length:
                matched_template = template_name
                max_match_length = len(template_name_lower)

    if not matched_template:
        print(f"No matching template found in filename: {file_name}")
        return None, None

    # Now fetch the row using the matched template
    filtered_row = config_csv[config_csv['template_internal_name'] == matched_template]
    # display(filtered_row)

    # Step 2: Extract the 'work_area' column value from the matched row
    if not filtered_row.empty:
        raw_value = filtered_row.iloc[0]['auto_conversion_properties']

        # Step 3: Safely parse the value to a dictionary
        props = None
        if isinstance(raw_value, str):
            try:
                props = ast.literal_eval(raw_value)  # safer than eval()
            except Exception:
                try:
                    props = json.loads(raw_value)
                except:
                    props = {}

        # Step 4: Extract values from parsed dictionary
        work_area = props.get('work_area') if isinstance(props, dict) else None
        footer_config = props.get('font_body_color') if isinstance(props, dict) else None
        footer_font = props.get('footer_font') if isinstance(props, dict) else None

        if footer_config and footer_font:
            footer_config["footer_font"] = footer_font

        print("footer_info:", footer_config)
    else:
        work_area = None
        footer_config = None

    print("work_area:", work_area)
    return work_area, footer_config


def get_work_area(template_type: str):
    filtered_row = config_csv[config_csv['template_internal_name'] == template_type]
    # Step 2: Extract the 'work_area' column value from the matched row
    if not filtered_row.empty:
        raw_value = filtered_row.iloc[0]['auto_conversion_properties']
        # Step 3: Safely parse the value to a dictionary
        props = None
        if isinstance(raw_value, str):
            try:
                props = ast.literal_eval(raw_value)  # safer than eval()
            except Exception as e:
                try:
                    props = json.loads(raw_value)
                except:
                    props = {}

        # Step 4: Extract work_area key from the parsed dictionary
        work_area = props.get('work_area') if isinstance(props, dict) else None
        footer_config = props.get('font_body_color') if isinstance(props, dict) else None
        footer_font = props.get('footer_font') if isinstance(props, dict) else None
        if footer_config and footer_font:
            footer_config["footer_font"] = footer_font
        print("footer_info ", footer_config)
    else:
        work_area = None
        footer_config = None

    print("work_area:", work_area)
    return work_area, footer_config

#  ============= Add final Footer shape ===============

def add_final_footer_shape(slide, footer_shape, footer_config, font_size, text_list, add_footer_row_wise):
    text = "\n".join(text_list)
    if add_footer_row_wise:
        add_rectangle_box(slide,footer_shape['x'],footer_shape['y'],footer_shape['width'],footer_shape['height'],footer_config,text,1,font_size)
    else:
        equal_width = (footer_shape['width']-5*(len(text_list))) / len(text_list)
        spacing = 0
        for text in text_list:
            add_rectangle_box(new_slide,footer_shape['x'] + spacing,footer_shape['y'],equal_width,footer_shape['height'],footer_config,text,1,font_size)
            spacing += equal_width + 5
        


# ========== Add Superscript References ==========
from aspose.slides import Portion, PortionFormat

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
                    superscript.text = f"[{count}]"
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


# ========== Remove Footer Shapes ==========
def remove_present_footer(slide):
    qualifying_keywords = ["footer", "source goes here", "goes here", "disclaimer place holder", "edit source",
                           "footnotes", "foot"]
    for shape in list(slide.shapes):
        shape_text = shape.text_frame.text.strip().lower() if hasattr(shape, "text_frame") and shape.text_frame else ""
        name = shape.name.lower() if shape.name else ""

        # Check if duplicate is footer-like
        if any(keyword in shape_text for keyword in qualifying_keywords):
            print("removing shape")
            slide.shapes.remove(shape)


# ========== Build Slide Payload ==========
def build_slide_payload(input_text, source_bucket, source_path, construct,
                        target_bucket, target_path, template_name, slide_type,
                        all_ml_words, config, lang="english", category="title", hub_title="", graph_data=None):
    payload = {
        "source": {
            "s3_bucket": source_bucket,
            "s3_path": source_path,
            "construct": construct,
        },
        "target": {
            "s3_bucket": target_bucket,
            "s3_path": target_path
        },
        "template": template_name,
        "slide_type": slide_type,
        "all_ml_words": all_ml_words,
        "input_text": input_text,
        # "config": config,
        "hub_title": hub_title,
        "graph_data": graph_data,
        "lang": lang,
        "category": category
    }
    return payload


url = "http://108.181.157.3:3008/sb/slide-generator/replace_text"
s3_bucket = "prez-kiran-superhawk-userdetails"
s3_path = "sagar_text_fill/input/"
targer_path = "sagar_text_fill/output/"
addintional_data = " random" * 0
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


def s3_key_exists(bucket, key):
    try:
        boto3_client.head_object(Bucket=bucket, Key=key)
        return True
    except ClientError:
        return False


# ========== Processing Logic ==========
def process_ppt(uploaded_file, footer_text: str):
    try:
        filename = uploaded_file.name
        input_s3_key = f"{s3_path}{filename}"
        output_filename = filename  # os.path.splitext(filename)[0] + "_processed" + os.path.splitext(filename)[1]
        output_s3_key = f"{targer_path}{output_filename}"
        # Save the uploaded file locally first
        temp_dir = "/tmp/"
        local_upload_path = os.path.join(temp_dir, filename)
        with open(local_upload_path, "wb") as f:
            f.write(uploaded_file.read())

        # Upload only if not already in S3
        if s3_key_exists(s3_bucket, input_s3_key):
            print(f"File already exists in S3: {input_s3_key}")
            # st.info(f"File already exists in S3:")
        else:
            upload_file_to_s3(s3_bucket, input_s3_key, temp_dir, filename)
            # st.success(f"Uploaded to S3: {input_s3_key}")
            print(f"Uploaded to S3: {input_s3_key}")
        work_area, footer_config = get_work_area_from_file_name(filename)
        logger.info(f"Work area: {work_area}, Footer config: {footer_config}")
        logger.info(f"Processing file: {filename}")
        config = {"template_config": {"work_area": work_area}}
        payload = build_slide_payload(
            input_text=input_text,
            source_bucket=s3_bucket,
            source_path=input_s3_key,
            construct="list",
            target_bucket=s3_bucket,
            target_path=output_s3_key,
            template_name="servicenow_light_251",
            config=config,
            slide_type="normal",
            all_ml_words='app::object,board::object,circle::shape,computer::object,front::object,list::construct,mobile::object,person::object,plant::object,powerpoint::object,presentation::object,screen::object,social_media::secondary,steps::object,tech::tertiary,vertical::flow,wheelchair::object,woman::object',
            category="non_title",
        )
        print(payload)
        print("Calling Text fill")
        url = "http://localhost:3008/sb/slide-generator/replace_text"
        response = requests.post(url, json=payload)
        print("response from text fill ", response)
        output_dir = "/home/sagar/project/slide-generator/test/text_fill_slides/"
        download_file_from_s3(s3_bucket, output_s3_key, output_dir, filename)
        print(f"Downloaded processed file to: {output_dir}{filename}")
        textfill_slide = f"{output_dir}{filename}"
        pres = slides.Presentation(textfill_slide)
        text_list = footer_text.split('\n')
        # work_area, footer_config = get_work_area_from_file_name(textfill_slide)
        # logger.info(f"Work area: {work_area}, Footer config: {footer_config}")
        # logger.info(f"Processing file: {textfill_slide}")
        original_slides = list(pres.slides)
        decision_message = "No decision made"
        MIN_FONT_SIZE = 4
        MAX_FONT_SIZE = 7

        if work_area:
            for slide in original_slides:
                all_shapes_df, layout_df_master,layout_df = find_all_the_shapes(slide)
                slide_width = slide.presentation.slide_size.size.width
                slide_height = slide.presentation.slide_size.size.height
                print("slide width and height ", slide_width,slide_height)
        #         find_all_the_shapes(slide)
                print("initial rendering of slide")
                # render_ppt(pres,slide)
                split_y, bottom_shapes, footer_shape_list = add_footer_shape_df(slide,work_area, padding=0, collision_threshold=2)
                print("split_y ", split_y)
                print("footer shape ",footer_shape_list)
                max_footer_font_1 = -1
                max_footer_font_2 = -1
                resulted_footer_shape = None
                if footer_shape_list:
                    for footer_shape in footer_shape_list:
                        if footer_shape['x'] >= work_area["left"]:
                            #max_footer_font_1 = min(max_footer_font_1,get_the_max_font_in_column_or_row_wise(new_slide,min_font,max_font,footer_shape,text_list,footer_config))
                            resulted_font_size1, add_footer_row_wise1 = get_the_max_font_in_column_or_row_wise(slide,min_font,max_font,footer_shape,text_list,footer_config)
        #                     max_footer_font_1 = min(max_footer_font_1,resulted_font_size1)
                            if resulted_font_size1 > max_footer_font_1:
                                max_footer_font_1 = resulted_font_size1
                                resulted_footer_shape = footer_shape
                            #add_rectangle_box(new_slide,footer_shape['x'],footer_shape['y'],footer_shape['width'],footer_shape['height'],footer_config,text,1,font_size)
                if split_y > work_area["bottom"] + 5:
                    results = find_max_footer_area_df(slide,work_area,split_y,all_shapes_df,layout_df_master,layout_df,min_width=20,min_height=10,max_height=20,)
        #             results = find_max_footer_area(new_slide_2, work_area,split_y)
        #         results = find_best_footer_area(slide,work_area["abbvie_aquipta"]["work_area"]["left"], split_y)
                    print("results ",results)
                    if results:
        #                 remove_present_footer(slide)
                        resulted_font_size2, add_footer_row_wise2 = get_the_max_font_in_column_or_row_wise(slide,min_font,max_font,results,text_list,footer_config)
                        if resulted_font_size2 > max_footer_font_2:
                            max_footer_font_2 = resulted_font_size2
                    print("max_footer_font_1, max_footer_font_2 ",max_footer_font_1,max_footer_font_2)
                if max_footer_font_1 == -1 and max_footer_font_2 == -1:
                    print("No footer found in both slides")
                elif max_footer_font_1 >= max_footer_font_2:
                    print("font. Use slide 2 As final footer solution")
                    add_final_footer_shape(slide, resulted_footer_shape, footer_config, max_footer_font_1, text_list, add_footer_row_wise1)
                elif max_footer_font_2 > max_footer_font_1:
                    add_final_footer_shape(slide, results, footer_config, max_footer_font_2, text_list, add_footer_row_wise2)
                    print("font. Use slide 1 As final footer solution")
                else:
                    print("No solution fount")
                print("final rendering of slide")
                # for shape_info in merged_results:
                #     remove_shapes_by_info(slide,shape_info)
        else:
            decision_message = "Work Area not found in template configuration."
        output_stream = io.BytesIO()
        pres.save(output_stream, slides.export.SaveFormat.PPTX)
        output_stream.seek(0)
        logger.info(f"Processed file: {uploaded_file.name}")
        return output_stream, decision_message

    except Exception as e:
        logger.error(f"Error processing file {uploaded_file.name}: {e}")
        return None, "Error processing the PPT"


# ========== UI ==========
st.set_page_config(layout="wide")
st.title("Footer and Citation Generator for PowerPoint")

# --- Initialize Session State ---
if "uploaded_files_data" not in st.session_state:
    st.session_state.uploaded_files_data = []
if "processing_results" not in st.session_state:
    st.session_state.processing_results = {}
if "ui_messages" not in st.session_state:
    st.session_state.ui_messages = []
# if "last_uploader_key" not in st.session_state:
#     st.session_state.last_uploader_key = None
if "results_table" not in st.session_state:
    st.session_state.results_table = []
if "footer_text" not in st.session_state:
    st.session_state.footer_text = text

# --- File Upload Section ---
st.header("1. Upload Files")

uploaded_files_widget = st.file_uploader(
    "Upload up to 5 .pptx files",
    type=["pptx"],
    accept_multiple_files=True,
    key="main_ppt_uploader"
)

# --- Input Field for Footer Text ---
st.header("2. Footer Text")
footer_text_input = st.text_area("Enter Footer/Citation Text", height=200)
st.session_state.footer_text = footer_text_input

if uploaded_files_widget:
    current_widget_filenames = {f.name for f in uploaded_files_widget}
    stored_session_filenames = {f_data['name'] for f_data in
                                st.session_state.uploaded_files_data} if st.session_state.uploaded_files_data else set()

    if current_widget_filenames != stored_session_filenames:
        st.session_state.uploaded_files_data = []
        st.session_state.processing_results = {}
        st.session_state.ui_messages = []
        st.session_state.results_table = []

        for uploaded_file in uploaded_files_widget:
            st.session_state.uploaded_files_data.append({
                "name": uploaded_file.name,
                "data": uploaded_file.getvalue()
            })

        st.session_state.ui_messages.append(f"📁 {len(st.session_state.uploaded_files_data)} file(s) loaded.")
        st.rerun()
else:
    if st.session_state.uploaded_files_data and not uploaded_files_widget:
        st.session_state.uploaded_files_data = []
        st.session_state.processing_results = {}
        st.session_state.results_table = []
        st.session_state.ui_messages = ["No files uploaded. Please upload files to begin."]
        st.rerun()

# --- Sidebar Info ---
if st.session_state.uploaded_files_data:
    st.sidebar.subheader("Loaded Files:")
    for file_info in st.session_state.uploaded_files_data:
        st.sidebar.write(f"- {file_info['name']}")
    if len(st.session_state.uploaded_files_data) > 5:
        st.warning("You have loaded more than 5 files. Only the first 5 will be processed.")

# --- Processing Options ---
# dropdown_option = st.selectbox(
#     "Choose processing option:",
#     ["abbvie_aquipta", "sanofi_corp_251", "sanofi_regeneron", "prezent_corporate_2022", "cisco_corp_appdynamics_light", "abbvie_iae_sa_dark_251" , "abbvie_one_pd_251"]
# )

# --- Process Files ---
if st.button("Process Files"):
    if not st.session_state.uploaded_files_data:
        st.session_state.ui_messages.append("⚠️ Please upload files before clicking Process.")
        st.rerun()
    else:
        st.session_state.processing_results = {}
        st.session_state.ui_messages = ["Starting file processing..."]
        st.session_state.results_table = []

        files_to_process = st.session_state.uploaded_files_data[:5]
        # logger.info(f"Processing {len(files_to_process)} files for template: {dropdown_option}")
        logger.info(f"Processing {len(files_to_process)} files.")

        with st.spinner("Processing in progress... Please wait."):
            for file_info in files_to_process:
                uploaded_file_name = file_info["name"]
                uploaded_file_bytes = file_info["data"]

                file_io = io.BytesIO(uploaded_file_bytes)
                file_io.name = uploaded_file_name

                # if dropdown_option.lower() not in uploaded_file_name.lower():
                #     msg = f"Filename '{uploaded_file_name}' does not contain template name '{dropdown_option}'"
                #     st.session_state.ui_messages.append(f"❌ {uploaded_file_name}: {msg}")
                #     logger.warning(msg)
                #     continue

                processed_file_io, decision_message = process_ppt(file_io, st.session_state.footer_text)

                if processed_file_io:
                    st.session_state.processing_results[uploaded_file_name] = {
                        "file_data": processed_file_io.getvalue(),
                        "decision": decision_message
                    }
                    st.session_state.results_table.append({
                        "Presentation": uploaded_file_name,
                        "Slide Number": decision_message or "No decision"
                    })
                else:
                    st.session_state.ui_messages.append(
                        f"❌ Failed to process: {uploaded_file_name} - {decision_message}")

        st.session_state.ui_messages.append("✅ Processing complete!")
        st.rerun()

# --- UI Messages ---
st.header("Status & Results")
for msg in st.session_state.ui_messages:
    st.markdown(msg)

# --- Results Table ---
if st.session_state.results_table:
    st.header("📊 Processing Summary")
    st.dataframe(pd.DataFrame(st.session_state.results_table), use_container_width=True)

# --- Download Buttons ---
if st.session_state.processing_results:
    st.header("Download Processed Files")
    with st.expander("Click to view download links"):
        num_files = len(st.session_state.processing_results)
        cols = st.columns(min(num_files, 3))

        for i, (filename, result) in enumerate(st.session_state.processing_results.items()):
            col = cols[i % 3]
            with col:
                st.download_button(
                    label=f"Download: {filename}",
                    data=result["file_data"],
                    file_name=f"processed_{filename}",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key=f"download_button_{filename}"
                )

# --- Reset ---
if st.button("Reset All", key="reset_button"):
    st.session_state.uploaded_files_data = []
    st.session_state.processing_results = {}
    st.session_state.results_table = []
    st.session_state.ui_messages = ["App reset. Please upload files to start."]
    st.rerun()

    # {'source': {'s3_bucket': 'prez-kiran-superhawk-userdetails', 's3_path': 'sagar_text_fill/input/background_menarini_bd_external_deck1_0183.pptx', 'construct': 'list'},
    #  'target': {'s3_bucket': 'prez-kiran-superhawk-userdetails', 's3_path': 'sagar_text_fill/output/background_menarini_bd_external_deck1_0183.pptx'},
    #  'template': 'servicenow_light_251', 'slide_type': 'normal',
    #    'all_ml_words': 'app::object,board::object,circle::shape,computer::object,front::object,list::construct,mobile::object,person::object,plant::object,powerpoint::object,presentation::object,screen::object,social_media::secondary,steps::object,tech::tertiary,vertical::flow,wheelchair::object,woman::object',
    #    'input_text': {'title': ["Ferrari's Grand Entrance: Roaring into the Indian Luxury Market "], 'subtitle': [], 'content': ['Established 3 state-of-the-art showrooms in key Indian cities', 'Introduced personalized Tailor Made program for Indian customers', 'Hosted Ferrari Challenge racing events to drive brand enthusiasm', 'Announced plans to double the dealer network within 3 years', 'Announced plans to double the dealer network within 3 years'], 'section_header': ['Exclusive Dealerships', 'Customization Offerings', 'Motorsport Engagement', 'Expanding Footprint', 'Expanding Footprint']},
    #    'config': {'template_config': {'work_area': {'left': 59.91241455078125, 'top': 214.94989013671875, 'right': 1523.0880737304688, 'bottom': 813.8257446289062}}},
    #    'hub_title': '', 'graph_data': None, 'lang': 'english', 'category': 'title'}