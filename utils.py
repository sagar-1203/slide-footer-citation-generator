import os
import boto3
import logging
import pickle
import time

logger = logging.getLogger()

try:
    APP_AWS_ACCESS_KEY = os.getenv("APP_AWS_ACCESS_KEY")
    APP_AWS_SECRET_KEY = os.getenv("APP_AWS_SECRET_KEY")
    REGION_NAME = os.getenv("APP_AWS_REGION", "us-west-2")
    boto3_resource = boto3.resource('s3', aws_access_key_id=APP_AWS_ACCESS_KEY,
                                    aws_secret_access_key=APP_AWS_SECRET_KEY)
    boto3_client = boto3.client('s3', aws_access_key_id=APP_AWS_ACCESS_KEY, aws_secret_access_key=APP_AWS_SECRET_KEY)
except Exception as er:
    logger.error(f"Error creating boto s3 object :: {er}")
    raise Exception(f"Error creating boto s3 object :: {er}")


def create_dir(local_path: str):
    """
        create directory if not present
    """
    try:
        if not os.path.isdir(local_path):
            os.makedirs(local_path)
        else:
            logger.info(f"directory {local_path} already exist")
    except Exception as er:
        raise er


def download_file_from_s3(bucket: str, key: str, file_local_path: str, local_file_name: str):
    """
        sample: s3://prez-production-calyrex-dev-images/20230208/124457/prezent/dates_prezent_corporate_2021_deck1_054.pptx
        download file from s3
        input:
            key:  20230208/124457/prezent/dates_prezent_corporate_2021_deck1_054.pptx
            file_local_path: /tmp/
            local_file_name: f8245ca5-fc78-4928-93ad-04cb54c48af0.pptx
    """
    try:
        logger.info(
            f"received download request for file from s3: {bucket}, {key}, {file_local_path}, {local_file_name}")
        # Handling regional s3 path files
        regional_file = get_regional_bucket_path(bucket, key)
        logger.info(
            f"downloading from s3: {regional_file.get('regional_bucket', bucket)}, {regional_file.get('regional_path', key)}, {file_local_path}, {local_file_name}")

        create_dir(file_local_path)
        boto3_resource.Bucket(regional_file.get("regional_bucket", bucket)). \
            download_file(regional_file.get("regional_path", key), "".join([file_local_path, local_file_name]))
        logger.info(f"File downloaded - {local_file_name}")
    except Exception as er:
        logger.error(f"Failed to download file {key} from s3 {er}")
        return er


def upload_file_to_s3(bucket: str, key: str, file_local_path: str, local_file_name: str):
    if key.startswith("/"):
        key = key[1:]
    else:
        key = key
    """
        sample: s3://prez-production-calyrex-dev-images/20230208/124457/prezent/dates_prezent_corporate_2021_deck1_054.pptx
        upload file to s3
        input:
            key:  20230208/124457/prezent/dates_prezent_corporate_2021_deck1_054.pptx
            file_local_path: /tmp/
            local_file_name: f8245ca5-fc78-4928-93ad-04cb54c48af0.pptx
    """
    try:
        logger.info("uploading file to s3")
        boto3_resource.meta.client.upload_file("".join([file_local_path, local_file_name]),
                                               Bucket=bucket, Key=key)
        logger.info(f"File uploaded - {key}")
    except Exception as er:
        logger.error(f"Failed to upload file {local_file_name} to s3 {er}")
        return er


def copy_file_in_s3(bucket: str, old_key: str, new_key: str):
    try:
        # Handle regional path conversion for source (old_key)
        regional_source = get_regional_bucket_path(bucket, old_key)
        source_bucket = regional_source.get("regional_bucket", bucket)
        source_key = regional_source.get("regional_path", old_key)

        # Handle regional path conversion for destination (new_key) if needed
        regional_dest = get_regional_bucket_path(bucket, new_key)
        dest_bucket = regional_dest.get("regional_bucket", bucket)
        dest_key = regional_dest.get("regional_path", new_key)

        logger.info(f"Copying file from s3: {source_bucket}/{source_key} to {dest_bucket}/{dest_key}")

        boto3_resource.meta.client.copy_object(
            Bucket=dest_bucket,
            CopySource={'Bucket': source_bucket, 'Key': source_key},
            Key=dest_key
        )
        logger.info(f"File copied successfully - {dest_key}")
        return None  # Return None on success
    except Exception as er:
        error_msg = f"Failed to copying file {old_key} to s3 {er}"
        logger.error(error_msg)
        return error_msg


def upload_image_to_s3(bucket: str, key: str, file_local_path, local_file_name):
    try:

        logger.info("uploading file to s3")
        print("uploading file to s3")

        boto3_resource.meta.client.upload_file("".join([file_local_path, local_file_name]), bucket, key,
                                               ExtraArgs=dict(ContentType='image/png'))

    except Exception as er:
        logger.error(f"Failed to upload file {local_file_name} to s3 {er}")
        raise Exception(er)


# function to derive right regional path and bucket name, based on UI input
def get_regional_bucket_path(bucket, path):
    regional_bucket = bucket
    regional_path = path
    regional_abb = None

    # geographical regions
    regional_paths = ["eu"]

    # regional shifting
    for region in regional_paths:
        if path.startswith(region + '/'):
            regional_path = path[3:]  # remove the region prefix
            regional_abb = region
            if not bucket.endswith('-' + region):
                regional_bucket = bucket + '-' + region  # add region
            break
        elif bucket.endswith('-' + region):
            regional_abb = region
            break

    return {
        'regional_bucket': regional_bucket,
        'regional_path': regional_path,
        'regional_abb': regional_abb
    }


def remove_shapes_outside_slide_dicts(shapes, slide_width, slide_height, threshold=5):
    """
    Remove shapes that are completely outside slide boundaries unless 'isfillable' is True.

    Parameters:
    shapes (list of dict): Each dict should contain 'x', 'y', 'width', 'height', and optionally 'isfillable'.
    slide_width (float): Width of the slide.
    slide_height (float): Height of the slide.

    Returns:
    list of dict: Filtered list of shapes.
    """
    filtered_shapes = []

    for shape in shapes:
        x = shape.x
        y = shape.y
        w = shape.width
        h = shape.height
        isfillable = False

        left = x
        right = x + w
        top = y
        bottom = y + h

        # Check if completely outside horizontally or vertically
        outside_horizontally = (left < 0) or (right > slide_width + threshold)
        outside_vertically = (top < 0) or (bottom > slide_height + threshold)

        # Only exclude if it's outside and not fillable
        if not (outside_horizontally or outside_vertically) or isfillable:
            filtered_shapes.append(shape)

    return filtered_shapes

def is_fillable(shape):
    if hasattr(shape, 'text_frame'):
        if shape.text_frame is not None:
            return len(shape.text_frame.text.strip()) > 0
    return False