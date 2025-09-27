import boto3
from botocore.exceptions import NoCredentialsError, ClientError
import uuid
import os
import logging

logger = logging.getLogger(__name__)

# Load env. variable for upload strategy
UPLOAD_STRATEGY = os.environ.get("UPLOAD_STRATEGY", "LOCAL")

# Checks value of env. variable
if UPLOAD_STRATEGY == "LOCAL":
    logger.info("Local upload strategy set.")

# Loads required env. variables for S3 upload strategy
elif UPLOAD_STRATEGY == "S3":
    AWS_ACCESS_KEY = os.environ.get("AWS_ACCESS_KEY")
    AWS_SECRET_ACCESS_KEY = os.environ.get("AWS_SECRET_ACCESS_KEY")
    AWS_REGION = os.environ.get('AWS_REGION')
    S3_BUCKET = os.environ.get("S3_BUCKET")
    if not AWS_REGION:
        logger.error("Missing AWS_REGION env. variable")
    elif not AWS_ACCESS_KEY:
        logger.error("Missing AWS_ACCESS_KEY env. variable")
    elif not AWS_SECRET_ACCESS_KEY:
        logger.error("Missing AWS_SECRET_ACCESS_KEY env. variable")
    elif not S3_BUCKET:
        logger.error("Missing S3_BUCKET env. variable")
    else:
        logger.info("S3 upload strategy set, all required env. variable provided.")

else:
    logger.error("Invalid upload strategy, set either to LOCAL or S3")

def generate_unique_object_name(suffix):
    """Generate a unique object name using UUID and preserve the file extension.

    :return: Unique object name with the same file extension
    """

    # Generate a UUID
    unique_id = str(uuid.uuid4())
    # Combine UUID and extension
    unique_object_name = f"{unique_id}.{suffix}"

    return unique_object_name

def upload_file(file_object, suffix):
    """Upload a file to an S3 bucket and return a pre-signed URL valid for 1 hour.

    :param file_object: File-like object to upload
    :return: Pre-signed URL string if successful, else None
    """

    object_name = generate_unique_object_name(suffix)

    if UPLOAD_STRATEGY == "LOCAL":
        return upload_to_local_folder(file_object, object_name)
    elif UPLOAD_STRATEGY == "S3":
        return upload_to_s3(file_object, object_name)
    else:
        return "No upload strategy set, presentation cannot be created."

def upload_to_s3(file_object, file_name):

    # Create an S3 client
    s3_client = boto3.client('s3', region_name=AWS_REGION, aws_access_key_id=AWS_ACCESS_KEY,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY, endpoint_url=f'https://s3.{AWS_REGION}.amazonaws.com')

    if "pptx" in file_name:
        content_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    elif "docx" in file_name:
        content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif "xlsx" in file_name:
        content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif "eml" in file_name:
        content_type = "application/octet-stream"
    else:
        raise ValueError("Unknown file type")

    try:
        # Upload the file to S3
        s3_client.upload_fileobj(Fileobj=file_object, Bucket=S3_BUCKET, Key=file_name, ExtraArgs={'ContentType': content_type})

        # Generate a pre-signed URL valid for 1 hour (3600 seconds)
        url = s3_client.generate_presigned_url('get_object',
                                               Params={'Bucket': S3_BUCKET,
                                                       'Key': file_name},
                                               ExpiresIn=3600)

        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for 1 hour."

    except FileNotFoundError:
        print(f"The file {file_object} was not found.")
        return None
    except NoCredentialsError:
        print("AWS credentials are not available.")
        return None
    except ClientError as e:
        print(f"Client error: {e}")
        return None

def upload_to_local_folder(file_object, file_name):

    """save_path = f'/app/output/{file_name}'"""
    save_path = f'{file_name}'

    with open(save_path, 'wb') as f:
        f.write(file_object.read())

    return f"Inform user that the document {file_name} was saved to his output folder."