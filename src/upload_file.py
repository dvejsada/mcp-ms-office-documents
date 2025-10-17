import boto3
from botocore.exceptions import NoCredentialsError, ClientError
from google.cloud import storage
from google.cloud.exceptions import GoogleCloudError
import uuid
import os
import logging
from datetime import timedelta

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

# Loads required env. variables for GCS upload strategy
elif UPLOAD_STRATEGY == "GCS":
    GCS_BUCKET = os.environ.get("GCS_BUCKET")
    GCS_CREDENTIALS_PATH = os.environ.get("GCS_CREDENTIALS_PATH")
    if not GCS_BUCKET:
        logger.error("Missing GCS_BUCKET env. variable")
    elif not GCS_CREDENTIALS_PATH:
        logger.error("Missing GCS_CREDENTIALS_PATH env. variable")
    else:
        logger.info("GCS upload strategy set, all required env. variable provided.")

else:
    logger.error("Invalid upload strategy, set either to LOCAL, S3 or GCS")

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
    """Upload a file to an S3 bucket, GCS bucket or local folder and return appropriate response.

    :param file_object: File-like object to upload
    :param suffix: File extension (e.g., 'pptx', 'docx', 'xlsx', 'eml')
    :return: Status message with download URL or save location
    """

    object_name = generate_unique_object_name(suffix)

    if UPLOAD_STRATEGY == "LOCAL":
        return upload_to_local_folder(file_object, object_name)
    elif UPLOAD_STRATEGY == "S3":
        return upload_to_s3(file_object, object_name)
    elif UPLOAD_STRATEGY == "GCS":
        return upload_to_gcs(file_object, object_name)
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

    save_dir = '/app/output'
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, file_name)

    with open(save_path, 'wb') as f:
        f.write(file_object.read())

    return f"Inform user that the document {file_name} was saved to his output folder."

def upload_to_gcs(file_object, file_name):
    """Upload a file to a GCS bucket and return a signed URL valid for 1 hour.

    :param file_object: File-like object to upload
    :param file_name: Name of the file to upload
    :return: Message with signed URL if successful, else error message
    """

    # Determine content type based on file extension
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
        # Create a GCS client with credentials from the environment variable
        storage_client = storage.Client.from_service_account_json(GCS_CREDENTIALS_PATH)
        
        # Get the bucket
        bucket = storage_client.bucket(GCS_BUCKET)
        
        # Create a blob (object) in the bucket
        blob = bucket.blob(file_name)
        
        # Upload the file to GCS
        file_object.seek(0)  # Reset file pointer to beginning
        blob.upload_from_file(file_object, content_type=content_type)
        
        # Generate a signed URL valid for 1 hour (3600 seconds)
        url = blob.generate_signed_url(
            version="v4",
            expiration=timedelta(hours=1),
            method="GET"
        )
        
        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for 1 hour."

    except GoogleCloudError as e:
        logger.error(f"Google Cloud error: {e}")
        return None
    except Exception as e:
        logger.error(f"Error uploading to GCS: {e}")
        return None