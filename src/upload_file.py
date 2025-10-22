import boto3
from botocore.exceptions import NoCredentialsError, ClientError
from google.cloud import storage
from google.cloud.exceptions import GoogleCloudError
import uuid
import os
import logging
from datetime import timedelta, datetime, timezone

logger = logging.getLogger(__name__)

# Load env. variable for upload strategy
UPLOAD_STRATEGY = os.environ.get("UPLOAD_STRATEGY", "LOCAL").upper()

# Configurable signed URL validity (seconds), default 3600s
try:
    SIGNED_URL_EXPIRES_IN = int(os.environ.get("SIGNED_URL_EXPIRES_IN", "3600"))
    if SIGNED_URL_EXPIRES_IN <= 0:
        raise ValueError
except ValueError:
    logger.warning("Invalid SIGNED_URL_EXPIRES_IN, falling back to 3600 seconds")
    SIGNED_URL_EXPIRES_IN = 3600

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

# Loads required env. variables for Azure Blob Storage upload strategy
elif UPLOAD_STRATEGY == "AZURE":
    AZURE_STORAGE_ACCOUNT_NAME = os.environ.get("AZURE_STORAGE_ACCOUNT_NAME")
    AZURE_STORAGE_ACCOUNT_KEY = os.environ.get("AZURE_STORAGE_ACCOUNT_KEY")
    AZURE_CONTAINER = os.environ.get("AZURE_CONTAINER")
    AZURE_BLOB_ENDPOINT = os.environ.get("AZURE_BLOB_ENDPOINT")  # optional, e.g. for sovereign clouds
    if not AZURE_STORAGE_ACCOUNT_NAME:
        logger.error("Missing AZURE_STORAGE_ACCOUNT_NAME env. variable")
    elif not AZURE_STORAGE_ACCOUNT_KEY:
        logger.error("Missing AZURE_STORAGE_ACCOUNT_KEY env. variable")
    elif not AZURE_CONTAINER:
        logger.error("Missing AZURE_CONTAINER env. variable")
    else:
        logger.info("Azure Blob upload strategy set, all required env. variable provided.")

else:
    logger.error("Invalid upload strategy, set either to LOCAL, S3, GCS or AZURE")

def generate_unique_object_name(suffix):
    """Generate a unique object name using UUID and preserve the file extension.

    :return: Unique object name with the same file extension
    """

    # Generate a UUID
    unique_id = str(uuid.uuid4())
    # Combine UUID and extension
    unique_object_name = f"{unique_id}.{suffix}"

    return unique_object_name

def get_content_type(file_name):
    """Determine content type based on file extension.

    :param file_name: Name of the file
    :return: MIME type string
    :raises ValueError: If file type is unknown
    """
    if "pptx" in file_name:
        return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    elif "docx" in file_name:
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif "xlsx" in file_name:
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif "eml" in file_name:
        return "application/octet-stream"
    else:
        raise ValueError("Unknown file type")

def upload_file(file_object, suffix):
    """Upload a file to an S3 bucket, GCS bucket, Azure Blob Storage, or local folder and return appropriate response.

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
    elif UPLOAD_STRATEGY == "AZURE":
        return upload_to_azure(file_object, object_name)
    else:
        return "No upload strategy set, presentation cannot be created."

def upload_to_s3(file_object, file_name):

    # Create an S3 client
    s3_client = boto3.client('s3', region_name=AWS_REGION, aws_access_key_id=AWS_ACCESS_KEY,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY, endpoint_url=f'https://s3.{AWS_REGION}.amazonaws.com')

    content_type = get_content_type(file_name)

    try:
        # Upload the file to S3
        file_object.seek(0)
        s3_client.upload_fileobj(Fileobj=file_object, Bucket=S3_BUCKET, Key=file_name, ExtraArgs={'ContentType': content_type})

        # Generate a pre-signed URL valid for configured duration
        url = s3_client.generate_presigned_url('get_object',
                                               Params={'Bucket': S3_BUCKET,
                                                       'Key': file_name},
                                               ExpiresIn=SIGNED_URL_EXPIRES_IN)

        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for {SIGNED_URL_EXPIRES_IN} seconds."

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

    file_object.seek(0)
    with open(save_path, 'wb') as f:
        f.write(file_object.read())

    return f"Inform user that the document {file_name} was saved to his output folder."

def upload_to_gcs(file_object, file_name):
    """Upload a file to a GCS bucket and return a signed URL valid for configured duration.

    :param file_object: File-like object to upload
    :param file_name: Name of the file to upload
    :return: Message with signed URL if successful, else error message
    """

    content_type = get_content_type(file_name)

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
        
        # Generate a signed URL valid for configured duration
        url = blob.generate_signed_url(
            version="v4",
            expiration=timedelta(seconds=SIGNED_URL_EXPIRES_IN),
            method="GET"
        )
        
        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for {SIGNED_URL_EXPIRES_IN} seconds."

    except GoogleCloudError as e:
        logger.error(f"Google Cloud error: {e}")
        return None
    except Exception as e:
        logger.error(f"Error uploading to GCS: {e}")
        return None

def upload_to_azure(file_object, file_name):
    """Upload a file to Azure Blob Storage and return a SAS URL valid for configured duration.

    :param file_object: File-like object to upload
    :param file_name: Name of the file to upload
    :return: Message with SAS URL if successful, else error message
    """
    try:
        # Import here to avoid requiring azure-storage-blob unless AZURE strategy is used
        from azure.storage.blob import (
            BlobServiceClient,
            generate_blob_sas,
            BlobSasPermissions,
            ContentSettings,
        )
    except ImportError:
        logger.error("azure-storage-blob is not installed. Please add it to requirements and install.")
        return None

    content_type = get_content_type(file_name)

    account_name = os.environ.get("AZURE_STORAGE_ACCOUNT_NAME")
    account_key = os.environ.get("AZURE_STORAGE_ACCOUNT_KEY")
    container_name = os.environ.get("AZURE_CONTAINER")
    endpoint = os.environ.get("AZURE_BLOB_ENDPOINT") or f"https://{account_name}.blob.core.windows.net"

    try:
        # Create a BlobServiceClient
        blob_service_client = BlobServiceClient(account_url=endpoint, credential=account_key)
        container_client = blob_service_client.get_container_client(container_name)

        # Upload the blob
        blob_client = container_client.get_blob_client(file_name)
        file_object.seek(0)
        blob_client.upload_blob(
            file_object,
            overwrite=True,
            content_settings=ContentSettings(content_type=content_type)
        )

        # Generate a SAS token for read access
        expiry_time = datetime.now(timezone.utc) + timedelta(seconds=SIGNED_URL_EXPIRES_IN)
        sas_token = generate_blob_sas(
            account_name=account_name,
            container_name=container_name,
            blob_name=file_name,
            account_key=account_key,
            permission=BlobSasPermissions(read=True),
            expiry=expiry_time,
        )

        url = f"{endpoint}/{container_name}/{file_name}?{sas_token}"
        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for {SIGNED_URL_EXPIRES_IN} seconds."

    except Exception as e:
        logger.error(f"Error uploading to Azure Blob Storage: {e}")
        return None
