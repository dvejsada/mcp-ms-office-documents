import boto3
from botocore.exceptions import NoCredentialsError, ClientError
from google.cloud import storage
from google.cloud.exceptions import GoogleCloudError
import uuid
import os
import logging
from datetime import timedelta, datetime, timezone

from config import get_config

logger = logging.getLogger(__name__)

# Load centralized configuration
cfg = get_config()

# Convenience aliases
UPLOAD_STRATEGY = cfg.storage.strategy
SIGNED_URL_EXPIRES_IN = cfg.storage.signed_url_expires_in
OUTPUT_DIR = cfg.storage.output_dir

# Strategy announcement logs
if UPLOAD_STRATEGY == "LOCAL":
    logger.info("Local upload strategy set.")
elif UPLOAD_STRATEGY == "S3":
    logger.info("S3 upload strategy set.")
elif UPLOAD_STRATEGY == "GCS":
    logger.info("GCS upload strategy set.")
elif UPLOAD_STRATEGY == "AZURE":
    logger.info("Azure Blob upload strategy set.")
else:
    logger.error("Invalid upload strategy, set either to LOCAL, S3, GCS or AZURE")


def generate_unique_object_name(suffix: str) -> str:
    """Generate a unique object name using UUID and preserve the file extension."""
    unique_id = str(uuid.uuid4())
    return f"{unique_id}.{suffix}"


def get_content_type(file_name: str) -> str:
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


def upload_file(file_object, suffix: str):
    """Upload a file to configured backend and return appropriate response.

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


def upload_to_s3(file_object, file_name: str):
    s3cfg = cfg.storage.s3
    if not s3cfg:
        logger.error("S3 configuration not provided")
        return None

    # Create an S3 client
    s3_client = boto3.client(
        's3',
        region_name=s3cfg.region,
        aws_access_key_id=s3cfg.access_key,
        aws_secret_access_key=s3cfg.secret_key,
        endpoint_url=f'https://s3.{s3cfg.region}.amazonaws.com'
    )

    content_type = get_content_type(file_name)

    try:
        # Upload the file to S3
        file_object.seek(0)
        s3_client.upload_fileobj(Fileobj=file_object, Bucket=s3cfg.bucket, Key=file_name, ExtraArgs={'ContentType': content_type})

        # Generate a pre-signed URL valid for configured duration
        url = s3_client.generate_presigned_url(
            'get_object',
            Params={'Bucket': s3cfg.bucket, 'Key': file_name},
            ExpiresIn=SIGNED_URL_EXPIRES_IN
        )

        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for {SIGNED_URL_EXPIRES_IN} seconds."

    except FileNotFoundError:
        logger.error(f"The file {file_object} was not found.")
        return None
    except NoCredentialsError:
        logger.error("AWS credentials are not available.")
        return None
    except ClientError as e:
        logger.error(f"Client error: {e}")
        return None


def upload_to_local_folder(file_object, file_name: str):
    save_dir = OUTPUT_DIR
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, file_name)

    file_object.seek(0)
    with open(save_path, 'wb') as f:
        f.write(file_object.read())

    return f"Inform user that the document {file_name} was saved to his output folder."


def upload_to_gcs(file_object, file_name: str):
    """Upload a file to a GCS bucket and return a signed URL valid for configured duration."""

    gcscfg = cfg.storage.gcs
    if not gcscfg:
        logger.error("GCS configuration not provided")
        return None

    content_type = get_content_type(file_name)

    try:
        # Create a GCS client with credentials from the configured path
        storage_client = storage.Client.from_service_account_json(gcscfg.credentials_path)

        bucket = storage_client.bucket(gcscfg.bucket)
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


def upload_to_azure(file_object, file_name: str):
    """Upload a file to Azure Blob Storage and return a SAS URL valid for configured duration."""

    azcfg = cfg.storage.azure
    if not azcfg:
        logger.error("Azure configuration not provided")
        return None

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

    account_name = azcfg.account_name
    account_key = azcfg.account_key
    container_name = azcfg.container
    endpoint = azcfg.endpoint or f"https://{account_name}.blob.core.windows.net"

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
