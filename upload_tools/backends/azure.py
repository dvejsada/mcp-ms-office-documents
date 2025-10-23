import logging
from datetime import timedelta, datetime, timezone
from ..utils import get_content_type

logger = logging.getLogger(__name__)


def upload_to_azure(file_object, file_name: str, azcfg, signed_url_expires_in: int):
    """Upload a file to Azure Blob Storage and return a SAS URL valid for configured duration."""

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
        expiry_time = datetime.now(timezone.utc) + timedelta(seconds=signed_url_expires_in)
        sas_token = generate_blob_sas(
            account_name=account_name,
            container_name=container_name,
            blob_name=file_name,
            account_key=account_key,
            permission=BlobSasPermissions(read=True),
            expiry=expiry_time,
        )

        url = f"{endpoint}/{container_name}/{file_name}?{sas_token}"
        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for {signed_url_expires_in} seconds."

    except Exception as e:
        logger.error(f"Error uploading to Azure Blob Storage: {e}")
        return None

