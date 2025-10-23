import logging
from datetime import timedelta
from ..utils import get_content_type

logger = logging.getLogger(__name__)


def upload_to_gcs(file_object, file_name: str, gcscfg, signed_url_expires_in: int):
    """Upload a file to a GCS bucket and return a signed URL valid for configured duration."""

    if not gcscfg:
        logger.error("GCS configuration not provided")
        return None

    # Lazy import to avoid requiring google-cloud-storage unless GCS strategy is used
    try:
        from google.cloud import storage  # type: ignore
        from google.cloud.exceptions import GoogleCloudError  # type: ignore
    except Exception:
        logger.error("google-cloud-storage is not installed. Please add it to requirements and install.")
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
            expiration=timedelta(seconds=signed_url_expires_in),
            method="GET"
        )

        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for {signed_url_expires_in} seconds."

    except GoogleCloudError as e:  # type: ignore[name-defined]
        logger.error(f"Google Cloud error: {e}")
        return None
    except Exception as e:
        logger.error(f"Error uploading to GCS: {e}")
        return None
