import logging
from config import get_config
from .utils import generate_unique_object_name
from .backends.local import upload_to_local_folder
from .backends.s3 import upload_to_s3
from .backends.gcs import upload_to_gcs
from .backends.azure import upload_to_azure

logger = logging.getLogger(__name__)

# Load centralized configuration
cfg = get_config()

# Convenience aliases
UPLOAD_STRATEGY = cfg.storage.strategy
SIGNED_URL_EXPIRES_IN = cfg.storage.signed_url_expires_in

# Strategy announcement logs
if UPLOAD_STRATEGY == "LOCAL":
    logger.info("Local upload strategy set.")
elif UPLOAD_STRATEGY == "S3":
    logger.info("S3 upload strategy set.")
elif UPLOAD_STRATEGY == "GCS":
    logger.info("GCS upload strategy set.")
elif UPLOAD_STRATEGY == "AZURE":
    logger.info("Azure Blob upload strategy set.")


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
        return upload_to_s3(file_object, object_name, cfg.storage.s3, SIGNED_URL_EXPIRES_IN)
    elif UPLOAD_STRATEGY == "GCS":
        return upload_to_gcs(file_object, object_name, cfg.storage.gcs, SIGNED_URL_EXPIRES_IN)
    elif UPLOAD_STRATEGY == "AZURE":
        return upload_to_azure(file_object, object_name, cfg.storage.azure, SIGNED_URL_EXPIRES_IN)
    else:
        return "No upload strategy set, presentation cannot be created."
