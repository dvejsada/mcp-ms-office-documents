from .local import upload_to_local_folder
from .s3 import upload_to_s3
from .gcs import upload_to_gcs
from .azure import upload_to_azure
from .minio import upload_to_minio

__all__ = [
    "upload_to_local_folder",
    "upload_to_s3",
    "upload_to_gcs",
    "upload_to_azure",
    "upload_to_minio",
]
