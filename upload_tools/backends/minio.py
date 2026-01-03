import logging

from ..utils import get_content_type

logger = logging.getLogger(__name__)


def upload_to_minio(file_object, file_name: str, minicfg, signed_url_expires_in: int):
    """Upload a file to a private MinIO bucket and generate a presigned URL."""

    if not minicfg:
        logger.error("MinIO configuration not provided")
        return None

    try:
        import boto3  # type: ignore
        from botocore.config import Config as BotoConfig  # type: ignore
        from botocore.exceptions import BotoCoreError, ClientError, NoCredentialsError  # type: ignore
    except ImportError:
        logger.error("boto3/botocore are required for MinIO uploads")
        return None

    content_type = get_content_type(file_name)

    try:
        addressing_style = "path" if minicfg.path_style else "auto"
        endpoint_is_https = minicfg.endpoint.lower().startswith("https")
        boto_cfg = BotoConfig(signature_version="s3v4", s3={"addressing_style": addressing_style})
        s3_client = boto3.client(
            "s3",
            aws_access_key_id=minicfg.access_key,
            aws_secret_access_key=minicfg.secret_key,
            region_name=minicfg.region,
            endpoint_url=minicfg.endpoint,
            use_ssl=endpoint_is_https,
            verify=minicfg.verify_ssl if endpoint_is_https else False,
            config=boto_cfg,
        )

        file_object.seek(0)
        extra_args = {"ContentType": content_type}
        s3_client.upload_fileobj(file_object, minicfg.bucket, file_name, ExtraArgs=extra_args)

        url = s3_client.generate_presigned_url(
            "get_object",
            Params={"Bucket": minicfg.bucket, "Key": file_name},
            ExpiresIn=signed_url_expires_in,
        )

        return (
            "Link to created document to be shared with user in markdown format: "
            f"{url} . Link is valid for {signed_url_expires_in} seconds."
        )
    except (BotoCoreError, ClientError, NoCredentialsError) as err:
        logger.error("MinIO upload error: %s", err)
        return None
    except Exception as exc:
        logger.error("Unexpected MinIO upload error: %s", exc)
        return None
