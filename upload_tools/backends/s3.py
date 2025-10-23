import logging
from ..utils import get_content_type

logger = logging.getLogger(__name__)


def upload_to_s3(file_object, file_name: str, s3cfg, signed_url_expires_in: int):
    if not s3cfg:
        logger.error("S3 configuration not provided")
        return None

    # Lazy import to avoid requiring boto3 unless S3 strategy is used
    try:
        import boto3  # type: ignore
        from botocore.exceptions import NoCredentialsError, ClientError  # type: ignore
    except Exception as e:
        logger.error("boto3/botocore are not installed. Please add them to requirements and install.")
        return None

    content_type = get_content_type(file_name)

    try:
        # Create an S3 client
        s3_client = boto3.client(
            's3',
            region_name=s3cfg.region,
            aws_access_key_id=s3cfg.access_key,
            aws_secret_access_key=s3cfg.secret_key,
            endpoint_url=f'https://s3.{s3cfg.region}.amazonaws.com'
        )

        # Upload the file to S3
        file_object.seek(0)
        s3_client.upload_fileobj(Fileobj=file_object, Bucket=s3cfg.bucket, Key=file_name, ExtraArgs={'ContentType': content_type})

        # Generate a pre-signed URL valid for configured duration
        url = s3_client.generate_presigned_url(
            'get_object',
            Params={'Bucket': s3cfg.bucket, 'Key': file_name},
            ExpiresIn=signed_url_expires_in
        )

        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for {signed_url_expires_in} seconds."

    except FileNotFoundError:
        logger.error(f"The file {file_object} was not found.")
        return None
    except NoCredentialsError:
        logger.error("AWS credentials are not available.")
        return None
    except ClientError as e:
        logger.error(f"Client error: {e}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error uploading to S3: {e}")
        return None
