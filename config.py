"""Centralized configuration and logging setup for the MCP Office Documents server.

This module is the single source of truth for reading and validating environment
variables. No other module should access os.environ directly.

Highlights:
- Reads all env vars and constructs a typed Config instance (Pydantic v2).
- Validates required settings based on chosen upload strategy (LOCAL/S3/GCS/AZURE).
- Configures global logging (format and level) exactly once on first access.
- Exposes get_config() to retrieve a singleton Config across the app.

Environment variables (see .env.example for full list):
- Logging: DEBUG (true/false)
- Storage generic: UPLOAD_STRATEGY, SIGNED_URL_EXPIRES_IN
- Strategy specific: AWS_*, GCS_*, AZURE_*
"""

from __future__ import annotations

import logging
import os
from enum import Enum
from typing import Optional
from pydantic import BaseModel, Field, ValidationError, model_validator


class LogLevel(str, Enum):
    """Application log levels (restricted to INFO and DEBUG)."""
    DEBUG = "DEBUG"
    INFO = "INFO"


class LoggingSettings(BaseModel):
    """Logging configuration simplified to a single DEBUG flag.

    Behavior:
    - If DEBUG env var is truthy (1/true/on), logging is DEBUG.
    - Otherwise logging is INFO.

    Exposes convenience properties used across the app:
    - level_no: numeric logging level
    - mcp_level_str: lower-case string for FastMCP's `log_level` argument
    """
    debug: bool = Field(default=False, description="True to enable DEBUG level, False for INFO")

    @property
    def level_no(self) -> int:
        return logging.DEBUG if self.debug else logging.INFO

    @property
    def mcp_level_str(self) -> str:
        """Return lower-case string for FastMCP `log_level` argument."""
        return "debug" if self.debug else "info"


class S3Settings(BaseModel):
    """Required credentials and configuration for AWS S3 uploads."""
    access_key: str
    secret_key: str
    region: str
    bucket: str

    @model_validator(mode="after")
    def _non_empty(self) -> "S3Settings":
        """Ensure all S3 fields are non-empty, raising a helpful error otherwise."""
        missing = [
            name for name, val in (
                ("AWS_ACCESS_KEY", self.access_key),
                ("AWS_SECRET_ACCESS_KEY", self.secret_key),
                ("AWS_REGION", self.region),
                ("S3_BUCKET", self.bucket),
            ) if not str(val).strip()
        ]
        if missing:
            raise ValueError(f"Missing required S3 settings: {', '.join(missing)}")
        return self


class GCSSettings(BaseModel):
    """Required configuration for Google Cloud Storage uploads."""
    bucket: str
    credentials_path: str

    @model_validator(mode="after")
    def _non_empty(self) -> "GCSSettings":
        """Ensure all GCS fields are non-empty, raising a helpful error otherwise."""
        missing = [
            name for name, val in (
                ("GCS_BUCKET", self.bucket),
                ("GCS_CREDENTIALS_PATH", self.credentials_path),
            ) if not str(val).strip()
        ]
        if missing:
            raise ValueError(f"Missing required GCS settings: {', '.join(missing)}")
        return self


class AzureSettings(BaseModel):
    """Required configuration for Azure Blob Storage uploads.

    Note: `endpoint` is optional; if empty, defaults to
    https://<account>.blob.core.windows.net
    """
    account_name: str
    account_key: str
    container: str
    endpoint: Optional[str] = None

    @model_validator(mode="after")
    def _non_empty(self) -> "AzureSettings":
        """Ensure all required Azure fields are non-empty."""
        missing = [
            name for name, val in (
                ("AZURE_STORAGE_ACCOUNT_NAME", self.account_name),
                ("AZURE_STORAGE_ACCOUNT_KEY", self.account_key),
                ("AZURE_CONTAINER", self.container),
            ) if not str(val).strip()
        ]
        if missing:
            raise ValueError(f"Missing required Azure settings: {', '.join(missing)}")
        return self


class MinioSettings(BaseModel):
    """Configuration for self-hosted MinIO (S3-compatible) uploads."""

    endpoint: str = Field(description="Base URL of the MinIO server, e.g., http://minio:9000")
    access_key: str
    secret_key: str
    bucket: str
    region: str = Field(default="us-east-1", description="Region to report to boto3; defaults to us-east-1")
    verify_ssl: bool = Field(default=True, description="Whether to verify SSL certificates when connecting")
    path_style: bool = Field(default=True, description="Use path-style addressing (recommended for MinIO)")

    @model_validator(mode="after")
    def _non_empty(self) -> "MinioSettings":
        missing = [
            name for name, val in (
                ("MINIO_ENDPOINT", self.endpoint),
                ("MINIO_ACCESS_KEY", self.access_key),
                ("MINIO_SECRET_KEY", self.secret_key),
                ("MINIO_BUCKET", self.bucket),
            ) if not str(val).strip()
        ]
        if missing:
            raise ValueError(f"Missing required MinIO settings: {', '.join(missing)}")
        return self


class StorageStrategy(str, Enum):
    """Supported upload backends for produced documents."""
    LOCAL = "LOCAL"
    S3 = "S3"
    GCS = "GCS"
    AZURE = "AZURE"
    MINIO = "MINIO"


class StorageSettings(BaseModel):
    """Generic storage configuration plus strategy-specific nested settings.

    Note: The LOCAL strategy always writes to the working folder ./app/upload;
    there is no configurable output directory for LOCAL.
    """
    strategy: StorageStrategy = Field(default=StorageStrategy.LOCAL)
    signed_url_expires_in: int = Field(default=3600, gt=0, description="TTL for S3/GCS/Azure download links in seconds")

    # Optional nested settings depending on strategy
    s3: Optional[S3Settings] = None
    gcs: Optional[GCSSettings] = None
    azure: Optional[AzureSettings] = None
    minio: Optional[MinioSettings] = None

    @model_validator(mode="after")
    def validate_strategy_requirements(self) -> "StorageSettings":
        """Ensure required nested settings exist for chosen strategy."""
        if self.strategy == StorageStrategy.S3:
            if not self.s3:
                raise ValueError("S3 settings are required for S3 strategy")
        elif self.strategy == StorageStrategy.GCS:
            if not self.gcs:
                raise ValueError("GCS settings are required for GCS strategy")
        elif self.strategy == StorageStrategy.AZURE:
            if not self.azure:
                raise ValueError("Azure settings are required for AZURE strategy")
        elif self.strategy == StorageStrategy.MINIO:
            if not self.minio:
                raise ValueError("MinIO settings are required for MINIO strategy")
        return self


class Config(BaseModel):
    """Top-level configuration container used by the whole application."""
    logging: LoggingSettings
    storage: StorageSettings

    @staticmethod
    def _parse_bool(value: Optional[str]) -> bool:
        """Interpret common truthy representations used in env vars."""
        if value is None:
            return False
        return value.strip().lower() in {"1", "true", "yes", "y", "on"}

    @classmethod
    def from_env(cls) -> "Config":
        """Construct Config from environment variables with sensible defaults and validation.

        This does not configure logging by itself; see configure_logging().
        """
        # Logging: only use DEBUG env var (truthy -> DEBUG, falsy -> INFO)
        debug = cls._parse_bool(os.environ.get("DEBUG"))
        logging_settings = LoggingSettings(debug=debug)

        # Storage
        raw_strategy = (os.environ.get("UPLOAD_STRATEGY", "LOCAL")).upper()
        strategy = raw_strategy if raw_strategy in {e.value for e in StorageStrategy} else "LOCAL"

        # Signed URL expiry (fallback to 3600 on invalid input)
        try:
            expires_in = int(os.environ.get("SIGNED_URL_EXPIRES_IN", "3600"))
            if expires_in <= 0:
                raise ValueError
        except ValueError:
            expires_in = 3600

        # Strategy-specific settings (only populate the relevant one)
        s3_settings = None
        gcs_settings = None
        azure_settings = None
        minio_settings = None

        if strategy == StorageStrategy.S3.value:
            s3_settings = S3Settings(
                access_key=os.environ.get("AWS_ACCESS_KEY", ""),
                secret_key=os.environ.get("AWS_SECRET_ACCESS_KEY", ""),
                region=os.environ.get("AWS_REGION", ""),
                bucket=os.environ.get("S3_BUCKET", ""),
            )
        elif strategy == StorageStrategy.GCS.value:
            gcs_settings = GCSSettings(
                bucket=os.environ.get("GCS_BUCKET", ""),
                credentials_path=os.environ.get("GCS_CREDENTIALS_PATH", ""),
            )
        elif strategy == StorageStrategy.AZURE.value:
            azure_settings = AzureSettings(
                account_name=os.environ.get("AZURE_STORAGE_ACCOUNT_NAME", ""),
                account_key=os.environ.get("AZURE_STORAGE_ACCOUNT_KEY", ""),
                container=os.environ.get("AZURE_CONTAINER", ""),
                endpoint=os.environ.get("AZURE_BLOB_ENDPOINT"),
            )
        elif strategy == StorageStrategy.MINIO.value:
            minio_settings = MinioSettings(
                endpoint=os.environ.get("MINIO_ENDPOINT", ""),
                access_key=os.environ.get("MINIO_ACCESS_KEY", ""),
                secret_key=os.environ.get("MINIO_SECRET_KEY", ""),
                bucket=os.environ.get("MINIO_BUCKET", ""),
                region=os.environ.get("MINIO_REGION", "us-east-1") or "us-east-1",
                verify_ssl=cls._parse_bool(os.environ.get("MINIO_VERIFY_SSL", "true")),
                path_style=cls._parse_bool(os.environ.get("MINIO_PATH_STYLE", "true")),
            )

        storage_settings = StorageSettings(
            strategy=StorageStrategy(strategy),
            signed_url_expires_in=expires_in,
            s3=s3_settings,
            gcs=gcs_settings,
            azure=azure_settings,
            minio=minio_settings,
        )

        try:
            return cls(logging=logging_settings, storage=storage_settings)
        except ValidationError as e:
            # Wrap Pydantic validation errors in a simpler exception for callers
            raise ValueError(f"Invalid configuration: {e}")


# Singleton instance and guard for one-time logging configuration
_CONFIG: Optional[Config] = None
_LOGGING_CONFIGURED: bool = False


def configure_logging(config: Config) -> None:
    """Configure root logger format and level exactly once.

    - Uses a more verbose format (file:line) in DEBUG level.
    - Keeps concise formatting otherwise.
    """
    global _LOGGING_CONFIGURED
    if _LOGGING_CONFIGURED:
        return

    level = config.logging.level_no
    root = logging.getLogger()
    if not root.handlers:
        handler = logging.StreamHandler()
        # Use the debug flag (simplified API)
        if config.logging.debug:
            fmt = "%(asctime)s | %(levelname)s | %(name)s:%(lineno)d | %(message)s"
        else:
            fmt = "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
        handler.setFormatter(logging.Formatter(fmt))
        root.addHandler(handler)
    root.setLevel(level)
    _LOGGING_CONFIGURED = True


def get_config() -> Config:
    """Return the process-wide Config singleton and ensure logging is configured."""
    global _CONFIG
    if _CONFIG is None:
        cfg = Config.from_env()
        configure_logging(cfg)
        _CONFIG = cfg
    return _CONFIG
