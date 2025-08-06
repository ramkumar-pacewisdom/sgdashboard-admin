from google.cloud import storage
from google.oauth2 import service_account
import os
import logging

# Configure logging for better debugging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def upload_file_to_gcs_and_get_directory(bucket_name, source_file_path, destination_blob_name, private_key_path):
    """
    Uploads a file to a Google Cloud Storage bucket and returns the public URL for the folder.

    Args:
        bucket_name (str): Name of the GCS bucket (e.g., 'mentoring-dev-storage-private').
        source_file_path (str): Path to the local file to upload.
        destination_blob_name (str): Name for the file in the bucket (e.g., 'folder/file.txt').
        private_key_path (str): Path to the GCP service account JSON key file.

    Returns:
        str: Public URL for the folder containing the uploaded file, or None if an error occurs.
    """
    try:
        # Validate input file paths
        if not os.path.exists(source_file_path):
            logger.error(f"Source file not found: {source_file_path}")
            return None
        if not os.path.exists(private_key_path):
            logger.error(f"Private key file not found: {private_key_path}")
            return None

        # Initialize GCS client with service account credentials
        logger.info("Initializing GCS client with service account credentials")
        credentials = service_account.Credentials.from_service_account_file(
            private_key_path,
            scopes=['https://www.googleapis.com/auth/cloud-platform']
        )

        # Initialize GCS client and get the bucket
        storage_client = storage.Client(credentials=credentials)
        bucket = storage_client.bucket(bucket_name)

        # Upload the file to the specified bucket
        logger.info(f"Uploading {source_file_path} to {bucket_name}/{destination_blob_name}")
        blob = bucket.blob(destination_blob_name)
        blob.upload_from_filename(source_file_path)

        # Make the uploaded file publicly accessible
        logger.info(f"Making file {destination_blob_name} publicly accessible")
        blob.make_public()

        # Extract the folder path (prefix) from the destination_blob_name
        folder_path = os.path.dirname(destination_blob_name)
        if not folder_path:
            folder_path = ""  # Use empty string for root of the bucket

        # Construct the public folder URL
        public_folder_url = f"https://storage.googleapis.com/{bucket_name}/{folder_path}"
        logger.info(f"Generated public folder URL: {public_folder_url}")

        # Verify if the file is publicly accessible
        if blob.public_url:
            logger.info(f"Public URL for file: {blob.public_url}")
            return public_folder_url
        else:
            logger.error("File is not publicly accessible")
            return None

    except Exception as e:
        logger.error(f"Failed to upload file or generate public URL: {str(e)}")
        return None

