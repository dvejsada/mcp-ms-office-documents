import os
import logging

logger = logging.getLogger(__name__)


def upload_to_local_folder(file_object, file_name: str):
    """
    Save the provided file-like object into the working upload folder: ./app/upload

    This function no longer accepts an external output directory and always
    writes to a fixed location relative to the current working directory.
    """
    # Fixed working upload folder (relative to the process CWD)
    save_dir = os.path.join(os.getcwd(), "output")
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, file_name)

    try:
        file_object.seek(0)
        with open(save_path, 'wb') as f:
            f.write(file_object.read())

        logger.info("Saved file to %s", save_path)
        return f"Document saved to {save_path}"
    except Exception as e:
        logger.exception("Failed to save file locally")
        return f"Error saving document locally: {e}\n"
