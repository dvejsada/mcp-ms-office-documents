import os
import logging

logger = logging.getLogger(__name__)


def upload_to_local_folder(file_object, file_name: str, output_dir: str):
    save_dir = output_dir
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, file_name)

    file_object.seek(0)
    with open(save_path, 'wb') as f:
        f.write(file_object.read())

    return f"Inform user that the document {file_name} was saved to his output folder."

