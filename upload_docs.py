import os
from pyRM import files_in_folder, upload_to_rm
from tqdm import tqdm
import re

def tidy_filename(filepath):
    full_path = os.path.abspath(filepath)
    folder = os.sep.join(full_path.split(os.sep)[:-1])
    filename = full_path.split(os.sep)[-1]
    # just me and extension
    new_filename = filename.replace("granatm_", "").lower().removesuffix(".pdf")
    # snake language
    new_filename = new_filename.replace(" ", "_")
    new_filename = new_filename.replace("-", "_")
    new_filename = re.sub("_+", "_", new_filename)
    # hungarian characters
    new_filename = new_filename.replace("é", "e") 
    new_filename = new_filename.replace("á", "a")
    new_filename = new_filename.replace("ó", "o")
    new_filename = new_filename.replace("ő", "o")
    new_filename = new_filename.replace("ö", "o")
    new_filename = new_filename.replace("ú", "u")
    new_filename = new_filename.replace("ü", "u")
    new_filename = new_filename.replace("ű", "u")
    new_filename = new_filename.replace("í", "i")
    new_filename = re.sub(r'[^a-zA-Z0-9_]', '', new_filename) + ".pdf"

    os.rename(os.path.abspath(f"{folder}/{filename}"), os.path.abspath(f"{folder}/{new_filename}"))

    return new_filename

def upload_pdfs(folder = "attachments") -> None:
    filenames = [_ for _ in os.listdir(folder) if _.endswith(".pdf") and not _.startswith("~")]
    existing_files = [_ + ".pdf" for _ in files_in_folder(rm_folder=folder)]
    filenames = list(set(filenames) - set(existing_files))

    for filename in tqdm(filenames, desc = f"uploading from {folder}"):
        new_filename = tidy_filename(f"{folder}/{filename}")
        filepath = os.path.join(folder, new_filename)
        upload_to_rm(filepath, rm_folder=folder)