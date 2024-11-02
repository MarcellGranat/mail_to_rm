from download_attachments import main as download_attachments
from convert_to_pdf import convert_folder
from upload_docs import upload_pdfs


if __name__ == "__main__":
    download_attachments(days=5)
    convert_folder(folder="attachments")
    upload_pdfs(folder="attachments")
