import os

class DocFileDeleter:
    def __init__(self):
        pass
    
    def delete_doc_files(self):
        """
        delete_doc_files deletes any doc and docx files in the current directory
        """
        files = os.listdir()
        for file in files:
            if file.endswith('.doc') or file.endswith('.docx'):
                os.remove(file)

if __name__ == "__main__":
    deleter = DocFileDeleter()
    deleter.delete_doc_files()