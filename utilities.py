import os


# Use os.path.basename to extract the filename from a full path string
# e.g. path = c:\Users\work\tkinter\utilities.py
#      fn = utilities.py
def extract_file(path):
    fn = os.path.basename(path)
    return fn


# extracts file extension, i.e. file type
# e.g. fn = customs.pdf
#      ext = pdf
def extract_file_extension(fn):
    ext = os.path.splitext(fn)[1]
    return ext


# extracts file name
# e.g. fn = customs.pdf
#      fn_output = customs
def extract_file_name(fn):
    fn_output = os.path.splitext(fn)[0]
    return fn_output


def extract_directory(fn):
    dir = os.path.dirname(fn)
    return dir


