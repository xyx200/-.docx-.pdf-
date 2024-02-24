# Import the modules
import os
import docx
import fitz # PyMuPDF

# Define a function to check if a character is Chinese
def is_chinese(char):
    # Check the Unicode range of the character
    return 0x4E00 <= ord(char) <= 0x9FFF

# Define a function to count the Chinese characters in a string
def count_chinese(string):
    # Initialize the counter
    count = 0
    # Loop through the string
    for char in string:
        # If the character is Chinese, increment the counter
        if is_chinese(char):
            count += 1
    # Return the counter
    return count

# Define a function to count the Chinese characters in a file
def count_chinese_in_file(file):
    # Initialize the counter
    count = 0
    # Get the file extension
    ext = os.path.splitext(file)[1]
    # If the file is a docx file
    if ext == ".docx":
        # Open the file
        doc = docx.Document(file)
        # Loop through the paragraphs
        for para in doc.paragraphs:
            # Count the Chinese characters in the paragraph text
            count += count_chinese(para.text)
    # If the file is a pdf file
    elif ext == ".pdf":
        # Open the file
        doc = fitz.open(file)
        # Loop through the pages
        for page in doc:
            # Get the page text
            text = page.get_text()
            # Count the Chinese characters in the page text
            count += count_chinese(text)
    # Return the counter
    return count

# Define a function to count the Chinese characters in a folder and its subfolders
def count_chinese_in_folder(folder):
    # Initialize the counter
    count = 0
    # Loop through the files and subfolders in the folder
    for root, dirs, files in os.walk(folder):
        # Loop through the files
        for file in files:
            # Get the full file path
            file_path = os.path.join(root, file)
            # Count the Chinese characters in the file
            count += count_chinese_in_file(file_path)
    # Return the counter
    return count

# Test the function
folder = "C:\\Users\\Administrator\\Desktop\\test" # Change this to your folder path
result = count_chinese_in_folder(folder)
print(f"There are {result} Chinese characters in the folder {folder} and its subfolders.")
