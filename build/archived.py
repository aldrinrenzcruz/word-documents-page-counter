import os
import win32com.client
wdStatisticPages = 2
# get the folder path from the user
folder_path = input("Enter the folder path containing the Word documents: ")

# create a new instance of Word
word = win32com.client.Dispatch("Word.Application")

# initialize the total page count
total_page_count = 0

# create a dictionary to store the document names and page counts
doc_dict = {}

# loop through all files in the folder
for filename in os.listdir(folder_path):
    # check if the file is a Word document
    if filename.endswith(('.doc', '.docx')):
        # open the document and get its page count
        doc = word.Documents.Open(os.path.join(folder_path, filename))
        # 2 corresponds to wdStatisticPages
        page_count = doc.ComputeStatistics(wdStatisticPages)
        print(f'{filename} has {page_count} pages')

        # add the document name and page count to the dictionary
        doc_dict[filename] = page_count

        # add the page count to the total
        total_page_count += page_count

        # close the document
        doc.Close()

# close the Word instance
word.Quit()

# output the document names and page counts
for doc, page_count in doc_dict.items():
    print(f'\n{doc} [{page_count} pages]')

# output the total page count
print(f'\nTotal pages: {total_page_count}\n\n')

# keep the console open
input("Press Enter to exit...")
