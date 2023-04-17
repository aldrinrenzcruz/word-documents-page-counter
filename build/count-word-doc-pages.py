import os
import win32com.client  # requires 'pywin32' package to be installed

folder_path = input("Enter the folder path containing the Word documents: ")

word_app = win32com.client.Dispatch("Word.Application")
word_app.Visible = False

total_pages = 0
page_counts = {}

for file_name in os.listdir(folder_path):
    if file_name.endswith(".doc") or file_name.endswith(".docx"):
        file_path = os.path.join(folder_path, file_name)
        doc = word_app.Documents.Open(file_path, ReadOnly=True)
        # 2 = wdStatisticPages (page count)
        num_pages = doc.ComputeStatistics(2)
        total_pages += num_pages
        page_counts[file_name] = num_pages
        doc.Close()

word_app.Quit()

# create a text file with the page count and breakdown in the same folder
report_file_path = os.path.join(folder_path, "_page_count_report.txt")
with open(report_file_path, "w") as f:
    for file_name, num_pages in page_counts.items():
        page_info = f"{file_name} [{num_pages} pages]\n"
        f.write(page_info)
        print(page_info, end='')
    f.write(f"\nTotal pages: {total_pages}\n")
print(f"\nTotal pages: {total_pages}")

# keep the console open
print(f"\nGenerating the report in the folder path provided...")
input("\nPress Enter to exit...")