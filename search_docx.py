import os
import glob
import docx

def flatten_list(l):
    return [item for sublist in l for item in sublist]


def get_docx_files():
    path = os.getcwd()
    docx_files = []
    for folder,_,_ in os.walk(path):
        docx_files.append((glob.glob(os.path.join(folder, "*.docx"))))
    return flatten_list(docx_files)

def search_through_document(str, doc, file_name):
    for paragraph_number in range(len((doc.paragraphs))):
        if str in doc.paragraphs[paragraph_number].text.lower():
            print(f'File: {file_name}; paragraph number {paragraph_number+1}')

def main():
    docx_files = get_docx_files()
    while True:
        input_string = input('What are you looking for: ').strip().lower()
        if input_string == 'kraj rada':
            break
        for docx_file in docx_files:
            file_name = docx_file.split("\\")[-2:]
            try:
                document = docx.Document(docx_file)
                search_through_document(input_string, document, file_name)
            except PermissionError:
                print(f'NO PERMISSINO TO OPEN {file_name}')
            except ValueError:
                print(f'FAILED TO OPEN {file_name}')
            except:
                print(f'UNKNOWN ERROR FOR {file_name}')
            


if __name__ == "__main__":
    main()