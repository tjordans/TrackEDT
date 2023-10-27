from bs4 import BeautifulSoup
import os
import re
from pathlib import Path
import shutil
import zipfile
import time
from tkinter import filedialog, Tk, messagebox
import csv
import pandas

# if a document is open while running the code, it creates issues. Make an error window pop up when this issue happens.

def select_directory():
    root = Tk()
    root.withdraw()  # Hide the root window
    directory = filedialog.askdirectory(title="Select a Folder")
    for file_name in os.listdir(directory):
        if file_name.startswith('~$'):
            root = Tk()
            root.withdraw()
            messagebox.showerror("Word doc in directory is open. Close it and try again. Exiting.")
            print("Close Word Document. Exiting.")
            root.mainloop()
            exit()
    return directory

#function for unzipping all files in the new coem folder
def unzip_files(directory):
    for file_name in os.listdir(directory):
        if file_name.endswith('.zip'):
            file_path = os.path.join(directory, file_name)
            print(file_path)
            # Create a subfolder with the original filename without the '.zip' extension
            subfolder_name = os.path.splitext(file_name)[0]
            subfolder_path = create_subfolder(directory, subfolder_name)
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(subfolder_path)
            os.remove(file_path)

#function for creation a subfolder (will be coem subfolder within the directory the user selects)
def create_subfolder(parent_dir, subfolder_name):
    subfolder_path = os.path.join(parent_dir, subfolder_name)
    if not os.path.exists(subfolder_path):
        os.makedirs(subfolder_path)
    return subfolder_path

#function for copying all .docx in the directory the user selects into a new coem folder and then replacing the extensions in the copied files with .zip
def copy_and_rename_files(source_dir, dest_dir):
    name_list = []
    for file_name in os.listdir(source_dir):
        if file_name.endswith('.docx'):
            # print(file_name)
            src_path = os.path.join(source_dir, file_name)
            # print(src_path)
            dest_path = os.path.join(dest_dir, file_name)
            # print(dest_path)
            shutil.copy2(src_path, dest_path)
            new_name = dest_path.replace('.docx', '.zip')
            print(new_name)
            os.rename(dest_path, new_name)
            name_list.append(new_name)
            if not os.path.exists(new_name):
                shutil.copy2(src_path, dest_path)
                os.rename(dest_path, new_name)
                name_list.append(new_name)
                print(name_list)
    return name_list

def split_into_sentences(text):
    # sentence-splitting solution credit: https://stackoverflow.com/questions/4576077/how-can-i-split-a-text-into-sentences
    alphabets = "([A-Za-z])"
    prefixes = "(Mr|St|Mrs|Ms|Dr|Prof|Capt|Cpt|Lt|Mt|et al|p|para)[.]"
    suffixes = "(Inc|Ltd|Jr|Sr|Co)"
    starters = "(Mr|Mrs|Ms|Dr|Prof|Capt|Cpt|Lt|He\s|She\s|It\s|They\s|Their\s|Our\s|We\s|But\s|However\s|That\s|This\s|Wherever)"
    acronyms = "([A-Z][.][A-Z][.](?:[A-Z][.])?)"
    websites = "[.](com|net|org|io|gov|me|edu)"
    digits = "([0-9])"

    xml_sentEndings = "()"

    text = " " + text + "  "
    text = text.replace("\n"," ")
    text = re.sub(prefixes,"\\1<prd>",text)
    text = re.sub(websites,"<prd>\\1",text)
    text = re.sub(digits + "[.]" + digits,"\\1<prd>\\2",text)
    if "..." in text: text = text.replace("...","<prd><prd><prd>")
    if "Ph.D" in text: text = text.replace("Ph.D.","Ph<prd>D<prd>")
    if "..." in text: text = text.replace("...", "<prd><prd><prd>")
    text = re.sub("\s" + alphabets + "[.] "," \\1<prd> ",text)
    text = re.sub(acronyms+" "+starters,"\\1<stop> \\2",text)
    text = re.sub(alphabets + "[.]" + alphabets + "[.]" + alphabets + "[.]","\\1<prd>\\2<prd>\\3<prd>",text)
    text = re.sub(alphabets + "[.]" + alphabets + "[.]","\\1<prd>\\2<prd>",text)
    text = re.sub(" "+suffixes+"[.] "+starters," \\1<stop> \\2",text)
    text = re.sub(" "+suffixes+"[.]"," \\1<prd>",text)
    text = re.sub(" " + alphabets + "[.]"," \\1<prd>",text)
    if "”" in text: text = text.replace(".”","”.")
    if "\"" in text: text = text.replace(".\"","\".")
    if "!" in text: text = text.replace("!\"","\"!")
    if "?" in text: text = text.replace("?\"","\"?")
    text = text.replace(".",".<stop>")
    text = text.replace(".<stop></moveFrom>", ".</moveFrom><stop>") #added by Jordan
    text = text.replace(".<stop></moveTo>", ".</moveTo><stop>")  # added by Jordan
    text = text.replace("?","?<stop>")
    text = text.replace("!","!<stop>")
    text = text.replace("<prd>",".")
    sentences = text.split("<stop>")
    sentences = sentences[:-1]
    sentences = [s.strip() for s in sentences]
    return sentences #This function

def read_file(path):
    with open(path, "r") as f:
        data = f.read()
    return data

def parse_data(data):
    soup = BeautifulSoup(data, "xml")
    found_tags = soup.find_all(['w:ins', 'w:del', 'w:moveTo', 'w:moveFrom', 'w:rPrChange'])
    for t in found_tags:
        if not t.is_empty_element:
            t.replace_with(BeautifulSoup(f'<{t.name}>&lt;{t.name}&gt;{t.text.strip()}&lt;/{t.name}&gt;</{t.name}>', 'html.parser'))
    extractedText = (soup.get_text(strip=True, separator=' '))
    sent_list = split_into_sentences(extractedText)
    cleanSentences = [i for i in sent_list]
    author = [t['w:author'] for t in found_tags if 'w:author' in t.attrs]
    return cleanSentences, author

def write_header(outpath):
    header = "NAME FILE" + "\t" + "EDIT TYPE" + "\t" + "MATCH" + "\t" + "SENTENCE" + "\t" + "AUTHOR"+ "\t" + "LENGTH" + "\n"
    with open(os.path.join(outpath, "_results.txt"), "a") as f:
        f.write(header)

def write_rows(outpath, src_path, cleanSentences, authors, listItem):
    for sentence in cleanSentences:
        print(sentence)
        tags = re.findall(r'<(del|ins|moveTo|moveFrom|rPrChange)>(.*?)</\1>', sentence)
        for tag, author in zip(tags, authors):
            theTag = tag[1]
            # print("Inside the loop")
            with open(os.path.join(outpath, "_results.txt"), "a") as f:
                p = Path(src_path[listItem])
                f.write(p.parts[-1] + "\t")
            #these if-statements work, except that it also goes through the "else" statement and replaces it with that 'row'
            if tag[0] == 'rPrChange':
                row = "format" + "\t" + f'{tag[1]}' + "\t" + sentence + "\t" + author + "\t" + str(len(tag[1])) + "\n"
            else:
                print(theTag)
                row = tag[0] + "\t" + f'{tag[1]}' + "\t" + sentence + "\t" + author + "\t" + str(len(tag[1])) + "\n"
                print(row)
            print(row)
            with open(os.path.join(outpath, "_results.txt"), "a") as f:
                print(row)
                f.write(row)

# to convert the text file into a csv file with everything delimited correctly

def txt_to_csv(selected_folder):
    with open(os.path.join(selected_folder, "_results.txt"), "r") as f:    
        dataframe = pandas.read_csv(f, delimiter="\t")
    with open(os.path.join(selected_folder, "_results.csv"), "a") as f:
        dataframe.to_csv(f, encoding='utf-8', index=False)


#function for running previous functions for unzipping files
def main():
    selected_folder = select_directory()
    if not selected_folder:
        print("No folder selected. Exiting.")
        return

    xml_folder = create_subfolder(selected_folder, 'XML_folder')
    src_path = copy_and_rename_files(selected_folder, xml_folder)
    print(src_path)
    src_pathType = type(src_path)
    print(src_pathType)
    # print(src_path)
    unzip_files(xml_folder)
    write_header(selected_folder)
    listItem = 0

    # check out what this for loop is doing to debug it
    for folder in os.listdir(xml_folder):
        working_folder = os.path.join(xml_folder, folder, "word/document.xml")
        # print(working_folder)
        data = read_file(working_folder)
        # print(data)
        cleanSentences, author = parse_data(data)
        # print(cleanSentences, author)
        write_rows(selected_folder, src_path, cleanSentences, author, listItem)
        listItem = listItem + 1

    txt_to_csv(selected_folder)

    print(f"Processed files in '{xml_folder}'.")

if __name__ == "__main__":
    main()
