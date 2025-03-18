import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os
import re
import requests
import json
import colorama
from colorama import Fore, Style
import concurrent.futures
from urllib.parse import urlparse, parse_qs
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from subprocess import CREATE_NO_WINDOW
import subprocess
from tkinter.scrolledtext import ScrolledText
import sys
import webbrowser
from tkinter import scrolledtext
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import fitz  # PyMuPDF for PDFs
import docx
import pptx
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Menu
import subprocess
import platform



colorama.init()



RESUME_FOLDER = "Resume_Download"

global AuthorizationToken_Created
AuthorizationToken_Created = None

class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        self.text_widget.insert(tk.END, text, ("color",))
        self.text_widget.see(tk.END)  # Scroll to the end of the text
    def flush(self):
        # This method is required to satisfy the sys.stdout interface.
        # However, it is not needed for this specific implementation.
        pass

#def ModuleInstall():
    #print(Fore.GREEN +"Install selenium, requests, colorama, sphinx, futures"+Style.RESET_ALL)
    #print(Fore.GREEN +"Run Command on CMD/Terminal: "+Style.RESET_ALL+"pip3 install selenium requests colorama sphinx futures")

def send_get_request(url, headers):
    response = requests.get(url, headers=headers, verify=False)
    return response

def read_urls_from_file(file_path):
    with open(file_path, "r") as file:
        return [line.strip() for line in file]

def read_urls_from_file1(file_path):
    with open(file_path, "r") as file:
        return [line.strip() for line in file]
        file.close()
    
def downloadfile(url, headers):
    response = requests.get(url, headers=headers)
    return response

def login_to_microsoft():
    #ModuleInstall()
    CreateFolder()
    global AuthorizationToken_Created
    regex = r'(\"secret\":\")(.*?)\"'

    #chromedriver_path = '.\\Driver\\chromedriver.exe'
    #os.environ['PATH'] += os.pathsep + chromedriver_path
    # Get the current directory where your Python script is located
    current_directory = os.path.dirname(os.path.abspath(__file__))
    # Set the binary_location to the path of chrome.exe in the same directory
    chrome_executable_path = os.path.join(current_directory,'chrome-win', 'chrome.exe')
    chromedriver_executable_path = os.path.join(current_directory,'chrome', 'chromedriver.exe')
    os.environ['PATH'] += os.pathsep + chromedriver_executable_path
    
    
    #chrome_options = Options()
    #chrome_options.binary_location = chrome_executable_path
    #print(chrome_executable_path)
    #print(chromedriver_executable_path)
    #driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=chrome_options)

    chrome_options = webdriver.ChromeOptions()
    chrome_options.binary_location = chrome_executable_path
    driver = webdriver.Chrome(options=chrome_options)

    
    

    try:
        driver.get('https://people.deloitte/profile/')  # Open the login page
        #time.sleep(5)

        #input(Fore.GREEN + "Please enter your details in the browser window. Press Enter after entering OTP."+Style.RESET_ALL)
        #input((Fore.GREEN + "Please enter your details in the browser window. Press Enter after entering OTP."+Style.RESET_ALL))
        #time.sleep(2)  # Give some time for the page to load and complete the login process
        #root.protocol("WM_DELETE_WINDOW", lambda: var.set(1))
        
        #button.wait_variable(var)
        #print("Login successful!")
        while AuthorizationToken_Created is None:
            session_storage_data = driver.execute_script("""var sessionStorageData = {};for (var i = 0; i < sessionStorage.length; i++) {var key = sessionStorage.key(i);var value = sessionStorage.getItem(key);sessionStorageData[key] = value;}return sessionStorageData;""")
            session_storage_data = str(session_storage_data)
            # Print all session storage data
            match = re.search(regex, session_storage_data, re.MULTILINE)
            if match:
                #first_match = match.group()
                AuthorizationTOken = match.group(2)
                AuthorizationToken_Created =  AuthorizationTOken
                print(AuthorizationToken_Created)
                break
                #driver.quit()
            time.sleep(2)
    except Exception as e:
        print("An error occurred:", e)

    finally:
        if driver:
            driver.quit()



def Download_Resume():
    global AuthorizationToken_Created
    if AuthorizationToken_Created:
        print("Login Successful..")
        print("AuthorizationToken: ",AuthorizationToken_Created)
        authorization_token = AuthorizationToken_Created
        headers = {"Authorization": f"Bearer {authorization_token}"}
        Resume_URL = "https://apim.people.deloitte/personresumes?email="
        
        print(Fore.GREEN+f"Initiate Downloading...."+Style.RESET_ALL)
        #authorization_token = input("Please enter your authorization token: ")
        authorization_token = AuthorizationToken_Created
        headers = {"Authorization": f"Bearer {authorization_token}"}
        Resume_URL = "https://apim.people.deloitte/personresumes?email="

        print("Reading Emails From EmailId.txt")
        Email_ID = open(email_id_file, 'r')
        Email_ID1 = Email_ID.readlines()
        #print(Email_ID1)
        Email_URL_Write = open('Email_URL_Write.txt', 'a')
        Email_URL_read = open('Email_URL_Write.txt', 'r')
        existing_emails = set(Email_URL_read.read().splitlines())
        Email_URL_read.close()
        for Email in Email_ID1:
            Email = Email.strip()
            Resume_URL_with_email = str(Resume_URL+Email)
            #print(Resume_URL_with_email)
            #print(Resume_URL_with_email)
            if Resume_URL_with_email not in existing_emails:
                Email_URL_Write.write(Resume_URL_with_email)
                Email_URL_Write.write("\n")
                existing_emails.add(Resume_URL_with_email)
            #else:
                #print("Email Exist in the file..")
            
        Email_ID.close()    
        Email_URL_Write.close()
    
        urls_file = "Email_URL_Write.txt"
        max_workers = 20
        urls = read_urls_from_file(urls_file)
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_url = {executor.submit(send_get_request, url, headers): url for url in urls}
            for future in concurrent.futures.as_completed(future_to_url):
                url = future_to_url[future]
                try:
                    response = future.result()
                    # You can handle the response here (e.g., save to a file, process data, etc.)
                    #print(f"Response from {url}: {response.status_code}")
                    ResumeDownloadURL = open('ResumeDownloadURL.txt', 'a')
                    json_Resume_data = (response.json())
                    for item in json_Resume_data["data"]:
                        fileName = item["fileName"]
                        creationDate = item["creationDate"]
                        documentUrl = item["documentUrl"]
                        #print("Resume_Name:", fileName)
                        #print("creationDate:", creationDate)
                        #print("documentUrl:", documentUrl)
                        documentUrl = documentUrl.strip()
                        resume_documentUrl = str(documentUrl)
                        #print(documentUrl)
                        ResumeDownloadURL.write(resume_documentUrl)
                        ResumeDownloadURL.write("\n")
                        #print("\n")
                    ResumeDownloadURL.close()
                except Exception as e:
                    print(f"Error accessing {url}: {e}")

        
        ResumeDownloadurls_file = "ResumeDownloadURL.txt"
        ResumeDownloadurls_urls = read_urls_from_file1(ResumeDownloadurls_file)
    
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_url1 = {executor.submit(downloadfile, url, headers): url for url in ResumeDownloadurls_urls}
            for future in concurrent.futures.as_completed(future_to_url1):
                url = future_to_url1[future]
                try:
                    response = future.result()
                    #print(f"Response from {url}: {response.status_code}")
                    headerfilename = (response.headers.get("content-disposition"))
                    downloadresponse = response.content
                    filename_pattern = r'filename="([^"]+)"'
                    match = re.search(filename_pattern, headerfilename)
                    if match:
                        filename = match.group(1)
                        print(filename)
                    
                        save_path = os.path.join("Resume_Download", filename)
                        #print(save_path)
                        with open(save_path, 'wb') as f:
                            #print(save_path)
                            f.write(response.content)
                            print("PDF file downloaded successfully and saved to Resume_Download.")
    
                except Exception as e:
                    print(f"Error accessing {url}: {e}")


    else:
        print("Login Failed..")

    print("Download complete..")
    os.remove("ResumeDownloadURL.txt")
    os.remove("Email_URL_Write.txt")
    
    #input(Fore.GREEN + "Press enter to close the terminal:"+Style.RESET_ALL)
    
 
def CreateFolder():
    path = ('Resume_Download')
    if not os.path.exists(path):
        os.mkdir(path)
        print("Folder %s created!" % path)
    #else:
        #print("Folder %s already exists" % path)
    

def browse_file():
    global email_id_file
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    email_id_entry.delete(0, tk.END)
    email_id_entry.insert(tk.END, file_path)
    email_id_file = email_id_entry.get()

def on_closing():
    #os.remove("ResumeDownloadURL.txt")
    #os.remove("Email_URL_Write.txt")
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.destroy()


def extract_text_from_pdf(filepath):
    text = ""
    try:
        doc = fitz.open(filepath)
        for page in doc:
            text += page.get_text("text") + "\n"
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
    return text

def extract_text_from_docx(filepath):
    try:
        doc = docx.Document(filepath)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return ""

def extract_text_from_pptx(filepath):
    try:
        presentation = pptx.Presentation(filepath)
        text = []
        for slide in presentation.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text.append(shape.text)
        return "\n".join(text)
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return ""

def extract_text_from_excel(filepath):
    try:
        df = pd.read_excel(filepath, sheet_name=None)
        text = "\n".join([df[sheet].to_string() for sheet in df])
        return text
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return ""

def extract_text(filepath):
    ext = filepath.split(".")[-1].lower()
    if ext == "pdf":
        return extract_text_from_pdf(filepath)
    elif ext == "docx":
        return extract_text_from_docx(filepath)
    elif ext == "pptx":
        return extract_text_from_pptx(filepath)
    elif ext in ["xls", "xlsx"]:
        return extract_text_from_excel(filepath)
    else:
        print(f"Unsupported format: {filepath}")
        return ""

def boolean_search(text, query):
    query = format_search_query(query)  # ✅ Ensure it's formatted correctly
    query = query.replace("AND", "&").replace("OR", "|").replace("NOT", "~")

    words = re.findall(r'"[^"]+"|\w+', query)  # Extract quoted phrases and single words

    for word in words:
        word_cleaned = word.strip('"')  # Remove surrounding quotes for actual matching
        if word_cleaned.lower() not in text.lower():
            query = query.replace(word, "False")
        else:
            query = query.replace(word, "True")

    try:
        return eval(query)  # Evaluate boolean expression safely
    except:
        return False

def on_select(event):
    try:
        result_text.tag_remove("highlight", "1.0", tk.END)  # Remove old selection
        index = result_text.index(tk.CURRENT)
        line_start = f"{index.split('.')[0]}.0"
        line_end = f"{index.split('.')[0]}.end"

        result_text.tag_add("highlight", line_start, line_end)  # Add highlight tag
        result_text.tag_config("highlight", background="lightblue")  # Blue highlight
    except Exception as e:
        print(f"Error selecting file: {e}")

def open_file(event):
    try:
        # Find the selected line based on highlight
        index = result_text.index(tk.CURRENT)
        line_start = f"{index.split('.')[0]}.0"
        line_end = f"{index.split('.')[0]}.end"
        selected_text = result_text.get(line_start, line_end).strip()

        if selected_text:
            filepath = os.path.join(RESUME_FOLDER, selected_text)
            if os.path.exists(filepath):
                if platform.system() == "Windows":
                    os.startfile(filepath)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", filepath])
                else:  # Linux
                    subprocess.run(["xdg-open", filepath])
    except Exception as e:
        print(f"Error opening file: {e}")

def format_search_query(query):
    words = query.split()  # Split query into words
    operators = {"AND", "OR", "NOT"}  # Logical operators
    formatted_query = []
    temp_phrase = []

    for word in words:
        if word.upper() in operators:  # If it's an operator, keep it
            if temp_phrase:
                formatted_query.append(f'"{" ".join(temp_phrase)}"')  # Wrap phrase in quotes
                temp_phrase = []
            formatted_query.append(word)  # Add operator
        else:
            temp_phrase.append(word)  # Add word to phrase

    if temp_phrase:
        formatted_query.append(f'"{" ".join(temp_phrase)}"')  # Wrap last phrase in quotes

    final_query = " ".join(formatted_query)

    # ✅ If there's NO logical operator, wrap the **entire query** in quotes
    if not any(op in final_query for op in operators) and " " in final_query:
        final_query = f'"{final_query}"'

    return final_query

def search_resumes(result_text, query):
    formatted_query = format_search_query(query)  # ✅ Auto-format query
    print("Formatted Query:", formatted_query)  # Debugging

    if not query.strip():
        messagebox.showerror("Error", "Please enter a search query.")
        return
    
    matching_files = []
    for file in os.listdir(RESUME_FOLDER):
        filepath = os.path.join(RESUME_FOLDER, file)
        if os.path.isfile(filepath):
            text = extract_text(filepath)
            if boolean_search(text, formatted_query):  # ✅ Use formatted query
                matching_files.append(file)
    
    result_text.delete("1.0", tk.END)
    if matching_files:
        result_text.insert(tk.END, "\n".join(matching_files))
    else:
        result_text.insert(tk.END, "No matching resumes found.")

def browse_folder():
    global RESUME_FOLDER, folder_label
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        RESUME_FOLDER = folder_selected
        folder_label.config(text=f"Folder: {RESUME_FOLDER}")  # Updates folder label



def search_resumes(result_text, query):
    if not query:
        messagebox.showerror("Error", "Please enter a search query.")
        return
    
    matching_files = []
    for file in os.listdir(RESUME_FOLDER):
        filepath = os.path.join(RESUME_FOLDER, file)
        if os.path.isfile(filepath):
            text = extract_text(filepath)
            if boolean_search(text, query):
                matching_files.append(file)
    
    # Display results in the new window
    result_text.delete("1.0", tk.END)
    if matching_files:
        result_text.insert(tk.END, "\n".join(matching_files))
    else:
        result_text.insert(tk.END, "No matching resumes found.")

def append_operator(op, search_entry):
    """Insert a logical operator (AND, OR, NOT) into the search entry field."""
    text = search_entry.get().strip()

    if text and not text.endswith(("AND", "OR", "NOT")):  # Avoid duplicate operators
        search_entry.insert(tk.END, f" {op} ")
    elif not text:  # Prevent starting with an operator
        messagebox.showerror("Error", "Please enter a search keyword before adding an operator.")

def on_hover(event):
    try:
        result_text.tag_remove("hover", "1.0", tk.END)  # Remove previous highlights
        index = result_text.index(f"@{event.x},{event.y}")  # Get cursor position
        line_start = f"{index.split('.')[0]}.0"
        line_end = f"{index.split('.')[0]}.end"

        result_text.tag_add("hover", line_start, line_end)  # Apply hover highlight
        result_text.tag_config("hover", background="lightblue")  # Light blue hover effect
    except Exception as e:
        print(f"Error in hover effect: {e}")
        
def Search_Resume():
    search_window = tk.Toplevel(root)
    search_window.title("Search String")
    search_window.geometry("500x400")
    search_window.configure(bg="#ddc8cf")

    try:
        search_window.iconbitmap(r".\icon1.ico")
    except tk.TclError:
        print("Icon file 'icon1.ico' not found or not a valid icon file.")

    # Search Entry Field
    tk.Label(search_window, text="Enter search keywords:", bg="#ddc8cf").pack(pady=5)
    search_entry = tk.Entry(search_window, width=50)
    search_entry.pack(pady=5)

    # Operator Buttons (AND, OR, NOT)
    button_frame = tk.Frame(search_window, bg="#ddc8cf")
    button_frame.pack(pady=5)

    tk.Button(button_frame, text="AND", command=lambda: append_operator("AND", search_entry)).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="OR", command=lambda: append_operator("OR", search_entry)).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="NOT", command=lambda: append_operator("NOT", search_entry)).pack(side=tk.LEFT, padx=5)

    # Search Button
    tk.Button(search_window, text="Search", command=lambda: search_resumes(result_text, search_entry.get())).pack(pady=5)

    # Folder Display
    global folder_label  
    folder_label = tk.Label(search_window, text=f"Folder: {RESUME_FOLDER}", bg="#ddc8cf")
    folder_label.pack(pady=5)

    # Change Folder Button
    tk.Button(search_window, text="Change Folder", command=browse_folder).pack(pady=5)

    # Result Text Area
    global result_text  
    result_text = scrolledtext.ScrolledText(search_window, height=10, width=60)
    result_text.pack(pady=5)

    # Bind mouse motion for hover effect
    result_text.bind("<Motion>", on_hover)  # ✅ Highlights the file under the cursor
    result_text.bind("<Double-Button-1>", open_file)  # ✅ Double-click to open file




# Create the GUI
root = tk.Tk()
#root.overrideredirect(True)
root.title("Deloitte Resume Downloader v1.0")


root.geometry("400x500")
root.configure(bg="#ddc8cf")

def open_email():
    webbrowser.open("mailto:rishabhsharma96@deloitte.com")

try:
    root.iconbitmap(r".\icon.ico")
except tk.TclError:
    print(f"Icon file '{icon_file_path}' not found or not a valid icon file.")


label = tk.Label(root, text="Select Email txt File:",bg="#ddc8cf")
label.pack(pady=10)

email_id_entry = tk.Entry(root, width=40)
email_id_entry.pack(pady=5)

label = tk.Label(root, text="Follow steps one by one",bg="#ddc8cf")
label.pack(pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.pack(pady=5)

sign_in_button = tk.Button(root, text="Sign In", command=login_to_microsoft)
sign_in_button.pack(pady=5)


#var = tk.IntVar()
#button = tk.Button(root, text="Login Done", command=lambda: var.set(1))
#button.pack(pady=5)

download_button = tk.Button(root, text="Download", command=Download_Resume)
download_button.pack(pady=5)

search_button = tk.Button(root, text="Search", command=Search_Resume)
search_button.pack(pady=5)

close_button = tk.Button(root, text="Close", command=on_closing)
close_button.pack(pady=5)

#label = tk.Label(root, text="Install package:\npip3 install selenium requests colorama sphinx futures")
#label.pack(pady=5)

label = tk.Label(root, text="Created By Rishabh Sharma",fg="red",bg="#ddc8cf")
label.pack(pady=2)


label = tk.Label(root, text="For any feedback: rishabhsharma96@deloitte.com",fg="red", cursor="hand2",font=("Arial",10,"underline"),bg="#ddc8cf")
label.pack(pady=2)
label.bind("<Button-1>", lambda e: open_email())

output_text = tk.Text(root, wrap="word", height=15)
output_text.pack(fill=tk.BOTH, expand=True)
output_text.tag_configure("color", foreground="green")
sys.stdout = TextRedirector(output_text)

root.mainloop()
