import sys
import speech_recognition as sr
import win32com.client
import os
import webbrowser
import datetime
import subprocess
import time
import psutil
import queue
from collections import deque
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Voice = speaker.GetVoices().Item(1)

def takeCmnd():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-US")
            print(f"User Said: {query}")
            return query
        except Exception as ex:
            return "Error Occurred.Sorry."

def say(text):
    print("Speak.")
    speaker.Speak(text)
def detect_and_open_file(query):
    extension_mapping = {
        "txt": ".txt",
        "text": ".txt",
        "docx": ".docx",
        "document": ".docx",
        "pdf": ".pdf",
        "PDF": ".pdf",
        "PPT": ".ppt",
        "ppt": ".ppt",
        "PPTX": ".pptx",
        "pptx": ".pptx"
    }

    query_lower = query.lower()
    base_filename = query_lower
    extension = None

    
    for keyword, ext in extension_mapping.items():
        if keyword in query_lower:
            base_filename = query_lower.replace(keyword, "").strip()
            extension = ext
            break  

    if base_filename:
        matching_files = []

        for root, dirs, files in os.walk(os.path.expanduser('~/Documents')):
            for file in files:
                if extension and file.lower().startswith(base_filename.lower()) and file.lower().endswith(extension.lower()):
                    file_path = os.path.join(root, file)
                    process = subprocess.Popen(["start", "", file_path], shell=True)
                    matching_files.append(file_path)
                    print(f"Opening {file} at {file_path}")
                    say(f"Opening {base_filename}.")
                    return matching_files

                elif not extension and file.lower().startswith(base_filename.lower()):
                    file_path = os.path.join(root, file)
                    process = subprocess.Popen(["start", "", file_path], shell=True)
                    matching_files.append(file_path)
                    print(f"Opening {file} at {file_path}")
                    say(f"Opening {base_filename}.")

        if matching_files:
            return matching_files
        else:
            if extension:
                print(f"No {extension} files with base filename {base_filename} found.")
            else:
                print(f"No files with base filename {base_filename} found.")
            return []

    return False

  




def createstatespace(directory):
    state_space = {}

    for root, dirs, files in os.walk(directory):
        node = state_space
        path_parts = os.path.relpath(root, directory).split(os.path.sep)

        for part in path_parts:
            node = node.setdefault(part, {})

        for file in files:
            node[file] = {}

    return state_space


def bfs(query, state_space, root_directory):
    query_lower = query.lower()
    matching_files = []

    queue = deque()
    queue.append((root_directory, state_space))

    while queue:
        current_directory, current_state = queue.popleft()

        for item in os.listdir(current_directory):
            item_path = os.path.join(current_directory, item)

            if os.path.isfile(item_path) and query_lower in item.lower():
                process = subprocess.Popen(["start", "", item_path], shell=True)
                matching_files.append(item_path)
                print(f"Opening {item} at {item_path}")

            elif os.path.isdir(item_path) and item in current_state:
                queue.append((item_path, current_state[item]))

    if matching_files:
        return matching_files
    else:
        print(f"No files matching '{query}' found.")
        return []






special_character_mapping = {
    "dot": ".",
    "comma": ",",
    "exclamation": "!",
    "number": "#",
    "dash": "-",
    " ": ""
}
command_phrases = [
    "open file",
    "file open",
    "please open file",
    "I want you to open",
    "I wish to work on",
    "open",
    
]

def replace_symbols(text):
    words = text.split()
    replaced_words = []

    for word in words:
        if word in special_character_mapping:
            replaced_word = special_character_mapping[word]
            replaced_words.append(replaced_word)
        else:
            replaced_words.append(word)

    return " ".join(replaced_words)

#todo: opening any website on the internet.

def open_website(website_name):
    url = f"https://www.{website_name}.com"
    webbrowser.open(url)
    print(f"Opening {website_name}: {url}")

#todo: main.

if __name__ == '__main__':
    directo = os.path.expanduser('~/Desktop') 
    ss = createstatespace(directo)
    bfs('fuck', ss, directo)
    print('Welcome.')
    say("Hello. I am Emma")
    while True:
        print("Listening...")
        query = takeCmnd()
        if "using BFS" in query:
            base_filename = query.replace("using BFS", "").strip()
            bfs(base_filename, ss, directo)
            print(f"Opening {base_filename} with breadth first search algorithm.")
            
        for command_phrase in command_phrases:
            if command_phrase in query:
                base_filename = query.replace(command_phrase, "").strip()
                if "symbol" in base_filename:
                       replaced_query = replace_symbols(base_filename)
                       base_filename = replaced_query.replace("symbol", "").strip().replace(" ", "")
                       detect_and_open_file(base_filename)
                else:
                       detect_and_open_file(base_filename)

            
        if "close file" in query.lower():
            base_filename = query.replace("close file", "").strip()
            print(f"Closing {base_filename}.")
            say(f"Closing {base_filename}.")
            os.system(f"taskkill /im {base_filename}.exe /f")
            base_filename = None

        # todo: opening any browsers
        elif "open website" in query.lower():
            website_name = query.replace("open website", "").strip()
            say(f"Opening {website_name}.")
            open_website(website_name)

        # todo: telling time
        elif "the time" in query:
            Time = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"The time is {Time}")

        #todo: stopping Emma

        elif "stop" in query:
            say("Good Bye.")
            sys.exit()

        # say(query)



