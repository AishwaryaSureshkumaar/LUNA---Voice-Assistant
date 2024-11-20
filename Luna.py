import speech_recognition as sr
import pyttsx3
import pywhatkit
import datetime
import wikipedia
import pyjokes
import math
import sys
import webbrowser
import requests
from bs4 import BeautifulSoup
from googlesearch import search
import pandas as pd
from openpyxl import load_workbook
from fuzzywuzzy import fuzz, process
import os
from termcolor import colored

# Initialize the speech recognition and text-to-speech engine
listener = sr.Recognizer()
engine = pyttsx3.init()
voices = engine.getProperty('voices')
if len(voices) > 1:
    engine.setProperty('voice', voices[1].id)
else:
    engine.setProperty('voice', voices[0].id)

# Helper function to speak and print output
def engine_ttalk(text):
    print(colored(f"Luna is saying: {text}", "blue"))
    engine.say(text)
    engine.runAndWait()

# Capture user commands
def user_commands():
    try:
        with sr.Microphone() as source:
            listener.adjust_for_ambient_noise(source)
            print(colored("Start Speaking !!", "green"))
            voice = listener.listen(source)
            command = listener.recognize_google(voice).lower()
            if 'Luna' in command:
                command = command.replace('Luna', '').strip()
                print(colored(f"User Said: {command}", "green"))
                return command
    except Exception as e:
        print(f"Error: {e}")
        engine_ttalk("I did not catch that. Please speak again.")
    return ""

# Perform a Google search
def google_search(query):
    try:
        search_query = ' '.join(query.split())
        url = f"https://www.google.com/search?q={search_query}"
        webbrowser.open_new_tab(url)
        engine_ttalk(f"Here are the search results for {search_query}.")
    except Exception as e:
        print(f"Error performing Google search: {e}")
        engine_ttalk("Sorry, I encountered an error while performing the search.")

# Fuzzy matching function for student names
def find_closest_match(input_name, name_list):
    match, score = process.extractOne(input_name, name_list)
    return match, score

# Read Excel file and retrieve student details
def read_excel(file_path, query):
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found at path: {file_path}")

        df = pd.read_excel(file_path, sheet_name='Placements')
        df['Normalized Name'] = df['Name'].str.lower()
        student_name = query.replace("details", "").strip().lower()
        
        column_mappings = {
            "register number": "Register Number", "aadhar number": "Aadhar Number",
            "date of birth": "Date of Birth", "email id": "Email ID", "address": "Address",
            "hsc percentage": "HSC Percentage", "hsc institution": "HSC Institution",
            "ssc percentage": "SSLC Percentage", "ssc institution": "SSLC Institution",
            "ug percentage": "UG Percentage", "pg percentage": "PG Percentage",
            "department": "Department", "contact number": "Contact Number",
            "personal mail id": "Personal Mail ID", "day scholar": "DAY Scholar",
            "hosteller": "Hosteller", "hostel name": "Hostel Name",
            "father name": "Father Name", "mother name": "Mother Name"
        }
        
        names_list = df['Normalized Name'].tolist()
        closest_match, score = find_closest_match(student_name, names_list)
        
        if score > 80:
            matched_row = df[df['Normalized Name'] == closest_match].iloc[0]
            details = []
            
            if 'details' in query:
                for attribute, col_name in column_mappings.items():
                    if col_name in matched_row:
                        details.append(f"{col_name}: {matched_row[col_name]}")
                full_details = '\n'.join(details)
                engine_ttalk(f"Here are the full details for {matched_row['Name']}")
                print(colored(full_details, "blue"))
                for detail in details:
                    engine_ttalk(detail)
            else:
                for attribute, col_name in column_mappings.items():
                    if attribute in query and col_name in matched_row:
                        engine_ttalk(f"{col_name}: {matched_row[col_name]}")
                        return
        else:
            engine_ttalk("No close match found for the student name.")
    except Exception as e:
        engine_ttalk("Sorry, I encountered an error reading the Excel sheet.")
        print(f"Error reading Excel: {e}")

# Wikipedia summary fetch
def get_wikipedia_summary(name):
    try:
        summary = wikipedia.summary(name, sentences=2)
        return summary
    except wikipedia.exceptions.DisambiguationError as e:
        return f"Multiple results for '{name}': {e.options}"
    except wikipedia.exceptions.PageError:
        return f"No Wikipedia page for '{name}'."
    except Exception as e:
        return f"An error occurred: {e}"

# Run the voice assistant
def run_Luna():
    command = user_commands()
    if command:
        if any(greet in command for greet in ['hi', 'hello', 'hey']):
            engine_ttalk("Hiii..! How can I help you today?")
        elif 'play' in command:
            song = command.replace('play', '').strip()
            engine_ttalk('Playing ' + song)
            pywhatkit.playonyt(song)
        elif 'time' in command:
            time = datetime.datetime.now().strftime('%I:%M %p')
            engine_ttalk('The current time is ' + time)
        elif 'who is' in command:
            name = command.replace('who is', '').strip()
            info = get_wikipedia_summary(name)
            engine_ttalk(info)
            print(colored(info, "blue"))
        elif 'search' in command:
            query = command.replace('search', '').strip()
            google_search(query)
        elif 'placement excel' in command:
            query = command.replace('placement excel', '').strip()
            file_path = r'C:/Project Luna/placements.xlsx'
            read_excel(file_path, query)
        elif 'joke' in command:
            engine_ttalk(pyjokes.get_joke())
        else:
            engine_ttalk("Sorry, I don't know that command.")

if __name__ == "__main__":
    while True:
        run_Luna()
