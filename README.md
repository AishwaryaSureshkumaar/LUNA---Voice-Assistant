# Voice Assistant - "LUNA"

## Overview

This project is a voice-activated assistant, LUNA, designed to streamline access to student documents using Python. 
LUNA can recognize voice commands and provide specific student information such as contact details, academic records, and other attributes directly from an Excel file.
It can also handle commands related to "Playing song,Google search,Setting Remainders and Alarms,Wikipedia search,Telling Jokes and Fun-facts and Telling current Date and Time.

## Installation

1. Clone the repository:
    ```
    git clone https://github.com/AishwaryaSureshkumaar/LUNA---Voice-Assistant.git  
    cd LUNA---Voice-Assistant  
    ```

2. Create a virtual environment and activate it:
    ```
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3. Install the dependencies:
    ```
    pip install -r requirements.txt
    ```

4. Run the application:
    ```
    python LUNA.py
    ```

## Usage

Starting the Assistant

-> Run the script 
-> Speak using the wake word "LUNA" followed by your command " Hi Luna" or "Hey Luna"

## Example Commands
### these are the example commands that are based on the data in the excel sheet which i provided.

1. Query by Student Name:
"LUNA, tell me the contact number for Aishwarya."
"LUNA, is Abirami a day scholar?"

2. Query by Attribute:

"LUNA, provide the PG percentage of Janani J."
"LUNA, what's the admission number of Sree?"

3. General Queries:

"LUNA, who is a hosteller among the students?"
"LUNA, list students with UG percentages above 80%."


## Features
- Voice recognition powered by Python libraries like speech_recognition.
- Natural language understanding for querying specific student details.
- Integration with Excel files for dynamic data retrieval.
- Supports custom commands for seamless access to student information.

## Troubleshooting

- Ensure your microphone is working properly.
- Verify the Excel file path and formatting match the expected structure.
- If you encounter errors, refer to the log file `LUNA.log` for more details.

