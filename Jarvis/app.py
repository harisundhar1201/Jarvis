from flask import Flask, render_template, jsonify
import speech_recognition as sr
import pyttsx3
import pyaudio
from datetime import datetime
import webbrowser
import comtypes.client
import pythoncom
import os
import pygame
from urllib.parse import quote
from googletrans import Translator


app = Flask(__name__)

def speak(text):
    pythoncom.CoInitialize()  # Initialize COM library
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()
    pythoncom.CoUninitialize()  # Uninitialize COM library

def wish_me():
    hour = datetime.now().hour
    if 0 <= hour < 12:
        return "Good Morning Boss..."
    elif 12 <= hour < 17:
        return "Good Afternoon Master..."
    else:
        return "Good Evening Sir..."

@app.route('/')
def index():
    speak("Initializing JARVIS...")
    greeting = wish_me()
    speak(greeting)
    return render_template('index.html')

@app.route('/listen')
def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        try:
            # Adjust the timeout as needed
            audio = recognizer.listen(source, timeout=10)  # Increase timeout to 10 seconds
            print("Processing...")
            message = recognizer.recognize_google(audio, language='en-in').lower()
            response = execute_command(message)
            speak(response)  # Add this line to provide voice output
        except sr.UnknownValueError:
            response = "Sorry, I did not understand that."
            speak(response)
        except sr.RequestError:
            response = "Sorry, I am unable to process your request right now."
            speak(response)
        except sr.WaitTimeoutError:
            response = "Sorry, listening timed out. Please try again."
            speak(response)
    
    return jsonify({'message': response})

def translate_text(text, dest_language='en'):
    translator = Translator()
    translated = translator.translate(text, dest=dest_language)
    return translated.text

def open_application(app_name):
    applications = {
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "paint": "mspaint.exe",
        "word": "winword.exe",
        "excel": "excel.exe",
        "powerpoint": "powerpnt.exe",
        "outlook": "outlook.exe",
        "chrome": "chrome.exe",
        "firefox": "firefox.exe",
        "cmd": "cmd.exe",
        "explorer": "explorer.exe",
        "vlc": "vlc.exe",
        "spotify": "spotify.exe",
        "telegram": "telegram.exe",
        "whatsapp": "whatsapp.exe",
        "instagram": "instagram.exe",
        "file manager": "explorer.exe", 
        "microsoft edge": "msedge.exe",
        "teams": "teams.exe",
        "visual studio": "devenv.exe"
        # Add other applications as needed
    }
    app_path = applications.get(app_name.lower())
    if app_path:
        os.system(f"start {app_path}")
        return f"Opening {app_name}."
    else:
        return "Application not found."
    
def close_application(app_name):
    applications = {
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "paint": "mspaint.exe",
        "word": "winword.exe",
        "excel": "excel.exe",
        "powerpoint": "powerpnt.exe",
        "outlook": "outlook.exe",
        "chrome": "chrome.exe",
        "firefox": "firefox.exe",
        "cmd": "cmd.exe",
        "explorer": "explorer.exe",
        "vlc": "vlc.exe",
        "spotify": "spotify.exe",
        "telegram": "telegram.exe",
        "whatsapp": "whatsapp.exe",
        "instagram": "instagram.exe",
        "file manager": "explorer.exe", 
        "microsoft edge": "msedge.exe",
        "teams": "teams.exe",
        "visual studio": "devenv.exe"
        # Add other applications as needed
    }
    app_path = applications.get(app_name.lower())
    if app_path:
        os.system(f"taskkill /f /im {app_path}")
        return f"Closing {app_name}."
    else:
        return "Application not found."
def execute_command(message):
    if 'hey' in message or 'hello' in message:
        return "Hello Sir, How May I Help You?"
    elif 'play music' in message:
        song_name = message.replace("play music", "").strip()
        query = quote(song_name)  # Encode the song name for URL
        url = f"https://open.spotify.com/search/{query}"
        webbrowser.open(url)
        return f"Opening Spotify to search for {song_name}..."
    elif 'open google' in message:
        webbrowser.open("https://google.com")
        return "Opening Google..."
    elif 'open youtube' in message:
        webbrowser.open("https://youtube.com")
        return "Opening Youtube..."
    elif 'open facebook' in message:
        webbrowser.open("https://facebook.com")
        return "Opening Facebook..."
    elif 'what is' in message or 'who is' in message or 'what are' in message:
        query = message.replace(" ", "+")
        webbrowser.open(f"https://www.google.com/search?q={query}")
        return f"This is what I found on the internet regarding {message}"
    elif 'wikipedia' in message:
        query = message.replace("wikipedia", "").strip()
        webbrowser.open(f"https://en.wikipedia.org/wiki/{query}")
        return f"This is what I found on Wikipedia regarding {message}"
    elif 'time' in message:
        time = datetime.now().strftime("%I:%M %p")
        return f"The current time is {time}"
    elif 'date' in message:
        date = datetime.now().strftime("%B %d")
        return f"Today's date is {date}"
    elif 'calculator' in message:
        webbrowser.open('Calculator:///')
        return "Opening Calculator"
    elif 'go to sleep' in message:
        speak("Going to sleep. Goodbye!")
        # Here you might add logic to close specific apps or pages
        # For example, closing Chrome browser (if it's open):
        os.system("taskkill /IM chrome.exe /F")
        sys.exit()  # Exit the application

    elif 'translate' in message:
        parts = message.split(' to ')
        text_to_translate = parts[0].replace("translate", "").strip()
        dest_language = parts[1].strip() if len(parts) > 1 else 'en'
        speak(translate_text)
        return translate_text(text_to_translate, dest_language)
    
    elif 'open' in message:
        app_name = message.replace("open", "").strip()
        return open_application(app_name)
    elif 'close' in message:
        app_name = message.replace("close", "").strip()
        return close_application(app_name)
    else:
        query = message.replace(" ", "+")
        webbrowser.open(f"https://www.google.com/search?q={query}")
        return f"I found some information for {message} on Google"

if __name__ == "__main__":
    app.run(debug=True)
