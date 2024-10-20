import os
import time
import smtplib
import speech_recognition as sr
import pyttsx3
import webbrowser
import pyjokes
import pyautogui
import subprocess
from datetime import datetime
import requests
from bs4 import BeautifulSoup
from forex_python.converter import CurrencyRates
from googletrans import Translator
from selenium import webdriver
from PIL import Image
import pytesseract
import keyboard

# Initialize speech recognition and text-to-speech engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Function for speaking text
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Voice command listening function that recognizes voice commands and ignores background noise
def take_command():
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        print("Listening for your command...")
        audio = recognizer.listen(source)
        
        try:
            command = recognizer.recognize_google(audio)
            print(f"You said: {command}")
        except sr.UnknownValueError:
            speak("Sorry, I didn't catch that.")
            return None
        return command.lower()

# Open applications and websites via voice commands
def open_application_or_website(command):
    if 'open browser' in command:
        webbrowser.open("http://google.com")
    elif 'open youtube' in command:
        webbrowser.open("http://youtube.com")
    elif 'open facebook' in command:
        webbrowser.open("http://facebook.com")
    elif 'open instagram' in command:
        webbrowser.open("http://instagram.com")
    elif 'open gmail' in command:
        webbrowser.open("http://mail.google.com")
    elif 'open excel' in command:
        os.startfile("C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE")
    elif 'open word' in command:
        os.startfile("C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE")
    elif 'open powerpoint' in command:
        os.startfile("C:/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE")
    elif 'open camera' in command:
        os.system("start microsoft.windows.camera:")
    elif 'open calculator' in command:
        os.system("calc")
    elif 'open notepad' in command:
        os.system("notepad")
    elif 'open control panel' in command:
        os.system("control")
    elif 'open my pc' in command:
        os.system("explorer")
    elif 'open chrome' in command:
        os.startfile("C:/Program Files/Google/Chrome/Application/chrome.exe")

# Function to control tabs and browser actions
def browser_control(command):
    if 'open tab' in command:
        pyautogui.hotkey('ctrl', 't')
    elif 'close tab' in command:
        pyautogui.hotkey('ctrl', 'w')
    elif 'open new window' in command:
        pyautogui.hotkey('ctrl', 'n')
    elif 'close browser' in command:
        pyautogui.hotkey('alt', 'f4')

# Control all desktop application actions: minimize, maximize, close
def desktop_application_control(command):
    if 'minimize' in command:
        pyautogui.hotkey('win', 'down')
    elif 'maximize' in command:
        pyautogui.hotkey('win', 'up')
    elif 'close' in command:
        pyautogui.hotkey('alt', 'f4')

# Function to control keyboard keys via voice command
def keyboard_control(command):
    if 'press' in command:
        key = command.split("press")[-1].strip()
        keyboard.press_and_release(key)
    elif 'type' in command:
        text_to_type = command.split("type")[-1].strip()
        keyboard.write(text_to_type)

# Open or close applications and software by name
def manage_applications(command):
    if 'open' in command:
        app_name = command.split('open')[-1].strip()
        speak(f"Opening {app_name}")
        os.system(f"start {app_name}")
    elif 'close' in command:
        app_name = command.split('close')[-1].strip()
        speak(f"Closing {app_name}")
        os.system(f"taskkill /f /im {app_name}.exe")

# Function to search on YouTube
def search_youtube(command):
    if 'search' in command and 'youtube' in command:
        search_query = command.split('search')[-1].replace('on youtube', '').strip()
        speak(f"Searching for {search_query} on YouTube.")
        webbrowser.open(f"https://www.youtube.com/results?search_query={search_query}")

# Function to handle volume and music controls
def music_control(command):
    if 'volume up' in command:
        pyautogui.press('volumeup')
    elif 'volume down' in command:
        pyautogui.press('volumedown')
    elif 'mute' in command:
        pyautogui.press('volumemute')

# Main loop to listen and execute commands
if __name__ == "__main__":
    speak("Voice assistant is now active. How can I assist you?")
    while True:
        command = take_command()
        if command:
            if 'open' in command or 'close' in command:
                open_application_or_website(command)
                manage_applications(command)
            elif 'minimize' in command or 'maximize' in command or 'close' in command:
                desktop_application_control(command)
            elif 'press' in command or 'type' in command:
                keyboard_control(command)
            elif 'search' in command and 'youtube' in command:
                search_youtube(command)
            elif 'volume' in command:
                music_control(command)
            elif 'open tab' in command or 'close tab' in command:
                browser_control(command)
            else:
                speak("I didn't understand the command. Please try again.")
        else:
            speak("Waiting for your command...")
