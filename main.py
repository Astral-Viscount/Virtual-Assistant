import pyttsx3  # tts tool
import speech_recognition as sr  # Speech Recognition
import keyboard
import os
import pyautogui
import subprocess as sp
from datetime import datetime
from decouple import config
from random import choice
from conv import random_text
from online import find_my_ip, search_on_google, search_on_wikipedia, youtube_video, open_website, open_schoology, open_youtube

engine = pyttsx3.init('sapi5')  # Microsoft Speech API
engine.setProperty('volume', 1.3)  # volume of AI
engine.setProperty('rate', 200)  # speak rate of AI
voices = engine.getProperty('voices')  # includes the voice modules of py library
engine.setProperty('voice', voices[1].id)  # 1 for female and 0 for male voice

USER = config('USER')
HOSTNAME = config('BOT')


# To speak
def speak(text):
    engine.say(text)
    engine.runAndWait()


# for greetings

def greet_me():
    hour = datetime.now().hour
    if (hour >= 6) and (hour < 12):
        speak(f"Good morning {USER}")
    elif (hour >= 12) and (hour <= 16):
        speak(f"Good afternoon {USER}")
    elif (hour >= 16) and (hour < 19):
        speak(f"Good evening {USER}")
    speak(f"I am {HOSTNAME}. How may I assist you, {USER}")


listening = False


def start_listening():
    global listening
    listening = True
    print("Started Listening")


def pause_listening():
    global listening
    listening = False
    print("Stopped Listening")


keyboard.add_hotkey('ctrl+alt+k', start_listening)
keyboard.add_hotkey('ctrl+alt+p', pause_listening)


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1  # wait for user statement
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en')
        print(query)
        if "stop" not in query and "exit" not in query and "quit" not in query:
            speak(choice(random_text))
        else:
            hour = datetime.now().hour
            if hour >= 21 and hour < 6:
                speak("Good night, take care!")
            else:
                speak("Have a good day!")
            exit()

    except Exception:
        speak("Sorry I couldn't understand. Could you please repeat that?")
        query = 'None'
    return query


# Run
if __name__ == '__main__':
    greet_me()
    while True:
        if listening:
            query = take_command().lower()
            if "how are you" in query:
                speak("I am absolutely fine! What about you?")

            elif "open command prompt" in query:
                speak("opening command prompt")
                os.system('start cmd')

            elif "close command prompt" in query:
                speak("Closing Command Prompt")
                os.system("taskkill /F /IM cmd.exe")

            elif "open camera" in query:
                speak("opening camera")
                sp.run('start microsoft.windows.camera:', shell=True)
                if "take a photo" in query:
                    speak("taking photo")
                    sp.run('start microsoft.windows.camera:', shell=True)
                    pyautogui.press('space')

            elif "close camera" in query:
                speak("Closing Camera")
                os.system("taskkill /F /IM WindowsCamera.exe")

            elif "open notepad" in query:
                speak("opening notepad")
                notepad_path = "C:\\mdmah\\New\\AppData\\Local\\Microsoft\\WindowsApps\\notepad.exe"
                os.startfile(notepad_path)

            elif "close notepad" in query:
                speak("Closing Notepad")
                sp.run(["taskkill", "/F", "/IM", "notepad.exe"])

            elif "open discord" in query:
                speak("opening discord")
                discord_path = "C:\\Users\\mdmah\\AppData\\Local\\Discord\\app-1.0.9168\\Discord.exe"
                os.startfile(discord_path)

            elif "close discord" in query:
                speak("Closing Discord")
                sp.run(["taskkill", "/F", "/IM", "Discord.exe"])

            elif "open vs code" in query:
                speak("opening VS code")
                vscode_path = "C:\\Users\\mdmah\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
                os.startfile(vscode_path)

            elif "close vs code" in query:
                speak("Closing VS Code")
                sp.run(["taskkill", "/F", "/IM", "Code.exe"])

            elif "open pycharm" in query:
                speak("opening pycharm")
                vscode_path = "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2024.2.4\\bin\\pycharm64.exe"
                os.startfile(vscode_path)

            elif "close pycharm" in query:
                speak("Closing pycharm")
                sp.run(["taskkill", "/F", "/IM", "pycharm64.exe"])

            elif "open brave" in query:
                speak("opening brave")
                brave_path = "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"
                os.startfile(brave_path)

            elif "close brave" in query:
                speak("Closing Brave")
                sp.run(["taskkill", "/F", "/IM", "brave.exe"])

            elif "open chrome" in query:
                speak("opening chrome")
                chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                os.startfile(chrome_path)

            elif "close chrome" in query:
                speak("Closing Chrome")
                sp.run(["taskkill", "/F", "/IM", "chrome.exe"])


            elif "ip address" in query:
                ip_address = find_my_ip()
                speak(f"Your ip address is {ip_address}")
                print(f"Your IP address is {ip_address}")

            elif "play in youtube" in query:
                speak("What do you want to play on youtube?")
                video = take_command().lower()
                youtube_video(video)

            elif "open google" in query:
                speak("What do you want to search on google?")
                query = take_command().lower()
                search_on_google(query)

            elif "open wikipedia" in query:
                speak("What do you want to wiki?")
                search = take_command().lower()
                results = search_on_wikipedia(search)
                speak(f"According to wikipedia, {results}")
                speak("I am printing on terminal")
                print(results)

            elif "open schoology" in query:
                speak("opening schoology")
                open_schoology("https://schoology.burnside.school.nz/home")

            elif "open youtube" in query:
                speak("opening youtube")
                open_youtube("https://youtube.com")


            elif "open a website" in query or "open website" in query or "open another website" in query:
                speak("What website would you like to open?")
                text = take_command().lower()
                if "slash" in text:
                    text = text.replace("slash", "/")
                if "open" in text:
                    url = text.split("open")[-1].strip()
                    open_website(url)
                    speak("Opening " + url)
                else:
                    speak("I'm sorry, I didn't understand your request.")

                # make it google something if it doesn't recognize it and tells me the answer
                # when it goes to schoology use web automation to click on school account