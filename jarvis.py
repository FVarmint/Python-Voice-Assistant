from genericpath import getatime
from typing import Mapping
import pyttsx3
import datetime
import wikipedia
import webbrowser
import os
import smtplib , ssl
# import PyAudio
import speech_recognition as sr
# from espeak import espeak
from win32com.client import Dispatch

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[0].id)
engine.setProperty('voice' , voices[1].id)

def speak(audio):
    # print(audio)
    # speak = Dispatch(('SAPI.SpVoice'))
    # speak.Speak(str)
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)

    if hour>=0 and hour<12:
        speak("Good Morning Sir")
    
    elif hour>=12 and hour<18:
        speak("Good Afternoon Sir")
    
    else:
        speak("Good Evening Sir")

    speak("How May I help you")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        # r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        # query = r.recognize(audio , Language="en-in")
        query = r.recognize_google(audio , language="en-in")
        print(query)
       

    except Exception as e:
        # print(e)
        print("Say that again please")
        takeCommand()
        return "None"
    return query

def sendEmail(to , content):
    try:
        server = smtplib.SMTP("smtp.gmail.com" , 587)
        server.ehlo() # Can be omitted
        server.starttls() # Secure the connection
        server.ehlo() # Can be omitted
        server.login("internst0@gmail.com", 'abc@123#ST')
        server.sendmail("internst0@gmail.com" , to , content)
    except Exception as e:
        # Print any error messages to stdout
        print(e)
    finally:
        server.close() 
    
# if '__name__' == '__main__':
#     speak('Dishan is a good boy')

# speak('the program is running')
if __name__ == '__main__':
    wishMe()
    while True:
        query = takeCommand().lower()

        #logic for executing commands based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia')
            query = query.replace('wikipedia' , "")
            results = wikipedia.summary(query , sentences=2)
            print(results)
            speak("According to Wikipedia")
            speak(results)
        
        elif 'open youtube' in query:
            speak("Opening Youtube")
            webbrowser.open("youtube.com")

        elif 'open gmail' in query:
            speak("Opening Gamil")
            webbrowser.open("gmail.com")

        elif 'open stackoverflow' in query:
            speak("Opening stackoverflow")
            webbrowser.open("stackoverflow.com")

        elif 'open amazon india' in query:
            speak("Opening amazon")
            webbrowser.open("amazon.in")

        elif 'open amazon us' in query:
            speak("opening amazon")
            webbrowser.open("amazon.com")

        elif 'play music' in query:
            speak("playing music")
            music_dir = "D:\\Music"
            songs = os.listdir(music_dir)
            for song in songs:
                if song.endswith(".mp3"):
                    print(songs)
                    os.startfile(os.path.join(music_dir , songs[0]))

        elif 'the time' in query:
           strTime = datetime.datetime.now().strftime("%H:%M:%S")
           speak(f"Sir , the time is {strTime}")

        elif 'visual studio code' in query:
            vscPath = "C:\\Users\\chuta\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(vscPath)

        elif 'pycharm' in query:
            pycharmPath = "C:\\Program Files\\JetBrains\\PyCharm 2019.2.3\\bin\\pycharm64.exe"
            os.startfile(pycharmPath)

        elif 'android studio' in query:
            androidStudioPath = "C:\\Program Files\\Android\\Android Studio\\bin\\studio64.exe"
            os.startfile(androidStudioPath)

        elif 'send email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "chutanidishan@gmail.com"
                sendEmail(to , content)
                speak("The mail has been sent")
            except Exception as e:
                print(e)
                speak("Sorry sir , I am not able to send email at the moment") 

        elif 'thank you' in query:
            speak("You're welcome sir")

        elif 'shut down' in query:
            speak("shutting down")
            exit()

        elif 'command prompt' in query:
            speak("Starting CMD")
            os.system("start cmd")

        elif 'discord' in query:
            speak('opening discord')
            discordPath = 'C:\\Users\\chuta\\AppData\\Local\\Discord\\app-0.0.309\\Discord.exe'
            os.startfile(discordPath)
        

