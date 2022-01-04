import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import pyaudio

print("Initializing Grey....")

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)



MASTER = "User"

def speak(str):
    from win32com.client import Dispatch
    speak= Dispatch("SAPI.SpVoice")
    speak.Speak(str)

# This function will make assistant greet me
def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning" + MASTER + "How are you?")
    elif hour >= 12 and hour<17:
        speak("Good Afternoon" + MASTER + "How are you?")
    else:
        speak("Good Evening" + MASTER + "How are you?")

    speak("I am Grey. How may I help you?")



# This function will take command
def takeCommand():
    speak("Listening...")
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        audio = r.listen(source)
        r.pause_threshold = 0.5
        r.energy_threshold = 100

    try:
        print("Recoganizing...")
        query = r.recognize_google(audio)
        print(f"user said: {query}\n")

    except Exception as e:
        print("Pardon! Can u please repeat...")
        speak("Pardon! Can u please repeat...")
        takeCommand()
        query = None

    return query


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('waytomayankdawar@gmail.com', 'Charger@123')
    server.sendmail('parfait2209@gmail.com', to, content)
    server.close()


# Main program Starts


if __name__ == "__main__":
    speak("Initializing Grey")
    wishMe()

    cond = True
    while cond:
        Str = takeCommand()
        if Str == None:
            query = "default"
        else:
            query = Str.lower()

        if 'wikipedia' in query:
            speak('Searching wikipedia...')
            query = query.replace("search", "")
            query = query.replace("on", "")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            print(results)
            speak(results)

        elif 'how are you' in query:
            speak("I'm perfectly fine !")

        elif 'open youtube' in query:
            webbrowser.open('https://www.youtube.com/')

        elif 'open google' in query:
            webbrowser.open('https://www.google.com/')

        elif 'music' in query:
            songs_dir = "F:\\Music"
            songs = os.listdir(songs_dir)
            os.startfile(os.path.join(songs_dir, songs[0]))
            os.startfile(os.path.join(songs_dir, songs[1]))

        elif 'email' in query:
            try:
                speak("What should I send!")
                content = takeCommand()
                to = "parfaite2209@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent successfully")
            except Exception as e:
                print(e)

        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"{MASTER} the time if {strTime}")

        elif 'android studio' in query:
            codePath = "C:\\Program Files\\Android\\Android Studio\\bin\\studio64.exe"
            os.startfile(codePath)

        elif 'java' in query:
            codePath = "C:\\Program Files\\JetBrains\\IntelliJ IDEA 2019.3\\bin\\idea64.exe"

        elif 'shutdown' in query:
            print("Shutting Down ...Have a nice day!")
            speak("Shutting Down ...Have a nice day!")
            break
            cond = False

        elif 'open email' in query:
            webbrowser.open('https://mail.google.com/mail/u/0/#inbox')

        elif 'default' in query:
            speak("Pardon! I am still learning, can you please repeat this")

        else:
            url = "https://www.google.com.tr/search?q={}".format(query)
            webbrowser.open_new_tab(url)

exit()
