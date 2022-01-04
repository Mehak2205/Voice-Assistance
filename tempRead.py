def speak(str, win32com=None):
    from win32com.client import Dispatch

    speak= Dispatch("SAPI.SpVoice")

    speak.Speak(str)

if __name__ == "__main__":
    speak("Mehak is a good person")
