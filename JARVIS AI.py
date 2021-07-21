import pyttsx3  # use "pip install speech recognition" in command prompt
import speech_recognition as sr  # use "pip install speech recognition" in command prompt
import datetime  # use "pip install datetime" in command prompt
import wikipedia  # use "pip install wikipedia" in command prompt
import webbrowser  # use "pip install webbrowser" in command prompt
import os  # use "pip install os" in command prompt
import smtplib  # use "pip install smtplib" in command prompt
from playsound import playsound  # use "pip install playsound" in command prompt
import win32gui  # use "pip install win32gui" in command prompt
import win32con  # use "pip install win32con" in command prompt
import requests  # use "pip install requests" in command prompt
from youtube_search import YoutubeSearch  # use "pip install youtube search" in command prompt
import json  # use "pip install json" in command prompt

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')

# print(voices)
# for male
engine.setProperty('voice', voices[1].id)


# for female
# engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


# play youtube videos


def ytlink():
    ytsearch = SearchVideos("despacito", offset=1,
                            mode="json", max_results=1)
    vidli = ytsearch.result()
    # parse x:
    ydata = json.loads(vidli)
    for ytv in ydata['search_result']:
        ytmli = ytv['link']

    webbrowser.open(ytmli)


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

    speak("I am Luna Sir and Madame. Please tell me how may I help you")


def takeCommand():
    # It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"You said: {query}\n")

    except Exception:
        # print(e)
        print("Say that again please...")
        return "None"
    return query


if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("sir,")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'open github' in query:
            webbrowser.open("github.com")

        elif 'open facebook' in query:
            webbrowser.open("facebook.com")

        # test me
        elif 'who is' in query or 'how to' in query or 'what is' in query:
            speak('Searching Wikipedia...')
            resultsw = wikipedia.summary(query, sentences=3)
            speak("sir,")
            print(resultsw)
            speak(resultsw)
            speak("that's it")

        elif 'play music' in query:
            # added  misic dir location
            music_dir = 'C:\\Users\\Laiku\\Music\\Favorate songs'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[0]))

        elif 'the time' in query or 'what is time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'open code' in query:
            # added vscode exe locaion
            speak("opening vs code.")
            codePath = "C:\\Users\\User\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"  # add your vs code path
            os.startfile(codePath)

        elif 'open sublime' in query:
            # added sublime exe locaion
            speak("opening Sublime Text ")
            codePaths = "C:\\Program Files\\Sublime Text 3"  # add sublime text path
            os.startfile(codePaths)

        elif 'open firefox' in query:
            # added sublime exe locaion
            speak("opening firefox browser ")
            codePathf = "C:\\Users\\Laiku\\AppData\\Local\\Mozilla Firefox"  # add firefox browser path
            os.startfile(codePathf)

        elif 'hide window' in query or 'hide work' in query or 'change window' in query or 'minimise window' in query:
            # close in window
            speak("ok.")
            Minimize = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(Minimize, win32con.SW_MINIMIZE)

        elif 'full window' in query or 'full screen window' in query or 'fullscreen' in query or 'maximize window' in query:
            # full in window
            speak("sure.")
            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

        elif 'who are you' in query or 'about you' in query or "your details" in query:
            speak("i am your work partner. sir, i work for you.")

        elif 'how are you' in query or 'fill you' in query or "filling" in query:
            speak("i am your fine. sir, can i help you.")

        elif 'exit' in query or 'goodbye' in query or 'good bye' in query or 'bye' in query:
            speak("thank you sir. good bye .")
            quit('while True :')

        elif 'thank you' in query or 'thanks' in query:
            speak("No problem sir.")

        elif "hello" in query or "hello Luna" in query:
            hel = "Hello  Sir ! How May i Help you.."
            print(hel)
            speak(hel)

        elif 'clean' in query:
            speak("ok.")


            def clear():
                return os.system('cls')


            clear()
        elif "play on youtube" in query:
            query = query.replace("play on youtube", "")
            speak(query + "playing youtube")
            ytlink()

        elif "it is fathers day" in query or "It is Fathers Day" in query:
            print(speak)
            speak(
                "Happy father's day to sir's daddy - from Luna. Daddy you are the best dad in the whole world. We can't describe how we love you and how much you love us. Thanks dad for all you have done for us. Here is a simple quote from us daddy (A father is neither an anchor to hold us back nor a sail to take us somewhere, but a guiding light whose love shows us the way.")

        elif "open photo" in query:
            webbrowser.open("file:///C:/Users/Laiku/Pictures/Saved%20Pictures/IMG_2529%20(1).jpg")

        elif "what is the weather" in query or "weather" in query:
            webbrowser.open(
                "https://www.google.com/search?q=weather&rlz=1C1OKWM_enUS772US772&oq=weather&aqs=chrome..69i57j0i131i433j0j0i131i433l2j0i433j0i131i433i457j0i402l2j0i131i433.2547j0j7&sourceid=chrome&ie=UTF-8")

        elif 'email to sai' in query or 'email to sahi':
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "layakarasai476@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry my friend Laya. I couldn't send the Email.")
