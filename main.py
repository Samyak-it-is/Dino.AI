import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import pywhatkit
import smtplib
import win32com.client

engine= pyttsx3.init()
engine.setProperty('voice', 'Microsoft Sam')

'''speaker=win32com.client.Dispatch("SAPI.SpVoice")
s=input()
speaker.Speak(s)'''


def sendEmail(to,content):
    server = smtplib.SMTP("smtp.gmail.com",587)
    server.ehlo()
    server.starttls()
    server.login("samyak@gmail.com","123")
    server.sendmail("samyak@gmail.com",to,content)
    server.close()


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon")
    else:
        speak("Good Evening")
    speak("I am Dino Please Let Me know if you need anything!!")

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)    
        print("Say that again please...")  
        return "None"
    return query


if __name__=="__main__":
    #speak("Hey Dont lose hope and complete this project!! no matter HOW")

    wishMe()
    while True:
        query = takeCommand().lower()

        if 'play music' in query:
            music_dir = "D:\\FavMusic"
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir,songs[0]))

        elif "the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strfTime}")

        elif "open code" in query:
            codePath = "D:\Microsoft VS Code\Code.exe"
            os.startfile(codePath)

        elif "email to " in query:
            try:
                speak("What do you want to send email about")
                content = takeCommand()
                to ="samyak1@gmail.com"
                sendEmail(to,content)
                speak("Email has been sent")
            except Exception as e:
                print(e)
                speak("Sorry,email not sent")


        elif "youtube" in query.lower():
            text=query.replace("on youtube","")
            speak("OKAY!!!, playing that on youtube")
            pywhatkit.playonyt(query)
        elif "google" in query.lower():
            query=query.replace("on google","")
            speak("OKAY!!!, searching that on google")
            pywhatkit.search(query)

        elif "whatsapp" in query.lower():
            text=query.lower()
            phone="+919098765432"
            message="Hi this is an automated message"
            speak("Message will be sent")
            pywhatkit.sendwhatmsg(phone ,message,18,54)


                                  
        elif "thank you" in query:
            exit()

        