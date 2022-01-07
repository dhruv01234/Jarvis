import speech_recognition as sr
import webbrowser
import datetime
import wikipedia
import os
import re
import smtplib
import random
import pywhatkit
import pyautogui
import shutil

def get_audio(): 
    r = sr.Recognizer()
    with sr.Microphone()  as source:
        print("listning....")
        # r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("recognising....")
        query = r.recognize_google(audio,language="en-in").lower()
        print(query)
    except Exception as e:
        print(e)
        return "none"
    return query

def speak(query):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(query)

def starting():
    time = int(datetime.datetime.now().hour)
    if time >= 4 and time <=12:
        speak("Good Morning sir how may i help you")
    elif time>12 and time<=17:
        speak("Good Afternoon sir how may i help you")
    elif time>17 and time<=23:
        speak("Good Evening sir how may i help you")
    else:
        speak("are you not sleeping sir tell me how may i help you")

def time():
    h = str(datetime.datetime.now())[11:13]
    m = str(datetime.datetime.now())[14:16]
    t = int(datetime.datetime.now().hour)
    if t>=0 and t<12:
        speak(f"{h} {m} AM")
    else:
        h = int(h)
        h -= 12
        speak(f"{h} {m} PM")

def date():
    d = str(datetime.datetime.now())[8:10]
    m = (str(datetime.datetime.now())[5:7])
    y = str(datetime.datetime.now())[0:4]
    datetime_object = datetime.datetime.strptime(m, "%m")
    month_name = datetime_object.strftime("%B")
    speak(f"Today date is {d} {month_name} {y}")

def findlink(query):
    r = re.findall(r"[a-z]*[.][a-z]*|[a-z]* dot [a-z]",query)
    return r[0]

def sendemail(email,content):
    server = smtplib.SMTP("smtp.gmail.com",587)
    server.ehlo()
    server.starttls()
    e = input("enter your email : ")
    p = input("enter your password : ")
    server.login(e,p)
    server.sendmail(e,email,content)
    server.close()

starting()
i = 1
while(True): 
    query = get_audio()
    if "exit" in query:
        speak("thank you welcome back again")
        break
    elif "open youtube" in query:
        speak("opening youtube")
        webbrowser.open("https:\\youtube.com")
    elif "name" in query:
        speak("Sir your name is Dhruv Modi")
    elif "open whatsapp" in query:
        speak("opening whatsapp")
        webbrowser.open("https:\\web.whatsapp.com")
    elif "who are you" in query:
        speak("I am jarvis")
    elif "time" in query:
        time()
    elif "date" in query:
        date()
    elif "google" in query:
        webbrowser.open("https:\\google.com")
    elif " code" in query:
        os.startfile("C:\\Users\\HP Pavilion\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe")
    elif "." in query:
        link = findlink(query)
        webbrowser.open(f"https:\\{link}")
        speak("if your link is not open yet please write your link and if it opens press enter")
        link = input()
        if link !="\n":
            webbrowser.open(f"https:\\{link}")
    elif "mail" in query:
        try:
            speak("tell the email id to which you want to send email")
            email = input()
            speak("to send the email you want to speak or want to write the content")
            q = get_audio()
            if "write" in q:
                content = input()
            elif "speak" in q:
                content = get_audio()
            sendemail(email,content)
            speak("email has been sent succesfully")
        except Exception as e:
            print(e)
            speak("sorry sir i am unable to send the email try again")
    elif "roll" in query:
        number = random.randint(1,6)
        speak(f"you got {number}")
    elif "toss" in query:
        n = random.randint(1,2)
        if n==1:
            speak(f"you got head")
        else:
            speak(f"you got taill")
    elif "snake" in query:
        os.startfile("C:\\Users\\HP Pavilion\\Desktop\\games_python\\first_game\\first_game")
    elif ("create a folder" in query) or ("directory" in query):
        speak("to create a new folder please write the folder path name and i will create it for you")
        name = input()
        os.mkdir(name)
    elif "create a file" in query:
        speak("to create a new file please write the file path name and i will create it for you")
        f = open(input(),"w")
        f.close()
    elif (("remove" in query) or ("delete" in query)) and (("folder" in query) or ("directory" in query)):
        speak("to delete a folder write folder path name")
        folder = input()
        os.removedirs(folder)
    elif (("remove" in query) or ("delete" in query)) and ("file" in query):
        speak("to delete a file write file path name")
        fl = input()
        os.remove(fl)
    elif "are you mad" in query:
        speak("you make me if you are mad then i am also mad hhahhahhahhahha")
    elif "how are you" in query:
        speak("i am fine and what about you")
        me = get_audio()
        if "not good" in me or "not fine" in me:
            speak("ohhh that is bad you can play pubg or mini militia to make your mood feel better hope you enjoy")
        elif "fine" in me or "good" in me:
            speak("that is nice sir hope your coming time also be good")
        else:
            speak("ok but i think you should fresh your mind")
    elif "who make you" in query:
        speak("dhruv modi make me")
    elif "hello" in query:
        speak("hello sir do you want any help just say it i will do it for you if i can")
    elif "i love you" in query:
        speak("ohh thank you sir but i think you should find a better than me because i am not a human being")
    elif "are you married" in query:
        speak("no sir you hit my heart please find or make a girl AI for me")
    elif "your favourite colour" in query:
        speak("my favourite color is red and blue")
    elif "what is love" in query:
        speak("according to me love has no defination it is a feeling and a wireless connection of two hearts that has infinite range")
    elif "map" in query:
        speak("opening map")
        webbrowser.open("https:\\maps.google.co.in")
    elif "best date" in query:
        speak("5 jan is the best date")
    elif "wikipedia of" in query:
        li = query.split("of ")
        actor = li[1]
        speak(wikipedia.summary(actor,sectences = 3))
    elif "spell" in query:
        word = query.split("spell ")[1]
        word_list = list(word)
        for i in word_list:
            speak(i)
    elif ("play music" in query) or ("play song" in query):
        speak("which song you want to play")
        song = get_audio()
        pywhatkit.playonyt(song)
    elif "download" in query:
        download = "download" + query.split("download")[1]
        webbrowser.open(f"https:\\www.google.com\\search?q={download}")
    elif "search on youtube" in query:
        speak("for what content should i search on youtube")
        content = get_audio()
        webbrowser.open(f"https:\\www.youtube.com\\search?q={content}")
    elif "search" in query:
        search = query.split("search ")[1]
        webbrowser.open(f"https:\\www.google.com\\search?q={search}")
    elif "repeat" in query:
        speak("what you want to repeat and if does not want to repeat just say cancel")
        repeat = get_audio()
        while("cancel" not in repeat):
            speak(repeat)
            repeat = get_audio()
    elif "play" in query:
        song = query.split("play ")[1]
        pywhatkit.playonyt(song)
    elif "open c folder" in query:
        os.startfile("C:\\")
    elif "open d folder" in query:
        os.startfile("D:\\")
    elif "open e folder" in query:
        os.startfile("E:\\")
    elif "open f folder" in query or "open app folder" in query:
        os.startfile("F:\\")
    elif "open g folder" in query:
        os.startfile("G:\\")
    elif "close" in query:
        program = query.split("close ")[1]
        os.system(f"taskkill/im {program}.exe")
    elif "open" in query:
        app = query.split("open ")[1]
        os.startfile(f"{app}.exe")
    elif "take screenshot" in query:
        ss = pyautogui.screenshot()
        n = f"screenshot{i}.png"
        ss.save(n)
        # shutil.copyfile(f'C:\Users\HP Pavilion\Desktop\project_python\{n}',f'C:\Users\HP Pavilion\Pictures\Screenshots\{n}.png')
        i += 1