import os
import time
import sys
import ctypes
import random
import win32com.client as wincl     #pip install pypiwin32
import speech_recognition as sr     #pip install speechrecognition
import webbrowser                   #pip install webbrowser
from bs4 import BeautifulSoup       #pip install beautifullsoup/Beautifullsoup
import requests                     #pip install requests
import winshell                     #pip install winshell
from urllib.request import urlopen  #pip install urllib
requests.packages.urllib3.disable_warnings()
leave=0
videos=['Ed Sheeran - Shape of You [Official Video].mp4','Ellie Goulding - Burn.mp4'
        ,'Agar Tum Saath Ho  FULL VIDEO SONG  Tamasha  Ranbir Kapoor & Deepika Padukone.mp4'
        ,'Afreen Afreen, Rahat Fateh Ali Khan & Momina Mustehsan, Episode 2, Coke Studio Season 9.mp4'
        ,'bulleya.mp4','love me like u do.mp4','boulevards of broken dreams.mp4','numb.mp4','what makes u beautifull.mp4']
speak = wincl.Dispatch("SAPI.SpVoice")
yt=0
while(leave==0 and yt==0):
    again=0
    yt=0
    r = sr.Recognizer()
    speak.Speak("sir,speak what should i do for you and If you want to leave just say Close the program")
    with sr.Microphone() as source:
#for mic use sr.Microphone() as source and r.listen instead of r.record
#print("Say something! I'll try to recognize it")   use sr.AudioFile("woman1_wb.wav") as source for reading audio file and r.record instead of r.listen
        r.adjust_for_ambient_noise(source, duration = 0.5)  
        audio = r.listen(source)
    try:
        put=r.recognize_google(audio, key = None, language = "en-US", show_all = False)
        print("You said:" + put)
    except sr.UnknownValueError:
        again=1
        print("I tried hard but could not understand what you spoke")
    except sr.RequestError as e:
        again=1
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
    except:
        again=1
        print("Unknown exception occurred!")
    if(again==1):
        again=0
        continue
    put=put.lower()
    link=put.split()
    if(put.startswith('open')):
        try:
            webbrowser.open('http://www.'+link[1]+'.com')
        except:
            speak.Speak("sorry,no internet connection")
    elif(put.startswith('play')):
        try:
            link = '+'.join(link[1:])
            voice = link.replace('+', ' ')
            url = 'https://www.youtube.com/results?search_query='+link
            source = requests.get(url,timeout=50)
            txt = source.text
            soup = BeautifulSoup(txt, "html.parser")
            allsongs = soup.findAll('div', {'class': 'yt-lockup-video'})
            parsong = allsongs[0].contents[0].contents[0].contents[0]
            toplay = parsong['href']
            speak.Speak("playing "+voice)
            webbrowser.open('https://www.youtube.com'+toplay)
            yt=1
        except:
            speak.Speak('Sorry, No internet connection!')
    elif put.startswith('empty'):
        try:
            winshell.recycle_bin().empty(confirm=False,show_progress=False, sound=True)
            speak.Speak("Sir the Recycle Bin is Empty!!")
            print('recycle bin empty .You can check it')
        except:
            print("Unknown Error")
    elif put.startswith('lock'):
        try:
            speak.Speak("Sir i am Putting this device in sleep mode")
            ctypes.windll.user32.LockWorkStation()
        except Exception as e:
            print(str(e))
    elif put.endswith('bored'):
        try:
            speak.Speak('''Sir, I\'m playing a music video since you're bored''')
            song = random.choice(videos)
            os.startfile(song)
            quit
        except Exception as e:
            print(str(e))
    elif('close' in link):
        leave=1
    elif('time' in link):
        speak.Speak('sir the time is')
        speak.Speak(time.strftime("%I:%M:%S"))
    elif('date' in link):
        speak.Speak('sir the date is')
        speak.Speak(time.strftime("%x"))
if(yt!=1):
    speak.Speak("closing the program")
quit
