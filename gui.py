from tkinter import *
import tkinter as tk
import speech_recognition as sr
import win32com.client as wincl
import time
import wikipedia
import os
import webbrowser
from string import *
from PyDictionary import PyDictionary


#<------time task--------->
def Time(q):
    if q=="t":
        speak.Speak("time is "+time.strftime("%I:%M:%S"))
        

#<-----------dictionary----------->
def dic():
    dictionary=PyDictionary(li[23:])
    if li[0:22]=="give me the meaning of":
          print (dictionary.getMeanings())
          speak.Speak(dictionary.getMeanings())
    elif li[0:22]=="give me the synonyms of":
          print (dictionary.getSynonyms())
          speak.Speak(dictionary.getSynonyms())
    



        
#<------greeting------->
def greeting():
  global gr
  gr =""
  if "hi" in li[0:2]:
      gr="Hi."
  if int(time.strftime("%H"))>=5 &  int(time.strftime("%H"))>12:
       gr+="Good Morning"
       #if li in["good evening","good night","good afternoon","good noon"]:
          #gr+=". it is before 12 o clock so i thing it is morning"
  if int(time.strftime("%H"))>=12 &  int(time.strftime("%H"))>16:
       gr+="Good Afternoon"
       #if li in ["hi","good morning","good evening","good night"]:
        #  gr+=". it is before 4 o clock so i thing it is afternoon"
  if int(time.strftime("%H"))>=16:
       gr+="Good evening"
       #if li in ["hi","good morning","good evening","good night","good afternoon","good noon"]:
        #  gr+=". it is after 4 so i thing it is  evening"
  if  li=="good night":
         print ("Are You Going To Sleep Good Night\nbye")
         speak.Speak("Are You Going To Sleep Good Night. bye")
  else:
       print (gr)
       speak.Speak(gr)
       
      



#<------wiki task-------->
def wiki():
   speak = wincl.Dispatch("SAPI.SpVoice")
   if li=="tell me about yourself":
       print ("my name is bob.and i am here to help u")
       speak.Speak("my name is bob.and i am here to help u")
   else:
    print("working on "+li[14:]+"....")
    speak.Speak("working on "+li[14:])
    print (wikipedia.summary(li[14:],sentences=1))
    speak.Speak(wikipedia.summary(li[14:],sentences=1))
   


#<--------search---------->
def search():
     if li[0:23]=="search in YouTube about":
        webbrowser.open_new("https://www.youtube.com/results?search_query="+li[24:])
     elif li[0:22]=="search in Google about":
         webbrowser.open_new("https://www.google.co.in/search?q="+li[23:])
     else:
        webbrowser.open_new("https://www."+li[7:]+".com")
     
     

#<----------open---------->
def open():
  if  li =="open calculator":
    os.system('calc.exe')
  elif len(li)<15:
      if capwords(li[5:]) in ['Desktop','Documents','Downloads','Music','Videos','Pictures']:
           webbrowser.open("C:/Users/"+os.getenv('username')+"/"+capwords(li[5:]))
      else:
          print ("we are unable to detect location "+li[5:]+"\n please try again by say \n open desktop,documents,vedios,pictures,music \n or \n u can say open a file in destop,documents.....")
          speak.Speak("we are unable to detect location "+li[5:]+". please try again by say.  open desktop,documents,vedios,pictures,music . or u can say open a file in desktop,documents and so on")
          start();
  elif li[0:14]=="open a file in":
        ne=os.listdir("C:/Users/"+os.getenv('username')+"/"+capwords(li[15:]))
        for i in ne:
          print(ne.index(i)+1, end=' ')
          print(" ",i)
        try:
          with sr.Microphone() as source:
            time.sleep(10)
            speak.Speak("pronounce the number in the list that u want to open like one two .....")
            print ("listening......")
            audio = r.listen(source)
            nu=r.recognize_google(audio)
            print(nu)
            try:
               webbrowser.open("C:/Users/"+os.getenv('username')+"/"+capwords(li[15:]+"/"+ne[int(nu)-1]))
            except:
               print ("we have found some illigal number can u please type the number that u want to open in the list")
               print ("we have found some illigal number can u please type the number that u want to open in the list")
               speak.Speak("we have found some illigal number can u please type the number that u want to open in the list")
               nu=input()
               webbrowser.open("C:/Users/"+os.getenv('username')+"/"+capwords(li[15:]+"/"+ne[int(nu)-1]))
        except sr.UnknownValueError:
             print("Google Speech Recognition could not understand audio")
             speak.Speak("we have found some illigal number can u please type the number that u want to open in the list")
             nu=input()
             webbrowser.open("C:/Users/"+os.getenv('username')+"/"+capwords(li[15:]+"/"+ne[int(nu)-1]))
             #precld();
        except sr.RequestError as e:
             print("Could not request results from Google Speech Recognition service; {0}".format(e))
  else:
          print ("we are unable to detect location "+li[5:]+"\n please try again by say \n open desktop,documents,vedios,pictures,music \n or \n u can say open a file in destop,documents.....")
          speak.Speak("we are unable to detect location "+li[5:]+". please try again by say.  open desktop,documents,vedios,pictures,music . or u can say open a file in desktop,documents and so on")
               
  
'''
#<----------conclude---------->
def precld():
 time.sleep(1)
 speak.Speak("your task is done sucessfully . to continue enter y")
 d = input ("do u want to speak again press y else n\n")
 if d=="y":
     start();
 else:
     Quit();'''

#<---------------what--------->
def what():
     print("open desktop,documents,vedios,pictures,music .\n open a file in desktop,documents and so on. \n give me the meaning of\n tell me about\n search\n what is the time\n play music etc")
     speak.Speak("u can order open desktop,documents,vedios,pictures,music . or u can say open a file in desktop,documents and so on. and u can ask give me the meaning of, tell me about, search, what is the time,play music etc")
               

#<---------play---------->
def play():
     webbrowser.open("C:\\Program Files (x86)\\Windows Media Player\\wmplayer.exe")
     speak.Speak("here is the player please select and enjoy the "+li[5:])
     #precld();


#<--------Quit----------->
def Quit():
    speak.Speak("bye have a great day")
         
      
def display():
  T = tk.Label(text=li).pack() 
   
  
          

#<-----start------>
def start():
 try:
   with sr.Microphone() as source:
    print ("listening......")
    global r
    r = sr.Recognizer()
    r.adjust_for_ambient_noise(source)
    speak.Speak("hi. how can i help u")
    audio = r.listen(source)
    global li 
    li=r.recognize_google(audio)
    sys.stdout.flush()
    display();
    global gl
    gl=["hi good morning","hi good evening","hi good night","hi good afternoon"," hi good noon","hi","good morning","good evening","good night","good afternoon","good noon"]
#<-----TASK ASSIGNMENT------->
    if li in ["time","what is the time"]:
       Time(q="t");
    elif li in gl:
        greeting();
    elif li[0:22] in ["give me the meaning of","give me the synonyms of"]:
           dic();
    elif li[0:13]==("tell me about"):
        wiki();
    elif li[0:6]=="search":
         search();
    elif li[0:4] in ["open","view"]:
          open();
    elif li in ["play music","play movies"]:
          play();
    elif li in ["what can you do"]:
          what();
    elif li in ["quit","shutdown","close"]:
          Quit();
    else:
         print("let me search in google")
         speak.Speak("let me search in google")
         webbrowser.open_new("https://www.google.co.in/search?q="+li)
         #precld();
 except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
        speak.Speak("Google Speech Recognition could not understand audio")
 except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; please check your internet {0}".format(e))
        speak.Speak("Could not request results from Google Speech Recognition service. please check your internet") 





 


def center_window(width=300, height=200):
    # get screen width and height
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # calculate position x and y coordinates
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    window.geometry('%dx%d+%d+%d' % (width, height, x, y))


window = tk.Tk()
center_window(500, 400)
photo=PhotoImage(file="1.png")

title = tk.Label(window, text="welcome to bob assistance",  fg="skyblue", font=("Helvetica", 16)).pack()

button1 = tk.Button(image=photo ,width="100",height="70",command=start).pack()

speak = wincl.Dispatch("SAPI.SpVoice")
window .mainloop()

