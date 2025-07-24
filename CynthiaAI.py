import speech_recognition as sr
import os
import win32com.client
import webbrowser
import random
import datetime
import google.generativeai as genai
import time
import requests
import re
import pyautogui
import keyboard
import pygetwindow as gw
import pyttsx3
import pywhatkit as kit
import json
import sys
import wolframalpha
import math


###############################                                 Initialize chat string                                  ###############################

chatStr = "" 


###############################                            Function to convert text to speech                           ###############################


def say(text): 
   
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    time.sleep(0.5)
    speaker.Speak(text)
    time.sleep(0.2)


###############################                                 Function to greet                                  ###############################

def greet():
    hour = int(datetime.datetime.now().hour)
    if hour < 12:
        greeting = "Good Morning! sir" 
        say(greeting)
    elif hour < 18:
        greeting = "Good Afternoon! sir" 
        say(greeting)
    else:
        greeting = "Good Evening! sir" 
        say(greeting)   


#############################                           Function to take voice input from the user                       ###############################


def takeCommand():   
    r =sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=0.5)
        r.energy_threshold = 300
        r.phrase_time_limit = 5
        print("Listening for your response...")
        audio = r.listen(source)
        try:
            print("Recognizing....")
            query = r.recognize_google( audio , language="en-in")
            print(f"Input: {query}")
            time
            return query
        except Exception as e:
            print("Sorry, I didn't catch that. Please try again.")
            return ""




###############################                                 Function to get AI response                                 ###############################
def ai(prompt): 
    try:
        
        genai.configure(api_key="AIzaSyAtqGTPeXKZ689vY8TrBFoQTYsUJFaYXYo")

        text = f"Gemini response for Prompt: {prompt}\n*\n\n"

    
        model = genai.GenerativeModel('gemini-1.5-flash')  
        response = model.generate_content(prompt)

        text += response.text
        text = text.replace("", "").replace("*", "").replace("_", "")

       
        if not os.path.exists("Gemini"):
            os.mkdir("Gemini")

        file_name = f"Gemini/{prompt.replace(' ', '_')[:50]}.txt"
        with open(file_name, "w", encoding='utf-8') as f:
            f.write(text)

        print(f"Response saved to: {file_name}")

        

    except Exception as e:
        print("An error occurred:", e)


################################                                 Function to search Google                                 ###############################

def searchgoogle(query):
    try:
        spoken_query = query.strip()  
        encoded_query = '+'.join(spoken_query.split())  

        url = f"https://www.google.com/search?q={encoded_query}"
        webbrowser.open(url)

        print(f"Searching Google for: {spoken_query}")  
        say(f"Searching Google for {spoken_query}")     

    except Exception as e:
        print("Sorry, an error occurred:", e)
        say("Sorry, I couldn't search Google.")


################################                                 Function to play first video on youtube                                 ###############################

def playfirstvideo(query):
    try:
        search_query = '+'.join(query.strip().split())
        url = f"https://www.youtube.com/results?search_query={search_query}"
        response = requests.get(url)

        # Extract video URLs using regex (match video ID format)
        video_ids = re.findall(r"watch\?v=(\S{11})", response.text)

        if video_ids:
            video_url = f"https://www.youtube.com/watch?v={video_ids[0]}"
            webbrowser.open(video_url)
            print(f"Playing: {video_url}")
            say(f"Playing the top result for {query} on YouTube.")
        else:
            say("Sorry, I couldn't find any videos.")
            print("No video ID found.")

    except Exception as e:
        print("Error:", e)
        say("An error occurred while trying to play the video.")

################################                                 Function to search Youtube                                 ###############################
 
def searchyoutube(query):
    try:
        spoken_query = query.strip()  
        encoded_query = '+'.join(spoken_query.split())  

        url = f"https://www.youtube.com/results?search_query={encoded_query}"
        webbrowser.open(url)

        print(f"Searching Youtube for: {spoken_query}")  
        say(f"Searching Youtube for {spoken_query}")     
    except Exception as e:
        print("Sorry, an error occurred:", e)
        say("Sorry, I couldn't search Youtube.")


################################                                 Function to control youtube                                ###############################

def focus_youtube_tab():
    try:
        for window in gw.getWindowsWithTitle('YouTube'):
            if window.isActive == False:
                window.activate()
                time.sleep(0.5)
                return True
        say("YouTube window not found.")
        return False
    except Exception as e:
        print("Error focusing YouTube:", e)
        return False

def youtube_control(command):
    try:
        if command == "pause":
            pyautogui.press("k")
            say("Paused the video.")
        elif command == "play":
            pyautogui.press("k")
            say("Playing the video.")
        elif command == "next":
            pyautogui.hotkey("shift", "n")
            say("Playing the next video.")
        elif command == "previous":
            pyautogui.hotkey("shift", "p")
            say("Playing the previous video.")
        elif command == "mute":
            pyautogui.press("m")
            say("Muted the video.")
        elif command == "unmute":
            pyautogui.press("m")
            say("Unmuted the video.")
        elif command == "fullscreen":
            pyautogui.press("f")
            say("Entered fullscreen mode.")
        elif command == "exit fullscreen":
            pyautogui.press("f")
            say("Exited fullscreen mode.")
        elif command == "volume up":
            pyautogui.press("up")
            say("Increased the volume.")
        elif command == "volume down":
            pyautogui.press("down")
            say("Decreased the volume.")
        elif command == "skip":
            for i in range(10):
                pyautogui.press("right")
                time.sleep(0.1)
            say("Skipped the video.")
        elif command == "forward":
            pyautogui.press("right")
            say("Forwarded the video.")
        elif command == "rewind":
            pyautogui.press("left")
            say("Rewound the video.")
        elif command == "stop":
            pyautogui.press("k")
            say("Stopped the video.")
        elif command == "miniplayer":
            pyautogui.press("i")
            say("Entered miniplayer mode.")
        elif command == "close":
            pyautogui.hotkey("ctrl", "w")
            say("Closed the YouTube tab.")

        else:
            say("Invalid command for YouTube control.")
    except Exception as e:
        print("Error controlling YouTube:", e)


################################                                 Function for calculator                                       ###############################

def WolfRamAlpha(query):
    api_key = "2GTYW2-2EQPYKHVHJ"  # Replace with your API key
    requester = wolframalpha.Client(api_key)
    try:
        requested = requester.query(query)
        answer = next(requested.results).text
        return answer
    except:
        say("The value is not answerable")
        return None

def Calc(query):
    term = str(query).lower()
    spoken_term = term
    term = term.replace("jarvis", "")
    term = term.replace("multiply", "*")
    term = term.replace("plus", "+")
    term = term.replace("minus", "-")
    term = term.replace("divide", "/")
    term = term.replace("x", "*")
    final = term.strip()
    try:
        result = WolfRamAlpha(final)
        if result:
            print(f"Result: {result}")
            say(f"The result is {result}")
        else:
            say("The value is not answerable")
    except Exception as e:
        print("Calculation error:", e)
        say("Sorry, I couldn't calculate that.")



################################                                 Function to search Wikipedia                                 ###############################

def searchwikipedia(query):
    try:
        spoken_query = query.strip()  
        encoded_query = '+'.join(spoken_query.split())  

        url = f"https://en.wikipedia.org/wiki/{encoded_query}"
        webbrowser.open(url)

        print(f"Searching Wikipedia for: {spoken_query}")  
        say(f"Searching Wikipedia for {spoken_query}")     
    except Exception as e:
        print("Sorry, an error occurred:", e)
        say("Sorry, I couldn't search Wikipedia.")


################################                                 Function to get temperature                                 ###############################

def get_temperature(city):
    api_key = "9f205388470490d51c6aae00b7849d12"
    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    complete_url = base_url + "q=" + city + "&appid=" + api_key + "&units=metric"
    response = requests.get(complete_url)
    data = response.json()
    
    if data.get("main") and data.get("weather"):
        main = data["main"]
        temperature = main["temp"]
        humidity = main["humidity"]
        pressure = main["pressure"]
        weather_description = data["weather"][0]["description"]
        return temperature, humidity, pressure, weather_description
    else:
        return None, None, None, "Sorry, I couldn't retrieve the weather information."




        
################################                                 Function to chat with AI                                 ###############################



def chat(query):
    global chatStr
    print(chatStr)
    try:
        genai.configure(api_key="AIzaSyAtqGTPeXKZ689vY8TrBFoQTYsUJFaYXYo")
        chatStr += f"Input: {query}\nCynthia: "

        model = genai.GenerativeModel('gemini-1.5-flash')
        response = model.generate_content(query)

        response = response.text.strip()
        response = response.replace("", "").replace("*", "").replace("_", "")
        say(response)
        chatStr += response + "\n"

        return response 
    except Exception as e:
        print("An error occurred:", e)
        error_msg = "Sorry, I couldn't process that."
        say(error_msg)
        return error_msg



###############################                                 Main Function to run the AI                                 ###############################



if __name__ == '__main__':
    print("Your Personal AI Responsing")
    say(" Hi User I am your personal AI Cynthia , how may I help you")
    while True:
        print("Listening.....")
        query = takeCommand()
        # say(query)

################################                                 Open websites and apps                                 ###############################

        sites = [
        ["youtube", "https://www.youtube.com/"],
        ["wikipedia", "https://www.wikipedia.org/"],
        ["google", "https://www.google.com/"],
        ["facebook", "https://www.facebook.com/"],
        ["twitter", "https://www.twitter.com/"],
        ["instagram", "https://www.instagram.com/"],
        ["linkedin", "https://www.linkedin.com/"],
        ["github", "https://www.github.com/"],
        ["stackoverflow", "https://stackoverflow.com/"],
        ["reddit", "https://www.reddit.com/"],
        ["netflix", "https://www.netflix.com/"],
        ["amazon", "https://www.amazon.com/"],
        ["flipkart", "https://www.flipkart.com/"],
        ["coursera", "https://www.coursera.org/"],
        ["khan academy", "https://www.khanacademy.org/"],
        ["spotify", "https://www.spotify.com/"],
        ["gmail", "https://mail.google.com/"],
        ["drive", "https://drive.google.com/"]
        ]

        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
        
        apps = {
            "camera": r"C:\Users\asus\OneDrive\Desktop\Camera.lnk",
            "chrome": r"C:\Users\asus\OneDrive\Desktop\Adarsh - Chrome.lnk",
            "notepad": r"C:\Windows\system32\notepad.exe",
            "calculator": r"C:\Windows\System32\calc.exe",
            }
        # apps = {
        #     "camera": r"C:\Users\asus\OneDrive\Desktop\Camera.lnk",
        #     "chrome": r"C:\Users\asus\OneDrive\Desktop\Adarsh - Chrome.lnk",
        #     "notepad": r"C:\Windows\system32\notepad.exe",
        #     "calculator": r"C:\Windows\System32\calc.exe",
        # }

        process_names = {
            "camera": "WindowsCamera",
            "chrome": "chrome.exe",
            "notepad": "notepad.exe",
            "calculator": "Calculator.exe",  
        }

        for app in apps:
            if f"open {app}" in query.lower():
                say(f" Opening {app} ...")
                os.system(f'"{apps[app]}"')
                

            
            elif f"close {app.lower()}" in query.lower():
                process_name = os.path.basename(apps[app]) 
                say(f"Closing {app}...")
                os.system(f'taskkill /f /im "{process_name}"')
                break



################################                                      Play music                                 ###############################


        if "play music".lower() in query.lower():
            musicFolder = r"D:\Songs\songs mix\Hindi songs_1"
            musicFiles = [f for f in os.listdir(musicFolder) if f.endswith('.mp3')]

            if musicFiles:
                randomSong = random.choice(musicFiles)
                songPath = os.path.join(musicFolder, randomSong)
                print(f"Playing music...")
                os.startfile(songPath)
                say(f"Enjoy your music...")
            else:
                say("No mp3 files found in the folder.")


################################                                  The date and time                                 ###############################
     
     
        elif "the date and time".lower() in query.lower():  
            strfDateTime = datetime.datetime.now().strftime("%Y-%m-%d  %H:%M:%S")
            year = datetime.datetime.now().year
            month = datetime.datetime.now().month
            day = datetime.datetime.now().day
            hour = datetime.datetime.now().hour
            minute = datetime.datetime.now().minute
            second = datetime.datetime.now().second

            print(f"The Date and Time is {strfDateTime}")
            say(f"The Date is day {day} month {month} year {year} and the time is {hour} hours {minute} minutes and {second} seconds and today is {datetime.datetime.now().strftime('%A')}") 

        elif "the date".lower() in query.lower():
            strfDate = datetime.datetime.now().strftime("%Y-%m-%d  %D  %A")
            year = datetime.datetime.now().year
            month = datetime.datetime.now().month
            day = datetime.datetime.now().day
            weekday = datetime.datetime.now().strftime("%A")
            print(f"The Date is {strfDate}")
            say(f"The Date is day {day} month {month} year {year} and today is {weekday}")

        elif "the time".lower() in query.lower():
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            hour = datetime.datetime.now().hour
            minute = datetime.datetime.now().minute
            second = datetime.datetime.now().second
            print(f"The Time is {strfTime}")
            say(f"The Time is {hour} hours {minute} minutes and {second} seconds")
        
  


#################################                                 Ask with AI                                 ###############################
        elif "Panther".lower() in query.lower():
                ai(prompt=query)
                say("I have saved the response to a file")
                print("Response saved to a file")
        


#################################                                 Search with Google                                 ###############################
        elif "google" in query.lower():
            cleaned_query = query.lower().replace("searching google for:", "")
            cleaned_query = cleaned_query.replace("google", "").replace("for", "", 1).strip()
            searchgoogle(cleaned_query)

            

#################################                                 Search On Youtube                                ###############################
        elif "search youtube" in query.lower():
            cleaned_query = query.lower().replace("search youtube", "").replace("for ","",1).strip()
            searchyoutube(cleaned_query)


##################################                                 Play first video on youtube                                 ###############################
        elif "play video".lower() in query.lower():
            cleaned_query = query.lower().replace("play video", "").replace("for ","",1).strip()
            playfirstvideo(cleaned_query)


#################################                                 Control Youtube                                 ###############################

        elif "youtube" in query.lower():
            if "pause" in query:
                youtube_control("pause")
            elif "play" in query:
                youtube_control("play")
            elif "mute" in query:
                youtube_control("mute")
            elif "volume up" in query:
                youtube_control("volume up")
            elif "volume down" in query:
                youtube_control("volume down")
            elif "forward" in query or "skip" in query:
                youtube_control("forward")
            elif "rewind" in query:
                youtube_control("rewind")
            elif "fullscreen" in query:
                youtube_control("fullscreen")
            elif "miniplayer" in query:
                youtube_control("miniplayer")
            elif "exit fullscreen" in query:
                youtube_control("exit fullscreen")
            elif "next" in query:
                youtube_control("next")
            elif "previous" in query:
                youtube_control("previous")
            elif "stop" in query:
                youtube_control("stop")
            elif "close" in query:
                youtube_control("close")
            elif "unmute" in query:
                youtube_control("unmute")
            elif "play" in query:
                youtube_control("play")
            elif "skip" in query:
                youtube_control("skip")
            elif "rewind" in query:
                youtube_control("rewind")

            elif "search" in query:
                cleaned_query = query.lower().replace("search", "").replace("for ","",1).strip()
                searchyoutube(cleaned_query)
            # elif "open" in query:
            #     cleaned_query = query.lower().replace("open", "").replace("for ","",1).strip()
            #     searchyoutube(cleaned_query)

                
##################################                                 Search Wikipedia                                 ###############################

        elif "search wikipedia" in query.lower():
            cleaned_query = query.lower().replace("search wikipedia", "").replace("for ","",1).strip()
            searchwikipedia(cleaned_query)


##################################                                Greeting                                             ###############################

        elif "wake up".lower() in query.lower():
            greet()
            say("I am awake now, how may I help you")


#################################                                 Temperature                                 ###############################

        elif "temperature" in query.lower() or "weather" in query.lower():
            city = query.replace("temperature", "").replace("in", "").strip()
            temperature, humidity, pressure, weather_description = get_temperature(city)

            if temperature is None:
                say(weather_description)
            else:
                print(f"The temperature in {city} is {temperature}°C with {humidity}% humidity and {pressure} hPa pressure. Weather description: {weather_description}.")
                say(f"The temperature in {city} is {temperature}°C with {humidity}% humidity and {pressure} hPa pressure. Weather description: {weather_description}.")
                

#################################                                 Calculator                                 ###############################
        elif "calculate" in query.lower():
            cleaned_query = query.lower().replace("calculate", "").replace("jarvis", "").strip()
            Calc(cleaned_query)


#################################                                 Exit the program                                 ###############################
        elif "exit".lower() in query.lower():
            say("Goodbye my friend, have a nice day, see you soon")
            exit()


###############################                                 Reset the chat history                                 ###############################
        elif "reset".lower() in query.lower():
            say("Resetting the chat history")
            chatStr = ""


###############################                                   Shutdown                                               ##################################
        elif "shutdown system" in query.lower():
            say("Are you sure you want to shut down the computer? Please say yes or no.")
            confirmation = takeCommand().lower()
            print(f"Confirmation captured: {confirmation}")

            if "yes" in confirmation:
                say("Shutting down the system. Goodbye.")
                os.system("shutdown /s /t 1")
            elif "no" in confirmation:
                say("Shutdown cancelled.")
            else:
                say("I didn't understand your response. Shutdown aborted.")


#################################                                 Restart System                                 ###############################

        elif "restart system" in query.lower() or "restart computer" in query.lower():
            say("Are you sure you want to restart the computer? Please say yes or no.")
            time.sleep(1)
            confirmation = takeCommand().lower().strip()
            print(f"Confirmation captured: {confirmation}")

            if "yes" in confirmation:
                say("Restarting the system now. Please wait.")
                os.system("shutdown /r /t 1")
            elif "no" in confirmation:
                say("Restart cancelled.")
            elif confirmation == "":
                say("I didn't catch that. Please try again.")
            else:
                say("I didn't understand your response. Restart aborted.")


############################                                       Chat with AI                                 ###############################
        else:
            print("Chatting with AI...")
            chat(query)