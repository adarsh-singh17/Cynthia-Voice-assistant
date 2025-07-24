import threading
import os
import win32com.client
import random
import datetime
import google.generativeai as genai
import speech_recognition as sr
from queue import Queue

chatStr = ""  # Global variable to keep track of the conversation

# This function makes the AI speak the response
def say(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

# Function to listen to the user through the microphone
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing....")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Sorry, I couldn't understand that. Please try again."

# Function to interact with the chatbot and generate responses
def chat(query):
    global chatStr
    print(f"Chat history:\n{chatStr}")

    try:
        # Initialize Google AI API configuration (use the actual API key here)
        genai.configure(api_key="AIzaSyAtqGTPeXKZ689vY8TrBFoQTYsUJFaYXYo")  # Replace with your API key

        # Update chat history with the user's query
        chatStr += f"User: {query}\nCynthia: "

        model = genai.GenerativeModel('gemini-1.5-flash')  # Using Gemini model
        response = model.generate_content(query)

        # Clean up the response by removing unwanted symbols like **, *, and _
        response_text = response.text.strip()
        response_text = response_text.replace("**", "").replace("*", "").replace("_", "")

        # Send the response to the main thread for speaking
        response_queue.put(response_text)

        chatStr += response_text + "\n"  # Add the response to chat history

        return response_text

    except Exception as e:
        print("An error occurred:", e)
        error_msg = "Sorry, I couldn't process that."
        response_queue.put(error_msg)
        return error_msg

# Function to continuously listen for commands and process them
def listen_for_commands():
    while True:
        query = takeCommand()
        if query.lower() == "stop":
            print("Stop command received.")
            say("Stopping now, listening for your next command.")
            break  # Stop the loop when "stop" is heard
        elif query.lower() == "quit":
            say("Goodbye, have a nice day!")
            exit()  # Exit the program when "quit" is said
        else:
            chat(query)  # Call chat function for further responses

# Function to handle speaking in the main thread
def handle_speaking():
    while True:
        text_to_speak = response_queue.get()  # Get the response text from the queue
        if text_to_speak:
            say(text_to_speak)

# Entry point to run the chatbot
if __name__ == "__main__":
    print("Your Personal AI is Ready to Chat!")

    # Create a queue to handle communication between threads
    response_queue = Queue()

    # Start the speaking thread
    threading.Thread(target=handle_speaking, daemon=True).start()

    # Greeting message to the user
    say("Hi User, I am your personal AI Cynthia. How may I help you?")

    while True:
        # Start listening for user's input
        print("Listening.....")
        query = takeCommand()

        # If the user asks to quit, the program will exit
        if "quit" in query.lower():
            say("Goodbye, have a nice day!")
            break

        # If the user asks to reset the conversation history
        elif "reset" in query.lower():
            say("Resetting the chat history.")
    
            chatStr = ""  # Clear chat history

        # Chat with the AI if it's neither quit nor reset
        else:
            print("Chatting....")
            chat(query)
