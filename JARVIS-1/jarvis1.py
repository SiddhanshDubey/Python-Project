import pyttsx3
import speech_recognition as sr
from selenium import webdriver

# Initialize text-to-speech engine
engine = pyttsx3.init()

# Set properties
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)  # You can change index to select different voice
engine.setProperty('rate', 150)  # Speed of speech

# Function to speak
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Function to listen to voice command
def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.pause_threshold = 0.5
        audio = recognizer.listen(source)
    try:
        print("Recognizing...")
        query = recognizer.recognize_google(audio, language='en-us')
        print(f"User said: {query}\n")
    except Exception as e:
        print("Say that again please...")
        return "None"
    return query.lower()

# Main program
if __name__ == "__main__":
    speak("Hello! What can I do for you?")
    driver = webdriver.Chrome()  # Initialize WebDriver instance
    while True:
        query = listen().lower()
        if 'open google' in query or 'open chrome' in query or 'open google chrome' in query:
            speak("Opening Google Chrome")
            driver.get('https://www.google.com')
        elif 'open youtube' in query or 'open yt' in query:
            speak("Opening YouTube")
            driver.get('https://www.youtube.com/')
        elif 'open github' in query:
            speak("Opening Git Hub")
            driver.get('https://https://github.com/')
        elif 'open my page' in query:
            speak("Opening your Website")
            driver.get('https://thefuneducator.wordpress.com/')
        elif 'exit' in query:
            speak("exiting")
            driver.quit()  # Quit the browser session
            break