import speech_recognition as sr
import os
import win32com.client
import webbrowser
import glob
import anthropic
import datetime


# Other initializations
r = sr.Recognizer()
speaker = win32com.client.Dispatch("SAPI.SpVoice")
 

def ai(prompt):
    text = f"Anthropic response for: {prompt}\n********************************************\n\n"
    messages = [
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": prompt
                }
            ]
        }
    ]
    
# the following block of code  is used to make Anthropic's AI respond via the enviorment variable 
# although it did not worked for me always gives me error of ANTHROPIC_API_KEY environment variable not set, 
#so I removed that part and hardcoded the api key
    '''api_key = os.environ.get('ANTHROPIC_API_KEY')
    if api_key:
        client = anthropic.Client(api_key=api_key)
    else:
        raise ValueError("ANTHROPIC_API_KEY environment variable not set")'''

    client = anthropic.Client(api_key='YPUR_API_KEY_HERE')
    response = client.messages.create(
        model="claude-2.1",
        max_tokens=1000,
        temperature=0,
        messages=messages
    )

    response_text = ''
    if isinstance(response.content, list):
        for item in response.content:
            response_text += item.text
    else:    
        response_text = response.content.decode('utf-8')

    print(response_text)
    text += response_text

    if not os.path.exists("Anthropic"):    
        os.makedirs("Anthropic")

    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"{prompt.replace(' ', '_')}_{timestamp}.txt"
    filepath = os.path.join("Anthropic", filename)

    with open(filepath, "w", encoding="utf-8") as f:    
        f.write(text)


def say(text):
    speaker.Speak(text)

def takeCommand():
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language="en-in")
        print(f"User said: {query}")
        return query
    except sr.UnknownValueError:
        print("Sorry, I could not understand that.")
        return ""
    except sr.RequestError as e:
        print(f"Speech recognition error: {e}")
        return ""

# Main function to run the program
if __name__ == "__main__":
    say("Hello! I am your personal assistant.")
    while True:
        print("Listening...")
        query = takeCommand()

        # AI usage
        if "using ai" in query.lower():
            ai(prompt=query)
            continue

        # For dealing with web related activities
        sites = [
            ["youtube", "https://youtube.com/"],
            ["wikipedia", "https://wikipedia.com"],
            ["google", "https://google.com"],
            ["my site", "https://thefuneducator.wordpress.com"]
        ]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} Sir...")
                webbrowser.open(site[1])

        # For handling the music player functionality (for downloaded songs only)
        if "play " in query:
            song_name = query.replace("play ", "").strip()
            playlists_path = "C:/Users/gamin/Music/Playlists" #change this path to yours 
            search_pattern = os.path.join(playlists_path, f"*{song_name}*.mp3")
            song_files = glob.glob(search_pattern)
            if song_files:
                song_to_play = song_files[0]
                say(f"Playing '{song_name}', Sir.")
                os.startfile(song_to_play)
            else:
                say(f"No song found for '{song_name}', Sir.")

        # For opening applications
        apps = [
            ["ms word", "C:/Users/gamin/Documents/PATHforAppsForJARVIS2/Word.lnk"],
            ["powerpoint", "C:/Users/gamin/Documents/PATHforAppsForJARVIS2/PowerPoint.lnk"],
            ["publisher", "C:/Users/gamin/Documents/PATHforAppsForJARVIS2/Publisher.lnk"],
            ["excel", "C:/Users/gamin/Documents/PATHforAppsForJARVIS2/Excel.lnk"],
            ["onenote", "C:/Users/gamin/Documents/PATHforAppsForJARVIS2/OneNote.lnk"],
            ["my computer", "C:/Users/gamin/Documents/PATHforAppsForJARVIS2/ThisPC.lnk"]
        ]
        for app in apps:
            if f"Open {app[0]}".lower() in query.lower():
                say(f"Opening {app[0]} Sir...")
                os.startfile(app[1])
        if 'exit' in query:
            say("exiting")
            break
