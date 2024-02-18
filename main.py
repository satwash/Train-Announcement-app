import pandas as pd
import pyttsx3
import time
import pygame
# ----------------------------------------------------------------------------------------------
# Initialize text-to-speech engine
engine = pyttsx3.init()
voices = engine.getProperty('voices')
selected_voice = voices[1]  # Index 1 corresponds to the second voice
# Set selected voice
engine.setProperty('voice', selected_voice.id)
engine.setProperty('rate', 90)

# Initialize Pygame for audio playback
pygame.init()
# Read data from Excel spreadsheet
schedule = pd.read_excel('train_schedule.xlsx')
# ----------------------------------------------------------------------------------------------
# Define audio file
chime = '1_chime.mp3'
tada = '2_tada.mp3'
attention = 'attention.mp3'
arriving = 'arriving.mp3'
departing = 'depart.mp3'
safe = 'safe.mp3'
melody_file = 'speech.mp3'

# Iterate over rows and generate announcement for each train
for index, row in schedule.iterrows():
    train_status = row['Train Status']
    train_number = row['Train Number']
    train_name = row['Train Name']
    platform_number = row['Platform Number']

    if f"{train_status}" == "Departure":

        pygame.mixer.music.load(chime)
        pygame.mixer.music.play()
        chime_length = pygame.mixer.Sound(chime).get_length()
        time.sleep(chime_length)
        # ----------------------------------------------------------------------------------------------
        pygame.mixer.music.load(tada)
        pygame.mixer.music.play()
        tada_length = pygame.mixer.Sound(tada).get_length()
        time.sleep(tada_length)
        # ----------------------------------------------------------------------------------------------
        pygame.mixer.music.load(attention)
        pygame.mixer.music.play()
        attention_length = pygame.mixer.Sound(attention).get_length()
        time.sleep(attention_length)
        # ----------------------------------------------------------------------------------------------

        # Generate announcement text
        announcement_text = f"  {train_number}    {train_name}"
        engine.say(announcement_text)
        engine.runAndWait()
        # ----------------------------------------------------------------------------------------------
        pygame.mixer.music.load(departing)
        pygame.mixer.music.play()
        departing_length = pygame.mixer.Sound(departing).get_length()
        time.sleep(departing_length)
        # ----------------------------------------------------------------------------------------------
        PF = f" {platform_number}"
        engine.say(PF)
        engine.runAndWait()

        pygame.mixer.music.load(safe)
        pygame.mixer.music.play()
        safe_length = pygame.mixer.Sound(safe).get_length()
        time.sleep(safe_length)
        continue
# ----------------------------------------------------------------------------------------------

    if f"{train_status}"== "Arrival":

        pygame.mixer.music.load(chime)
        pygame.mixer.music.play()
        chime_length = pygame.mixer.Sound(chime).get_length()
        time.sleep(chime_length)
    # ----------------------------------------------------------------------------------------------
        pygame.mixer.music.load(tada)
        pygame.mixer.music.play()
        tada_length = pygame.mixer.Sound(tada).get_length()
        time.sleep(tada_length)
    # ----------------------------------------------------------------------------------------------
        pygame.mixer.music.load(attention)
        pygame.mixer.music.play()
        attention_length = pygame.mixer.Sound(attention).get_length()
        time.sleep(attention_length)
    # ----------------------------------------------------------------------------------------------

        # Generate announcement text
        announcement_text = f"  {train_number}    {train_name}"
        engine.say(announcement_text)
        engine.runAndWait()
    # ----------------------------------------------------------------------------------------------
        pygame.mixer.music.load(arriving)
        pygame.mixer.music.play()
        arriving_length = pygame.mixer.Sound(arriving).get_length()
        time.sleep(arriving_length)
    # ----------------------------------------------------------------------------------------------
        PF = f" {platform_number}"
        engine.say(PF)
        engine.runAndWait()
        continue

# ----------------------------------------------------------------------------------------------
pygame.mixer.quit()
