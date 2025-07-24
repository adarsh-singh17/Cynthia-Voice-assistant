def say(text, lang=None): 
   
    # speaker = win32com.client.Dispatch("SAPI.SpVoice")
    # time.sleep(0.5)
    # speaker.Speak(text)
    # time.sleep(0.2)
    
    try:
        # Detect language from text if not provided
        if lang is None:
            lang = detect(text)

        print(f"Detected language: {lang}")

        # Generate speech
        tts = gTTS(text=text, lang=lang)

        # Save and play using pygame
        with tempfile.NamedTemporaryFile(delete=True) as fp:
            tts.save(fp.name + ".mp3")
            pygame.init()
            pygame.mixer.init()
            pygame.mixer.music.load(fp.name + ".mp3")
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy():
                continue
            pygame.quit()
    except Exception as e:
        print("TTS Error:", e)