import pyttsx3

engine = pyttsx3.init()

# Получаем все доступные голоса
voices = engine.getProperty('voices')
print("Доступные голоса:")
for i, voice in enumerate(voices):
    print(f"{i}: {voice.name} ({voice.id})")
    print(f"   Язык: {voice.languages}")

# Пробуем разные голоса
engine.say("Привет, это тест звука")
engine.runAndWait()
engine.stop()