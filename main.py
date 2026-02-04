import os 
import time 
import speech_recognition as sr 
from fuzzywuzzy import fuzz
import datetime 
import win32com.client
import webbrowser
import random
import urllib.parse

comands = {
    "alias": ('орион',),
    "tbr": ('скажи', 'расскажи', 'покажи', 'включи', 'открой', 'найди'),
    "cmds": {
        "ctime": ('который час', 'текущее время', 'сейчас время', 'сколько времени'),
        "date": ('какое сегодня число', 'какая дата', 'текущая дата'),
        "open_browser": ('открой браузер', 'запусти браузер', 'интернет'),
        "open_calculator": ('открой калькулятор', 'калькулятор'),
        "search": ('найди в интернете', 'поищи', 'гугл', 'найти информацию', 'поиск'),
        "weather": ('какая погода', 'прогноз погоды', 'погода на улице'),
        "system_info": ('системная информация', 'информация о системе', 'характеристики'),
        "shutdown": ('выключи компьютер', 'заверши работу', 'выключение'),
        "note": ('запиши заметку', 'сделай запись', 'запомни'),
    }
}

#
jokes = [
    "Почему программисты так плохо танцуют? У них нет алгоритма!",
    "Как называют программиста, который боится женщин? Гитхаб!",
    "Почему Python стал таким популярным? Потому что у него есть змеиный шарм!",
    "Что сказал один бит другому? Давай встретимся на байте!",
    "Почему компьютер пошел к врачу? У него был вирус!",
]

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def speak(what):
    print(f"[Орион]: {what}")
    speaker.Speak(what)

def recognize_cmd(cmd):
    RC = {'cmd': '', 'percent': 0} 
    for c, v in comands['cmds'].items():
        for x in v: 
            vrt = fuzz.ratio(cmd, x)
            if vrt > RC['percent']:
                RC['cmd'] = c 
                RC['percent'] = vrt 
    return RC 

def extract_search_query(text):
    """Извлекает поисковый запрос из текста"""
    remove_words = ['орион', 'скажи', 'расскажи', 'покажи', 'найди', 'поищи', 
                   'в интернете', 'информацию', 'про', 'гугл', 'поиск', 'найти']
    
    for word in remove_words:
        text = text.replace(word, '')
    
    text = ' '.join(text.split())
    return text.strip()

def execute_cmd(cmd, search_query=""):
    """Выполняет команду"""
    if cmd == 'ctime':
        now = datetime.datetime.now()
        speak(f"Сейчас {now.hour} часов {now.minute} минут")
        
    elif cmd == 'date':
        now = datetime.datetime.now()
        months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
                 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
        speak(f"Сегодня {now.day} {months[now.month-1]} {now.year} года")
        
    elif cmd == 'joke':
        joke = random.choice(jokes)
        speak(joke)
        
    elif cmd == 'open_browser':
        speak("Открываю браузер")
        webbrowser.open("https://www.google.com")
        
    elif cmd == 'open_calculator':
        speak("Открываю калькулятор")
        os.system('calc.exe')
        
    elif cmd == 'search':
        if search_query:
            clean_query = extract_search_query(search_query)
            if clean_query:
                speak(f"Ищу в интернете: {clean_query}")
                encoded_query = urllib.parse.quote(clean_query)
                webbrowser.open(f"https://www.google.com/search?q={encoded_query}")
            else:
                speak("Что именно нужно найти? Повторите запрос.")
        else:
            speak("Что нужно найти?")
            
    elif cmd == 'weather':
        speak("Открываю погоду")
        webbrowser.open("https://yandex.ru/pogoda")
        
    elif cmd == 'system_info':
        import platform
        info = f"Система: {platform.system()} {platform.release()}. "
        info += f"Процессор: {platform.processor()[:50]}..."
        speak(info)
        
    elif cmd == 'shutdown':
        speak("Для выключения компьютера скажите: да, выключить")
        
    elif cmd == 'note':
        speak("Что записать? У вас 15 секунд.")
        try:
            with sr.Microphone() as mic:
                
                audio = r.listen(mic, timeout=15, phrase_time_limit=15)
                note_text = r.recognize_google(audio, language="ru-RU")
                with open("заметки.txt", "a", encoding="utf-8") as f:
                    f.write(f"{datetime.datetime.now()}: {note_text}\n")
                speak("Заметка сохранена")
        except Exception as e:
            speak("Не удалось записать заметку")
            print(f"[log] Ошибка записи заметки: {e}")
        
    else:
        speak("Команда не распознана")

def listen_longer(recognizer, microphone, timeout=10, phrase_time=10):
    """
    Слушает микрофон дольше с настройками timeout и phrase_time
    timeout - максимальное время ожидания начала речи (сек)
    phrase_time - максимальная длина фразы (сек)
    """
    try:
        print(f"[log] Слушаю {timeout} секунд...")
        audio = recognizer.listen(
            microphone, 
            timeout=timeout, 
            phrase_time_limit=phrase_time
        )
        return audio
    except sr.WaitTimeoutError:
        print("[log] Время ожидания истекло, не слышно речи")
        return None
    except Exception as e:
        print(f"[log] Ошибка прослушивания: {e}")
        return None

def main():
    r = sr.Recognizer()
    
#Индекс твоего микрофона    
    try:
        m = sr.Microphone(device_index=3)
    except:
        m = sr.Microphone()
    
    print("=" * 60)
    print("Голосовой помощник Орион")
    print("=" * 60)
    print("НАСТРОЙКИ ПРОСЛУШИВАНИЯ:")
    print("- Ожидание речи: 15 секунд")
    print("- Максимальная длина фразы: 20 секунд")
    print("- Пауза между словами: 1 секунда")
    print("=" * 60)
    
    
    r.pause_threshold = 1.0  
    r.non_speaking_duration = 0.5  
    
    # Калибровка
    with m as source:
        print("[log] Калибровка микрофона... (помолчите 2 секунды)")
        r.adjust_for_ambient_noise(source, duration=2)
        print("[log] Калибровка завершена")
    
    speak("Добрый день! Я голосовой помощник Орион.")

    
    
    while True:
        try:
            print("\n[log] Ожидаю команду... (до 15 секунд)")
            
            with m as source:
                
                audio = r.listen(
                    source, 
                    timeout=15,  
                    phrase_time_limit=20  
                )
            
            if audio:
                print("[log] Речь обнаружена, распознаю...")
                voice = r.recognize_google(audio, language="ru-Ru").lower()
                print(f'[log] Распознано: {voice}')
                
                
                if len(voice) < 3:
                    print('[log] Слишком короткая фраза')
                    continue
                
                
                exit_words = ['выход', 'стоп', 'заверши', 'до свидания', 'пока', 'хватит']
                if any(word in voice for word in exit_words):
                    speak("До свидания! Рад был помочь.")
                    break
                
                
                if 'помощь' in voice or 'команды' in voice or 'что ты умеешь' in voice:
                    speak("Я умею: говорить время и дату, рассказывать шутки, открывать браузер и калькулятор.")
                    speak("Могу искать в интернете, показывать погоду, записывать заметки.")
                    speak("Просто скажите: Орион, и нужную команду.")
                    continue
                
                
                starts_with_alias = any(voice.startswith(alias) for alias in comands["alias"])
                
                if starts_with_alias:
                    original_voice = voice
                    cmd_text = voice
                    
                    
                    for alias in comands["alias"]:
                        if voice.startswith(alias):
                            cmd_text = cmd_text.replace(alias, "", 1).strip()
                            break
                    
                    
                    for trigger in comands["tbr"]:
                        if cmd_text.startswith(trigger):
                            cmd_text = cmd_text.replace(trigger, "", 1).strip()
                    
                    print(f'[log] Команда после обработки: "{cmd_text}"')
                    
                    if cmd_text:
                        recognized_cmd = recognize_cmd(cmd_text)
                        print(f'[log] Найдена команда: {recognized_cmd["cmd"]} (уверенность: {recognized_cmd["percent"]}%)')
                        
                        if recognized_cmd['percent'] > 40:  
                            if recognized_cmd['cmd'] == 'search':
                                execute_cmd(recognized_cmd['cmd'], original_voice)
                            elif recognized_cmd['cmd'] == 'note':
                                execute_cmd(recognized_cmd['cmd'])
                            else:
                                execute_cmd(recognized_cmd['cmd'])
                        else:
                            print(f'[log] Команда не распознана')
                            
                else:
                    
                    recognized_cmd = recognize_cmd(voice)
                    if recognized_cmd['percent'] > 60:
                        if recognized_cmd['cmd'] == 'search':
                            execute_cmd(recognized_cmd['cmd'], voice)
                        else:
                            execute_cmd(recognized_cmd['cmd'])
                    else:
                        print('[log] Не начинается с ключевого слова')
                        
            
            else:
                print("[log] Речь не обнаружена")
                
        except sr.UnknownValueError:
            print('[log] Речь не распознана')
            print("Не расслышал, повторите пожалуйста")
        except sr.WaitTimeoutError:
            print('[log] Время ожидания истекло')
            
            continue
        except sr.RequestError as e:
            print(f'[log] Ошибка сервиса распознавания: {e}')
            speak("Проблема с интернет-соединением")
        except KeyboardInterrupt:
            print("\n[log] Выход по запросу пользователя")
            speak("До свидания!")
            break
        except Exception as e:
            print(f'[log] Неожиданная ошибка: {e}')
            

if __name__ == "__main__":
    main()