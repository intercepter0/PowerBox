import os
from pyowm import OWM
import speech_recognition as sr
from fuzzywuzzy import fuzz
import pyttsx3
import datetime
import webbrowser
import random
import pyperclip
import getpass
import ui_hook
import threading
import keyboard


opts = {
    "alias": ('слушай', 'эй'),
    "tbr": ('скажи', 'расскажи', 'покажи', 'произнеси'),
    "cmds": {
        "ctime": ('текущее время', 'сейчас времени', 'который час', 'сколько время'),
        "weather": ('какая погода', 'погода сегодня'),
        "count": ('сколько будет', 'посчитай'),
        "write": ('напечатай', 'напиши', 'напечатать', 'написать'),
        "lock": ('заблокировать', 'выйти из сеанса'),
        "copy": ('скопировать', 'скопируй'),
        "paste": ('вставить', 'вставь'),
        "line": ('открыть строку', 'открыть поиск'),
        "settings": ('открыть настройки', 'настройки'),
        "shutdown": ('завершение работы', 'выключить компьютер'),
        "shutdown_cancel": ('отмена', 'не выключать компьютер'),
        "restart": ('перезагрузить', 'перезагрузить компьютер', 'перезагрузить компьютер', 'перезагрузка компьютера'),
        "thanks": ('благодарю', 'спасибо', 'классно'),
        "browser_search": ('искать в браузере', 'искать в интернете', 'ищи в браузере', 'ищи в интернете', 'поиск в интернете')
    }
}
ambient_noise_adjusted = False


# Functions
def listen():
    global audio
    global ambient_noise_adjusted
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.2)
        print('Say something...')
        r.pause_threshold = 0.5
        r.non_speaking_duration = 0.5
        audio = r.listen(source)
        callback(r, audio)


def speak(message):
    print("[log]: Speak: ", message)
    speak_engine.say(message)
    speak_engine.runAndWait()
    ui_hook.append_log("PowerBox >> " + str(message))
    speak_engine.stop()


def callback(recognizer, audio):
    try:
        voice = recognizer.recognize_google(audio, language="ru-RU").lower()
        ui_hook.append_log(getpass.getuser()+((8-len(list(getpass.getuser())))*' ')+" >> "+voice.capitalize())
        print("[log] Распознано: "+voice)

        cmd = voice

        for x in opts['tbr']:
            cmd = cmd.replace(x, "").strip()

        cmd = recognize_cmd(cmd)
        if cmd['percent'] > 70:
            execute_cmd(cmd['cmd'], voice)
            print("Execute finished!")
        else:
            speak("Простите, я вас не понимаю.")

    except sr.UnknownValueError:
        print("[log] Голос не распознан!")
    except sr.RequestError as e:
        print("[log] Ошибка сети")
    print("Starting listen...")
    listen()


def recognize_cmd(cmd):
    RC = {'cmd': '', 'percent': 0}
    for c, v in opts['cmds'].items():
        for x in v:
            vrt = fuzz.ratio(cmd, x)
            if vrt > RC['percent']:
                RC['cmd'] = c
                RC['percent'] = vrt
            if x in cmd:
                RC['cmd'] = c
                RC['percent'] = 95
                print(cmd," contains ",x,"!")
    print("Ratio: ", RC['percent'])
    return RC


def execute_cmd(cmd, parameter):
    print("Exec: ", cmd)
    if cmd == 'ctime':
        # Сказать текущее время
        now = datetime.datetime.now()
        speak('Сейчас ' + str(now.hour) + ':' + (str(now.minute) if len(str(now.minute)) > 1 else '0'+str(now.minute)))
    elif cmd == 'shutdown':
        os.system('shutdown -s')
        speak('Выключаю. Вы можете отменить это действие.')

    elif cmd == 'shutdown_cancel':
        os.system('shutdown -a')
        speak('Отменяю завершение работы.')

    elif cmd == 'restart':
        os.system('shutdown -r')
        speak('Перезагружаю. Вы можете отменить это действие.')

    elif cmd == 'lock':
        os.system('rundll32 user32.dll LockWorkStation')

    elif cmd == 'copy':
        pass

    elif cmd == 'paste':
        pyperclip.paste()

    elif cmd == 'weather':
        weather = dict()

        # Init pyowm lib
        owm = OWM('e3b69afc0e2c681527fee196cf54144d', language="RU")
        obs = owm.weather_at_place('Севастополь')
        w = obs.get_weather()

        weather['temp'] = w.get_temperature(unit='celsius')['temp']
        weather['status'] = w.get_detailed_status()
        weather['pressure'] = (str(round(w.get_pressure()['press']*0.750063755419211)))  # гПа -> мм.р.с
        p1 = [0, 5, 6, 7, 8, 9]
        if int(weather['pressure']) % 10 in p1:
            weather['pressure'] = str(weather['pressure']).__add__(" миллиметров")
        elif int(weather['pressure']) % 10 == 1:
            weather['pressure'] = str(weather['pressure']).__add__(" миллиметр")
        else:
            weather['pressure'] = str(weather['pressure']).__add__(" миллиметрa")
            weather['pressure'] = weather['pressure'].__add__(" ртутного столба")

        weather['humidity'] = str(w.get_humidity()) + "%"
        speak("В Вашем городе сейчас {1} , температура: {0}°. Давление: {2}, влажность: {3}".format(weather['temp'],
                weather['status'], weather['pressure'], weather['humidity']))

    elif cmd == 'write':
        w_request = parameter
        for i in opts['cmds'].get("write"):
            w_request = w_request.replace(i, '')
        keyboard.write(w_request.strip())

    elif cmd == 'line':
        jar_path = os.path.realpath(__file__).replace("core.py", "") + "Powerline.jar"
        os.system('java -jar "'+jar_path+'"')

    elif cmd == 'settings':
        ui_hook.init()

    elif cmd == 'count':
        removed = ('сколько будет', 'посчитай')
        math_request = parameter.replace(removed[0], '').replace(removed[1], '').replace('разделить на', '/').replace('умножить на', '*').replace('в квадрате', '** 2').replace('в кубе', '** 3').replace('в степени', '**').replace('плюс', '+').replace('минус', '-').replace("х", "*").replace('два', '2').replace('три', '3').replace('четыре', '4').strip()
        try:
            speak(str(round(eval(math_request), 4)))
        except Exception as e:
            speak("Простите, я не могу это посчитать.")
            print(math_request, " ", e)

    elif cmd == 'browser_search':
        ress = list(parameter.split(' '))
        not_request = ('искать', 'поиск', 'интернете', 'браузере')
        for i in ress:
            for j in not_request:
                if fuzz.ratio(i, j) > 60:
                    for h in range(0, ress.index(i)+1):
                        ress[h] = ''
        web_request = "https://yandex.ru/search/?text="
        for l in ress:
            web_request += l+"%20"
        webbrowser.open(web_request)

    elif cmd == 'thanks':
        speak('И Вам спасибо')


# Run
r = sr.Recognizer()
mic = sr.Microphone(device_index=1)

with mic as source:
    r.adjust_for_ambient_noise(source)

speak_engine = pyttsx3.init()
ui_hook.pre_init()
speak("Приветствую. Я Вас слушаю")

listen()
event = threading.Event()
#ui_thread = threading.Thread(target=listen)
#ui_thread.start()

