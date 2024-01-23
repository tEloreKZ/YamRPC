import asyncio
import psutil
import pypresence
import time
import ctypes
import win32gui
from enum import Enum
from winsdk.windows.media.control import GlobalSystemMediaTransportControlsSessionManager as MediaManager
from yandex_music import Client
from itertools import permutations
from PIL import Image as PILImage
import threading
from pystray import MenuItem as item
from pystray import Icon as icon
from pystray import Menu as menu
import sys

# Идентификатор клиента Discord для Rich Presence
client_id = '1199408232771366962'

# Флаг для поиска трека с 100% совпадением названия и автора. Иначе будет найден близкий результат.
strong_find = True

# Переменная для хранения предыдущего трека и избежания дублирования обновлений.
name_prev = str()

# Enum для статуса воспроизведения мультимедийного контента.
class PlaybackStatus(Enum):
    Unknown = 0, 1
    Opened = 2
    Paused = 3
    Playing = 4
    Stopped = 5

# Асинхронная функция для получения информации о мультимедийном контенте через Windows SDK.
async def get_media_info():
    sessions = await MediaManager.request_async()
    current_session = sessions.get_current_session()
    if current_session:
        info = await current_session.try_get_media_properties_async()
        info_dict = {song_attr: info.__getattribute__(song_attr) for song_attr in dir(info) if
                     song_attr[0] != '_'}
        info_dict['genres'] = list(info_dict['genres'])
        playback_status = PlaybackStatus(current_session.get_playback_info().playback_status)
        info_dict['playback_status'] = playback_status.name
        return info_dict
    raise Exception('Музыка сейчас не играет.')

# Класс для работы с Rich Presence в Discord.
class Presence:
    def __init__(self) -> None:
        self.client = None
        self.currentTrack = None
        self.rpc = None
        self.running = False
        self.paused = False
        self.pause_start_time = None  # Добавляем переменную для отслеживания времени начала паузы
        self.rpc_hidden = False  # Добавляем переменную для отслеживания статуса скрытого RPC
        self.last_pause_notification_time = None  # Добавляем переменную для отслеживания времени последнего уведомления о длительной паузе

    def start(self):
        # Скрываем окно
        ctypes.windll.user32.ShowWindow(win32gui.GetForegroundWindow(), 0)  # 0 соответствует SW_HIDE
        
        # Запуск обновления Presence в отдельном потоке
        presence_thread = threading.Thread(target=self.start_presence)
        presence_thread.start()

        # Запуск значка в трее в главном потоке
        self.create_system_tray_icon()
        
        presence_thread.join()  # Ожидаем завершение потока обновления Presence после закрытия значка в трее

    def exit_application(self, icon, item):
        icon.stop()
        self.running = False
        sys.exit(0)  # Полностью завершить приложение

    def start_presence(self):
        if "Discord.exe" not in (p.name() for p in psutil.process_iter()):
            print("[YamRPC] -> Discord не запущен")
            self.exit_application(None, None)
            return

        self.rpc = pypresence.Presence(client_id)
        self.rpc.connect()
        self.client = Client().init()
        self.running = True
        self.currentTrack = None

        while self.running:
            currentTime = time.time()

            if "Discord.exe" not in (p.name() for p in psutil.process_iter()):
                print("[YamRPC] -> Discord закрыт")
                self.exit_application(None, None)
                return

            ongoing_track = self.get_track()

            if self.currentTrack != ongoing_track:
                if ongoing_track['success']:
                    if self.currentTrack is not None and 'label' in self.currentTrack and self.currentTrack[
                        'label'] is not None:
                        if ongoing_track['label'] != self.currentTrack['label']:
                            print(f"[YamRPC] -> Изменил трек на {ongoing_track['label']}")
                            self.rpc_hidden = False  # Сброс статуса скрытого RPC при смене трека
                    else:
                        print(f"[YamRPC] -> Изменил трек на {ongoing_track['label']}")
                        self.rpc_hidden = False  # Сброс статуса скрытого RPC при смене трека

                    trackTime = currentTime
                    remainingTime = ongoing_track['durationSec'] - 2 - (currentTime - trackTime)
                    self.rpc.update(
                        details=ongoing_track['label'],
                        end=currentTime + remainingTime,
                        large_image=ongoing_track['og-image'],
                        large_text='Яндекс Музыка',
                        buttons=[{'label': 'Слушать', 'url': ongoing_track['link']}]
                    )
                else:
                    if not self.rpc_hidden:
                        self.rpc.clear()
                        print("[YamRPC] -> Чистим RPC")
                        self.rpc_hidden = True  # Установка статуса скрытого RPC при отсутствии трека

                self.currentTrack = ongoing_track

            else:
                if ongoing_track['success'] and ongoing_track["playback"] != PlaybackStatus.Playing.name and not self.paused:
                    self.paused = True
                    self.pause_start_time = currentTime  # Начало паузы

                    print(f"[YamRPC] -> Трек {ongoing_track['label']} на паузе")

                    if ongoing_track['success']:
                        trackTime = currentTime
                        remainingTime = ongoing_track['durationSec'] - 2 - (currentTime - trackTime)
                        self.rpc.update(
                            details=ongoing_track['label'],
                            state="На паузе",
                            large_image=ongoing_track['og-image'],
                            large_text='Яндекс Музыка',
                            buttons=[{'label': 'Слушать', 'url': ongoing_track['link']}]
                        )

                elif ongoing_track['success'] and ongoing_track["playback"] == PlaybackStatus.Playing.name and self.paused:
                    print(f"[YamRPC] -> Трек {ongoing_track['label']} на паузе.")
                    self.paused = False
                    self.pause_start_time = None  # Очистка времени начала паузы
                    self.rpc_hidden = False  # Сброс статуса скрытого RPC при воспроизведении трека

            # Проверка на длительную паузу
            if self.paused and self.pause_start_time and currentTime - self.pause_start_time > 300:
                if not self.rpc_hidden and (self.last_pause_notification_time is None or
                                             currentTime - self.last_pause_notification_time > 300):
                    print("[YamRPC] -> Трек на паузе больше 5 минут. Чистим RPC.")
                    self.rpc.clear()
                    self.rpc_hidden = True  # Установка статуса скрытого RPC при длительной паузе
                    self.last_pause_notification_time = currentTime

            # Проверка на воспроизведение после длительной паузы
            if not self.paused and self.rpc_hidden:
                print("[YamRPC] -> Возобновление воспроизведения. Обновляем RPC.")
                self.rpc_hidden = False
                if ongoing_track['success']:
                    trackTime = currentTime
                    remainingTime = ongoing_track['durationSec'] - 2 - (currentTime - trackTime)
                    self.rpc.update(
                        details=ongoing_track['label'],
                        end=currentTime + remainingTime,
                        large_image=ongoing_track['og-image'],
                        large_text='Яндекс Музыка',
                        buttons=[{'label': 'Слушать', 'url': ongoing_track['link']}]
                    )

            time.sleep(3)

    def create_system_tray_icon(self):
        # Создание значка в трее
        image_path = "img/ym.jpg"  # Укажите путь к вашей иконке
        image_icon = PILImage.open(image_path)

        menu_exit = item('Выход', self.exit_application)
        tray_menu = menu(menu_exit)
        tray_icon = icon("name", image_icon, menu=tray_menu)
    
        # Запуск значка в трее
        tray_icon.run()
        
    def exit_application(self, icon, item):
        icon.stop()
        self.running = False
        sys.exit(0)  # Полностью завершить приложение

    def get_track(self) -> dict:
        try:
            current_media_info = asyncio.run(get_media_info())
            name_current = current_media_info["artist"] + " - " + current_media_info["title"]
            global name_prev
            global strong_find
            if str(name_current) != name_prev:
                print("[YamRPC] -> Сейчас слушаем: " + name_current)
            else:
                current_track_copy = self.currentTrack.copy()
                current_track_copy["playback"] = current_media_info['playback_status']
                return current_track_copy

            name_prev = str(name_current)
            search = self.client.search(name_current, True, "all", 0, False)

            if not search.best:
                print(f"[YamRPC] -> Не можем найти песню: {name_current}")
                return {'success': False}
            if search.best.type not in ['music', 'track', 'podcast_episode']:
                print(
                    f"[YamRPC] -> Не можем найти песню: {name_current}, наилучший результат имеет неправильный тип")
                return {'success': False}
            find_track_name = ', '.join([str(elem) for elem in search.best.result.artists_name()]) + " - " + \
                              search.best.result.title

            # Авторы могут отличаться положением, поэтому делаем все возможные варианты их порядка.
            artists = search.best.result.artists_name()
            all_variants = list(permutations(artists))
            all_variants = [list(variant) for variant in all_variants]
            find_track_names = []
            for variant in all_variants:
                find_track_names.append(', '.join([str(elem) for elem in variant]) + " - " + search.best.result.title)
            # Также может отличаться регистр, так что приведём всё в один регистр.
            bool_name_correct = any(name_current.lower() == element.lower() for element in find_track_names)

            if strong_find and not bool_name_correct:
                print(
                    f"[YamRPC] -> Не можем найти песню (strong_find). Сейчас играет: {name_current}. Но мы нашли: {find_track_name}")
                return {'success': False}

            track = search.best.result
            track_id = track.trackId.split(":")

            if track:
                return {
                    'success': True,
                    'label': f"{', '.join(track.artists_name())} - {track.title}",
                    'duration': "Duration: None",
                    'link': f"https://music.yandex.ru/album/{track_id[1]}/track/{track_id[0]}/",
                    'durationSec': track.duration_ms // 1000,
                    'playback': current_media_info['playback_status'],
                    'og-image': "https://" + track.og_image[:-2] + "400x400"
                }

        except Exception as exception:
            print(f"[YamRPC] -> Вызывай фиксиков! У нас 404!: {exception}")
            return {'success': False}

def WaitAndExit():
    time.sleep(3)
    exit()

if __name__ == '__main__':
    presence = Presence()
    presence.start()
