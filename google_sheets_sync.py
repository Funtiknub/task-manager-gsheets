import json
import webbrowser
import pyperclip
import keyboard
import os
from datetime import datetime, timedelta
import calendar
import time

class GoogleSheetsSync:
    def __init__(self):
        self.settings_file = 'sheet_settings.json'
        self.колонки = [chr(i) for i in range(65, 91)]  # A-Z
        self.доп_колонки = []
        # Генерируем колонки AA-ZZ
        for i in self.колонки:
            for j in self.колонки:
                self.доп_колонки.append(i + j)
        self.все_колонки = self.колонки + self.доп_колонки

    def настроить_доступ(self):
        try:
            print("\n----- Настройка доступа к Google Таблице -----")
            
            # Читаем текущие настройки
            текущие_настройки = {}
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as file:
                    текущие_настройки = json.load(file)
                    print("\nТекущие настройки:")
                    print(f"URL таблицы: {текущие_настройки.get('spreadsheet_id', 'Не задан')}")
                    print(f"Начальная дата: {текущие_настройки.get('start_date', 'Не задана')}")
                    print(f"Начальная ячейка: {текущие_настройки.get('start_cell', 'Не задана')}")

            изменить = input("\nХотите изменить настройки? (да/нет): ").lower()
            
            if изменить == 'да':
                url = input("\nВведите URL таблицы (Enter - оставить текущий): ").strip()
                if url:
                    текущие_настройки['spreadsheet_id'] = url
                
                дата = input("Введите дату начала в формате ДД.ММ.ГГГГ (Enter - оставить текущую): ").strip()
                if дата:
                    текущие_настройки['start_date'] = дата
                
                ячейка = input("Введите номер ячейки для начальной даты (Enter - оставить текущую): ").strip()
                if ячейка:
                    текущие_настройки['start_cell'] = ячейка
                
                with open(self.settings_file, 'w', encoding='utf-8') as file:
                    json.dump(текущие_настройки, file, ensure_ascii=False, indent=2)
                
                print("\nНастройки успешно обновлены!")
            else:
                print("\nНастройки остались без изменений")
            
        except Exception as e:
            print(f"Ошибка при работе с настройками: {e}")

    def получить_букву_колонки(self, номер):
        if номер < len(self.все_колонки):
            return self.все_колонки[номер]
        return None

    def является_рабочим_днем(self, дата, настройки):
        """Проверка, является ли день рабочим с учетом исключений"""
        дата_str = дата.strftime('%d.%m.%Y')
        
        if дата_str in настройки['exceptions']['additional_workdays']:
            return True
        
        if дата_str in настройки['exceptions']['removed_workdays']:
            return False
        
        return дата.weekday() < 5

    def получить_колонку_для_даты(self, начальная_дата, целевая_дата, настройки):
        """Получение колонки для даты с учетом рабочих дней"""
        if not self.является_рабочим_днем(целевая_дата, настройки):
            return None

        # Получаем начальную колонку из настроек
        начальная_ячейка = настройки['start_cell']  # 'A1'
        начальная_колонка = ''.join(c for c in начальная_ячейка if c.isalpha())  # 'A'
        индекс_начальной_колонки = self.все_колонки.index(начальная_колонка)

        # Считаем номер недели и день в неделе
        разница_дней = (целевая_дата - начальная_дата).days
        номер_недели = разница_дней // 7
        день_в_неделе = разница_дней % 7

        # Каждая неделя занимает 15 колонок (5 дней по 3 колонки)
        смещение_колонок = номер_недели * 15 + день_в_неделе * 3

        # Вычисляем итоговый индекс колонки для даты
        итоговый_индекс = индекс_начальной_колонки + смещение_колонок

        if итоговый_индекс >= len(self.все_колонки):
            raise ValueError(f"Ошибка: требуемая колонка {итоговый_индекс} находится за пределами доступного диапазона")

        return self.все_колонки[итоговый_индекс]

    def получить_ячейки_для_задач(self, колонка_даты):
        """Возвращает диапазон ячеек для вставки текущих задач"""
        # Находим индекс колонки даты в списке всех колонок
        индекс = self.все_колонки.index(колонка_даты)
        
        # Получаем две следующие колонки для задач (справа от даты)
        колонка_начала = self.все_колонки[индекс + 1]
        колонка_конца = self.все_колонки[индекс + 2]
        
        # Возвращаем диапазон для строки 2
        return f"{колонка_начала}2:{колонка_конца}2"

    def форматировать_задачи(self, имя_пользователя, tasks_file='tasks.json'):
        """Форматирует все задачи из файла tasks.json"""
        try:
            with open(tasks_file, 'r', encoding='utf-8') as file:
                все_задачи = json.load(file)
                
            задачи = []
            for задача in все_задачи[имя_пользователя]:
                задачи.append(задача['задача'])
            
            return '\n'.join(f"{i+1}. {задача}" for i, задача in enumerate(задачи))
            
        except Exception as e:
            print(f"Ошибка при чтении задач: {e}")
            return ""

    def форматировать_выполненные_задачи(self, имя_пользователя, tasks_file='tasks.json'):
        """Форматирует только выполненные задачи из файла tasks.json"""
        try:
            with open(tasks_file, 'r', encoding='utf-8') as file:
                все_задачи = json.load(file)
                
            выполненные = []
            # Сохраняем оригинальные номера задач
            for i, задача in enumerate(все_задачи[имя_пользователя]):
                if задача['выполнено']:
                    # Используем i+1 для сохранения оригинального номера задачи
                    выполненные.append((i+1, задача['задача']))
            
            # Форматируем с оригинальными номерами
            return '\n'.join(f"{номер}. {задача}" for номер, задача in выполненные)
            
        except Exception as e:
            print(f"Ошибка при чтении выполненных задач: {e}")
            return ""

    def отправить_задачи(self, имя_пользователя, задачи=None):
        try:
            with open(self.settings_file, 'r', encoding='utf-8') as file:
                настройки = json.load(file)
                base_url = настройки['spreadsheet_id']
                начальная_дата = datetime.strptime(настройки['start_date'], '%d.%m.%Y')

            сегодня = datetime.now()
            
            целевая_колонка = self.получить_колонку_для_даты(начальная_дата, сегодня, настройки)
            if not целевая_колонка:
                print("\nСегодня не рабочий день!")
                return False

            диапазон_задач = self.получить_ячейки_для_задач(целевая_колонка)
            
            # Подготавливаем тексты задач заранее
            текст_задач = self.форматировать_задачи(имя_пользователя)
            текст_выполненных = self.форматировать_выполненные_задачи(имя_пользователя)

            url = f"{base_url}#range={диапазон_задач}"
            webbrowser.open(url)

            def on_f8_press(e):
                if e.event_type == keyboard.KEY_DOWN:
                    keyboard.send('f2')
                    time.sleep(0.1)
                    pyperclip.copy(текст_задач)
                    keyboard.send('ctrl+v')

            def on_f9_press(e):
                if e.event_type == keyboard.KEY_DOWN:
                    keyboard.send('f2')
                    time.sleep(0.1)
                    pyperclip.copy(текст_выполненных)
                    keyboard.send('ctrl+v')

            # Подключаем оба обработчика сразу
            keyboard.hook_key('f8', on_f8_press)
            keyboard.hook_key('f9', on_f9_press)
            keyboard.wait('esc')
            
            return True

        except Exception as e:
            print(f"Ошибка: {e}")
            return False 