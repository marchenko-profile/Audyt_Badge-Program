import os
import sys
import xlwings as xw
from datetime import datetime
from xlwings import utils
import tkinter as tk
from tkinter import messagebox, filedialog
import webbrowser

# Генерация имени файла
file_name = datetime.today().strftime('%d.%m.%Y') + '_AUDYT_BADGE_MAIN_GATE_3.xlsx'


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Файлы Excel", "*.xlsx")])
    if file_path:
        # Обновляем поле ввода для отображения выбранного пути к файлу
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)
        process_button.config(state=tk.NORMAL)  # Активируем кнопку "Обработать данные"

def process_data():
    # Получаем путь к выбранному файлу из поля ввода
    file_path = file_entry.get()

    if not file_path:
        messagebox.showerror("Ошибка", "Выберите файл перед обработкой данных.")
        return

    # Получаем путь к папке, в которой находится выбранный файл
    folder_path = os.path.dirname(file_path)

    if not file_path:
        messagebox.showerror("Ошибка", "Выберите файл перед обработкой данных.")
        return

        # Создаем новый лист для несовпадающих номеров
        unmatched_sheet = workbook.sheets.add('Номера')

        unmatched_numbers = []

    # Открываем файл в режиме без отображения
    app = xw.App(visible=False)
    workbook = app.books.open(file_path)


    # Выбираем активный лист
    worksheet = workbook.sheets.active

    # Запрашиваем имя у пользователя
    name = name_entry.get().upper()

    # Записываем имя в столбец
    for row in range(8, worksheet.range("B65536").end("up").row + 1):
        worksheet.range(f'E{row}').value = name

    # Записываем имя в столбец
    for row in range(8, worksheet.range("H65536").end("up").row + 1):
        worksheet.range(f'K{row}').value = name


    # Записываем имя в столбец
    for row in range(8, worksheet.range("N65536").end("up").row + 1):
        worksheet.range(f'Q{row}').value = name

    # Записываем имя в столбец
    for row in range(8, worksheet.range("T65536").end("up").row + 1):
        worksheet.range(f'W{row}').value = name

    # Запрашиваем номера карт VISITOR у пользователя
    cards_input_visitor = visitor_entry.get()

    # Разделяем номера карт по пробелам и сохраняем в список
    cards_visitor = cards_input_visitor.split()

    # Добавляем функцию для создания нового листа
    def create_new_sheet(workbook, sheet_name):
        workbook.sheets.add(sheet_name)
        workbook.sheets(sheet_name).api.Move(After=workbook.sheets[-1].api)

    # Получаем все номера из столбца B
    card_numbers = [float(cell.value) for cell in
                    worksheet.range("B8:B" + str(worksheet.range("B65536").end("up").row))]

    # Получаем уникальные номера, которые пользователь ввел
    user_numbers = set(map(float, cards_visitor))

    # Создаем новый лист "номера"
    create_new_sheet(workbook, "Brak kart")
    numbers_sheet = workbook.sheets["Brak kart"]

    row_index = 1

    # Перебираем номера, введенные пользователем
    for number in user_numbers:
        if number not in card_numbers:
            numbers_sheet.range("A" + str(row_index)).value = number
            row_index += 1

    # Обновляем столбец  и устанавливаем зеленый цвет для ячеек, которые совпадают
    for row in range(8, worksheet.range("B65536").end("up").row + 1):
        card_number = float(worksheet.range(f"B{row}").value)
        if card_number in user_numbers:
            worksheet.range(f"C{row}").value = "JEST"
            # Устанавливаем зеленый цвет для ячейки
            cell = worksheet.range(f"C{row}")
            cell.color = utils.rgb_to_int((0, 255, 0))  # Зеленый цвет

    # Запрашиваем номера карт TEMP у пользователя
    cards_input_temp = temp_entry.get()

    # Разделяем номера карт по пробелам и сохраняем в список
    cards_temp = cards_input_temp.split()

    # Сравниваем число в ячейках столбца H с каждым номером, введенным пользователем
    for row in range(8, worksheet.range("H65536").end("up").row + 1):
        card_number = float(worksheet.range(f"H{row}").value)
        if card_number in map(float, cards_temp):
            worksheet.range(f"I{row}").value = "JEST"
            # Устанавливаем зеленый цвет для ячейки
            cell = worksheet.range(f"I{row}")
            cell.color = utils.rgb_to_int((0, 255, 0))  # Зеленый цвет


    # Запрашиваем номера карт GUEST у пользователя
    cards_input_guest = guest_entry.get()

    # Разделяем номера карт по пробелам и сохраняем в список
    cards_guest = cards_input_guest.split()

    # Сравниваем число в ячейках столбца N с каждым номером, введенным пользователем
    for row in range(8, worksheet.range("N65536").end("up").row + 1):
        card_number = float(worksheet.range(f"N{row}").value)
        if card_number in map(float, cards_guest):
            worksheet.range(f"O{row}").value = "JEST"
            # Устанавливаем зеленый цвет для ячейки
            cell = worksheet.range(f"O{row}")
            cell.color = utils.rgb_to_int((0, 255, 0))  # Зеленый цвет


    # Запрашиваем номера карт VIP у пользователя
    cards_input_vip = vip_entry.get()

    # Разделяем номера карт по пробелам и сохраняем в список
    cards_vip = cards_input_vip.split()

    # Сравниваем число в ячейках столбца T с каждым номером, введенным пользователем
    for row in range(8, worksheet.range("T65536").end("up").row + 1):
        card_number = float(worksheet.range(f"T{row}").value)
        if card_number in map(float, cards_vip):
            worksheet.range(f"U{row}").value = "JEST"
            # Устанавливаем зеленый цвет для ячейки
            cell = worksheet.range(f"U{row}")
            cell.color = utils.rgb_to_int((0, 255, 0))  # Зеленый цвет

    # Записываем актуальную дату в столбец D от строки 8 до 75
    today = datetime.today().strftime('%d.%m.%Y')
    for row in range(8, worksheet.range("B65536").end("up").row + 1):
        worksheet.range(f'D{row}').value = today

    today = datetime.today().strftime('%d.%m.%Y')
    for row in range(8, worksheet.range("H65536").end("up").row + 1):
        worksheet.range(f'J{row}').value = today

    today = datetime.today().strftime('%d.%m.%Y')
    for row in range(8, worksheet.range("N65536").end("up").row + 1):
        worksheet.range(f'P{row}').value = today

    today = datetime.today().strftime('%d.%m.%Y')
    for row in range(8, worksheet.range("T65536").end("up").row + 1):
        worksheet.range(f'V{row}').value = today

    # Запрашиваем путь для сохранения файла
    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile=file_name,
                                                 filetypes=[("Файлы Excel", "*.xlsx")])
    if save_path:
        # Сохраняем изменения с актуальной датой в выбранном месте
        workbook.save(save_path)
        workbook.close()
        app.quit()
    else:
        messagebox.showinfo("Информация", "Сохранение отменено.")

    workbook.save(save_path)
    workbook.close()
    app.quit()

def open_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(tk.END, file_path)
        open_button.config(text="Plik wybrany!", state=tk.DISABLED)  # Изменяем текст кнопки и отключаем ее

def enable_process_button():
    file_path = file_entry.get()
    name = name_entry.get()
    cards_visitor = visitor_entry.get()
    cards_temp = temp_entry.get()
    cards_guest = guest_entry.get()
    cards_vip = vip_entry.get()

    if file_path and name and cards_visitor and cards_temp and cards_guest and cards_vip:
        process_button.config(state=tk.NORMAL)
    else:
        process_button.config(state=tk.DISABLED)




# Создаем графический интерфейс
root = tk.Tk()
root.title("Data Processing")
root.geometry("1000x372")


# Устанавливаем шрифт
font_style = ("Arial", 16)

# Создаем кнопку для открытия диалогового окна выбора файла
open_button = tk.Button(root, text="Otwórz plik", font=font_style, command=open_file)
open_button.grid(row=0, column=0, columnspan=2, pady=10)

# Создаем поле для отображения выбранного файла
file_entry = tk.Entry(root, font=font_style, show='*')
file_entry.grid(row=1, column=0, columnspan=2)

# Скрываем поле из видимости
file_entry.grid_remove()

# Создаем метку и поле для ввода имени
name_label = tk.Label(root, text="Nazwisko audytora:", font=font_style)
name_label.grid(row=2, column=0, sticky="w")
name_entry = tk.Entry(root, font=font_style, width=60)
name_entry.grid(row=2, column=1, pady=10)

# Создаем метку и поле для ввода номеров карт VISITOR
visitor_label = tk.Label(root, text="Wprowadź karty VISITOR:", font=font_style)
visitor_label.grid(row=3, column=0, sticky="w")
visitor_entry = tk.Entry(root, font=font_style, width=60)
visitor_entry.grid(row=3, column=1, pady=10)

# Создаем метку и поле для ввода номеров карт TEMP
temp_label = tk.Label(root, text="Wprowadź karty TEMP:", font=font_style)
temp_label.grid(row=4, column=0, sticky="w")
temp_entry = tk.Entry(root, font=font_style, width=60)
temp_entry.grid(row=4, column=1)

# Создаем метку и поле для ввода номеров карт GUEST
guest_label = tk.Label(root, text="Wprowadź karty GUEST:", font=font_style)
guest_label.grid(row=5, column=0, sticky="w")
guest_entry = tk.Entry(root, font=font_style, width=60)
guest_entry.grid(row=5, column=1, pady=10)

# Создаем метку и поле для ввода номеров карт VIP
vip_label = tk.Label(root, text="Wprowadź karty VIP:", font=font_style)
vip_label.grid(row=6, column=0, sticky="w")
vip_entry = tk.Entry(root, font=font_style, width=60)
vip_entry.grid(row=6, column=1)

# Создаем кнопку для обработки данных
process_button = tk.Button(root, text="Rozpocznij operację", font=font_style, command=process_data, state=tk.DISABLED)
process_button.grid(row=7, column=0, columnspan=2, pady=10)

# Set up the event bindings for the input fields
file_entry.bind("<KeyRelease>", lambda event: enable_process_button())
name_entry.bind("<KeyRelease>", lambda event: enable_process_button())
visitor_entry.bind("<KeyRelease>", lambda event: enable_process_button())
temp_entry.bind("<KeyRelease>", lambda event: enable_process_button())
guest_entry.bind("<KeyRelease>", lambda event: enable_process_button())
vip_entry.bind("<KeyRelease>", lambda event: enable_process_button())



# Функция для открытия ссылки
def open_link(event):
    webbrowser.open("https://www.marchenko-profile.com/")

# Функция для изменения цвета текста при наведении
def enter(event):
    developer_label.config(fg="black")

# Функция для восстановления цвета текста при выходе из наведения
def leave(event):
    developer_label.config(fg="gray70")

# Создаем метку с информацией о разработчике
developer_label = tk.Label(root, text="creator: Serhii Marchenko | tel.: +48 503 555 215 | www.marchenko-profile.com",
                           font=("Arial", 12), fg="gray70", cursor="hand2")
developer_label.grid(row=8, column=0, columnspan=2, pady=10)

# Привязываем обработчики событий к метке
developer_label.bind("<Button-1>", open_link)
developer_label.bind("<Enter>", enter)
developer_label.bind("<Leave>", leave)




# Создаем функцию для открытия окна с инструкцией
def open_instructions():
    instructions_window = tk.Toplevel(root)
    instructions_window.title("Instrukcja użytkowania")
    instructions_window.geometry("600x650")

    # Добавляем текстовую метку с инструкцией
    instructions_label = tk.Label(instructions_window, text="INSTRUKCJA DLA UŻYTKOWNIKA:\n" "\n""\n"
                                                            "1. Nacisnąć 'Otwórz plik'.\n" "\n"
                                                            "2. W folderze Audyt Badge [Program]\n" "proszę wybrać odpowiedni plik w formacie Excel \n" "w zależności od posterunku.\n" "\n"
                                                            "3. Napisz nazwisko audytora.\n" "\n"
                                                            "4. Wprowadź karty VISITOR \n" "które są w systemie i te karty, które są na rękach \n" "(pisz przez spację).\n" "\n"
                                                            "5. Wprowadź karty TEMP (pisz przez spację).\n" "\n"
                                                            "6. Wprowadź karty GUEST (pisz przez spację).\n""\n"
                                                            "7. Wprowadź karty VIP (pisz przez spację).\n""\n"
                                                            "8. Nacisnąć 'Rozpocznij operację'.\n""\n"
                                                            "9. Po zakończeniu operacji zapisz plik tam, gdzie potrzebujesz.\n""\n"
                                                            "10. Zamknij program.", font=font_style)
    instructions_label.pack(pady=10)

# Создаем кнопку "Instrukcja użytkowania" с настроенным размером, шрифтом и смайликом
instructions_button = tk.Button(root, text="Instrukcja", font=("Arial", 12), command=open_instructions, width=7, height=1)
instructions_button.grid(row=0, column=0, pady=10, padx=10, sticky="nw")





# Запускаем главный цикл обработки событий
root.mainloop()