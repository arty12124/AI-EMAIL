import customtkinter as ctk
from tkinter import messagebox, filedialog
import smtplib
from email.message import EmailMessage
import imaplib
import email
import threading
import openai
import csv
import pandas as pd
from PIL import Image, ImageTk
import os
import json  
import re    

# Настройка темы оформления
ctk.set_appearance_mode("System")      
ctk.set_default_color_theme("blue")    

# Настройки SMTP/IMAP
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
IMAP_SERVER = 'imap.gmail.com'

# Список клиентов (каждый элемент: {"name": ..., "email": ...})
clients = []

# ========== Функции для сохранения/загрузки настроек ==========
SETTINGS_FILE = "settings.json"

def save_settings(email, password, openai_key):
    """Сохранить настройки (Email, Пароль/App Password, API Key) в файл settings.json."""
    data = {
        "email": email,
        "password": password,
        "openai_key": openai_key
    }
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("Сохранение настроек", "Настройки успешно сохранены!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить настройки: {e}")

def load_settings(email_entry, pass_entry, api_entry):
    """Загрузить настройки из файла settings.json, если он существует,
    и заполнить поля Email, Пароль и API Key."""
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            email_entry.delete(0, "end")
            email_entry.insert(0, data.get("email", ""))
            pass_entry.delete(0, "end")
            pass_entry.insert(0, data.get("password", ""))
            api_entry.delete(0, "end")
            api_entry.insert(0, data.get("openai_key", ""))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить настройки: {e}")

# ========== Функция проверки корректности email ==========
def is_valid_email(email_str):
    """
    Проверяет, что email соответствует базовому формату:
    - Нет пробелов;
    - Есть символ "@" с непустой частью до и после;
    - После символа "@" есть точка, и доменное имя не пустое.
    """
    pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
    return bool(re.match(pattern, email_str))

# ========== Основной функционал приложения ==========
def send_email(to_email, subject, body, login_email, login_pass):
    """Отправка письма через SMTP (Gmail по умолчанию)."""
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = login_email
        msg['To'] = to_email
        msg.set_content(body)
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(login_email, login_pass)
            server.send_message(msg)
        return "Письмо отправлено успешно."
    except Exception as e:
        return f"Ошибка при отправке: {e}"

def check_and_reply(email_addr, password, api_key, log_box):
    """Подключается к почте через IMAP, ищет непрочитанные письма,
    отправляет ответ с помощью OpenAI (gpt-3.5-turbo)."""
    try:
        add_log(log_box, "🔄 Подключение к почте...", "info")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(email_addr, password)
        mail.select("inbox")
        status, messages = mail.search(None, 'UNSEEN')
        mail_ids = messages[0].split()
        if not mail_ids:
            add_log(log_box, "ℹ️ Новых писем нет.", "info")
            return
        openai.api_key = api_key
        for num in mail_ids:
            status, msg_data = mail.fetch(num, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    sender = msg['from']
                    subject = msg['subject']
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                body = part.get_payload(decode=True).decode()
                                break
                    else:
                        body = msg.get_payload(decode=True).decode()
                    add_log(log_box, f"📩 Новое письмо от {sender} | Тема: {subject}", "email")
                    completion = openai.ChatCompletion.create(
                        model="gpt-3.5-turbo",
                        messages=[
                            {"role": "system", "content": "Ты отвечаешь на письма клиентов вежливо и профессионально."},
                            {"role": "user", "content": body}
                        ]
                    )
                    reply_text = completion.choices[0].message.content
                    add_log(log_box, f"🤖 Ответ: {reply_text}", "reply")
                    send_email(sender, f"Re: {subject}", reply_text, email_addr, password)
                    add_log(log_box, f"✅ Ответ отправлен: {sender}", "success")
    except Exception as e:
        add_log(log_box, f"❌ Ошибка: {e}", "error")

def add_log(log_box, message, msg_type="info"):
    """Добавление строки в лог (с разным цветом)."""
    log_box.configure(state="normal")
    colors = {
        "info": "#808080",
        "email": "#0078D7",
        "reply": "#107C10",
        "success": "#10893E",
        "error": "#E81123"
    }
    log_box.insert("end", f"{message}\n", msg_type)
    log_box.tag_config(msg_type, foreground=colors.get(msg_type, "#000000"))
    log_box.see("end")
    log_box.configure(state="disabled")

def import_clients(client_listbox):
    """Импорт клиентов из CSV или Excel.
    Если в строке не найден стандартный ключ для email,
    пробегаемся по всем значениям строки и, если найдено значение с символом '@',
    используем его как email-адрес.
    """
    file_path = filedialog.askopenfilename(filetypes=[("CSV/Excel files", "*.csv *.xlsx")])
    if not file_path:
        return
    try:
        clients.clear()
        client_listbox.delete("1.0", "end")
        if file_path.endswith(".csv"):
            with open(file_path, newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    email_addr = row.get("email") or row.get("Email")
                    name = row.get("name") or row.get("Name") or "Клиент"
                    if not email_addr:
                        # Пробегаем по всем значениям строки
                        for value in row.values():
                            if isinstance(value, str) and "@" in value:
                                email_addr = value.strip()
                                break
                    if email_addr:
                        clients.append({"name": name, "email": email_addr})
        elif file_path.endswith(".xlsx"):
            df = pd.read_excel(file_path)
            for _, row in df.iterrows():
                email_addr = row.get("email") or row.get("Email")
                name = row.get("name") or row.get("Name") or "Клиент"
                if pd.isna(email_addr) or not isinstance(email_addr, str):
                    for item in row:
                        if isinstance(item, str) and "@" in item:
                            email_addr = item.strip()
                            break
                if email_addr:
                    clients.append({"name": name, "email": email_addr})
        for c in clients:
            client_listbox.insert("end", f"{c['name']} <{c['email']}>\n")
        messagebox.showinfo("Импорт", "Клиенты успешно загружены.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось загрузить файл: {e}")

def add_client_manually(client_listbox, entry):
    """Добавить клиента вручную из поля ввода, проверяя корректность email."""
    email_addr = entry.get().strip()
    if not email_addr:
        messagebox.showerror("Ошибка", "Пожалуйста, введите email-адрес.")
        return
    if not is_valid_email(email_addr):
        messagebox.showerror("Ошибка", "Некорректный формат email-адреса.\nУбедитесь, что адрес содержит '@' и домен.")
        return
    clients.append({"name": "Ручной клиент", "email": email_addr})
    client_listbox.insert("end", f"Ручной клиент <{email_addr}>\n")
    entry.delete(0, "end")

def send_to_all(subject, body_template, login_email, login_pass, log_box):
    """Отправка письма всем клиентам из списка."""
    if not clients:
        add_log(log_box, "⚠️ Список клиентов пуст! Импортируйте клиентов или добавьте вручную.", "error")
        return
    if not subject.strip():
        add_log(log_box, "⚠️ Не указана тема письма!", "error")
        return
    if not body_template.strip():
        add_log(log_box, "⚠️ Текст письма пустой!", "error")
        return
    add_log(log_box, "📤 Начинаем массовую рассылку...", "info")
    for client in clients:
        personalized_body = body_template.replace("{name}", client["name"])
        result = send_email(client["email"], subject, personalized_body, login_email, login_pass)
        if "успешно" in result:
            add_log(log_box, f"✅ {client['name']} <{client['email']}>: {result}", "success")
        else:
            add_log(log_box, f"❌ {client['name']} <{client['email']}>: {result}", "error")
    add_log(log_box, "✅ Рассылка завершена!", "success")

def toggle_password(entry, button):
    """Показать/скрыть пароль в поле ввода."""
    if entry.cget('show') == '•':
        entry.configure(show='')
        button.configure(text='🔒')
    else:
        entry.configure(show='•')
        button.configure(text='👁️')

def save_template(subject_entry, body_entry):
    """Сохранение текущего шаблона (темы и текста письма) в текстовый файл."""
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")],
        title="Сохранить шаблон письма"
    )
    if file_path:
        try:
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(f"SUBJECT: {subject_entry.get()}\n\n")
                file.write(body_entry.get("1.0", "end"))
            messagebox.showinfo("Сохранение", "Шаблон успешно сохранен!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить шаблон: {e}")

def load_template(subject_entry, body_entry):
    """Загрузка шаблона (темы и текста письма) из текстового файла."""
    file_path = filedialog.askopenfilename(
        filetypes=[("Text files", "*.txt")],
        title="Загрузить шаблон письма"
    )
    if file_path:
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.readlines()
                if content and content[0].startswith("SUBJECT:"):
                    subject = content[0][8:].strip()
                    subject_entry.delete(0, "end")
                    subject_entry.insert(0, subject)
                    body = ''.join(content[2:])
                    body_entry.delete("1.0", "end")
                    body_entry.insert("1.0", body)
                else:
                    body = ''.join(content)
                    body_entry.delete("1.0", "end")
                    body_entry.insert("1.0", body)
            messagebox.showinfo("Загрузка", "Шаблон успешно загружен!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить шаблон: {e}")

def clear_log(log_box):
    """Очистить содержимое логового окна."""
    log_box.configure(state="normal")
    log_box.delete("1.0", "end")
    log_box.configure(state="disabled")

def main_app():
    root = ctk.CTk()
    root.title("AI Почтовик Pro")
    root.geometry("900x750")
    root.resizable(False, False)  # Окно фиксированного размера

    # Создание вкладок
    tabview = ctk.CTkTabview(root)
    tabview.pack(fill="both", expand=True, padx=10, pady=10)
    
    tab_main = tabview.add("📧 Email")
    tab_settings = tabview.add("⚙️ Настройки")
    tab_help = tabview.add("❓ Помощь")
    
    # ====================== ВКЛАДКА EMAIL ======================
    left_frame = ctk.CTkFrame(tab_main)
    left_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
    
    template_frame = ctk.CTkFrame(left_frame, fg_color="transparent", corner_radius=0, border_width=0)
    template_frame.pack(fill="x", padx=5, pady=5)
    
    email_template_label = ctk.CTkLabel(template_frame, text="Шаблон письма", font=("Arial", 16, "bold"))
    email_template_label.pack(pady=5)

    # Кнопки для работы с шаблонами
    template_btn_frame = ctk.CTkFrame(template_frame, fg_color="transparent", corner_radius=0, border_width=0)
    template_btn_frame.pack(fill="x", pady=5)
    
    save_template_btn = ctk.CTkButton(template_btn_frame, text="💾 Сохранить шаблон", width=80)
    load_template_btn = ctk.CTkButton(template_btn_frame, text="📂 Загрузить шаблон", width=80)
    
    save_template_btn.grid(row=0, column=0, padx=5, pady=5)
    load_template_btn.grid(row=0, column=1, padx=5, pady=5)
    template_btn_frame.grid_columnconfigure(0, weight=1)
    template_btn_frame.grid_columnconfigure(1, weight=1)
    
    save_template_btn.configure(command=lambda: save_template(subject_entry, body_entry))
    load_template_btn.configure(command=lambda: load_template(subject_entry, body_entry))

    # Форма письма
    ctk.CTkLabel(template_frame, text="Тема письма:").pack(anchor="w", padx=5, pady=(10, 2))
    subject_entry = ctk.CTkEntry(template_frame, placeholder_text="Введите тему письма...")
    subject_entry.pack(fill="x", padx=5, pady=(0, 10))
    
    ctk.CTkLabel(template_frame, text="Текст письма (используйте {name} для имени клиента):").pack(anchor="w", padx=5, pady=(5, 2))
    body_entry = ctk.CTkTextbox(template_frame, height=200, text_color="#808080")
    body_entry.pack(fill="both", expand=True, padx=5, pady=(0, 10))
    
    # База клиентов
    clients_frame = ctk.CTkFrame(left_frame)
    clients_frame.pack(fill="both", expand=True, padx=5, pady=5)
    
    client_label = ctk.CTkLabel(clients_frame, text="База клиентов", font=("Arial", 16, "bold"))
    client_label.pack(pady=5)

    # Кнопка Импорт
    import_btn = ctk.CTkButton(clients_frame, text="📥 Импорт базы клиентов (CSV/Excel)",
                               command=lambda: import_clients(client_listbox))
    import_btn.pack(pady=5)

    # Фрейм для ручного добавления почты
    manual_add_frame = ctk.CTkFrame(clients_frame)
    manual_add_frame.pack(fill="x", padx=5, pady=(0, 5))
    
    manual_add_entry = ctk.CTkEntry(manual_add_frame, placeholder_text="Введите email...")
    manual_add_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
    manual_add_btn = ctk.CTkButton(manual_add_frame, text="Добавить",
                                   command=lambda: add_client_manually(client_listbox, manual_add_entry))
    manual_add_btn.pack(side="left")

    # Текстовое поле со списком клиентов
    client_listbox = ctk.CTkTextbox(clients_frame, height=150)
    client_listbox.pack(fill="both", expand=True, padx=5, pady=5)
    
    # Правая колонка
    right_frame = ctk.CTkFrame(tab_main)
    right_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)
    
    log_frame = ctk.CTkFrame(right_frame)
    log_frame.pack(fill="both", expand=True, padx=5, pady=5)
    
    log_label = ctk.CTkLabel(log_frame, text="Журнал активности", font=("Arial", 16, "bold"))
    log_label.pack(pady=5)
    
    log_box = ctk.CTkTextbox(log_frame, height=300, wrap="word")
    log_box.pack(fill="both", expand=True, padx=5, pady=5)
    log_box.configure(state="disabled")
    
    log_buttons_frame = ctk.CTkFrame(log_frame)
    log_buttons_frame.pack(fill="x", padx=5, pady=5)
    clear_log_btn = ctk.CTkButton(log_buttons_frame, text="🧹 Очистить лог", width=120,
                                  command=lambda: clear_log(log_box))
    clear_log_btn.pack(side="right", padx=5)
    
    actions_frame = ctk.CTkFrame(right_frame)
    actions_frame.pack(fill="x", padx=5, pady=5)
    
    actions_label = ctk.CTkLabel(actions_frame, text="Действия", font=("Arial", 16, "bold"))
    actions_label.pack(pady=5)
    
    send_all_btn = ctk.CTkButton(actions_frame, text="📤 Отправить всем из базы", height=40,
                                 fg_color="#107C10", hover_color="#0E6E0E",
                                 command=lambda: send_to_all(
                                     subject_entry.get(),
                                     body_entry.get("1.0", "end"),
                                     email_entry.get(),
                                     pass_entry.get(),
                                     log_box
                                 ))
    send_all_btn.pack(fill="x", padx=20, pady=5)
    
    reply_btn = ctk.CTkButton(actions_frame, text="🤖 Ответить на новые письма (AI)", height=40,
                              fg_color="#0078D7", hover_color="#0063B1",
                              command=lambda: threading.Thread(
                                  target=check_and_reply,
                                  args=(email_entry.get(), pass_entry.get(), api_entry.get(), log_box)
                              ).start())
    reply_btn.pack(fill="x", padx=20, pady=5)
    
    # ====================== ВКЛАДКА НАСТРОЙКИ ======================
    settings_frame = ctk.CTkFrame(tab_settings)
    settings_frame.pack(fill="both", expand=True, padx=30, pady=30)
    
    settings_label = ctk.CTkLabel(settings_frame, text="Настройки подключения", font=("Arial", 20, "bold"))
    settings_label.pack(pady=15)
    
    # Ваш Email
    ctk.CTkLabel(settings_frame, text="Ваш Email:").pack(anchor="w", padx=20, pady=(10, 5))
    email_entry = ctk.CTkEntry(settings_frame, placeholder_text="example@gmail.com", width=400)
    email_entry.pack(padx=20, pady=(0, 15), fill="x")
    
    # Пароль (или App Password)
    ctk.CTkLabel(settings_frame, text="Пароль (или App Password):").pack(anchor="w", padx=20, pady=(0, 5))
    pass_container = ctk.CTkFrame(settings_frame, fg_color="transparent", corner_radius=0)
    pass_container.pack(fill="x", padx=20, pady=(0, 15))
    pass_entry = ctk.CTkEntry(pass_container, placeholder_text="Ваш пароль или код приложения",
                              show="•", fg_color="transparent")
    pass_entry.pack(side="left", fill="x", expand=True)
    show_pass_btn = ctk.CTkButton(pass_container, text="👁️", width=40,
                                  command=lambda: toggle_password(pass_entry, show_pass_btn))
    show_pass_btn.pack(side="left", padx=5)
    
    # OpenAI API Key
    ctk.CTkLabel(settings_frame, text="OpenAI API Key:").pack(anchor="w", padx=20, pady=(0, 5))
    api_container = ctk.CTkFrame(settings_frame, fg_color="transparent", corner_radius=0)
    api_container.pack(fill="x", padx=20, pady=(0, 15))
    api_entry = ctk.CTkEntry(api_container, placeholder_text="sk-...", show="•", fg_color="transparent")
    api_entry.pack(side="left", fill="x", expand=True)
    show_api_btn = ctk.CTkButton(api_container, text="👁️", width=40,
                                 command=lambda: toggle_password(api_entry, show_api_btn))
    show_api_btn.pack(side="left", padx=5)
    
    # Секция "Внешний вид" полностью удалена
    
    # КНОПКА "Сохранить данные" (сохраняет Email, Пароль, API-ключ)
    save_data_btn = ctk.CTkButton(settings_frame, text="💾 Сохранить данные",
                                  command=lambda: save_settings(
                                      email_entry.get(),
                                      pass_entry.get(),
                                      api_entry.get()
                                  ))
    save_data_btn.pack(pady=10)

    # После создания всех полей — загружаем сохранённые настройки (если есть)
    load_settings(email_entry, pass_entry, api_entry)

    # ====================== ВКЛАДКА ПОМОЩЬ ======================
    help_frame = ctk.CTkFrame(tab_help)
    help_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    help_label = ctk.CTkLabel(help_frame, text="Руководство пользователя", font=("Arial", 20, "bold"))
    help_label.pack(pady=15)
    
    help_text = ctk.CTkTextbox(help_frame, height=600)
    help_text.pack(fill="both", expand=True, padx=20, pady=10)
    
    help_content = """# AI Почтовик Pro - Руководство пользователя

## 📧 Работа с письмами
1. **Шаблон письма** - Создайте тему и текст письма. Используйте {name} для подстановки имени клиента.
2. **База клиентов** - Импортируйте базу из CSV или Excel файла с колонками "name" и "email".
   Также можно вручную добавить клиента, используя поле "Введите email...".
3. **Отправка писем** - Нажмите "Отправить всем из базы" для массовой рассылки.

## 🤖 AI-ответы на письма
1. Настройте подключение к почте и API ключ OpenAI.
2. Нажмите "Ответить на новые письма (AI)" для автоматической обработки.
3. Бот прочитает непрочитанные письма и ответит на них с помощью ИИ.

## ⚙️ Настройка
1. **Email и пароль** - Для Gmail используйте пароль приложения, а не обычный пароль (особенно если включена 2FA).
2. **API ключ** - Укажите ваш ключ OpenAI API.
3. **Сохранить данные** - Позволяет сохранить текущие Email, пароль и API-ключ в файл для будущих запусков.

## 📝 Советы
- При ошибке AUTHENTICATIONFAILED проверьте правильность ввода Email/пароля.
- Сохраняйте успешные шаблоны писем для повторного использования.
- Проверяйте лог активности для мониторинга отправок.
"""
    help_text.insert("1.0", help_content)
    help_text.configure(state="disabled")
    
    add_log(log_box, "🚀 AI Почтовик Pro запущен и готов к работе!", "info")
    add_log(log_box, "ℹ️ Настройте Email, пароль и API ключ во вкладке 'Настройки'. Нажмите «Сохранить данные», чтобы сохранить настройки.", "info")
    
    root.mainloop()

if __name__ == "__main__":
    main_app()
