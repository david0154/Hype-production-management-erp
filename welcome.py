# welcome.py - First Run Password Setup and Login Screen with Modern UI and Enter Key Support

from tkinter import *
from tkinter import messagebox
from PIL import Image, ImageTk
import os
from utils import get_password, set_password, verify_password
from db import init_db
from main_ui import show_main_ui

APP_NAME = "Hype Production Management"

def show_welcome():
    root = Tk()
    root.title(APP_NAME)
    root.geometry("600x500")
    root.configure(bg="#f0faff")  # Light pastel background

    # --- Styles ---
    title_font = ("Helvetica", 18, "bold")
    subtitle_font = ("Helvetica", 12)
    input_font = ("Helvetica", 11)
    button_font = ("Helvetica", 11, "bold")

    # --- Logo (top) ---
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        img = Image.open(logo_path)
        img = img.resize((220, 90))
        logo = ImageTk.PhotoImage(img)
        logo_label = Label(root, image=logo, bg="#f0faff")
        logo_label.image = logo
        logo_label.pack(pady=15)

    # Title
    Label(root, text=f"Welcome to {APP_NAME}", font=title_font, fg="#1a237e", bg="#f0faff").pack(pady=5)

    container = Frame(root, bg="#f0faff")
    container.pack(pady=20)

    password_entry = Entry(container, show="*", width=30, font=input_font)
    password_entry.pack(pady=10)

    # --- First Run: Set Password ---
    if not get_password():
        Label(container, text="Set Your Admin Password", font=subtitle_font, bg="#f0faff", fg="#004d40").pack(pady=5)

        def save_password():
            new_pw = password_entry.get().strip()
            if not new_pw:
                messagebox.showerror("Error", "Password cannot be empty")
                return
            set_password(new_pw)
            messagebox.showinfo("Success", "Password set successfully!")
            root.destroy()
            init_db()
            show_main_ui()

        Button(container, text="Set Password & Continue", command=save_password,
               bg="#00bfa5", fg="white", font=button_font, bd=0, padx=10, pady=5).pack(pady=15)

        # Bind Enter key to Set Password
        root.bind('<Return>', lambda event: save_password())

    # --- Later Runs: Login ---
    else:
        Label(container, text="Enter Admin Password", font=subtitle_font, bg="#f0faff", fg="#1a237e").pack(pady=5)

        def verify():
            if verify_password(password_entry.get()):
                root.destroy()
                init_db()
                show_main_ui()
            else:
                messagebox.showerror("Error", "Incorrect password")

        Button(container, text="Login", command=verify,
               bg="#1976d2", fg="white", font=button_font, bd=0, padx=15, pady=5).pack(pady=15)

        # Bind Enter key to Login
        root.bind('<Return>', lambda event: verify())

        Label(root, text="Developed by David", font=("Arial", 10), fg="#888", bg="#f0faff").pack(side=BOTTOM, pady=10)

    root.mainloop()