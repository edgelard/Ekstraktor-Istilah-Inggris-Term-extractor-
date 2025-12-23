def main():
    import customtkinter as ctk
    from tkinter import filedialog
    import json
    import sys
    import os


    def resource_path(relative_path):
        if hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)

    SCRIPT_PATH = resource_path("extract_istilah_asing.py")
    VBA_SCRIPT_PATH = resource_path("vba_macro.py")
    CONFIG_PATH = resource_path("config.json")

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")

    existing_config = {}

    def select_address():
        file_path = filedialog.askopenfilename(
            # filetypes=[("Word Documents", "*.docx")]
            # filetypes=[("Word Documents")]
        )
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                existing_config = json.load(f)

        existing_config["input_docx"] = file_path

        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(existing_config, f, indent=2)

        if not file_path:
            return
        input_path_var.set(file_path)

    def min_word_length(value):
        if value == "":
            return True          
        return value.isdigit() and int(value) >= 1

    def only_var(value):
        if value == "":
            return True          
        return value.isdigit() and int(value) >= 0

    import vba_macro
    def run_vba_macro():
            vba_macro.run()

    import extract_istilah_asing
    def save_to_config():
        
        value1 = entry_min.get()
        value2 = entry_id_freq.get()
        value3 = entry_en_freq.get()

        if not value1 or int(value1) < 1:
            return
        if not value2 or int(value2) < 1:
            return
        if not value3 or int(value3) < 1:
            return

        kata_config = {
            "min_word_length": int(value1),
            "id_freq_Value": int(value2),
            "en_freq_Value": int(value3)
        }
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    existing_config = json.load(f)
            except json.JSONDecodeError:
                existing_config = {}

        # merge (new values override old ones)
        existing_config.update(kata_config)

        # write once
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(existing_config, f, indent=2)

        extract_istilah_asing.run()

    def choose_output_folder():
        global output_dir

        folder_path = filedialog.askdirectory(title="Select Output Folder")
        output_dir = folder_path
        output_path_var.set(folder_path)

        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                existing_config = json.load(f)

        existing_config["output_docx"] = folder_path

        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(existing_config, f, indent=2)

        if not folder_path:
            return

    # =========================
    app = ctk.CTk()
    frame = ctk.CTkFrame(app)
    frame.pack(padx=10, pady=10)
    frame2 = ctk.CTkFrame(app)
    frame2.pack(padx=10, pady=10)
    frame3 = ctk.CTkFrame(app)
    frame3.pack(padx=10, pady=10)
    frame3.grid_propagate(False)
    frame3.configure(width=480, height=140)
    frame3.grid_columnconfigure(0, weight=1)
    frame3.grid_columnconfigure(1, weight=1)
    frame4 = ctk.CTkFrame(app)
    frame4.pack(padx=10, pady=10)
    app.title("Qalbi Utilities")
    app.geometry("500x640")
    app.iconbitmap("sixthsense.ico")
    # =========================

    # ======================== Frame 1 ========================
    ctk.CTkLabel(
        frame,
        text="PENGEKSTRAK ISTILAH INGGRIS + ITALICIZER",
        font=ctk.CTkFont(weight="bold")
    ).grid(row=0, column=0)
    btn = ctk.CTkButton(
        frame,
        text="Pilih file path Document",
        width=220,
        height=40,
        command=select_address
    )
    btn.grid(row=1, column=0, pady=(10, 10))

    input_path_var = ctk.StringVar(value="No file selected")
    path_entry = ctk.CTkEntry(
        frame,
        textvariable=input_path_var,
        width=480,
        state="readonly"
    )
    path_entry.grid(row=2, column=0)


    # ======================== Frame 2 ========================
    ctk.CTkLabel(
        frame2,
        text="Minimum jumlah huruf yang di indeks (≥ 1)",
        width= 480
    ).grid(row=0, column=0, sticky="w")

    vcmd = app.register(min_word_length)
    min_word_length_path_var = ctk.IntVar(value=3)
    entry_min = ctk.CTkEntry(
        frame2,
        textvariable=min_word_length_path_var,
        validate="key",
        validatecommand=(vcmd, "%P")
    )
    entry_min.grid(row=1, column=0, pady=(10, 10))
    # ======================== Frame 3 ========================
    vcmd1 = app.register(only_var)
    ctk.CTkLabel(
        frame3,
        text="Frekuensi Jumlah Kata di Karya Tulis Indonesia",
        wraplength=190,
        font=ctk.CTkFont(weight="bold"),
        justify="center"
    ).grid(row=0, column=0, pady=(10, 0))
    vcmd1 = app.register(only_var)
    ctk.CTkLabel(
        frame3,
        text="(Semakin tinggi semakin banyak kata yang diambil)",
        wraplength=190,
        justify="center"
    ).grid(row=1, column=0, pady=(0, 0))

    id_freq_path_var = ctk.IntVar(value=2)
    entry_id_freq = ctk.CTkEntry(
        frame3,
        textvariable=id_freq_path_var,
        width= 50,
        validate="key",
        validatecommand=(vcmd1, "%P")
    )
    entry_id_freq.grid(row=2, column=0, pady=(5, 0))

    ctk.CTkLabel(
        frame3,
        text="Frekuensi Jumlah Kata di Karya Tulis Inggris",
        wraplength=190,
        font=ctk.CTkFont(weight="bold"),
        justify="center"
    ).grid(row=0, column=1, pady=(10, 0))
    ctk.CTkLabel(
        frame3,
        text="(Semakin tinggi semakin sedikit kata yang diambil)",
        wraplength=190,
        justify="center"
    ).grid(row=1, column=1, pady=(0, 0))

    en_freq_path_var = ctk.IntVar(value=3)
    entry_en_freq = ctk.CTkEntry(
        frame3,
        textvariable=en_freq_path_var,
        width= 50,
        validate="key",
        validatecommand=(vcmd1, "%P")
    )
    entry_en_freq.grid(row=2, column=1, pady=(5, 0))

    ctk.CTkLabel(
        frame3,
        text="Default id_freq= 2"
    ).grid(row=3, column=0, sticky="s", pady=(0, 40))
    ctk.CTkLabel(
        frame3,
        text="Default en_freq= 3"
    ).grid(row=3, column=1, sticky="s", pady=(0, 40))
    # ======================== Frame 4 ========================
    ctk.CTkButton(
        frame4,
        text="Pilih Ouput Folder",
        width=220,
        height=40,
        command=choose_output_folder
    ).grid(row=1, column=0, pady=(10, 10))

    output_path_var = ctk.StringVar(value="No path selected")
    path_entry = ctk.CTkEntry(
        frame4,
        textvariable=output_path_var,
        width=480,
        state="readonly"
    )
    path_entry.grid(row=2, column=0)
    ctk.CTkLabel(
        frame4,
        text="SEMUANYA HARUS DIISI",
        font=ctk.CTkFont(weight="bold")
    ).grid(row=3, column=0, sticky="s", pady=(5, 10))

    ctk.CTkButton(
        app,
        text="Run",
        command=save_to_config
    ).pack(pady=10)
    ctk.CTkButton(
        app,
        text="Ubah Otomatis (VBA)",
        command=run_vba_macro
    ).pack(pady=10)

    app.mainloop()

if __name__=="__main__":
    main()