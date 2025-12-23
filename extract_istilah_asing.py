def run():
    from docx import Document
    from wordfreq import zipf_frequency
    import re
    import os
    import json
    from main import resource_path

    # ====== KONFIGURASI ======
    with open(resource_path("config.json"), "r", encoding="utf-8") as f:
        config = json.load(f)

    INPUT_DOCX = config["input_docx"]
    OUTPUT_DOCX = config["output_docx"]

    if not INPUT_DOCX:
        raise ValueError("No DOCX file selected")

    print("Processing:", INPUT_DOCX)     # ganti dengan nama file kamu
    base_name = os.path.splitext(os.path.basename(INPUT_DOCX))[0]
    OUTPUT_TXT2 = os.path.join(
        OUTPUT_DOCX,
        f"istilah_Ver2_{base_name}.txt"
    )
    OUTPUT_TXT = os.path.join(
        OUTPUT_DOCX,
        f"istilah_{base_name}.txt"
    )

    MIN_WORD_LENGTH = config["min_word_length"]



    doc = Document(INPUT_DOCX)
    texts = []
    start = True
    # count_bab3 = 0

    for p in doc.paragraphs:
    #     if "BAB I" in p.text.upper(): # hanya mulai mengindeks jika teks "BAB I" ditemukan untuk kedua kalinya
    #         count_bab3 += 1
    #         if count_bab3 == 2:
    #             start = True

        if start:
            texts.append(p.text)
    text = " ".join(texts)
    words = re.findall(r"\b[A-Za-z]{3,}\b", text)
    terms = set()

    for w in words:
        w = w.strip().lower()

        if len(w) < MIN_WORD_LENGTH:
            continue

        # Jika kata TIDAK umum dalam bahasa Indonesia
        # tapi umum dalam bahasa Inggris → dianggap istilah asing
        id_freq = zipf_frequency(w, "id")
        en_freq = zipf_frequency(w, "en")

        en_freq_Value = config["en_freq_Value"]
        id_freq_Value = config["id_freq_Value"]
        if en_freq > en_freq_Value and id_freq < id_freq_Value:
            terms.add(w)

    with open(OUTPUT_TXT, "w", encoding="utf-8") as f:
        for t in sorted(terms):
            f.write(t + "\n")

    print(f"Selesai. {len(terms)} istilah asing disimpan ke {OUTPUT_TXT}")

    with open(OUTPUT_TXT2, "w", encoding="utf-8") as f:
        f.write("terms = Array(")
        for t in sorted(terms):
            f.write(f' "{t}",')
        f.write(")")

    print(f"Selesai2. {len(terms)} istilah asing disimpan ke {OUTPUT_TXT2}")

    CONFIG_PATH = resource_path("config.json")
    existing_config = {}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        existing_config = json.load(f)

    istilah_path = OUTPUT_TXT
    existing_config["istilah_path"] = istilah_path

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(existing_config, f, indent=2)

if __name__ == "__main__":
    run()