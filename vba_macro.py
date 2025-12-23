def run():
    import os
    import sys
    import json
    import win32com.client
    from main import resource_path


    with open(resource_path("config.json"), "r", encoding="utf-8") as f:
        config = json.load(f)


    with open(resource_path("vba_code.txt"), "r", encoding="utf-8") as f:
        vba_code = f.read()

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    INPUT_DOCX = os.path.normpath(config["input_docx"])

    if not os.path.exists(INPUT_DOCX):
        raise FileNotFoundError(INPUT_DOCX)
    doc = word.Documents.Open(INPUT_DOCX)
    # # 
    # doc = word.Documents.Add()

    vb_project = doc.VBProject
    vb_module = vb_project.VBComponents.Add(1)  # standard module
    vb_module.CodeModule.AddFromString(vba_code)


    istilah_path = config["istilah_path"]
    file_path = istilah_path
    word.Run("Italicize_From_Txt_Script", file_path)

if __name__ == "__main__":
    run()