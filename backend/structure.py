from pathlib import Path

files = [
    "main.py",
    "config.py",
    "excel_reader.py",
    "formatter.py",
    "word_generator.py",
    "pdf_generator.py",
    "logger.py",
    "README.md"
]

for file in files:
    Path(file).touch(exist_ok = True)

folders = [
    "data",
    "templates",
    "output",
    "output/docx",
    "output/pdf",
    "logs"
]

for folder in folders:
    Path(folder).mkdir(parents = True,exist_ok = True)


print("Structure created successfully")
