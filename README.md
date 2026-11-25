# FlashCards

A vibe-coded, Python-based flashcard study app that automatically generates an Excel file with sample questions on first launch. Customize questions and answers directly in the spreadsheet to create personalized study sessions.

## Download and use (Windows)

1. Open the [latest release](https://github.com/ahrenarcangel/flashcards/releases/latest).
2. Download `FlashCards.exe` from the release assets.
3. Save it in a folder you can write to, then double-click it to launch the app.

No Python installation or additional dependencies are required to use the executable.

## Customize your questions

On first launch, FlashCards automatically creates a `flashcards.xlsx` file in the same folder as the executable, containing initial sample questions for testing.

1. Close the app and open `flashcards.xlsx` in Excel or a compatible spreadsheet editor.
2. Edit the sample questions and answers or add your own in the workbook, keeping the existing worksheet names and column headers.
3. Save and close the workbook, then reopen FlashCards to study with your updated questions.

## Run from source (developers)

If you want to modify the app or run it from Python, clone the repository and install the dependency:

```sh
python -m pip install -r requirements.txt
```

Launch the desktop app:

```sh
python flashcards_gui.py
```

Alternatively, launch the command-line version:

```sh
python flashcards.py
```

Excel deck files and generated build outputs are kept local and are not tracked by Git.
