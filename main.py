from pathlib import Path
import win32com.client

input_dir = Path(
    "C:\\Users\\sawan\\OneDrive\\Desktop\\Assignment_project\\Chnagers"
)  # Directly specify the input directory
output_dir = (
    input_dir.parent / "output"
)  # Define the output directory as a subdirectory of the input directory

output_dir.mkdir(parents=True, exist_ok=True)

find = "Atharva"
replace = "Sawant"
wd_replace = 2
wd_find_wrap = 1

word_app = win32com.client.DispatchEx("Word.Application")
word_app.Visible = False
word_app.DisplayAlerts = False

for doc_file in input_dir.glob("*.*"):
    doc_path = str(doc_file)
    doc = word_app.Documents.Open(doc_path)
    word_app.Selection.Find.Execute(
        FindText=find,
        ReplaceWith=replace,
        Replace=wd_replace,
        Wrap=wd_find_wrap,
        Forward=True,
        MatchCase=True,
        MatchWholeWord=False,
        MatchWildcards=True,
        MatchSoundsLike=False,
        MatchAllWordForms=False,
        Format=True,
    )

    doc.Save()  # Save the changes
    doc.Close()  # Close the document

    output_path = output_dir / f"{doc_file.stem}_replaced{doc_file.suffix}"


word_app.Application.Quit()
