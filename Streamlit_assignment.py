import win32com.client
import pythoncom
import streamlit as st
from pathlib import Path


def main():
    st.title("Assignment0.1")
    st.image(
        "https://raw.githubusercontent.com/Tarikul-Islam-Anik/Animated-Fluent-Emojis/master/Emojis/Smilies/Hundred%20Points.png",
        width=150,
    )
    path = st.text_input("Path to the folder")
    input_dir = Path(path)
    output_dir = input_dir.parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    wd_replace = 2
    wd_find_wrap = 1
    pythoncom.CoInitialize()
    word_app = win32com.client.DispatchEx("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = False
    find_input = st.text_input("Enter words to find (separated by comma):")
    replace_input = st.text_input("Enter replacement words (separated by comma):")

    if st.button("Replace"):
        find_words = [word.strip() for word in find_input.split(",")]
        replace_words = [word.strip() for word in replace_input.split(",")]
        if all(find_words) and all(replace_words):
            if len(find_words) != len(replace_words):
                st.error("Number of find words must match number of replace words.")
            else:
                try:
                    with st.spinner("Performing find and replace operation..."):
                        for find, replace in zip(find_words, replace_words):
                            st.write(f"Replacing '{find}' with '{replace}'")
                            for doc_file in input_dir.glob("*.*"):
                                try:
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
                                    output_path = (
                                        input_dir / f"{doc_file.stem}{doc_file.suffix}"
                                    )
                                    st.write(
                                        f"Saved replaced document to: {output_path}"
                                    )
                                except Exception as e:
                                    st.error(
                                        f"Error processing document {doc_file}: {e}"
                                    )
                    word_app.Application.Quit()
                    st.success(
                        "Find and replace operation completed assignments done!!."
                    )
                    st.balloons()
                except Exception as e:
                    st.error(
                        f"An error occurred during the find and replace operation: {e}"
                    )
        else:
            st.error("Please enter both find and replace words.")


if __name__ == "__main__":
    main()
