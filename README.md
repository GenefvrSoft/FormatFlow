Clean_word_formatting.py

Removes all formatting from a Word document (.docx), leaving only plain text
with the default “Normal” style: no bold, italics, underlining, colors,
custom fonts, font sizes, highlighting, forced alignment, indents,
spacing, or borders. Works with paragraphs, tables (including nested ones), and,
optionally, headers and footers.

Usage (single file):
    python clean-word-formatting.py input.docx output.docx
    python clean-word-formatting.py input.docx output.docx --include-headers-footers
    python clean_word_formatting.py input.docx output.docx --keep-bold-italic
    python clean_word_formatting.py input.docx output.docx --remove-headers
    python clean_word_formatting.py input.docx output.docx --remove-text-boxes

Usage (multiple files at once):
    python clean_word_formatting.py input_folder/ output_folder/
    python clean_word_formatting.py input_folder/ output_folder/ --suffix _clean

    If “input” is a folder, it processes all .docx files it contains (it does not enter
    subfolders) and saves the results in “output” (created if it does not exist), with the
    same filename, or with the specified suffix if --suffix is used.
