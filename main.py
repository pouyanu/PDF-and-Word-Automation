#PDF
import re
import subprocess
import fitz
from tkinter import filedialog, messagebox
import os
from datetime import datetime
from tkinter import ttk
import openpyxl
import tkinter as tk
from docx import Document



highlight_values = {}
def highlight_pdf_words():

    entry_widgets = []
    root = tk.Tk()
    row_counter = [1]

    root.title("PDF Word Highlighter")
    window_width = 800
    window_height = 450
    root.geometry(f"{window_width}x{window_height}")

    # Calculate the center position
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2

    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    # Create 20 entries and dropdowns in two columns
    column_changed = False
    row = 0
    col = 0
    entry_widgets = []

    # Initialize color variables list
    color_vars = []

    column_changed = False
    col = 0
    row = 0

    for i in range(20):
        if not column_changed and i > 9:
            col = 1
            row = 0
            column_changed = True

        # Create an entry for the word
        word_label = tk.Label(root, text=f"Entry {i + 1}:")
        word_label.grid(row=row, column=col * 3, padx=10, pady=10)
        word_entry = tk.Entry(root)
        word_entry.grid(row=row, column=col * 3 + 1, padx=10, pady=10)

        # Create a dropdown for selecting color
        color_values = ["Yellow", "Red", "Green", "Blue", "Orange", "Purple", "Pink", "Teal", "Brown",
                       "Cyan", "Magenta", "Lime", "Turquoise", "Indigo", "SlateGray", "DarkOliveGreen",
                       "DarkSlateBlue", "Tomato", "Chocolate"]

        color_dropdown = ttk.Combobox(root, values=color_values, state="readonly")
        color_dropdown.set(color_values[0])  # Set default color
        color_dropdown.grid(row=row, column=col * 3 + 2, padx=10, pady=10)

        # Add color variable to the list
        color_vars.append(color_dropdown)

        row += 1
        entry_widgets.append(word_entry)

    def clear_fields():
        for i in range(20):
            text_value = entry_widgets[i].get()
            color_value = color_vars[i].get()

            # Clear the entry widget
            entry_widgets[i].delete(0, tk.END)

            # Reset the color to "Yellow"
            color_vars[i].set("Yellow")

            if text_value:
                print(f"Entry {i + 1}: Text: {text_value}, Color: {color_value}")
        run_highlights = {}


    def run_highlights(entry_widgets, color_vars):
        for i in range(20):

            text_value = entry_widgets[i].get()
            color_value = color_vars[i].get()
            if text_value:
                highlight_values[text_value] = color_value

        if highlight_values:
            run_highlight()
        else:
            messagebox.showinfo("Error", "Please enter values")

    # Create a button to highlight the word
    highlight_button = tk.Button(root, text="Highlight Word", command=lambda: run_highlights(entry_widgets, color_vars))
    highlight_button.grid(row=row, column=0, columnspan=6, pady=10)

    clear_button = tk.Button(root, text="Reset", command=clear_fields)
    clear_button.grid(row=row, column=2, columnspan=6, pady=10)


    def run_highlight():
        def select_pdf_file():
            root = tk.Tk()
            root.withdraw()
            file_path = filedialog.askopenfilename(title="Select a PDF file", filetypes=[("PDF files", "*.pdf")])
            return file_path


        def create_output_folder_if_not_exists():
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            output_folder_path = os.path.join(desktop_path, 'PDF file output')

            if not os.path.exists(output_folder_path):
                os.makedirs(output_folder_path)
                print(f"Created 'file output' folder on the desktop.")


            today_folder = datetime.now().strftime('%Y-%m-%d')
            today_folder_path = os.path.join(output_folder_path, today_folder)
            print(today_folder_path)

            if not os.path.exists(today_folder_path):
                os.makedirs(today_folder_path)
                print(f"Created folder for today's date: {today_folder}")

            return today_folder_path


        def highlight_word(page, word, pdf_document, color_name):

            color_dict = {
                'Red': (1, 0, 0),
                'Green': (0, 1, 0),
                'Blue': (0, 0, 1),
                'Yellow': (1, 1, 0),
                'Orange': (1, 0.5, 0),
                'Purple': (0.5, 0, 0.5),
                'Pink': (1, 0.5, 0.5),
                'Teal': (0, 0.5, 0.5),
                'Brown': (0.6, 0.3, 0),
                'Cyan': (0, 1, 1),
                'Magenta': (1, 0, 1),
                'Lime': (0.5, 1, 0),
                'Turquoise': (0, 1, 0.8),
                'Indigo': (0.29, 0, 0.51),
                'SlateGray': (0.44, 0.5, 0.56),
                'DarkOliveGreen': (0.33, 0.42, 0.18),
                'DarkSlateBlue': (0.28, 0.24, 0.55),
                'Tomato': (1, 0.39, 0.28),
                'Chocolate': (0.82, 0.41, 0.12)
            }

            color_rgb = color_dict.get(color_name)

            rect_list = page.search_for(word)
            text = page.get_text()
            words = text.split()
            exact_matches = []
            search_string = "CONFIRMATION"

            def get_letter_width(page):
                sample_text = 'X'  # You can use any character or string for measurement
                char_width = page.get_text('text', clip=(0, 0, 100, 100), clip_func='bbox').splitlines()[0]
                return len(char_width) / len(sample_text)

            # letter_width = get_letter_width(page)


            for word in words:
                if word == search_string:
                    exact_matches.append(word)

            printed_lines = set()

            if rect_list:

                for i, rect in enumerate(rect_list):



                    highlight = page.add_highlight_annot(rect_list)
                    highlight.set_colors(stroke=color_rgb)
                    highlight.set_opacity(0.2)

                    # Extract and print the whole line if not already printed
                    page_text = page.get_text("text")
                    lines = page_text.split('\n')

                    for line in lines:
                        if word.lower() in line.lower() and line not in printed_lines:

                            printed_lines.add(line)  # Mark the line as printed
                            break
            else:
                print("Error", f"Did not find any occurrences of the word: {word}")

            first_word_found = False

            page = pdf_document[0]

            # Get the text of the first page
            page_text = page.get_text("text")

            # Split the text into words
            words = page_text.split()

            if words:
                first_word = words[0]

                # Check if it's the first occurrence of a word
                if not first_word_found:

                    first_rect = page.search_for(first_word)[0]

                    first_highlight = page.add_rect_annot(first_rect)
                    first_highlight.set_opacity(0.4)
                    first_highlight.set_colors(stroke=(0, 1, 0), fill=(0, 1, 0))
                    page.delete_annot(first_highlight)
                    first_word_found = True

            if not first_word_found:
                print("Did not find any words on the first page")


        def print_pdf_content_and_highlight(pdf_path):
            pdf_document = fitz.open(pdf_path)

            for page_num in range(pdf_document.page_count):

                page = pdf_document[page_num]
                for word, color_name in highlight_values.items():

                    highlight_word(page, word, pdf_document, color_name)

            today_folder_path = create_output_folder_if_not_exists()
            file_name = os.path.splitext(os.path.basename(pdf_path))[0]

            output_path = os.path.join(today_folder_path, f"{file_name}.pdf")
            if os.path.exists(output_path):
                output_path = os.path.join(today_folder_path, f"{file_name}-Copy.pdf")

            pdf_document.save(output_path)
            pdf_document.close()

            response = messagebox.askyesno("Modified PDF saved at:",
                                           f"Modified PDF saved at: {output_path}\n\nWould you like to open the file?")

            if response:
                # Open the modified PDF using the default PDF viewer
                subprocess.Popen(['start', '', output_path], shell=True)


        def main():
            pdf_path = select_pdf_file()

            if pdf_path:
                print(f"Selected PDF file: {pdf_path}")
                print_pdf_content_and_highlight(pdf_path)


        if __name__ == "__main__":
            main()




    root.mainloop()
highlight_pdf_words()