import tkinter as tk
from tkinter import ttk
from services import XmlParserServices


def main():
    xml_parser_services = XmlParserServices()

    root = tk.Tk()
    root.title("Doc to Xml parser")

    # TextFields
    dir_text_field = tk.Entry(root, width=50)

    # Buttons
    stop_button = tk.Button(root, text="Stop", state=tk.DISABLED)
    start_button = tk.Button(root, text="Start")
    dir_button = tk.Button(root, text="File dir", command=lambda: xml_parser_services.choose_directory(dir_text_field))

    start_button.config(command=lambda: xml_parser_services.start_parsing_xml(stop_button, start_button, dir_text_field, progress_bar))
    stop_button.config(command=lambda: xml_parser_services.finish_parsing_xml(stop_button, start_button, dir_text_field, True))

    # Progress Bar
    global progress_bar
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=100, mode="determinate")

    # Arrange elements in the window
    dir_button.grid(row=0, column=0, padx=(10, 5), pady=10, sticky='w')
    dir_text_field.grid(row=0, column=1, padx=5, pady=10, sticky='ew')
    progress_bar.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky='ew')
    start_button.grid(row=2, column=1, padx=5, pady=10, sticky='se')
    stop_button.grid(row=2, column=0, padx=(10, 5), pady=10, sticky='sw')

    # Configure grid row and column weights for resizing
    root.grid_rowconfigure(1, weight=1)
    root.grid_columnconfigure(1, weight=1)

    # Main loop start
    root.mainloop()


if __name__ == "__main__":
    main()
