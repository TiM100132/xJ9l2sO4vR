import tkinter as tk
from tkinter import filedialog
from main import Analyzer
import time

def choose_file_and_run(root):
    file_path = filedialog.askopenfilename(parent=root)
    if file_path:
        return file_path

def create_gui():
    root = tk.Tk()
    root.title("Выбор файла для обработки")
    root.geometry("300x150")

    def run_analyzer():
        file_path = choose_file_and_run(root)
        if file_path:
            start = time.perf_counter()
            analyzer = Analyzer(file_path)
            analyzer.analyze()
            analyzer.file_handler.del_tmp_dir()
            end = time.perf_counter()
            print(f'Время обработки файла {round(end - start, 2)}')

    select_file_button = tk.Button(root, text="Выбрать файл", command=run_analyzer, width=20, height=2)
    select_file_button.pack(pady=30)
    
    exit_button = tk.Button(root, text="Выход", command=root.quit, width=20, height=2)
    exit_button.pack()

    root.mainloop()

if __name__ == "__main__":
    create_gui()