import tkinter as tk
from tkinter import ttk, filedialog
from folder_parser import process_folder_structure
from excel_writer import save_to_excel_form


class UserWindow:
    def __init__(self, root):
        self.root = root
        self.root.title('GSFX парсер')
        self.root.geometry('800x600')

        self.input_folder = ''
        self.output_folder = ''
        
        tk.Label(root, text='Выберите папку с объектами и папку для выходного документа', font=("Arial", 12, "bold")).pack(pady=5)

        ttk.Button(root, text="📂 Папка объектов", command=self.select_input_folder).pack()
        self.object_folder = tk.Label(root, text='Не выбрана', fg='gray')
        self.object_folder.pack()

        ttk.Button(root, text="📂 Папка сохранения документа", command=self.select_output_folder).pack()
        self.save_folder = tk.Label(root, text='Не выбрана', fg='gray')
        self.save_folder.pack()

        tk.Label(root, text='Имя выходного файла:').pack(pady=(10, 0))
        self.entry = tk.Entry(root, width=50)
        self.entry.insert(0, 'Отчет')
        self.entry.pack(pady=(0, 10))
        
        tk.Button(root, text="▶ Запустить", bg="green", fg="white", command=self.run_fill).pack(pady=15)


    def select_input_folder(self):
        path = filedialog.askdirectory(title="Выбор папки с объектами")
        if path:
            self.input_folder = path
            self.object_folder.config(text=path, fg="black")

    def select_output_folder(self):
        path = filedialog.askdirectory(title="Папка для сохранения")
        if path:
            self.output_folder = path
            self.save_folder.config(text=path, fg="black")

    def run_fill(self):
        try:
            input_path = self.input_folder
            output_path = self.output_folder
            filename = self.entry.get().strip()

            if not input_path or not output_path or not filename:
                tk.messagebox.showerror("Ошибка", "Выберите все папки и имя файла")
                return
            full_output_path = f'{output_path}/{filename}.xlsx'
            smeta_data, contractor_totals, object_totals = process_folder_structure(input_path)
            save_to_excel_form(smeta_data, contractor_totals, object_totals, full_output_path)

            tk.messagebox.showinfo("Готово ✅", f"Файл успешно сохранён:\n{full_output_path}")
            print(f'\n✅ Excel-файл успешно создан: {full_output_path}')

        except Exception as e:
            print(f"❌ Ошибка выполнения: {e}")
            tk.messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")