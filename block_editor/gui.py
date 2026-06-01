# gui.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
from autocad_handler import AutoCADHandler


class BlockEditorGUI:
    def __init__(self):
        self.handler = AutoCADHandler()
        self.current_block = None
        self.all_properties = []
        self.widgets = {}
        self.original_values = {}  # Для отслеживания изменений

        self.root = tk.Tk()
        self.root.title("Редактор динамических блоков AutoCAD")
        self.root.geometry("1120x800")
        self.root.configure(bg='#1e1e1e')

        self.create_widgets()

    def create_widgets(self):
        tk.Button(self.root, text="🔄 Выбрать блок в AutoCAD", command=self.select_block,
                  bg='#007acc', fg='white', font=('Segoe UI', 11, 'bold'), height=2).pack(fill='x', padx=10, pady=8)

        # Основные характеристики
        self.main_frame = tk.LabelFrame(self.root, text=" Основные характеристики ", fg='#4fc3f7', bg='#1e1e1e')
        self.main_frame.pack(fill='x', padx=10, pady=6)

        self.main_vars = {}
        main_params = ['Height', 'Depth', 'Width', 'ChangeHeight', 'ChangeWidth', 'ChangeDepth']

        for param in main_params:
            frame = tk.Frame(self.main_frame, bg='#1e1e1e')
            frame.pack(fill='x', padx=10, pady=5)
            tk.Label(frame, text=param, bg='#1e1e1e', fg='#cccccc', width=18, anchor='w').pack(side='left')
            var = tk.StringVar()
            entry = ttk.Entry(frame, textvariable=var, width=32)
            entry.pack(side='left', padx=5)
            self.main_vars[param] = (var, entry)

        # Кнопки
        btn_frame = tk.Frame(self.root, bg='#1e1e1e')
        btn_frame.pack(fill='x', padx=10, pady=8)

        tk.Button(btn_frame, text="▶ Все параметры", command=self.toggle_extra,
                  bg='#37474f', fg='white').pack(side='left', padx=5)

        tk.Button(btn_frame, text="✅ Применить все изменения", command=self.apply_all,
                  bg='#2e7d32', fg='white', font=('Segoe UI', 10, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="💾 Сохранить пресет", command=self.save_preset,
                  bg='#f57c00', fg='white').pack(side='left', padx=5)
        tk.Button(btn_frame, text="📂 Загрузить пресет", command=self.load_preset,
                  bg='#f57c00', fg='white').pack(side='left', padx=5)

        tk.Button(btn_frame, text="💾 Сохранить чертёж", command=self.save_drawing,
                  bg='#d32f2f', fg='white').pack(side='right', padx=5)

        # Дополнительные параметры
        self.extra_frame = tk.LabelFrame(self.root, text=" Все параметры блока ", fg='#ffab40', bg='#1e1e1e')

        self.canvas = tk.Canvas(self.extra_frame, bg='#1e1e1e', highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.extra_frame, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas, bg='#1e1e1e')

        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.canvas.pack(side="left", fill="both", expand=True, padx=8)
        self.scrollbar.pack(side="right", fill="y")

        self.status_label = tk.Label(self.root, text="Готов к работе", bg='#1e1e1e', fg='#66ff99',
                                     font=('Segoe UI', 10))
        self.status_label.pack(pady=8)

    def _on_mousewheel(self, event):
        if self.extra_frame.winfo_manager():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def toggle_extra(self):
        if self.extra_frame.winfo_manager():
            self.extra_frame.pack_forget()
        else:
            self.extra_frame.pack(fill='both', expand=True, padx=10, pady=5)

    def get_combo_options(self, prop_name, current_value):
        str_val = str(current_value).lower()
        if prop_name == "Doors" or "door" in str_val or "Doors_n_Accent" in prop_name:
            return ["1 door", "2 doors"]
        if prop_name == "opening_side" or "Side" in prop_name:
            return ["0", "1"] if str_val in ["0", "1"] else ["Left", "Right"]
        return None

    def select_block(self):
        self.current_block = self.handler.get_selected_block()
        if not self.current_block:
            messagebox.showwarning("Внимание", "Выберите динамический блок в AutoCAD")
            return

        self.all_properties = self.handler.get_all_properties(self.current_block)
        self.original_values = {p['name']: str(p['value']) for p in self.all_properties}

        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self.widgets.clear()

        self.update_interface()
        self.status_label.config(text=f"Блок: {getattr(self.current_block, 'EffectiveName', 'Unknown')}", fg='#66ff99')

    def update_interface(self):
        prop_dict = {p['name']: p for p in self.all_properties}

        for name, (var, _) in self.main_vars.items():
            if name in prop_dict:
                var.set(str(prop_dict[name]['value']))

        for prop in self.all_properties:
            frame = tk.Frame(self.scroll_frame, bg='#1e1e1e')
            frame.pack(fill='x', padx=10, pady=4)

            tk.Label(frame, text=prop['name'], width=28, anchor='w', bg='#1e1e1e', fg='#e0e0e0').pack(side='left')

            options = self.get_combo_options(prop['name'], prop['value'])

            if options:
                combo = ttk.Combobox(frame, values=options, width=30, state="readonly")
                combo.set(str(prop['value']))
                combo.pack(side='left', padx=5)
                if not prop['readonly']:
                    tk.Button(frame, text="Применить", bg='#2e7d32', fg='white',
                              command=lambda p=prop, c=combo: self.apply_combo(p, c)).pack(side='right')
                self.widgets[prop['name']] = combo
            else:
                var = tk.StringVar(value=str(prop['value']))
                entry = ttk.Entry(frame, textvariable=var, width=35)
                entry.pack(side='left', padx=5)
                if not prop['readonly']:
                    tk.Button(frame, text="Применить", bg='#2e7d32', fg='white',
                              command=lambda p=prop, v=var: self.apply_single(p, v)).pack(side='right')
                self.widgets[prop['name']] = var

    def _apply_property(self, prop, new_value):
        """Применяет только один параметр"""
        try:
            if str(new_value) == self.original_values.get(prop['name'], None):
                return  # ничего не изменилось

            if self.handler.set_property(prop, new_value):
                self.current_block.Update()
                self.handler.doc.Regen(2)
                self.handler.doc.Regen(1)
                self.original_values[prop['name']] = str(new_value)
                self.status_label.config(text=f"✅ {prop['name']} применён", fg='#66ff99')
            else:
                self.status_label.config(text=f"❌ Ошибка: {prop['name']}", fg='red')
        except Exception as e:
            self.status_label.config(text=f"❌ Ошибка: {prop['name']}", fg='red')
            print(f"Ошибка применения {prop['name']}: {e}")

    def apply_single(self, prop, var):
        self._apply_property(prop, var.get())

    def apply_combo(self, prop, combo):
        self._apply_property(prop, combo.get())

    def apply_all(self):
        """Применяет только изменённые параметры"""
        if not self.current_block:
            return
        count = 0
        for prop in self.all_properties:
            if prop['readonly']:
                continue
            if prop['name'] in self.main_vars:
                val = self.main_vars[prop['name']][0].get()
            elif prop['name'] in self.widgets:
                widget = self.widgets[prop['name']]
                val = widget.get() if hasattr(widget, 'get') else str(widget)
            else:
                continue

            # Применяем только если значение изменилось
            if str(val) != self.original_values.get(prop['name'], None):
                self._apply_property(prop, val)
                count += 1

        if count == 0:
            self.status_label.config(text="Нет изменений для применения", fg='#ffcc00')
        else:
            self.status_label.config(text=f"✅ Применено {count} изменений", fg='#66ff99')

    def save_preset(self):
        if not self.all_properties: return
        data = {p['name']: p['value'] for p in self.all_properties}
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            self.status_label.config(text="Пресет сохранён", fg='#66ff99')

    def load_preset(self):
        if not self.current_block: return
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            for prop in self.all_properties:
                if prop['name'] in data and not prop['readonly']:
                    self.handler.set_property(prop, data[prop['name']])
            self.apply_all()

    def save_drawing(self):
        if self.handler.save_drawing():
            self.status_label.config(text="💾 Чертёж сохранён", fg='#66ff99')
        else:
            self.status_label.config(text="❌ Ошибка сохранения", fg='red')

    def run(self):
        self.root.mainloop()