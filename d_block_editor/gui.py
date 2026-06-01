import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
from autocad_handler import AutoCADHandler


class BlockEditorGUI:
    def __init__(self):
        self.handler = AutoCADHandler()
        self.current_blocks = []          # для группового режима
        self.block_infos = []
        self.single_block = None
        self.single_properties = []
        self.single_all_params_visible = False

        self.root = tk.Tk()
        self.root.title("🏠 Редактор кухонных блоков AutoCAD")
        self.root.geometry("1280x900")
        self.root.configure(bg='#1e1e1e')

        self.create_widgets()
        self.update_status("Готов. Подключено к AutoCAD.")

    def create_widgets(self):
        # Верхняя панель
        top_frame = tk.Frame(self.root, bg='#1e1e1e')
        top_frame.pack(fill='x', padx=10, pady=8)

        self.refresh_btn = tk.Button(top_frame, text="🔄 Обновить выделение из AutoCAD",
                                     command=self.refresh_selection,
                                     bg='#007acc', fg='white', font=('Segoe UI', 11, 'bold'), height=1)
        self.refresh_btn.pack(side='left', padx=5)

        self.reset_btn = tk.Button(top_frame, text="❌ Сброс", command=self.reset_active_tab,
                                   bg='#ff5722', fg='white')
        self.reset_btn.pack(side='left', padx=5)

        self.save_dwg_btn = tk.Button(top_frame, text="💾 Сохранить чертёж", command=self.save_drawing,
                                      bg='#d32f2f', fg='white')
        self.save_dwg_btn.pack(side='right', padx=5)

        # Вкладки
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)

        # --- Вкладка 1: Групповое управление ---
        self.group_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.group_tab, text="📊 Группа (несколько блоков)")

        # Таблица выбранных блоков
        self.tree_frame = tk.Frame(self.group_tab, bg='#1e1e1e')
        self.tree_frame.pack(fill='both', expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(self.tree_frame, columns=('type', 'width', 'height', 'changeW', 'changeD'),
                                 show='headings', height=12)
        self.tree.heading('type', text='Тип')
        self.tree.heading('width', text='Ширина (мм)')
        self.tree.heading('height', text='Высота (мм)')
        self.tree.heading('changeW', text='Менять ширину')
        self.tree.heading('changeD', text='Менять высоту')
        self.tree.column('type', width=150)
        self.tree.column('width', width=100)
        self.tree.column('height', width=100)
        self.tree.column('changeW', width=100)
        self.tree.column('changeD', width=100)

        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        # Статистика
        stats_frame = tk.LabelFrame(self.group_tab, text="📐 Текущие размеры кухни", bg='#1e1e1e', fg='#4fc3f7')
        stats_frame.pack(fill='x', padx=5, pady=5)

        self.bottom_total_label = tk.Label(stats_frame, text="Длина нижних блоков: 0.0", bg='#1e1e1e', fg='#cccccc')
        self.bottom_total_label.pack(anchor='w', padx=10, pady=2)
        self.top_total_label = tk.Label(stats_frame, text="Длина верхних блоков: 0.0", bg='#1e1e1e', fg='#cccccc')
        self.top_total_label.pack(anchor='w', padx=10, pady=2)
        self.change_depth_heights_label = tk.Label(stats_frame, text="Высоты блоков с ChangeDepth:", bg='#1e1e1e', fg='#cccccc')
        self.change_depth_heights_label.pack(anchor='w', padx=10, pady=2)

        # Панель групповых настроек
        group_ctrl = tk.LabelFrame(self.group_tab, text="🎯 Групповое изменение", bg='#1e1e1e', fg='#ffab40')
        group_ctrl.pack(fill='x', padx=5, pady=10)

        # all_ширина (нижние)
        w_bottom_frame = tk.Frame(group_ctrl, bg='#1e1e1e')
        w_bottom_frame.pack(fill='x', padx=10, pady=8)
        tk.Label(w_bottom_frame, text="all_ширина (нижние блоки):", bg='#1e1e1e', fg='white', width=25, anchor='w').pack(side='left')
        self.target_width_bottom_var = tk.StringVar()
        ttk.Entry(w_bottom_frame, textvariable=self.target_width_bottom_var, width=12).pack(side='left', padx=10)
        self.apply_width_bottom_btn = tk.Button(w_bottom_frame, text="📏 Применить ширину (низ)",
                                                command=self.apply_target_width_bottom,
                                                bg='#2e7d32', fg='white')
        self.apply_width_bottom_btn.pack(side='left', padx=5)

        # all_ширина (верхние) + чекбокс синхронизации
        w_top_frame = tk.Frame(group_ctrl, bg='#1e1e1e')
        w_top_frame.pack(fill='x', padx=10, pady=8)
        tk.Label(w_top_frame, text="all_ширина (верхние блоки):", bg='#1e1e1e', fg='white', width=25, anchor='w').pack(side='left')
        self.target_width_top_var = tk.StringVar()
        ttk.Entry(w_top_frame, textvariable=self.target_width_top_var, width=12).pack(side='left', padx=10)
        self.apply_width_top_btn = tk.Button(w_top_frame, text="📏 Применить ширину (верх)",
                                             command=self.apply_target_width_top,
                                             bg='#2e7d32', fg='white')
        self.apply_width_top_btn.pack(side='left', padx=5)

        self.sync_width_var = tk.BooleanVar(value=False)
        self.sync_width_check = tk.Checkbutton(w_top_frame, text="Синхронно (низ+верх)",
                                               variable=self.sync_width_var,
                                               bg='#1e1e1e', fg='white', selectcolor='#1e1e1e')
        self.sync_width_check.pack(side='left', padx=10)

        # all_длинна (высота)
        h_frame = tk.Frame(group_ctrl, bg='#1e1e1e')
        h_frame.pack(fill='x', padx=10, pady=8)
        tk.Label(h_frame, text="all_длинна (высота для ChangeDepth=1):", bg='#1e1e1e', fg='white', width=25, anchor='w').pack(side='left')
        self.target_height_var = tk.StringVar()
        ttk.Entry(h_frame, textvariable=self.target_height_var, width=12).pack(side='left', padx=10)
        self.apply_height_btn = tk.Button(h_frame, text="⬆️ Применить высоту", command=self.apply_target_height,
                                          bg='#2e7d32', fg='white')
        self.apply_height_btn.pack(side='left', padx=5)

        # --- Вкладка 2: Одиночный блок ---
        self.single_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.single_tab, text="🔧 Один блок (свойства)")

        self.select_single_btn = tk.Button(self.single_tab, text="🎯 Выбрать блок в AutoCAD", command=self.select_single_block,
                                           bg='#007acc', fg='white', font=('Segoe UI', 10, 'bold'))
        self.select_single_btn.pack(pady=10)

        self.toggle_all_params_btn = tk.Button(self.single_tab, text="📋 Показать все параметры", command=self.toggle_all_params,
                                               bg='#37474f', fg='white')
        self.toggle_all_params_btn.pack(pady=5)

        # Фрейм для свойств одного блока (скроллируемый)
        self.single_canvas_frame = tk.Frame(self.single_tab, bg='#1e1e1e')
        self.single_canvas_frame.pack(fill='both', expand=True, padx=5, pady=5)

        self.single_canvas = tk.Canvas(self.single_canvas_frame, bg='#1e1e1e', highlightthickness=0)
        self.single_scrollbar = ttk.Scrollbar(self.single_canvas_frame, orient="vertical", command=self.single_canvas.yview)
        self.single_scroll_frame = tk.Frame(self.single_canvas, bg='#1e1e1e')

        self.single_scroll_frame.bind("<Configure>", lambda e: self.single_canvas.configure(scrollregion=self.single_canvas.bbox("all")))
        self.single_canvas.create_window((0, 0), window=self.single_scroll_frame, anchor="nw")
        self.single_canvas.configure(yscrollcommand=self.single_scrollbar.set)
        self.single_canvas.bind_all("<MouseWheel>", self._on_mousewheel_single)

        self.single_canvas.pack(side="left", fill="both", expand=True)
        self.single_scrollbar.pack(side="right", fill="y")

        # Статусная строка
        self.status_label = tk.Label(self.root, text="Готов", bg='#1e1e1e', fg='#66ff99', font=('Segoe UI', 10))
        self.status_label.pack(side='bottom', fill='x', pady=5)

    def _on_mousewheel_single(self, event):
        self.single_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ---------- Общий сброс ----------
    def reset_active_tab(self):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0:   # вкладка "Группа"
            self.reset_group()
        else:                   # вкладка "Один блок"
            self.reset_single()

    def reset_group(self):
        """Очищает все данные группового режима"""
        self.current_blocks = []
        self.block_infos = []
        self.bottom_blocks = []
        self.top_blocks = []
        self.change_depth_blocks = []
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.bottom_total_label.config(text="Длина нижних блоков: 0.0")
        self.top_total_label.config(text="Длина верхних блоков: 0.0")
        self.change_depth_heights_label.config(text="Высоты блоков с ChangeDepth:")
        self.target_width_bottom_var.set("")
        self.target_width_top_var.set("")
        self.target_height_var.set("")
        self.update_status("Групповые данные сброшены")

    def reset_single(self):
        """Очищает данные одиночного блока"""
        self.single_block = None
        self.single_properties = []
        self.single_widgets = {}
        # Очищаем фрейм со свойствами
        for w in self.single_scroll_frame.winfo_children():
            w.destroy()
        tk.Label(self.single_scroll_frame, text="Блок не выбран. Нажмите «Выбрать блок в AutoCAD».",
                 bg='#1e1e1e', fg='red').pack()
        self.update_status("Режим одного блока сброшен")

    # ---------- Групповой режим ----------
    def refresh_selection(self):
        self.current_blocks = self.handler.get_selected_blocks()
        if not self.current_blocks:
            messagebox.showinfo("Нет блоков", "Выделите хотя бы один динамический блок в AutoCAD.")
            self.update_status("Выделение пусто")
            self.reset_group()
            return

        self.block_infos = [self.handler.get_block_info(b) for b in self.current_blocks]
        self.update_tree()
        self.update_statistics()
        self.update_status(f"Загружено блоков: {len(self.block_infos)}")

    def update_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for info in self.block_infos:
            type_str = []
            if info['isBottom']:
                type_str.append("Нижний")
            if info['isTop']:
                type_str.append("Верхний")
            if not type_str:
                type_str.append("?")
            type_str = "+".join(type_str)
            self.tree.insert('', 'end', values=(
                f"{info['name']} ({type_str})",
                f"{info['width']:.1f}",
                f"{info['height']:.1f}",
                "Да" if info['changeWidth'] else "Нет",
                "Да" if info['changeDepth'] else "Нет"
            ))

    def update_statistics(self):
        self.bottom_blocks = [b for b in self.current_blocks if self.handler.get_is_bottom(b)]
        self.top_blocks = [b for b in self.current_blocks if self.handler.get_is_top(b)]
        self.change_depth_blocks = [b for b in self.current_blocks if self.handler.has_change_depth_flag(b)]

        bottom_total = self.handler.get_total_width(self.bottom_blocks)
        top_total = self.handler.get_total_width(self.top_blocks)

        self.bottom_total_label.config(text=f"📏 Длина нижних блоков: {bottom_total:.1f} мм")
        self.top_total_label.config(text=f"📏 Длина верхних блоков: {top_total:.1f} мм")

        heights_info = ", ".join(f"{self.handler.get_height(b):.1f}" for b in self.change_depth_blocks) or "нет"
        self.change_depth_heights_label.config(text=f"⬆️ Высоты блоков с ChangeDepth: {heights_info} мм")

    def apply_target_width_bottom(self):
        # Если синхронизация включена, проверяем/заполняем поле верхних
        if self.sync_width_var.get():
            self._sync_width_fields(from_bottom=True)
        self._apply_target_width_to_blocks(self.bottom_blocks, self.target_width_bottom_var.get(), "нижних")

    def apply_target_width_top(self):
        if self.sync_width_var.get():
            self._sync_width_fields(from_bottom=False)
        self._apply_target_width_to_blocks(self.top_blocks, self.target_width_top_var.get(), "верхних")

    def _sync_width_fields(self, from_bottom):
        """При синхронизации копирует значение из заполненного поля в пустое"""
        bottom_val = self.target_width_bottom_var.get().strip()
        top_val = self.target_width_top_var.get().strip()
        if from_bottom:
            if bottom_val and not top_val:
                self.target_width_top_var.set(bottom_val)
            elif not bottom_val and top_val:
                self.target_width_bottom_var.set(top_val)
            elif not bottom_val and not top_val:
                messagebox.showwarning("Нет значений", "Введите хотя бы одно значение ширины для синхронного применения")
                raise ValueError("Empty fields")
        else:
            if top_val and not bottom_val:
                self.target_width_bottom_var.set(top_val)
            elif not top_val and bottom_val:
                self.target_width_top_var.set(bottom_val)
            elif not top_val and not bottom_val:
                messagebox.showwarning("Нет значений", "Введите хотя бы одно значение ширины для синхронного применения")
                raise ValueError("Empty fields")

    def _apply_target_width_to_blocks(self, blocks, target_str, level_name):
        if not blocks:
            messagebox.showwarning(f"Нет {level_name} блоков", f"Нет блоков с isBottom/isTop для {level_name}")
            return
        try:
            target = float(target_str)
        except:
            messagebox.showerror("Ошибка", f"Введите число для all_ширина ({level_name})")
            return

        current_total = self.handler.get_total_width(blocks)
        if target <= current_total:
            messagebox.showinfo("Ничего не делаем", f"Целевая ширина {target} не больше текущей {current_total:.1f}. Изменения не требуются.")
            return

        changeable = [b for b in blocks if self.handler.has_change_width_flag(b)]
        if not changeable:
            messagebox.showwarning(f"Нет расширяемых {level_name} блоков", f"Среди {level_name} блоков нет ни одного с ChangeWidth=1")
            return

        current_changeable_total = sum(self.handler.get_width(b) for b in changeable)
        increase = target - current_total
        increment_per_block = increase / len(changeable)
        for b in changeable:
            new_w = self.handler.get_width(b) + increment_per_block
            self.handler.set_width(b, new_w)

        self.handler.doc.Regen(1)
        self.update_status(f"✅ Ширина {level_name} блоков увеличена до {target:.1f} (добавлено {increase:.1f})")
        self.refresh_selection()

    def apply_target_height(self):
        if not self.change_depth_blocks:
            messagebox.showwarning("Нет блоков", "Нет блоков с ChangeDepth=1")
            return
        try:
            target = float(self.target_height_var.get())
        except:
            messagebox.showerror("Ошибка", "Введите число для all_длинна")
            return

        count = 0
        for b in self.change_depth_blocks:
            cur_h = self.handler.get_height(b)
            if target > cur_h:
                self.handler.set_height(b, target)
                count += 1
        if count:
            self.handler.doc.Regen(1)
            self.update_status(f"✅ Для {count} блоков высота увеличена до {target}")
        else:
            self.update_status(f"Высота {target} не превышает текущие значения, изменений нет")
        self.refresh_selection()

    # ---------- Режим одного блока ----------
    def select_single_block(self):
        blocks = self.handler.get_selected_blocks()
        if len(blocks) != 1:
            messagebox.showwarning("Неверное выделение", "Выделите ровно один динамический блок")
            return
        self.single_block = blocks[0]
        self.single_properties = self.get_all_properties_single(self.single_block)
        self.build_single_editor()

    def get_all_properties_single(self, block):
        properties = []
        try:
            result = block.GetDynamicBlockProperties()
            if isinstance(result, tuple):
                dyn_props = result
            else:
                dyn_props = [result.Item(i) for i in range(result.Count)]
            for prop in dyn_props:
                properties.append({
                    'name': prop.PropertyName,
                    'value': prop.Value,
                    'object': prop,
                    'readonly': prop.ReadOnly
                })
        except Exception as e:
            self.update_status(f"Ошибка получения свойств: {e}")
        return properties

    def toggle_all_params(self):
        self.single_all_params_visible = not self.single_all_params_visible
        if self.single_all_params_visible:
            self.toggle_all_params_btn.config(text="📋 Скрыть все параметры")
        else:
            self.toggle_all_params_btn.config(text="📋 Показать все параметры")
        self.build_single_editor()

    def build_single_editor(self):
        for w in self.single_scroll_frame.winfo_children():
            w.destroy()

        if not self.single_block:
            tk.Label(self.single_scroll_frame, text="Блок не выбран. Нажмите «Выбрать блок в AutoCAD».",
                     bg='#1e1e1e', fg='red').pack()
            return

        tk.Label(self.single_scroll_frame, text=f"Блок: {getattr(self.single_block, 'EffectiveName', 'Unknown')}",
                 bg='#1e1e1e', fg='#ffab40', font=('Segoe UI', 12, 'bold')).pack(pady=5)

        self.single_widgets = {}

        core_params = {'Width', 'Length', 'Height', 'Depth', 'isTop', 'isBottom', 'ChangeWidth', 'ChangeDepth'}
        for prop in self.single_properties:
            if not self.single_all_params_visible and prop['name'] not in core_params:
                continue

            frame = tk.Frame(self.single_scroll_frame, bg='#1e1e1e')
            frame.pack(fill='x', padx=10, pady=4)

            tk.Label(frame, text=prop['name'], width=28, anchor='w', bg='#1e1e1e', fg='#e0e0e0').pack(side='left')

            options = None
            str_val = str(prop['value']).lower()
            if "door" in prop['name'].lower() or "doors" in str_val:
                options = ["1 door", "2 doors"]
            elif "side" in prop['name'].lower():
                options = ["0", "1"] if str_val in ["0", "1"] else ["Left", "Right"]

            if options and not prop['readonly']:
                combo = ttk.Combobox(frame, values=options, width=30, state="readonly")
                combo.set(str(prop['value']))
                combo.pack(side='left', padx=5)
                tk.Button(frame, text="Применить", bg='#2e7d32', fg='white',
                          command=lambda p=prop, c=combo: self.apply_single_property(p, c.get())).pack(side='right')
                self.single_widgets[prop['name']] = combo
            else:
                var = tk.StringVar(value=str(prop['value']))
                entry = ttk.Entry(frame, textvariable=var, width=35)
                entry.pack(side='left', padx=5)
                if not prop['readonly']:
                    tk.Button(frame, text="Применить", bg='#2e7d32', fg='white',
                              command=lambda p=prop, v=var: self.apply_single_property(p, v.get())).pack(side='right')
                self.single_widgets[prop['name']] = var

        tk.Button(self.single_scroll_frame, text="✅ Применить все изменения", command=self.apply_all_single,
                  bg='#2e7d32', fg='white', font=('Segoe UI', 10, 'bold')).pack(pady=10)

    def apply_single_property(self, prop, new_value):
        try:
            if prop['readonly']:
                self.update_status(f"Свойство {prop['name']} только для чтения")
                return
            if isinstance(prop['value'], (int, float)):
                new_value = float(new_value)
            prop['object'].Value = new_value
            self.handler.doc.Regen(1)
            self.update_status(f"✅ {prop['name']} = {new_value}")
        except Exception as e:
            self.update_status(f"❌ Ошибка: {e}")

    def apply_all_single(self):
        for prop in self.single_properties:
            if prop['readonly']:
                continue
            if prop['name'] in self.single_widgets:
                w = self.single_widgets[prop['name']]
                val = w.get() if hasattr(w, 'get') else str(w)
                self.apply_single_property(prop, val)

    # ---------- Общие методы ----------
    def save_drawing(self):
        if self.handler.save_drawing():
            self.update_status("💾 Чертеж сохранён")
        else:
            self.update_status("❌ Ошибка сохранения")

    def update_status(self, msg):
        self.status_label.config(text=msg)
        self.root.update_idletasks()

    def run(self):
        self.root.mainloop()