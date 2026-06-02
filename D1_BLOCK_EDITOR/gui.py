import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
from autocad_handler import AutoCADHandler
import win32com.client
import logging
import time
from datetime import datetime


logging.basicConfig(
    filename='debug.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

class BlockEditorGUI:
    def __init__(self):
        self.handler = AutoCADHandler()
        self.current_blocks = []
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

        self.validate_btn = tk.Button(top_frame, text="🔍 Проверить блоки",
                                      command=self.validate_blocks,
                                      bg='#ff9800', fg='white')
        self.validate_btn.pack(side='left', padx=5)

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
        self.tree.heading('changeW', text='ChangeWidth')
        self.tree.heading('changeD', text='ChangeHeight')  # исправлено название для ясности
        self.tree.column('type', width=200)
        self.tree.column('width', width=100)
        self.tree.column('height', width=100)
        self.tree.column('changeW', width=100)
        self.tree.column('changeD', width=100)

        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        # Статистика (новая логика)
        stats_frame = tk.LabelFrame(self.group_tab, text="📐 Текущие размеры кухни", bg='#1e1e1e', fg='#4fc3f7')
        stats_frame.pack(fill='x', padx=5, pady=5)

        self.bottom_total_label = tk.Label(stats_frame, text="📏 Нижняя линия (isBottom=1): 0.0 мм",
                                           bg='#1e1e1e', fg='#cccccc')
        self.bottom_total_label.pack(anchor='w', padx=10, pady=2)

        self.top_total_label = tk.Label(stats_frame, text="📏 Верхняя линия (isTop=1): 0.0 мм",
                                        bg='#1e1e1e', fg='#cccccc')
        self.top_total_label.pack(anchor='w', padx=10, pady=2)

        self.pure_bottom_label = tk.Label(stats_frame, text="🧩 Чистые нижние (для столешницы/плинтуса): 0.0 мм",
                                          bg='#1e1e1e', fg='#cccccc')
        self.pure_bottom_label.pack(anchor='w', padx=10, pady=2)

        self.pure_top_label = tk.Label(stats_frame, text="🧩 Чистые верхние (для верхней планки): 0.0 мм",
                                       bg='#1e1e1e', fg='#cccccc')
        self.pure_top_label.pack(anchor='w', padx=10, pady=2)

        self.full_blocks_label = tk.Label(stats_frame, text="🏢 Блоки полной высоты: 0",
                                          bg='#1e1e1e', fg='#cccccc')
        self.full_blocks_label.pack(anchor='w', padx=10, pady=2)

        # Панель групповых настроек
        group_ctrl = tk.LabelFrame(self.group_tab, text="🎯 Групповое изменение", bg='#1e1e1e', fg='#ffab40')
        group_ctrl.pack(fill='x', padx=5, pady=10)

        # Нижняя линия
        w_bottom_frame = tk.Frame(group_ctrl, bg='#1e1e1e')
        w_bottom_frame.pack(fill='x', padx=10, pady=8)
        tk.Label(w_bottom_frame, text="Целевая ширина (нижняя линия):", bg='#1e1e1e', fg='white', width=25,
                 anchor='w').pack(side='left')
        self.target_width_bottom_var = tk.StringVar()
        ttk.Entry(w_bottom_frame, textvariable=self.target_width_bottom_var, width=12).pack(side='left', padx=10)
        self.apply_width_bottom_btn = tk.Button(w_bottom_frame, text="📏 Применить к нижней линии",
                                                command=self.apply_target_width_bottom,
                                                bg='#2e7d32', fg='white')
        self.apply_width_bottom_btn.pack(side='left', padx=5)

        # Верхняя линия
        w_top_frame = tk.Frame(group_ctrl, bg='#1e1e1e')
        w_top_frame.pack(fill='x', padx=10, pady=8)
        tk.Label(w_top_frame, text="Целевая ширина (верхняя линия):", bg='#1e1e1e', fg='white', width=25,
                 anchor='w').pack(side='left')
        self.target_width_top_var = tk.StringVar()
        ttk.Entry(w_top_frame, textvariable=self.target_width_top_var, width=12).pack(side='left', padx=10)
        self.apply_width_top_btn = tk.Button(w_top_frame, text="📏 Применить к верхней линии",
                                             command=self.apply_target_width_top,
                                             bg='#2e7d32', fg='white')
        self.apply_width_top_btn.pack(side='left', padx=5)

        # Общая высота
        h_frame = tk.Frame(group_ctrl, bg='#1e1e1e')
        h_frame.pack(fill='x', padx=10, pady=8)
        tk.Label(h_frame, text="Целевая общая высота (Opening_Height):", bg='#1e1e1e', fg='white', width=25,
                 anchor='w').pack(side='left')
        self.target_height_var = tk.StringVar()
        ttk.Entry(h_frame, textvariable=self.target_height_var, width=12).pack(side='left', padx=10)
        self.apply_height_btn = tk.Button(h_frame, text="⬆️ Применить высоту",
                                          command=self.apply_overall_height,
                                          bg='#2e7d32', fg='white')
        self.apply_height_btn.pack(side='left', padx=5)

        # Новая кнопка "Применить все"
        all_frame = tk.Frame(group_ctrl, bg='#1e1e1e')
        all_frame.pack(fill='x', padx=10, pady=8)
        self.apply_all_btn = tk.Button(all_frame, text="🚀 Применить все (низ/верх + высота)",
                                       command=self.apply_all_operations,
                                       bg='#ff8c00', fg='white', font=('Segoe UI', 10, 'bold'))
        self.apply_all_btn.pack(side='left', padx=5)

        # Обновление спецэлементов
        update_frame = tk.Frame(group_ctrl, bg='#1e1e1e')
        update_frame.pack(fill='x', padx=10, pady=8)
        self.update_special_btn = tk.Button(update_frame, text="🔄 Обновить спецэлементы (столешница/планки)",
                                            command=self.update_special_elements,
                                            bg='#37474f', fg='white')
        self.update_special_btn.pack(side='left', padx=5)

        # --- Вкладка 2: Одиночный блок ---
        self.single_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.single_tab, text="🔧 Один блок (свойства)")

        self.select_single_btn = tk.Button(self.single_tab, text="🎯 Выбрать блок в AutoCAD",
                                           command=self.select_single_block,
                                           bg='#007acc', fg='white', font=('Segoe UI', 10, 'bold'))
        self.select_single_btn.pack(pady=10)

        self.toggle_all_params_btn = tk.Button(self.single_tab, text="📋 Показать все параметры",
                                               command=self.toggle_all_params,
                                               bg='#37474f', fg='white')
        self.toggle_all_params_btn.pack(pady=5)

        # Фрейм для свойств одного блока (скроллируемый)
        self.single_canvas_frame = tk.Frame(self.single_tab, bg='#1e1e1e')
        self.single_canvas_frame.pack(fill='both', expand=True, padx=5, pady=5)

        self.single_canvas = tk.Canvas(self.single_canvas_frame, bg='#1e1e1e', highlightthickness=0)
        self.single_scrollbar = ttk.Scrollbar(self.single_canvas_frame, orient="vertical",
                                              command=self.single_canvas.yview)
        self.single_scroll_frame = tk.Frame(self.single_canvas, bg='#1e1e1e')

        self.single_scroll_frame.bind("<Configure>", lambda e: self.single_canvas.configure(
            scrollregion=self.single_canvas.bbox("all")))
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

    # ---------- Валидация блоков ----------
    def validate_blocks(self):
        """Запускает проверку выделенных блоков"""
        if not self.current_blocks:
            messagebox.showinfo("Нет блоков", "Сначала обновите выделение из AutoCAD")
            return

        warnings = self.handler.validate_blocks(self.current_blocks)

        if warnings:
            msg = "🔍 Результаты проверки:\n\n" + "\n".join(warnings)
            messagebox.showwarning("Проверка блоков", msg)
        else:
            messagebox.showinfo("Проверка блоков",
                                "✅ Все блоки соответствуют логике!\n\nМожно безопасно выполнять операции.")

    # ---------- Обновление спецэлементов ----------
    def update_special_elements(self):
        """Обновляет столешницу, плинтус и верхнюю планку"""
        if not self.current_blocks:
            messagebox.showinfo("Нет блоков", "Сначала обновите выделение из AutoCAD")
            return

        result = self.handler.update_special_elements(self.current_blocks)

        if result['updated']:
            self.update_status(f"✅ Обновлены спецэлементы: {', '.join(result['updated'][:3])}")
            self.refresh_selection()
        else:
            self.update_status("ℹ️ Свойства спецэлементов не найдены в блоках")

    # ---------- Общий сброс ----------
    def reset_active_tab(self):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0:
            self.reset_group()
        else:
            self.reset_single()

    def reset_group(self):
        """Очищает все данные группового режима"""
        self.current_blocks = []
        self.block_infos = []
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.bottom_total_label.config(text="📏 Нижняя линия (isBottom=1): 0.0 мм")
        self.top_total_label.config(text="📏 Верхняя линия (isTop=1): 0.0 мм")
        self.pure_bottom_label.config(text="🧩 Чистые нижние (для столешницы/плинтуса): 0.0 мм")
        self.pure_top_label.config(text="🧩 Чистые верхние (для верхней планки): 0.0 мм")
        self.full_blocks_label.config(text="🏢 Блоки полной высоты: 0")
        self.target_width_bottom_var.set("")
        self.target_width_top_var.set("")
        self.target_height_var.set("")
        self.update_status("Групповые данные сброшены")

    def reset_single(self):
        """Очищает данные одиночного блока"""
        self.single_block = None
        self.single_properties = []
        self.single_widgets = {}
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

        # Автоматическая проверка при загрузке
        warnings = self.handler.validate_blocks(self.current_blocks)
        if warnings:
            self.update_status(f"⚠️ Проверка: {len(warnings)} предупреждений. Нажмите «Проверить блоки» для деталей")
        else:
            self.update_status(f"✅ Загружено блоков: {len(self.block_infos)}")

    def update_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for info in self.block_infos:
            # Пропускаем спецблоки из таблицы
            if info['type'] == 'special':
                continue
            self.tree.insert('', 'end', values=(
                f"{info['name']} ({info['type_display']})",
                f"{info['width']:.1f}",
                f"{info['height']:.1f}",
                "✅" if info['changeWidth'] else "❌",
                "✅" if info['changeHeight'] else "❌"
            ))

    def update_statistics(self):
        """Обновляет статистику по новой логике"""
        bottom_total = self.handler.get_bottom_line_total(self.current_blocks)
        top_total = self.handler.get_top_line_total(self.current_blocks)
        pure_bottom = self.handler.get_pure_bottom_total(self.current_blocks)
        pure_top = self.handler.get_pure_top_total(self.current_blocks)

        full_blocks = [b for b in self.current_blocks if self.handler.get_block_type(b) == 'full']
        full_count = len(full_blocks)

        # Получаем высоту из блока HeightsForCabs
        heights_block = self.handler.get_heights_block(self.current_blocks)
        opening_h = self.handler.get_opening_height(heights_block) if heights_block else None

        self.bottom_total_label.config(text=f"📏 Нижняя линия (isBottom=1): {bottom_total:.1f} мм")
        self.top_total_label.config(text=f"📏 Верхняя линия (isTop=1): {top_total:.1f} мм")
        self.pure_bottom_label.config(text=f"🧩 Чистые нижние (для столешницы/плинтуса): {pure_bottom:.1f} мм")
        self.pure_top_label.config(text=f"🧩 Чистые верхние (для верхней планки): {pure_top:.1f} мм")
        self.full_blocks_label.config(text=f"🏢 Блоки полной высоты: {full_count}")

        if opening_h:
            self.update_status(f"📐 Общая высота проёма: {opening_h:.1f} мм | " + self.status_label.cget("text"))

    def apply_target_width_bottom(self):
        if not self.current_blocks:
            messagebox.showinfo("Нет блоков", "Сначала обновите выделение из AutoCAD")
            return

        try:
            target = float(self.target_width_bottom_var.get())
        except:
            messagebox.showerror("Ошибка", "Введите число для целевой ширины нижней линии")
            return

        success, new_total, msg = self.handler.apply_target_width_bottom(self.current_blocks, target)

        if success:
            self.update_status(msg)
            self.refresh_selection()
            # Автоматически обновляем спецэлементы
            self.handler.update_special_elements(self.current_blocks)
            self.update_status(f"{msg} + автообновление спецэлементов")
        else:
            messagebox.showinfo("Ничего не изменено", msg)
            self.update_status(msg)

    def apply_target_width_top(self):
        if not self.current_blocks:
            messagebox.showinfo("Нет блоков", "Сначала обновите выделение из AutoCAD")
            return

        try:
            target = float(self.target_width_top_var.get())
        except:
            messagebox.showerror("Ошибка", "Введите число для целевой ширины верхней линии")
            return

        success, new_total, msg = self.handler.apply_target_width_top(self.current_blocks, target)

        if success:
            self.update_status(msg)
            self.refresh_selection()
            # Автоматически обновляем спецэлементы
            self.handler.update_special_elements(self.current_blocks)
            self.update_status(f"{msg} + автообновление спецэлементов")
        else:
            messagebox.showinfo("Ничего не изменено", msg)
            self.update_status(msg)

    def apply_overall_height(self):
        if not self.current_blocks:
            messagebox.showinfo("Нет блоков", "Сначала обновите выделение из AutoCAD")
            return

        try:
            target = float(self.target_height_var.get())
        except:
            messagebox.showerror("Ошибка", "Введите число для целевой общей высоты")
            return

        success, msg = self.handler.apply_overall_height(self.current_blocks, target)

        if success:
            self.update_status(msg)
            self.refresh_selection()
        else:
            messagebox.showinfo("Ничего не изменено", msg)
            self.update_status(msg)

    # ---------- НОВЫЙ МЕТОД: Применить все операции последовательно с таймером ----------
    def apply_all_operations(self):
        if not self.current_blocks:
            messagebox.showinfo("Нет блоков", "Сначала обновите выделение из AutoCAD")
            return

        # Фиксируем текущий список блоков (не перечитываем)
        blocks = self.current_blocks[:]
        self.update_status("🚀 Запуск пакетного изменения...")
        self.root.update()

        bottom_val = self.target_width_bottom_var.get().strip()
        top_val = self.target_width_top_var.get().strip()
        height_val = self.target_height_var.get().strip()

        start_time = time.time()

        # 1. Нижняя ширина (с обновлением спецэлементов внутри)
        if bottom_val:
            try:
                target = float(bottom_val)
                self.update_status(f"📏 Нижняя линия: {target} мм...")
                self.apply_target_width_bottom()  # внутри есть update_special_elements
                # time.sleep(0.2)
            except ValueError:
                self.update_status("❌ Нижняя линия: не число")

        # 2. Верхняя ширина (с обновлением спецэлементов внутри)
        if top_val:
            try:
                target = float(top_val)
                self.update_status(f"📏 Верхняя линия: {target} мм...")
                self.apply_target_width_top()  # внутри есть update_special_elements
                # time.sleep(0.2)
            except ValueError:
                self.update_status("❌ Верхняя линия: не число")

        # 3. Общая высота
        if height_val:
            try:
                target = float(height_val)
                self.update_status(f"⬆️ Высота: {target} мм...")
                # Прямой вызов метода handler без обновления GUI
                success, msg = self.handler.apply_overall_height(blocks, target)
                self.update_status(msg)
                # time.sleep(0.2)
            except ValueError:
                self.update_status("❌ Высота: не число")

        # 4. *** КЛЮЧЕВОЙ МОМЕНТ: принудительно обновляем спецэлементы после высоты ***
        self.update_status("🔄 Обновление спецэлементов (столешница, плинтус, планки)...")
        self.handler.update_special_elements(blocks)
        # time.sleep(0.1)

        # 5. Финальная регенерация и синхронизация GUI
        self.handler.doc.Regen(1)
        self.refresh_selection()  # обновляем таблицу и статистику

        elapsed = time.time() - start_time
        self.update_status(f"✅ Пакетное изменение завершено за {elapsed:.1f} сек")
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

        name = getattr(self.single_block, 'EffectiveName', 'Unknown')
        block_type = self.handler.get_block_type(self.single_block)
        type_display = {
            'bottom': 'Нижний',
            'top': 'Верхний',
            'full': 'Полной высоты (низ+верх)',
            'special': 'HeightsForCabs (специальный)',
            'unknown': 'Неизвестный'
        }.get(block_type, '?')

        tk.Label(self.single_scroll_frame, text=f"Блок: {name}",
                 bg='#1e1e1e', fg='#ffab40', font=('Segoe UI', 12, 'bold')).pack(pady=5)
        tk.Label(self.single_scroll_frame, text=f"Тип: {type_display}",
                 bg='#1e1e1e', fg='#4fc3f7').pack(pady=2)

        self.single_widgets = {}

        # Важные параметры для отображения в первую очередь
        core_params = {'Width', 'Length', 'Height', 'Depth', 'isTop', 'isBottom',
                       'ChangeWidth', 'ChangeDepth', 'Opening_Height', 'Upper_Height'}

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