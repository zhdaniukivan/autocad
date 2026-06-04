import win32com.client
import logging

logging.basicConfig(
    filename='debug.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)


class AutoCADHandler:
    def __init__(self):
        self.acad = None
        self.doc = None
        self.connect()

    def connect(self):
        logging.info("Попытка подключения к AutoCAD...")
        try:
            self.acad = win32com.client.GetActiveObject("AutoCAD.Application")
            self.doc = self.acad.ActiveDocument
            logging.info("✅ Подключено к запущенному AutoCAD")
        except:
            try:
                self.acad = win32com.client.Dispatch("AutoCAD.Application")
                self.acad.Visible = True
                self.doc = self.acad.ActiveDocument
                logging.info("✅ AutoCAD запущен")
            except Exception as e:
                logging.error(f"❌ Не удалось подключиться: {e}")

    # ---------- Получение выделенных блоков ----------
    def get_selected_blocks(self):
        try:
            ss = self.doc.PickfirstSelectionSet
            blocks = []
            for i in range(ss.Count):
                obj = ss.Item(i)
                is_dynamic = getattr(obj, 'IsDynamicBlock', False)
                if obj.ObjectName == "AcDbBlockReference" and is_dynamic:
                    blocks.append(obj)
            logging.info(f"Найдено динамических блоков: {len(blocks)}")
            return blocks
        except Exception as e:
            logging.error(f"Ошибка получения выделенных блоков: {e}")
            return []

    # ---------- Чтение/запись свойств ----------
    def get_dyn_prop_value(self, block, prop_name, default=None):
        try:
            props = block.GetDynamicBlockProperties()
            if isinstance(props, tuple):
                props_list = props
            else:
                props_list = [props.Item(i) for i in range(props.Count)]
            for p in props_list:
                if p.PropertyName == prop_name:
                    return p.Value
            return default
        except Exception as e:
            logging.debug(f"get_dyn_prop_value error for {prop_name}: {e}")
            return default

    def set_dyn_prop_value(self, block, prop_name, new_value):
        try:
            props = block.GetDynamicBlockProperties()
            if isinstance(props, tuple):
                props_list = props
            else:
                props_list = [props.Item(i) for i in range(props.Count)]
            for p in props_list:
                if p.PropertyName == prop_name and not p.ReadOnly:
                    p.Value = new_value
                    logging.debug(f"Установлено {prop_name} = {new_value} для блока {getattr(block, 'EffectiveName', '?')}")
                    return True
            logging.debug(f"Свойство {prop_name} не найдено или только для чтения")
            return False
        except Exception as e:
            logging.error(f"Ошибка установки {prop_name}={new_value}: {e}")
            return False

    # ---------- Классификация блоков ----------
    def get_is_bottom(self, block):
        return self.get_dyn_prop_value(block, 'isBottom') == 1.0

    def get_is_top(self, block):
        return self.get_dyn_prop_value(block, 'isTop') == 1.0

    def get_block_type(self, block):
        name = getattr(block, 'EffectiveName', '')
        if name == 'HeightsForCabs':
            return 'special'
        is_bottom = self.get_is_bottom(block)
        is_top = self.get_is_top(block)
        if is_bottom and is_top:
            return 'full'
        elif is_bottom and not is_top:
            return 'bottom'
        elif not is_bottom and is_top:
            return 'top'
        return 'unknown'

    # ---------- Ширина ----------
    def get_width(self, block):
        for name in ['Width', 'Length', 'OverallWidth']:
            val = self.get_dyn_prop_value(block, name)
            if val is not None:
                return float(val)
        return 0.0

    def set_width(self, block, new_width):
        for name in ['Width', 'Length', 'OverallWidth']:
            if self.get_dyn_prop_value(block, name) is not None:
                return self.set_dyn_prop_value(block, name, float(new_width))
        return False

    def has_change_width_flag(self, block):
        return self.get_dyn_prop_value(block, 'ChangeWidth') == 1.0

    # ---------- Высота ----------
    def get_height(self, block):
        for name in ['Height', 'Depth']:
            val = self.get_dyn_prop_value(block, name)
            if val is not None:
                return float(val)
        return 0.0

    def set_height(self, block, new_height):
        for name in ['Height', 'Depth']:
            if self.get_dyn_prop_value(block, name) is not None:
                return self.set_dyn_prop_value(block, name, float(new_height))
        return False

    def has_change_height_flag(self, block):
        return self.get_dyn_prop_value(block, 'ChangeHeight') == 1.0

    # ---------- Суммы ----------
    def get_bottom_line_total(self, blocks, log_details=False):
        total = 0.0
        details = []
        for b in blocks:
            if self.get_is_bottom(b):
                w = self.get_width(b)
                total += w
                if log_details:
                    name = getattr(b, 'EffectiveName', '?')
                    details.append(f"{name}: {w:.1f}")
        if log_details:
            logging.info(f"Нижняя линия: {total:.1f} из {details}")
        return total

    def get_top_line_total(self, blocks, log_details=False):
        total = 0.0
        details = []
        for b in blocks:
            if self.get_is_top(b):
                w = self.get_width(b)
                total += w
                if log_details:
                    name = getattr(b, 'EffectiveName', '?')
                    details.append(f"{name}: {w:.1f}")
        if log_details:
            logging.info(f"Верхняя линия: {total:.1f} из {details}")
        return total

    def get_pure_bottom_total(self, blocks, log_details=False):
        total = 0.0
        details = []
        for b in blocks:
            if self.get_block_type(b) == 'bottom':
                w = self.get_width(b)
                total += w
                if log_details:
                    name = getattr(b, 'EffectiveName', '?')
                    details.append(f"{name}: {w:.1f}")
        if log_details:
            logging.info(f"Чистые нижние: {total:.1f} из {details}")
        return total

    def get_pure_top_total(self, blocks, log_details=False):
        total = 0.0
        details = []
        for b in blocks:
            if self.get_block_type(b) == 'top':
                w = self.get_width(b)
                total += w
                if log_details:
                    name = getattr(b, 'EffectiveName', '?')
                    details.append(f"{name}: {w:.1f}")
        if log_details:
            logging.info(f"Чистые верхние: {total:.1f} из {details}")
        return total

    # ---------- Диагностика позиций ----------
    def get_block_position(self, block):
        try:
            ins_pt = block.InsertionPoint
            return (ins_pt[0], ins_pt[1], ins_pt[2])
        except Exception as e:
            logging.debug(f"Ошибка получения позиции: {e}")
            return None

    def get_bottom_blocks_x_range(self, blocks):
        data = []
        for b in blocks:
            if self.get_is_bottom(b):
                pos = self.get_block_position(b)
                if pos:
                    w = self.get_width(b)
                    name = getattr(b, 'EffectiveName', '?')
                    left = pos[0]
                    right = left + w
                    data.append({
                        'name': name,
                        'left': left,
                        'right': right,
                        'width': w,
                        'block': b
                    })
        if not data:
            return None, None, []
        min_x = min(d['left'] for d in data)
        max_x = max(d['right'] for d in data)
        return min_x, max_x, data

    # ---------- Применение ширины с диагностикой ----------
    def apply_target_width_bottom(self, blocks, target_width, regen=True):
        logging.info(f"=== Применение ширины к нижней линии: target={target_width} ===")

        # Логируем позиции ДО изменения
        min_x_before, max_x_before, bottom_data = self.get_bottom_blocks_x_range(blocks)
        if min_x_before is not None:
            logging.info(f"  ДО: левый край={min_x_before:.1f}, правый край={max_x_before:.1f}, общая ширина={max_x_before-min_x_before:.1f}")
            for d in bottom_data:
                logging.info(f"    {d['name']}: left={d['left']:.1f}, right={d['right']:.1f}, width={d['width']:.1f}")

        bottom_blocks = [b for b in blocks if self.get_is_bottom(b)]
        if not bottom_blocks:
            return False, 0, "Нет блоков с isBottom=1"
        current_total = self.get_bottom_line_total(blocks, log_details=True)
        if target_width <= current_total:
            return False, current_total, f"Целевая ширина {target_width:.1f} не больше текущей {current_total:.1f}"
        changeable = [b for b in bottom_blocks if self.has_change_width_flag(b)]
        if not changeable:
            return False, current_total, "Нет расширяемых блоков (ChangeWidth=1) среди нижних"
        increase = target_width - current_total
        increment = increase / len(changeable)
        for b in changeable:
            old_w = self.get_width(b)
            new_w = old_w + increment
            name = getattr(b, 'EffectiveName', '?')
            logging.info(f"  {name}: {old_w:.1f} → {new_w:.1f}")
            self.set_width(b, new_w)

        # Логируем позиции ПОСЛЕ изменения (до регенерации)
        min_x_after, max_x_after, _ = self.get_bottom_blocks_x_range(blocks)
        if min_x_after is not None:
            logging.info(f"  ПОСЛЕ (до регена): левый край={min_x_after:.1f}, правый край={max_x_after:.1f}, смещение левого края={min_x_after-min_x_before:+.1f}")

        if regen:
            self.doc.Regen(1)
            # Логируем после регена
            min_x_regen, max_x_regen, _ = self.get_bottom_blocks_x_range(blocks)
            if min_x_regen is not None:
                logging.info(f"  ПОСЛЕ регена: левый край={min_x_regen:.1f}, правый край={max_x_regen:.1f}, смещение={min_x_regen-min_x_before:+.1f}")

        new_total = self.get_bottom_line_total(blocks)
        logging.info(f"Результат: {current_total:.1f} → {new_total:.1f}")
        return True, new_total, f"✅ Нижняя линия: {current_total:.1f} → {new_total:.1f} (+{increase:.1f})"

    def apply_target_width_top(self, blocks, target_width, regen=True):
        logging.info(f"=== Применение ширины к верхней линии: target={target_width} ===")
        top_blocks = [b for b in blocks if self.get_is_top(b)]
        if not top_blocks:
            return False, 0, "Нет блоков с isTop=1"
        current_total = self.get_top_line_total(blocks, log_details=True)
        if target_width <= current_total:
            return False, current_total, f"Целевая ширина {target_width:.1f} не больше текущей {current_total:.1f}"
        changeable = [b for b in top_blocks if self.has_change_width_flag(b)]
        if not changeable:
            return False, current_total, "Нет расширяемых блоков (ChangeWidth=1) среди верхних"
        increase = target_width - current_total
        increment = increase / len(changeable)
        for b in changeable:
            old_w = self.get_width(b)
            new_w = old_w + increment
            name = getattr(b, 'EffectiveName', '?')
            logging.info(f"  {name}: {old_w:.1f} → {new_w:.1f}")
            self.set_width(b, new_w)
        if regen:
            self.doc.Regen(1)
        new_total = self.get_top_line_total(blocks)
        logging.info(f"Результат: {current_total:.1f} → {new_total:.1f}")
        return True, new_total, f"✅ Верхняя линия: {current_total:.1f} → {new_total:.1f} (+{increase:.1f})"

    # ---------- Общая высота ----------
    def get_heights_block(self, blocks):
        for b in blocks:
            if getattr(b, 'EffectiveName', '') == 'HeightsForCabs':
                return b
        return None

    def get_opening_height(self, heights_block):
        if not heights_block:
            return None
        return self.get_dyn_prop_value(heights_block, 'Opening_Height')

    def set_opening_height(self, heights_block, new_height):
        if not heights_block:
            return False
        return self.set_dyn_prop_value(heights_block, 'Opening_Height', float(new_height))

    def get_upper_height(self, heights_block):
        if not heights_block:
            return None
        return self.get_dyn_prop_value(heights_block, 'Upper_Height')

    def set_upper_height(self, heights_block, new_height):
        if not heights_block:
            return False
        return self.set_dyn_prop_value(heights_block, 'Upper_Height', float(new_height))

    def apply_overall_height(self, blocks, target_opening_height, regen=True):
        logging.info(f"=== Применение общей высоты: target={target_opening_height} ===")
        heights_block = self.get_heights_block(blocks)
        if not heights_block:
            return False, "Блок HeightsForCabs не найден в выделении"
        current = self.get_opening_height(heights_block)
        if current is None:
            return False, "Не удалось прочитать Opening_Height"
        delta = target_opening_height - current
        if delta <= 0:
            return False, f"Новая высота {target_opening_height:.1f} не больше текущей {current:.1f}"
        changed_blocks = []
        skipped_blocks = []
        for b in blocks:
            if self.get_block_type(b) == 'special':
                continue
            if self.has_change_height_flag(b):
                cur_h = self.get_height(b)
                if cur_h > 0:
                    new_h = cur_h + delta
                    name = getattr(b, 'EffectiveName', '?')
                    logging.info(f"  {name}: высота {cur_h:.1f} → {new_h:.1f} (ChangeHeight=1)")
                    self.set_height(b, new_h)
                    changed_blocks.append(f"{name} ({cur_h:.1f}→{new_h:.1f})")
            else:
                name = getattr(b, 'EffectiveName', '?')
                cur_h = self.get_height(b)
                logging.info(f"  {name}: высота НЕ изменена ({cur_h:.1f}, ChangeHeight=0)")
                skipped_blocks.append(name)
        current_upper = self.get_upper_height(heights_block)
        if current_upper is not None:
            new_upper = current_upper + delta
            logging.info(f"  Upper_Height: {current_upper:.1f} → {new_upper:.1f}")
            self.set_upper_height(heights_block, new_upper)
        self.set_opening_height(heights_block, target_opening_height)
        if regen:
            self.doc.Regen(1)
        msg = f"✅ Высота увеличена на {delta:.1f} мм\nИзменено блоков: {len(changed_blocks)}\nПропущено (ChangeHeight=0): {', '.join(skipped_blocks) if skipped_blocks else 'нет'}"
        logging.info(msg)
        return True, msg

    # ---------- Обновление спецэлементов (унифицированное имя плинтуса) ----------
    def update_special_elements(self, blocks, regen=True):
        logging.info("=== Обновление спецэлементов ===")
        pure_bottom_total = self.get_pure_bottom_total(blocks, log_details=True)
        pure_top_total = self.get_pure_top_total(blocks, log_details=True)
        result = {'countertop_w': pure_bottom_total, 'kickplate_w': pure_bottom_total,
                  'upper_filler_w': pure_top_total, 'updated': []}
        for b in blocks:
            name = getattr(b, 'EffectiveName', '?')
            # Столешница
            if self.get_dyn_prop_value(b, 'countertop_w') is not None:
                self.set_dyn_prop_value(b, 'countertop_w', pure_bottom_total)
                result['updated'].append(f"{name}.countertop_w={pure_bottom_total:.1f}")
            # Плинтус – проверяем оба возможных имени
            kick_val = self.get_dyn_prop_value(b, 'kickplate_w')
            if kick_val is None:
                kick_val = self.get_dyn_prop_value(b, 'kick_plate_w')
            if kick_val is not None:
                # Устанавливаем в то свойство, которое существует
                if self.get_dyn_prop_value(b, 'kickplate_w') is not None:
                    self.set_dyn_prop_value(b, 'kickplate_w', pure_bottom_total)
                else:
                    self.set_dyn_prop_value(b, 'kick_plate_w', pure_bottom_total)
                result['updated'].append(f"{name}.kickplate_w={pure_bottom_total:.1f}")
            # Верхняя планка
            if self.get_dyn_prop_value(b, 'upper_filler_w') is not None:
                self.set_dyn_prop_value(b, 'upper_filler_w', pure_top_total)
                result['updated'].append(f"{name}.upper_filler_w={pure_top_total:.1f}")
            # Recalculated_Width
            if self.get_dyn_prop_value(b, 'Recalculated_Width') is not None:
                if self.get_block_type(b) in ['bottom', 'full']:
                    self.set_dyn_prop_value(b, 'Recalculated_Width', pure_bottom_total)
                elif self.get_block_type(b) == 'top':
                    self.set_dyn_prop_value(b, 'Recalculated_Width', pure_top_total)
        if regen and result['updated']:
            self.doc.Regen(1)
        return result

    # ---------- Проверки ----------
    def validate_blocks(self, blocks):
        warnings = []
        if not blocks:
            warnings.append("⚠️ Нет выделенных блоков")
            return warnings
        logging.info("=== Детальная информация о блоках ===")
        for b in blocks:
            name = getattr(b, 'EffectiveName', '?')
            block_type = self.get_block_type(b)
            is_bottom = self.get_is_bottom(b)
            is_top = self.get_is_top(b)
            width = self.get_width(b)
            height = self.get_height(b)
            change_w = self.has_change_width_flag(b)
            change_h = self.has_change_height_flag(b)
            logging.info(f"Блок: {name}")
            logging.info(f"  Тип: {block_type}, isBottom={is_bottom}, isTop={is_top}")
            logging.info(f"  Width={width:.1f}, Height={height:.1f}")
            logging.info(f"  ChangeWidth={change_w}, ChangeHeight={change_h}")
        bottom_blocks = [b for b in blocks if self.get_is_bottom(b)]
        top_blocks = [b for b in blocks if self.get_is_top(b)]
        full_blocks = [b for b in blocks if self.get_block_type(b) == 'full']
        pure_bottom = [b for b in blocks if self.get_block_type(b) == 'bottom']
        pure_top = [b for b in blocks if self.get_block_type(b) == 'top']
        heights_block = self.get_heights_block(blocks)
        if not bottom_blocks and not top_blocks:
            warnings.append("⚠️ Нет блоков с isBottom=1 или isTop=1")
        if full_blocks:
            warnings.append(f"ℹ️ Найдено {len(full_blocks)} блоков полной высоты (участвуют в обеих линиях)")
        bottom_changeable = [b for b in bottom_blocks if self.has_change_width_flag(b)]
        top_changeable = [b for b in top_blocks if self.has_change_width_flag(b)]
        if bottom_blocks and not bottom_changeable:
            warnings.append("⚠️ Среди нижних блоков нет ни одного с ChangeWidth=1 — изменение ширины невозможно")
        if top_blocks and not top_changeable:
            warnings.append("⚠️ Среди верхних блоков нет ни одного с ChangeWidth=1 — изменение ширины невозможно")
        changeable_height = [b for b in blocks if self.has_change_height_flag(b) and self.get_block_type(b) != 'special']
        if changeable_height:
            names = [getattr(b, 'EffectiveName', '?') for b in changeable_height]
            warnings.append(f"ℹ️ Блоки с ChangeHeight=1 ({len(changeable_height)}): {', '.join(names)}")
        else:
            warnings.append("ℹ️ Нет блоков с ChangeHeight=1 — изменение общей высоты не затронет ни один блок")
        if not heights_block:
            warnings.append("⚠️ Блок HeightsForCabs не найден — изменение общей высоты недоступно")
        else:
            opening_h = self.get_opening_height(heights_block)
            upper_h = self.get_upper_height(heights_block)
            if opening_h is None:
                warnings.append("⚠️ В блоке HeightsForCabs нет свойства Opening_Height")
            if upper_h is None:
                warnings.append("⚠️ В блоке HeightsForCabs нет свойства Upper_Height")
        has_countertop = any(self.get_dyn_prop_value(b, 'countertop_w') is not None for b in blocks)
        has_kickplate = any(self.get_dyn_prop_value(b, 'kickplate_w') is not None or self.get_dyn_prop_value(b, 'kick_plate_w') is not None for b in blocks)
        has_upperfiller = any(self.get_dyn_prop_value(b, 'upper_filler_w') is not None for b in blocks)
        if pure_bottom and not (has_countertop or has_kickplate):
            warnings.append("ℹ️ Есть чистые нижние блоки, но не найдены свойства countertop_w/kickplate_w")
        if pure_top and not has_upperfiller:
            warnings.append("ℹ️ Есть чистые верхние блоки, но не найдено свойство upper_filler_w")
        return warnings

    # ---------- Получение информации о блоке для таблицы ----------
    def get_block_info(self, block):
        name = getattr(block, 'EffectiveName', 'Unknown')
        block_type = self.get_block_type(block)
        type_display = {
            'bottom': 'Нижний', 'top': 'Верхний', 'full': 'Полный (низ+верх)',
            'special': 'HeightsForCabs', 'unknown': '?'
        }.get(block_type, '?')
        return {
            'name': name, 'type': block_type, 'type_display': type_display,
            'isTop': self.get_is_top(block), 'isBottom': self.get_is_bottom(block),
            'width': self.get_width(block), 'height': self.get_height(block),
            'changeWidth': self.has_change_width_flag(block),
            'changeHeight': self.has_change_height_flag(block),
            'object': block
        }

    def align_bottom_line_left(self, blocks, original_min_x):
        """Выравнивает левый край нижней линии по заданной координате"""
        min_x_after, _, bottom_data = self.get_bottom_blocks_x_range(blocks)
        if min_x_after is None:
            return False
        delta = original_min_x - min_x_after
        if abs(delta) < 0.1:
            return False
        logging.info(f"Выравнивание нижней линии: сдвиг на {delta:.1f} мм вправо")
        for d in bottom_data:
            block = d['block']
            pos = self.get_block_position(block)
            if pos:
                new_x = pos[0] + delta
                block.InsertionPoint = (new_x, pos[1], pos[2])
                logging.debug(f"  Блок {d['name']}: позиция X {pos[0]:.1f} → {new_x:.1f}")
        self.doc.Regen(1)
        return True

    # ---------- Сохранение ----------
    def save_drawing(self):
        try:
            self.doc.Save()
            logging.info("Чертёж сохранён")
            return True
        except Exception as e:
            logging.error(f"Ошибка сохранения: {e}")
            return False