import win32com.client
import logging
from datetime import datetime

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
        """Возвращает список динамических блоков из текущего выделения"""
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

    # ---------- Чтение свойств блока (обобщённое) ----------
    def get_dyn_prop_value(self, block, prop_name, default=None):
        """Читает значение динамического свойства по имени"""
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
        except:
            return default

    def set_dyn_prop_value(self, block, prop_name, new_value):
        """Устанавливает значение динамического свойства"""
        try:
            props = block.GetDynamicBlockProperties()
            if isinstance(props, tuple):
                props_list = props
            else:
                props_list = [props.Item(i) for i in range(props.Count)]
            for p in props_list:
                if p.PropertyName == prop_name and not p.ReadOnly:
                    p.Value = new_value
                    return True
            return False
        except Exception as e:
            logging.error(f"Ошибка установки {prop_name}={new_value}: {e}")
            return False

    # ---------- Специфичные для кухонных блоков ----------
    def get_width(self, block):
        """Возвращает ширину блока (ищем Width, Length, OverallWidth)"""
        for name in ['Width', 'Length', 'OverallWidth']:
            val = self.get_dyn_prop_value(block, name)
            if val is not None:
                return float(val)
        return 0.0

    def set_width(self, block, new_width):
        """Устанавливает ширину блока"""
        for name in ['Width', 'Length', 'OverallWidth']:
            if self.get_dyn_prop_value(block, name) is not None:
                return self.set_dyn_prop_value(block, name, float(new_width))
        return False

    def get_height(self, block):
        """Возвращает высоту блока (Height, Depth)"""
        for name in ['Height', 'Depth']:
            val = self.get_dyn_prop_value(block, name)
            if val is not None:
                return float(val)
        return 0.0

    def set_height(self, block, new_height):
        """Устанавливает высоту блока"""
        for name in ['Height', 'Depth']:
            if self.get_dyn_prop_value(block, name) is not None:
                return self.set_dyn_prop_value(block, name, float(new_height))
        return False

    def get_is_top(self, block):
        val = self.get_dyn_prop_value(block, 'isTop')
        return val == 1.0

    def get_is_bottom(self, block):
        val = self.get_dyn_prop_value(block, 'isBottom')
        return val == 1.0

    def has_change_width_flag(self, block):
        val = self.get_dyn_prop_value(block, 'ChangeWidth')
        return val == 1.0

    def has_change_depth_flag(self, block):
        val = self.get_dyn_prop_value(block, 'ChangeDepth')
        return val == 1.0

    # ---------- Групповые операции ----------
    def get_total_width(self, blocks, filter_func=None):
        """Сумма ширин блоков, прошедших фильтр (filter_func принимает блок)"""
        total = 0.0
        for b in blocks:
            if filter_func is None or filter_func(b):
                total += self.get_width(b)
        return total

    def distribute_width_increase(self, blocks, target_total_width):
        """
        Распределяет увеличение ширины между блоками, у которых ChangeWidth=1.0.
        Возвращает количество изменённых блоков.
        """
        changeable = [b for b in blocks if self.has_change_width_flag(b)]
        if not changeable:
            return 0
        current_total = sum(self.get_width(b) for b in changeable)
        if target_total_width <= current_total:
            return 0  # только увеличение
        increase = target_total_width - current_total
        increment_per_block = increase / len(changeable)
        for b in changeable:
            new_w = self.get_width(b) + increment_per_block
            self.set_width(b, new_w)
        self.doc.Regen(1)
        return len(changeable)

    def set_height_for_change_depth(self, blocks, target_height):
        """
        Для всех блоков с ChangeDepth=1.0 устанавливает высоту = max(текущая, target_height)
        """
        count = 0
        for b in blocks:
            if self.has_change_depth_flag(b):
                cur_h = self.get_height(b)
                if target_height > cur_h:
                    self.set_height(b, target_height)
                    count += 1
        if count:
            self.doc.Regen(1)
        return count

    def get_block_info(self, block):
        """Возвращает словарь с основными параметрами для отображения в таблице"""
        name = getattr(block, 'EffectiveName', 'Unknown')
        is_top = self.get_is_top(block)
        is_bottom = self.get_is_bottom(block)
        width = self.get_width(block)
        height = self.get_height(block)
        change_width = self.has_change_width_flag(block)
        change_depth = self.has_change_depth_flag(block)
        return {
            'name': name,
            'isTop': is_top,
            'isBottom': is_bottom,
            'width': width,
            'height': height,
            'changeWidth': change_width,
            'changeDepth': change_depth,
            'object': block
        }

    def save_drawing(self):
        try:
            self.doc.Save()
            logging.info("Чертёж сохранён")
            return True
        except Exception as e:
            logging.error(f"Ошибка сохранения: {e}")
            return False