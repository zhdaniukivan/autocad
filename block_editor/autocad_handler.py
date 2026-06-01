# autocad_handler.py
import win32com.client
import logging
from datetime import datetime

# Логирование
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

    def get_selected_block(self):
        logging.info("Попытка получить выбранный блок...")
        try:
            ss = self.doc.PickfirstSelectionSet
            if ss.Count == 0:
                logging.warning("Выделение пустое")
                return None

            obj = ss.Item(0)
            is_dynamic = getattr(obj, 'IsDynamicBlock', False)
            logging.info(f"Выбран объект: {obj.ObjectName}, IsDynamicBlock={is_dynamic}")

            if obj.ObjectName == "AcDbBlockReference" and is_dynamic:
                effective_name = getattr(obj, 'EffectiveName', 'Unknown')
                logging.info(f"✅ Найден динамический блок: {effective_name}")
                return obj
            return None
        except Exception as e:
            logging.error(f"Ошибка получения блока: {e}")
            return None

    def get_all_properties(self, block):
        logging.info("=== Начало получения свойств ===")
        properties = []

        try:
            # Получаем свойства
            result = block.GetDynamicBlockProperties()
            logging.info(f"Тип возвращённого объекта: {type(result)}")

            # Обработка двух возможных типов
            if isinstance(result, tuple):
                dyn_props = result
                logging.info(f"Получен tuple с {len(dyn_props)} свойствами")
            else:
                # COM коллекция
                dyn_props = [result.Item(i) for i in range(result.Count)]
                logging.info(f"Получена COM коллекция с {len(dyn_props)} свойствами")

            for i, prop in enumerate(dyn_props):
                try:
                    prop_name = prop.PropertyName.strip()
                    value = prop.Value
                    readonly = prop.ReadOnly

                    logging.info(f"Свойство {i + 1:2d}: {prop_name} = {value} (ReadOnly={readonly})")

                    properties.append({
                        'name': prop_name,
                        'value': value,
                        'object': prop,
                        'readonly': readonly
                    })
                except Exception as prop_err:
                    logging.warning(f"Ошибка чтения свойства {i}: {prop_err}")

        except Exception as e:
            logging.error(f"Критическая ошибка при получении свойств: {e}")

        logging.info(f"Итого загружено свойств: {len(properties)}")
        return properties

    def set_property(self, prop_dict, new_value):
        logging.info(f"Изменение свойства: {prop_dict['name']} → {new_value}")
        try:
            if prop_dict['readonly']:
                logging.warning("Свойство только для чтения")
                return False

            # Автоопределение типа
            if isinstance(prop_dict['value'], (int, float)):
                try:
                    new_value = float(new_value)
                except:
                    new_value = int(new_value)

            prop_dict['object'].Value = new_value
            self.doc.Regen(1)
            logging.info("✅ Изменение успешно")
            return True
        except Exception as e:
            logging.error(f"Ошибка применения: {e}")
            return False

    def save_drawing(self):
        try:
            self.doc.Save()
            logging.info("Чертёж успешно сохранён")
            return True
        except Exception as e:
            logging.error(f"Ошибка сохранения: {e}")
            return False