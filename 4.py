import sys
import win32com.client


def connect_to_acad():
    """Подключается к AutoCAD через COM."""
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        print("✅ Подключено к AutoCAD")
        return acad
    except:
        print("❌ AutoCAD не запущен! Запустите AutoCAD и повторите.")
        return None


def select_object(acad):
    """Выбирает объект через SelectionSet (рабочий метод из 1.py)."""
    doc = acad.ActiveDocument

    print("\n👉 Пожалуйста, выберите объект в AutoCAD...")
    print("   (нажмите Enter после выбора)\n")

    try:
        # Создаём временную выборку
        ss_name = "TempSelect"
        # Удаляем если существует
        for i in range(doc.SelectionSets.Count):
            if doc.SelectionSets.Item(i).Name == ss_name:
                doc.SelectionSets.Item(i).Delete()
                break

        ss = doc.SelectionSets.Add(ss_name)
        ss.SelectOnScreen()

        if ss.Count == 0:
            print("❌ Объект не выбран.")
            ss.Delete()
            return None

        obj = ss.Item(0)
        ss.Delete()

        if obj.ObjectName != "AcDbBlockReference":
            print(f"❌ Выбранный объект - не блок!")
            print(f"   Тип объекта: {obj.ObjectName}")
            return None

        # Выводим имя блока
        try:
            name = obj.EffectiveName if hasattr(obj, 'EffectiveName') else obj.Name
            print(f"\n✅ Выбран блок: {name}")
        except:
            pass

        return obj

    except Exception as e:
        print(f"❌ Ошибка выбора: {e}")
        return None


def get_all_properties(obj):
    """Получает все динамические свойства блока."""
    properties = []
    try:
        dyn_props = obj.GetDynamicBlockProperties()
        # Перебираем через enumerate - работает для tuple и COM
        for i, prop in enumerate(dyn_props):
            try:
                prop_name = prop.PropertyName
                prop_value = prop.Value
                properties.append({
                    'index': i,
                    'name': prop_name,
                    'value': prop_value,
                    'prop_obj': prop
                })
            except Exception as e:
                print(f"   Ошибка при чтении свойства {i}: {e}")
    except Exception as e:
        print(f"   Ошибка чтения свойств: {e}")

    return properties


def display_properties(properties):
    """Показывает свойства с номерами."""
    if not properties:
        print("\n⚠️ Свойства не найдены!")
        return False

    print("\n" + "=" * 70)
    print("📋 ДОСТУПНЫЕ СВОЙСТВА ДЛЯ РЕДАКТИРОВАНИЯ")
    print("=" * 70)

    for prop in properties:
        value_str = str(prop['value'])
        if len(value_str) > 50:
            value_str = value_str[:47] + "..."
        print(f"   {prop['index'] + 1}. {prop['name']} = {value_str}")

    print("=" * 70)
    return True


def edit_property(prop, doc):
    """Редактирует выбранное свойство."""
    print(f"\n🔧 Редактирование свойства: {prop['name']}")
    print(f"   Текущее значение: {prop['value']}")

    # Проверяем, число ли это
    is_numeric = False
    try:
        float(prop['value'])
        is_numeric = True
    except:
        pass

    if is_numeric:
        try:
            new_value = input("   Новое значение (число): ").strip()
            if not new_value:
                print("   ❌ Изменение отменено")
                return False

            # Преобразуем в число
            if '.' in new_value:
                new_value = float(new_value)
            else:
                new_value = float(new_value)

            prop['prop_obj'].Value = new_value
            print(f"   ✅ Значение изменено на: {new_value}")
            return True

        except ValueError:
            print("   ❌ Ошибка: введите корректное число")
            return False
    else:
        new_value = input("   Новое значение (текст): ").strip()
        if not new_value:
            print("   ❌ Изменение отменено")
            return False

        prop['prop_obj'].Value = new_value
        print(f"   ✅ Значение изменено на: {new_value}")
        return True


def inspect_and_edit():
    """Основная функция инспекции и редактирования."""
    print("\n" + "=" * 70)
    print("🔧 ИНТЕРАКТИВНЫЙ РЕДАКТОР СВОЙСТВ БЛОКОВ")
    print("=" * 70)

    acad = connect_to_acad()
    if not acad:
        input("\nНажмите Enter для выхода...")
        return

    obj = select_object(acad)
    if not obj:
        input("\nНажмите Enter для продолжения...")
        return

    doc = acad.ActiveDocument

    # Краткая информация о блоке
    print("\n" + "=" * 60)
    print("📌 ИНФОРМАЦИЯ ОБ ОБЪЕКТЕ")
    print("=" * 60)
    print(f"\n🔷 ТИП ОБЪЕКТА: {obj.ObjectName}")

    if obj.ObjectName == "AcDbBlockReference":
        try:
            print(f"\n📦 БЛОК:")
            if hasattr(obj, 'EffectiveName'):
                print(f"   Имя (Effective): {obj.EffectiveName}")
            if hasattr(obj, 'Name'):
                print(f"   Имя (Name): {obj.Name}")
            if hasattr(obj, 'IsDynamicBlock'):
                print(f"   Динамический блок: {'Да' if obj.IsDynamicBlock else 'Нет'}")
            if hasattr(obj, 'HasAttributes'):
                print(f"   Имеет атрибуты: {'Да' if obj.HasAttributes else 'Нет'}")
            if hasattr(obj, 'InsertionPoint'):
                ip = obj.InsertionPoint
                print(f"\n📍 Точка вставки: X={ip[0]}, Y={ip[1]}, Z={ip[2]}")
        except Exception as e:
            print(f"   Ошибка при выводе информации: {e}")

    # Читаем свойства
    print("\n🔧 Чтение свойств блока...")
    properties = get_all_properties(obj)

    if not display_properties(properties):
        input("\nНажмите Enter для продолжения...")
        return

    # Цикл редактирования
    while True:
        print("\n" + "-" * 70)
        print("ДЕЙСТВИЯ:")
        print("   • Введите номер свойства для редактирования")
        print("   • 0 - обновить список свойств")
        print("   • n - выбрать другой объект")
        print("   • q - выход")
        print("-" * 70)

        action = input("\nВаш выбор: ").strip().lower()

        if action == 'q':
            print("\n👋 До свидания!")
            return

        if action == 'n':
            print("\n🔄 Возврат к выбору объекта...")
            break

        if action == '0':
            properties = get_all_properties(obj)
            display_properties(properties)
            continue

        try:
            prop_num = int(action)
            if 1 <= prop_num <= len(properties):
                selected_prop = properties[prop_num - 1]

                if edit_property(selected_prop, doc):
                    doc.Regen(1)
                    print("\n✅ Чертёж обновлён!")
                    # Обновляем список свойств
                    properties = get_all_properties(obj)
                    display_properties(properties)
                else:
                    print("\n⚠️ Изменение не применено")
            else:
                print(f"❌ Неверный номер! Введите число от 1 до {len(properties)}")

        except ValueError:
            print("❌ Неверный ввод! Введите номер свойства или команду (0/n/q)")


def main():
    while True:
        inspect_and_edit()
        print("\n" + "=" * 60)
        answer = input("\nХотите проверить другой блок? (y/n): ").lower()
        if answer != 'y':
            print("До свидания!")
            break


if __name__ == "__main__":
    main()
