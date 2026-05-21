import sys
from pyautocad import Autocad


def change_block_property_direct(acad, block_name, prop_name, new_value):
    """Изменяет свойство блока через прямой COM-доступ."""
    print(f"\n🔍 Поиск блоков '{block_name}'...")

    # Получаем активный документ через COM
    doc = acad.ActiveDocument
    model_space = doc.ModelSpace

    found_count = 0
    changed_count = 0

    # Перебираем объекты в пространстве модели
    for i in range(model_space.Count):
        obj = model_space.Item(i)

        # Проверяем, является ли объект блоком
        if obj.ObjectName == "AcDbBlockReference":
            # Получаем имя блока (для динамических блоков используем EffectiveName)
            try:
                if hasattr(obj, 'EffectiveName'):
                    obj_name = obj.EffectiveName
                else:
                    obj_name = obj.Name
            except:
                obj_name = ""

            if obj_name == block_name:
                found_count += 1
                print(f"\n📦 Найден блок #{found_count}: '{obj_name}'")
                print(f"   ID: {obj.ObjectID}")
                print(f"   Layer: {obj.Layer}")

                # Пробуем разные методы получения свойств
                properties_found = False

                # Метод 1: Прямой доступ к динамическим свойствам
                try:
                    dyn_props = obj.GetDynamicBlockProperties()
                    print(f"   Динамических свойств: {dyn_props.Count}")

                    for j in range(dyn_props.Count):
                        prop = dyn_props.Item(j)
                        prop_name_com = prop.PropertyName
                        prop_value = prop.Value

                        print(f"      - {prop_name_com} = {prop_value}")

                        # Если нашли нужное свойство - меняем
                        if prop_name_com.lower() == prop_name.lower():
                            print(f"      🎯 Найдено целевое свойство!")
                            print(f"         Старое значение: {prop_value}")
                            prop.Value = new_value
                            print(f"         ✅ Изменено на: {new_value}")
                            changed_count += 1
                            properties_found = True
                except Exception as e:
                    print(f"   ❌ Не удалось прочитать динамические свойства: {e}")

                # Метод 2: Через GetXData (если свойства хранятся там)
                if not properties_found:
                    try:
                        xdata = obj.GetXData("")
                        if xdata:
                            print(f"   XData найдено (может содержать параметры)")
                            # Обработка XData сложнее, но можно распарсить
                    except:
                        pass

                # Метод 3: Проверка атрибутов
                try:
                    atts = obj.GetAttributes()
                    if atts.Count > 0:
                        print(f"   Атрибутов: {atts.Count}")
                        for j in range(atts.Count):
                            att = atts.Item(j)
                            tag = att.TagString
                            val = att.TextString
                            print(f"      - {tag} = '{val}'")

                            if tag.lower() == prop_name.lower():
                                print(f"      🎯 Найден атрибут!")
                                print(f"         Старое значение: {val}")
                                att.TextString = str(new_value)
                                print(f"         ✅ Изменено на: {new_value}")
                                changed_count += 1
                                properties_found = True
                    else:
                        print(f"   Атрибутов: нет")
                except Exception as e:
                    print(f"   Ошибка при чтении атрибутов: {e}")

                if not properties_found:
                    print(f"   ⚠️ Свойство '{prop_name}' не найдено в этом блоке")
                    print(f"   Попробуйте проверить вручную через Properties в AutoCAD")

    return found_count, changed_count


def try_alternative_methods(acad, prop_name):
    """Пробует найти свойство в других местах чертежа."""
    print(f"\n🔍 Альтернативный поиск '{prop_name}'...")

    doc = acad.ActiveDocument

    # 1. Поиск среди определений блоков
    print("\n📚 Проверка определений блоков:")
    blocks = doc.Blocks

    for block in blocks:
        if block.Name.startswith("_") or block.Name.startswith("*"):
            continue

        try:
            # Проверяем вложенные объекты в определении блока
            for ent in block:
                if ent.ObjectName == "AcDbBlockReference":
                    try:
                        dyn_props = ent.GetDynamicBlockProperties()
                        for j in range(dyn_props.Count):
                            prop = dyn_props.Item(j)
                            if prop.PropertyName.lower() == prop_name.lower():
                                print(f"   Найдено в определении блока '{block.Name}'")
                                print(f"      Свойство: {prop.PropertyName} = {prop.Value}")
                                print(f"      Примечание: Это определение, нужно применить ATTSYNC")
                                return True
                    except:
                        pass
        except:
            continue

    # 2. Поиск в словарях расширенных данных
    print("\n📋 Проверка словарей чертежа:")
    try:
        # Получаем словарь именованных объектов
        named_dict = doc.Dictionaries
        for i in range(named_dict.Count):
            dict_item = named_dict.Item(i)
            print(f"   Словарь: {dict_item.Name}")
    except:
        pass

    return False


def main():
    BLOCK_NAME = "btm_cab_swing"
    PROPERTY_NAME = "btm_cab_width"
    NEW_VALUE = 650.0

    if len(sys.argv) > 1:
        try:
            NEW_VALUE = float(sys.argv[1])
        except:
            print(f"Использую значение по умолчанию: {NEW_VALUE}")

    print("🔄 Подключение к AutoCAD...")
    acad = Autocad(create_if_not_exists=True)
    print(f"✅ Подключено к AutoCAD версии {acad.app.Version}")

    # Основной поиск и изменение
    found, changed = change_block_property_direct(acad, BLOCK_NAME, PROPERTY_NAME, NEW_VALUE)

    # Если не нашли - пробуем альтернативные методы
    if found == 0:
        print(f"\n❌ Блок '{BLOCK_NAME}' не найден в пространстве модели.")
        try_alternative_methods(acad, PROPERTY_NAME)

        print("\n💡 Возможные причины:")
        print("   1. Блок находится на заблокированном слое")
        print("   2. Блок является анонимным динамическим блоком (имеет имя вида *U###)")
        print("   3. Блок находится внутри другого блока или внешней ссылки")
        print("   4. Блок ещё не синхронизирован после изменения определения")
    elif changed == 0:
        print(f"\n⚠️ Блок найден, но свойство '{PROPERTY_NAME}' не обнаружено.")
        print("\n💡 Попробуйте вручную в AutoCAD:")
        print("   1. Выделите блок")
        print("   2. Откройте Properties (Ctrl+1)")
        print("   3. Найдите в разделе 'Custom' свойство 'btm_cab_width'")
        print("   4. Убедитесь, что оно доступно для редактирования")

    # Обновляем чертёж
    if changed > 0:
        acad.app.Update()
        acad.app.ZoomExtents()
        print(f"\n🎉 Успешно изменено {changed} блоков!")
        print(f"   Новое значение: {PROPERTY_NAME} = {NEW_VALUE}")
    else:
        print(f"\n❌ Изменений не произведено.")

        # Дополнительная диагностика
        print("\n🛠️ Ручная диагностика:")
        print("   Выделите блок в AutoCAD и выполните в консоли AutoCAD:")
        print("   (vlax-dump-object (vlax-ename->vla-object (car (entsel))) T)")


if __name__ == "__main__":
    main()