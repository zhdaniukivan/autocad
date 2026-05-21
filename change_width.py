import sys
from pyautocad import Autocad


def change_dynamic_property(block, prop_name, new_value):
    """Изменяет динамическое свойство блока, если оно существует."""
    try:
        # Получаем все динамические свойства блока
        dyn_props = block.GetDynamicBlockProperties()
        for prop in dyn_props:
            if prop.PropertyName == prop_name:
                old_val = prop.Value
                print(f"  Динамическое свойство '{prop_name}': {old_val} -> {new_value}")
                prop.Value = new_value
                return True
    except Exception as e:
        # Не у всех блоков есть динамические свойства
        pass
    return False


def change_attribute(block, tag_name, new_value):
    """Изменяет атрибут блока с заданным тегом."""
    try:
        atts = block.GetAttributes()
        for att in atts:
            if att.TagString == tag_name:
                old_val = att.TextString
                print(f"  Атрибут '{tag_name}': {old_val} -> {new_value}")
                att.TextString = str(new_value)
                return True
    except:
        pass
    return False


def main():
    # Подключаемся к AutoCAD
    acad = Autocad(create_if_not_exists=True)
    acad.prompt("\n=== Поиск и замена btm_cab_width ===\n")

    target_name = "btm_cab_width"
    new_value = 650

    # Если передан аргумент командной строки — используем его
    if len(sys.argv) > 1:
        try:
            new_value = float(sys.argv[1])
        except:
            print(f"Использую значение по умолчанию: {new_value}")

    found = False
    # Перебираем все объекты в пространстве модели
    for obj in acad.iter_objects():
        # Проверяем, является ли объект блоком (AcDbBlockReference)
        if obj.ObjectName == "AcDbBlockReference":
            # Пытаемся изменить динамическое свойство
            if change_dynamic_property(obj, target_name, new_value):
                found = True
            # Если не сработало, пробуем атрибут
            elif change_attribute(obj, target_name, new_value):
                found = True

            # Можно добавить поиск во вложенных блоках, если нужно

    if found:
        acad.app.Update()
        print("\n✅ Изменения применены и экран обновлён.")
    else:
        print(f"\n❌ Блок или атрибут с именем '{target_name}' не найден.")
        print("   Убедитесь, что:")
        print("   - Вы находитесь в пространстве Model")
        print("   - Параметр является динамическим свойством или атрибутом блока")
        print("   - Блоки не находятся внутри других блоков (требуется рекурсивный обход)")


if __name__ == "__main__":
    main()