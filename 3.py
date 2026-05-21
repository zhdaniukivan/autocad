import sys
from pyautocad import Autocad


def find_and_change_in_nested_blocks(block_def, target_param="btm_cab_width", new_value=650.0, depth=0):
    """Рекурсивно ищет и изменяет параметры во всех вложенных блоках."""

    indent = "  " * depth
    found_count = 0

    try:
        # Перебираем все объекты в определении блока
        for i in range(block_def.Count):
            obj = block_def.Item(i)

            # Если это блок-ссылка внутри определения
            if obj.ObjectName == "AcDbBlockReference":
                block_name = ""
                try:
                    block_name = obj.EffectiveName if hasattr(obj, 'EffectiveName') else obj.Name
                except:
                    block_name = "Unknown"

                print(f"{indent}📦 Найден вложенный блок: {block_name}")

                # Пробуем изменить параметры в этом блоке
                try:
                    # Получаем динамические свойства блока
                    if hasattr(obj, 'GetDynamicBlockProperties'):
                        dyn_props = obj.GetDynamicBlockProperties()

                        for j in range(dyn_props.Count):
                            prop = dyn_props.Item(j)
                            prop_name = prop.PropertyName

                            # Если нашли нужный параметр
                            if prop_name == target_param:
                                old_val = prop.Value
                                print(f"{indent}  🎯 Найден параметр '{target_param}' в блоке '{block_name}'")
                                print(f"{indent}     Старое значение: {old_val}")

                                # Меняем значение
                                prop.Value = new_value
                                print(f"{indent}     ✅ Новое значение: {new_value}")
                                found_count += 1

                except Exception as e:
                    print(f"{indent}     ⚠️ Не удалось прочитать параметры: {e}")

                # Рекурсивно ищем параметры в определении этого блока
                try:
                    # Получаем определение вложенного блока
                    doc = acad.ActiveDocument
                    for block in doc.Blocks:
                        if block.Name == block_name or (
                                hasattr(block, 'EffectiveName') and block.EffectiveName == block_name):
                            print(f"{indent}  🔍 Ищем параметры в определении блока '{block_name}'...")
                            sub_found = find_and_change_in_nested_blocks(block, target_param, new_value, depth + 1)
                            found_count += sub_found
                            break
                except Exception as e:
                    print(f"{indent}     Ошибка при рекурсивном поиске: {e}")

            # Если это параметр блока
            elif obj.ObjectName == "AcDbBlockParameter":
                try:
                    param_name = obj.Name if hasattr(obj, 'Name') else "без имени"

                    if param_name == target_param:
                        old_val = obj.Value
                        print(f"{indent}🎯 Найден параметр-объект: {param_name}")
                        print(f"{indent}   Старое значение: {old_val}")

                        obj.Value = new_value
                        print(f"{indent}   ✅ Изменено на: {new_value}")
                        found_count += 1

                except Exception as e:
                    print(f"{indent}   Ошибка: {e}")

    except Exception as e:
        print(f"{indent}Ошибка при обходе блока: {e}")

    return found_count


def change_all_instances():
    """Изменяет параметр во всех вставках блоков в чертеже."""

    print("\n" + "=" * 70)
    print("🔧 ИЗМЕНЕНИЕ ПАРАМЕТРА ВО ВСЕХ ВСТАВКАХ БЛОКОВ")
    print("=" * 70)

    acad = Autocad(create_if_not_exists=True)
    doc = acad.ActiveDocument
    model_space = doc.ModelSpace

    target_param = "btm_cab_width"
    new_value = 650.0

    print(f"\n🔍 Поиск всех вставок блоков с параметром '{target_param}'...")
    print(f"🎯 Новое значение: {new_value}\n")

    found_count = 0
    changed_count = 0

    # Перебираем все объекты в пространстве модели
    for i in range(model_space.Count):
        obj = model_space.Item(i)

        if obj.ObjectName == "AcDbBlockReference":
            block_name = ""
            try:
                block_name = obj.EffectiveName if hasattr(obj, 'EffectiveName') else obj.Name
            except:
                block_name = "Unknown"

            # Пробуем получить динамические свойства
            try:
                if hasattr(obj, 'GetDynamicBlockProperties'):
                    dyn_props = obj.GetDynamicBlockProperties()

                    for j in range(dyn_props.Count):
                        prop = dyn_props.Item(j)
                        prop_name = prop.PropertyName

                        if prop_name == target_param:
                            found_count += 1
                            old_val = prop.Value
                            print(f"📦 Блок: {block_name}")
                            print(f"   ID: {obj.ObjectID}")
                            print(f"   {target_param}: {old_val} → {new_value}")

                            prop.Value = new_value
                            changed_count += 1
                            print(f"   ✅ Изменено!\n")

            except Exception as e:
                pass

    if found_count == 0:
        print("❌ Не найдено ни одной вставки блока с параметром 'btm_cab_width'")
        print("\nПроверяем определения блоков...")

        # Проверяем определения блоков
        for block in doc.Blocks:
            if not block.Name.startswith("*"):  # Пропускаем анонимные
                print(f"\n📚 Проверяем определение блока: {block.Name}")
                found = find_and_change_in_nested_blocks(block, target_param, new_value)
                if found > 0:
                    print(f"   ✅ Изменено {found} параметров в определении блока")
                    changed_count += found
    else:
        print(f"\n📊 ИТОГО:")
        print(f"   Найдено вставок с параметром: {found_count}")
        print(f"   Изменено: {changed_count}")

    if changed_count > 0:
        doc.Regen(1)
        print(f"\n🎉 Параметр '{target_param}' успешно изменён на {new_value}!")
        print("   Обновите чертёж командой REGEN, если изменения не видны")
    else:
        print(f"\n❌ Не удалось найти параметр '{target_param}'")


def inspect_all_parameters():
    """Показывает все параметры всех блоков в чертеже."""

    print("\n" + "=" * 70)
    print("📋 ПОИСК ВСЕХ ПАРАМЕТРОВ В ЧЕРТЕЖЕ")
    print("=" * 70)

    acad = Autocad(create_if_not_exists=True)
    doc = acad.ActiveDocument

    target_param = "btm_cab_width"
    print(f"\n🔍 Ищем параметр '{target_param}'...\n")

    found = False

    # 1. Проверяем вставки блоков в пространстве модели
    print("1. ПРОВЕРКА ВСТАВОК БЛОКОВ В ПРОСТРАНСТВЕ МОДЕЛИ:")
    model_space = doc.ModelSpace

    for i in range(model_space.Count):
        obj = model_space.Item(i)

        if obj.ObjectName == "AcDbBlockReference":
            try:
                if hasattr(obj, 'GetDynamicBlockProperties'):
                    dyn_props = obj.GetDynamicBlockProperties()

                    for j in range(dyn_props.Count):
                        prop = dyn_props.Item(j)
                        if prop.PropertyName == target_param:
                            block_name = obj.EffectiveName if hasattr(obj, 'EffectiveName') else obj.Name
                            print(f"   ✅ Найдено! Блок: {block_name}, Значение: {prop.Value}")
                            found = True
            except:
                pass

    # 2. Проверяем определения всех блоков
    print("\n2. ПРОВЕРКА ОПРЕДЕЛЕНИЙ БЛОКОВ:")

    for block in doc.Blocks:
        if block.Name.startswith("*"):
            continue

        print(f"\n   Блок: {block.Name}")

        for i in range(block.Count):
            obj = block.Item(i)

            # Проверяем параметры
            if obj.ObjectName == "AcDbBlockParameter":
                try:
                    if hasattr(obj, 'Name') and obj.Name == target_param:
                        print(f"      ✅ Параметр-объект: {obj.Name} = {obj.Value}")
                        found = True
                except:
                    pass

            # Проверяем динамические свойства вложенных блоков
            if obj.ObjectName == "AcDbBlockReference":
                try:
                    if hasattr(obj, 'GetDynamicBlockProperties'):
                        dyn_props = obj.GetDynamicBlockProperties()
                        for j in range(dyn_props.Count):
                            prop = dyn_props.Item(j)
                            if prop.PropertyName == target_param:
                                block_name = obj.EffectiveName if hasattr(obj, 'EffectiveName') else obj.Name
                                print(f"      ✅ Вложенный блок '{block_name}': {prop.PropertyName} = {prop.Value}")
                                found = True
                except:
                    pass

    if not found:
        print("\n❌ Параметр 'btm_cab_width' не найден нигде в чертеже")
        print("\nВозможно, параметр называется иначе. Вот все найденные параметры:")

        # Показываем все параметры для диагностики
        for block in doc.Blocks:
            if block.Name.startswith("*"):
                continue

            for i in range(block.Count):
                obj = block.Item(i)

                if obj.ObjectName == "AcDbBlockParameter":
                    try:
                        if hasattr(obj, 'Name'):
                            print(f"   {block.Name} → {obj.Name} = {obj.Value}")
                    except:
                        pass

                if obj.ObjectName == "AcDbBlockReference":
                    try:
                        if hasattr(obj, 'GetDynamicBlockProperties'):
                            dyn_props = obj.GetDynamicBlockProperties()
                            for j in range(dyn_props.Count):
                                prop = dyn_props.Item(j)
                                print(f"   {block.Name} → {obj.Name} → {prop.PropertyName} = {prop.Value}")
                    except:
                        pass


def main():
    while True:
        print("\n" + "=" * 70)
        print("🔧 ИНСТРУМЕНТ ДЛЯ РАБОТЫ С ПАРАМЕТРАМИ БЛОКОВ")
        print("=" * 70)
        print("\nВыберите действие:")
        print("1. Изменить btm_cab_width на 650 (во всех вставках)")
        print("2. Найти все параметры btm_cab_width в чертеже")
        print("3. Выйти")

        choice = input("\nВаш выбор (1-3): ").strip()

        if choice == "1":
            change_all_instances()
        elif choice == "2":
            inspect_all_parameters()
        elif choice == "3":
            print("\n👋 До свидания!")
            break
        else:
            print("Неверный выбор")


if __name__ == "__main__":
    main()