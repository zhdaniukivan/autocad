import sys
import win32com.client


def connect_to_acad():
    """Подключается к AutoCAD через чистый COM."""
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        print("✅ Подключено к AutoCAD")
        return acad
    except:
        print("❌ AutoCAD не запущен!")
        return None


def change_btm_cab_width():
    """Изменяет параметр btm_cab_width на 650."""

    print("\n" + "=" * 70)
    print("🔧 ИЗМЕНЕНИЕ ПАРАМЕТРА btm_cab_width")
    print("=" * 70)

    acad = connect_to_acad()
    if not acad:
        return

    doc = acad.ActiveDocument

    # Ищем блок Test1
    test1_block = None
    for block in doc.Blocks:
        if block.Name == "Test1":
            test1_block = block
            break

    if not test1_block:
        print("❌ Блок 'Test1' не найден!")
        return

    print(f"\n✅ Найден блок: Test1")

    # Ищем в блоке Test1 ссылку на btm_cab_swing
    found = False
    changed = False

    for i in range(test1_block.Count):
        obj = test1_block.Item(i)

        if obj.ObjectName == "AcDbBlockReference":
            try:
                # Проверяем, является ли этот блок btm_cab_swing
                eff_name = obj.EffectiveName if hasattr(obj, 'EffectiveName') else obj.Name

                if eff_name == "btm_cab_swing":
                    print(f"\n🎯 Найдена ссылка на btm_cab_swing в определении блока Test1")

                    # Получаем динамические свойства
                    dyn_props = obj.GetDynamicBlockProperties()

                    # Перебираем свойства
                    for j, prop in enumerate(dyn_props):
                        if hasattr(prop, 'PropertyName') and prop.PropertyName == "btm_cab_width":
                            old_value = prop.Value
                            print(f"   Текущее значение btm_cab_width: {old_value}")

                            # Меняем значение
                            prop.Value = 650.0
                            print(f"   ✅ Изменено на: 650.0")
                            changed = True
                            found = True
                            break

                    if not found:
                        print("   ⚠️ Параметр 'btm_cab_width' не найден в этом блоке")

                    break

            except Exception as e:
                print(f"   Ошибка: {e}")

    if changed:
        # Обновляем чертёж
        doc.Regen(1)
        print("\n" + "=" * 70)
        print("🎉 ПАРАМЕТР УСПЕШНО ИЗМЕНЁН!")
        print("   btm_cab_width = 650.0")
        print("=" * 70)
        print("\n⚠️ ВАЖНО: После изменения определения блока:")
        print("   1. Закройте редактор блоков (BCLOSE)")
        print("   2. Все вставки блока Test1 обновятся автоматически")
        print("   3. Если изменения не видны, выполните команду ATTSYNC")
    else:
        print("\n❌ Не удалось найти или изменить параметр")


def change_all_instances():
    """Изменяет параметр во всех существующих вставках блоков."""

    print("\n" + "=" * 70)
    print("🔧 ИЗМЕНЕНИЕ ВО ВСЕХ ВСТАВКАХ БЛОКОВ")
    print("=" * 70)

    acad = connect_to_acad()
    if not acad:
        return

    doc = acad.ActiveDocument
    model_space = doc.ModelSpace

    changed_count = 0

    print("\n🔍 Поиск всех вставок блоков...")

    for i in range(model_space.Count):
        obj = model_space.Item(i)

        if obj.ObjectName == "AcDbBlockReference":
            try:
                # Получаем эффективное имя блока
                eff_name = obj.EffectiveName if hasattr(obj, 'EffectiveName') else obj.Name

                # Проверяем, является ли блок btm_cab_swing
                if eff_name == "btm_cab_swing":
                    print(f"\n📦 Найден блок: {eff_name}")
                    print(f"   Позиция: {obj.InsertionPoint}")

                    # Получаем динамические свойства
                    dyn_props = obj.GetDynamicBlockProperties()

                    # Ищем и меняем параметр
                    for prop in dyn_props:
                        if hasattr(prop, 'PropertyName') and prop.PropertyName == "btm_cab_width":
                            old_val = prop.Value
                            prop.Value = 650.0
                            print(f"   btm_cab_width: {old_val} → 650.0")
                            changed_count += 1
                            break

            except Exception as e:
                print(f"   Ошибка: {e}")

    if changed_count > 0:
        doc.Regen(1)
        print(f"\n🎉 Изменено {changed_count} вставок блоков!")
    else:
        print("\n❌ Вставок блоков 'btm_cab_swing' не найдено")


def main():
    while True:
        print("\n" + "=" * 70)
        print("🔧 ИЗМЕНЕНИЕ btm_cab_width")
        print("=" * 70)
        print("\nВыберите действие:")
        print("1. Изменить в определении блока Test1 (рекомендуется)")
        print("2. Изменить во всех существующих вставках блоков")
        print("3. Выйти")

        choice = input("\nВаш выбор (1-3): ").strip()

        if choice == "1":
            change_btm_cab_width()
        elif choice == "2":
            change_all_instances()
        elif choice == "3":
            print("\n👋 До свидания!")
            break
        else:
            print("Неверный выбор")


if __name__ == "__main__":
    main()