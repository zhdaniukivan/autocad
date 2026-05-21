import sys
from pyautocad import Autocad


def explore_block(block, indent="", depth=0, max_depth=3):
    """Рекурсивно исследует блок и выводит все его свойства, атрибуты и динамические параметры."""
    if depth > max_depth:
        return

    block_name = ""
    try:
        block_name = block.EffectiveName if hasattr(block, 'EffectiveName') else block.Name
    except:
        block_name = "Unknown"

    print(f"{indent}📦 Блок: '{block_name}' (ID: {block.ObjectID})")

    # 1. Динамические свойства
    try:
        dyn_props = block.GetDynamicBlockProperties()
        if dyn_props.Count > 0:
            print(f"{indent}   🔧 Динамические свойства ({dyn_props.Count}):")
            for prop in dyn_props:
                try:
                    prop_name = prop.PropertyName
                    prop_val = prop.Value
                    print(f"{indent}      - {prop_name} = {prop_val}")
                except:
                    print(f"{indent}      - (ошибка чтения свойства)")
        else:
            print(f"{indent}   🔧 Динамические свойства: нет")
    except Exception as e:
        print(f"{indent}   🔧 Динамические свойства: недоступны ({str(e)[:50]})")

    # 2. Атрибуты (обычные текстовые поля)
    try:
        atts = block.GetAttributes()
        if atts.Count > 0:
            print(f"{indent}   🏷️ Атрибуты ({atts.Count}):")
            for att in atts:
                try:
                    tag = att.TagString
                    val = att.TextString
                    print(f"{indent}      - {tag} = '{val}'")
                except:
                    print(f"{indent}      - (ошибка чтения атрибута)")
        else:
            print(f"{indent}   🏷️ Атрибуты: нет")
    except Exception as e:
        print(f"{indent}   🏷️ Атрибуты: недоступны ({str(e)[:50]})")

    # 3. Вложенные объекты (рекурсивно)
    try:
        # Получаем все объекты внутри блока (если это блок-ссылка)
        # Для простоты используем обход через OwnerId — сложно.
        # Вместо этого проверим, есть ли у блока атрибуты типа "блок в блоке"
        # В AutoCAD можно получить все сущности внутри блока через блок определения (BlockTableRecord).
        # Но для основного пространства модели это работает иначе.
        pass
    except:
        pass


def explore_layout(layout, indent=""):
    """Исследует все объекты в указанном пространстве (модель или лист)."""
    print(f"\n📄 Пространство: {layout.Name}")
    print(f"   Количество объектов: {layout.Count}")
    found_any = False

    for i, obj in enumerate(layout):
        # Ограничим вывод первыми 50 объектами, чтобы не заваливать консоль
        if i >= 50:
            print(f"   ... и ещё {layout.Count - 50} объектов (показаны первые 50)")
            break

        try:
            obj_name = obj.ObjectName
            if obj_name == "AcDbBlockReference":
                explore_block(obj, indent="   ")
                found_any = True
            else:
                # Для не-блоков просто выводим тип
                print(f"   🧩 {obj_name} (ID: {obj.ObjectID})")
        except Exception as e:
            print(f"   ❌ Ошибка при чтении объекта: {e}")

    if not found_any:
        print("   ⚠️ В этом пространстве нет ни одного блока.")


def main():
    acad = Autocad(create_if_not_exists=True)
    acad.prompt("\n=== ДИАГНОСТИКА: поиск 'btm_cab_width' ===\n")

    target_name = "btm_cab_width"

    # Получаем активный документ
    doc = acad.ActiveDocument

    # 1. Пространство модели (Model Space)
    model_space = doc.ModelSpace
    explore_layout(model_space)

    # 2. Все листы (Paper Space / Layouts)
    print("\n" + "=" * 60)
    print("ПРОВЕРКА ЛИСТОВ (Layouts):")
    for layout in doc.Layouts:
        if layout.Name != "Model":  # Model уже проверили
            try:
                # Получаем блок-представление листа
                layout_block = layout.Block
                explore_layout(layout_block)
            except Exception as e:
                print(f"  Не удалось прочитать лист {layout.Name}: {e}")

    # 3. Дополнительно: поиск по всем блокам в чертеже (таблица блоков)
    print("\n" + "=" * 60)
    print("ПРОВЕРКА ТАБЛИЦЫ БЛОКОВ (BlockTable):")
    try:
        bt = doc.Blocks
        for block in bt:
            # Пропускаем специальные блоки (*Model_Space, *Paper_Space и т.д.) — они уже проверены
            if block.Name.startswith("*"):
                continue
            print(f"\n📚 Определение блока: '{block.Name}' (ID: {block.ObjectID})")
            # Исследуем объекты внутри определения блока
            count = 0
            for ent in block:
                if count > 20:
                    print(f"    ... и ещё {block.Count - 20} объектов")
                    break
                try:
                    if ent.ObjectName == "AcDbBlockReference":
                        explore_block(ent, indent="    ")
                    else:
                        print(f"    🧩 {ent.ObjectName}")
                except:
                    pass
                count += 1
    except Exception as e:
        print(f"Ошибка доступа к таблице блоков: {e}")

    # 4. Поиск по всем текстовым объектам (если параметр — просто текст)
    print("\n" + "=" * 60)
    print("ПОИСК ТЕКСТОВЫХ СТРОК, СОДЕРЖАЩИХ 'btm_cab_width':")
    found_text = False
    for obj in acad.iter_objects():
        try:
            if obj.ObjectName in ["AcDbText", "AcDbMText"]:
                txt = obj.TextString if hasattr(obj, 'TextString') else ""
                if target_name in txt:
                    print(f"  Найден текст: '{txt}' (ID: {obj.ObjectID}, слой: {obj.Layer})")
                    found_text = True
        except:
            pass
    if not found_text:
        print("  Текстовых вхождений не найдено.")

    print("\n" + "=" * 60)
    print("ДИАГНОСТИКА ЗАВЕРШЕНА.")
    print("Теперь вы видите все блоки, их динамические свойства и атрибуты.")
    print("Ищите в выводе строку, содержащую 'btm_cab_width' или близкую по смыслу.")
    acad.app.Update()


if __name__ == "__main__":
    main()