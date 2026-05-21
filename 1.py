import sys
from pyautocad import Autocad


def inspect_object():
    """Скрипт для инспекции выбранного объекта."""
    acad = Autocad(create_if_not_exists=True)

    print("\n" + "=" * 60)
    print("ИНСПЕКТОР ОБЪЕКТОВ AutoCAD")
    print("=" * 60)
    print("\n👉 Пожалуйста, выберите объект в AutoCAD...")

    try:
        # Запрашиваем выбор объекта
        acad.prompt("\nВыберите объект и нажмите Enter: ")

        # Получаем выбранный объект через COM
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        # Метод 1: через GetEntity (более надёжный)
        try:
            # Используем SelectionSet
            ss = doc.SelectionSets.Add("TempSet")
            ss.SelectOnScreen()

            if ss.Count == 0:
                print("❌ Объект не выбран.")
                ss.Delete()
                return

            obj = ss.Item(0)
            ss.Delete()

        except:
            # Альтернативный метод
            obj = acad.GetEntity("Выберите объект: ")

        print("\n" + "=" * 60)
        print(f"📌 ИНФОРМАЦИЯ ОБ ОБЪЕКТЕ")
        print("=" * 60)

        # Основная информация
        print(f"\n🔷 ТИП ОБЪЕКТА: {obj.ObjectName}")

        # Для блоков - дополнительная информация
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
            except:
                pass

            # Динамические свойства
            print(f"\n🔧 ДИНАМИЧЕСКИЕ СВОЙСТВА (Custom parameters):")
            try:
                dyn_props = obj.GetDynamicBlockProperties()
                if dyn_props.Count > 0:
                    for i in range(dyn_props.Count):
                        prop = dyn_props.Item(i)
                        prop_name = prop.PropertyName
                        prop_val = prop.Value
                        # Проверяем, видимо ли свойство
                        try:
                            is_visible = prop.Visible
                        except:
                            is_visible = True

                        print(f"   {i + 1}. {prop_name} = {prop_val}")
                else:
                    print("   ⚠️ Динамические свойства не найдены")
            except Exception as e:
                print(f"   ❌ Ошибка доступа: {e}")

            # Атрибуты
            print(f"\n🏷️ АТРИБУТЫ:")
            try:
                atts = obj.GetAttributes()
                if atts.Count > 0:
                    for i in range(atts.Count):
                        att = atts.Item(i)
                        tag = att.TagString
                        val = att.TextString
                        print(f"   {i + 1}. {tag} = '{val}'")
                else:
                    print("   ⚠️ Атрибуты не найдены")
            except Exception as e:
                print(f"   ❌ Ошибка доступа: {e}")

            # Пользовательские свойства через другие методы
            print(f"\n📋 ДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ:")
            try:
                # XData (расширенные данные)
                xdata = obj.GetXData("")
                if xdata:
                    print(f"   XData: {xdata}")

                # Hyperlink
                if hasattr(obj, 'Hyperlinks'):
                    if obj.Hyperlinks.Count > 0:
                        print(f"   Гиперссылки: {obj.Hyperlinks.Count}")

                # Layer
                if hasattr(obj, 'Layer'):
                    print(f"   Слой: {obj.Layer}")

                # Color
                if hasattr(obj, 'Color'):
                    print(f"   Цвет: {obj.Color}")

                # Linetype
                if hasattr(obj, 'Linetype'):
                    print(f"   Тип линии: {obj.Linetype}")

                # PlotStyleName
                if hasattr(obj, 'PlotStyleName'):
                    print(f"   Стиль печати: {obj.PlotStyleName}")
            except:
                pass

        # Общая информация для всех объектов
        print(f"\n📍 ПОЛОЖЕНИЕ:")
        try:
            if hasattr(obj, 'InsertionPoint'):
                ip = obj.InsertionPoint
                print(f"   Точка вставки: X={ip[0]}, Y={ip[1]}, Z={ip[2]}")
        except:
            pass

        print(f"\n🔄 ОБНОВЛЕНИЕ:")
        acad.app.Update()

    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        print("   Возможно, вы не выбрали объект или отменили операцию.")


def main():
    while True:
        inspect_object()
        print("\n" + "=" * 60)
        answer = input("\nХотите проверить другой объект? (y/n): ").lower()
        if answer != 'y':
            print("До свидания!")
            break


if __name__ == "__main__":
    main()