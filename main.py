from pyautocad import Autocad

acad = Autocad(create_if_not_exists=True)
acad.prompt("Консольное приложение успешно подключилось к AutoCAD.")
print(f"Версия AutoCAD: {acad.app.Version}")

from pyautocad import Autocad, APoint

acad = Autocad(create_if_not_exists=True)

# Итерируемся по всем объектам в пространстве модели
for obj in acad.iter_objects():
    # Проверяем, является ли объект линией
    if obj.ObjectName == "AcDbLine":
        print(f"Найдена линия! ID: {obj.ObjectID}")
        # Здесь можно добавить логику для выбора конкретной линии
        break