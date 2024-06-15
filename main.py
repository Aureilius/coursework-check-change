import dox, constructor


mode = int(input('Выберите режим:\n1. Проверка документа на соответствие оформлению\n2. Создание курсовой по всем правилам оформления\nОтвет: '))
if mode == 1:
    path = input('Введите полный путь до нужного файла: ')
    dox.pointer(path)
elif mode == 2:
    constructor.construct()