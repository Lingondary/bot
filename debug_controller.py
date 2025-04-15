from structure_handler import StructureHandler
from excel_handler import get_data_from_table


def main(file_path):
    selector = StructureHandler()

    while True:
        while True:
            # 1. Выбор дивизиона
            print("\n=== Выбор дивизиона ===")
            divisions = selector.get_all_divisions()
            for i, division in enumerate(divisions, 1):
                print(f"{i}. {division}")

            try:
                choice = int(input("Выберите номер дивизиона: ")) - 1
                if 0 <= choice < len(divisions):
                    selector.select_division(divisions[choice])
                    print(f"\nВыбран дивизион: {selector.selected_division}")
                    break
                else:
                    print("Некорректный номер. Попробуйте еще раз.")
            except ValueError:
                print("Введите число.")

        while True:
            # 2. Выбор юнита
            print("\n=== Выбор юнита ===")
            units = selector.get_units_with_options()
            for i, unit in enumerate(units, 1):
                print(f"{i}. {unit}")

            try:
                choice = int(input("Выберите номер: ")) - 1
                if 0 <= choice < len(units):
                    action = selector.select_unit(units[choice])

                    if action == "back":
                        break
                    elif action == "total":
                        result = selector.process_excel_data(file_path)
                        if result:
                            print("\n=== Результаты из Excel ===")
                            print(result[0])
                            print(result[1])
                        continue

                    print(f"\nВыбран юнит: {selector.selected_unit}")
                else:
                    print("Некорректный номер. Попробуйте еще раз.")
                    continue
            except ValueError:
                print("Введите число.")
                continue

            # 3. Обработка специальных юнитов (РК/СП/КИБ)
            if selector.selected_unit in ["РК", "СП", "КИБ"]:
                while True:
                    print("\n=== Выбор типа сегментов ===")
                    tags = selector.get_tags_with_options()
                    for i, tag in enumerate(tags, 1):
                        print(f"{i}. {tag}")

                    try:
                        choice = int(input("Выберите номер: ")) - 1
                        if 0 <= choice < len(tags):
                            action = selector.select_tag(tags[choice])

                            if action == "back":
                                break
                            elif action == "total":
                                result = selector.process_excel_data(file_path)
                                if result:
                                    print("\n=== Результаты из Excel ===")
                                    print(result[0])
                                    print(result[1])
                                continue

                            print(f"\nВыбран тип: {selector.selected_tag}")
                        else:
                            print("Некорректный номер. Попробуйте еще раз.")
                            continue
                    except ValueError:
                        print("Введите число.")
                        continue

                    # 4. Выбор сегмента с обработкой тега "Масс"
                    segments = selector.get_segments_for_unit()
                    if segments:
                        while True:
                            print("\n=== Выбор сегмента ===")
                            seg_options = selector.get_segments_with_options()
                            for i, seg in enumerate(seg_options, 1):
                                print(f"{i}. {seg}")

                            try:
                                choice = int(input("Выберите номер: ")) - 1
                                if 0 <= choice < len(seg_options):
                                    action = selector.select_segment(seg_options[choice])

                                    if action == "back":
                                        break
                                    elif action == "mass":
                                        print("\nТег 'Масс' добавлен")
                                        request_str = selector.form_request_string()
                                        print("\n=== Результаты из Excel ===")
                                        result = get_data_from_table(file_path, request_str)
                                        print(request_str)
                                        print(result)
                                        break
                                    elif action == "total":
                                        result = selector.process_excel_data(file_path)
                                        if result:
                                            print("\n=== Результаты из Excel ===")
                                            print(result[0])
                                            print(result[1])
                                        continue

                                    print(f"\nВыбран сегмент: {selector.selected_segment}")
                                    break
                                else:
                                    print("Некорректный номер. Попробуйте еще раз.")
                                    continue
                            except ValueError:
                                print("Введите число.")
                                continue

                        if action == "back":
                            continue

                    break

                if action == "back":
                    continue

            # 5. Обработка юнитов без тегов (Сервисы, ПРПА, ППП)
            elif selector.selected_unit in ["Сервисы", "ПРПА", "ППП"]:
                segments = selector.get_segments_for_unit()
                if segments:
                    while True:
                        print("\n=== Выбор сегмента ===")
                        seg_options = selector.get_segments_with_options()
                        for i, seg in enumerate(seg_options, 1):
                            print(f"{i}. {seg}")

                        try:
                            choice = int(input("Выберите номер: ")) - 1
                            if 0 <= choice < len(seg_options):
                                action = selector.select_segment(seg_options[choice])

                                if action == "back":
                                    break
                                elif action == "total":
                                    result = selector.process_excel_data(file_path)
                                    if result:
                                        print("\n=== Итоговая строка запроса ===")
                                        print("\n=== Результаты из Excel ===")
                                    continue

                                print(f"\nВыбран сегмент: {selector.selected_segment}")
                                break
                            else:
                                print("Некорректный номер. Попробуйте еще раз.")
                                continue
                        except ValueError:
                            print("Введите число.")
                            continue

                    if action == "back":
                        continue

            # 6. Поиск в Excel
            print("\n=== Результаты из Excel ===")
            result = selector.process_excel_data(file_path)

            restart = input("\nХотите выбрать другой дивизион? (да/нет): ").strip().lower()
            if restart != "да":
                print("Программа завершена.")
                return

            break