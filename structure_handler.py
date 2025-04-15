from typing import List, Tuple, Optional, Dict, Union
from structural_list import STRUCTURE


class StructureHandler:
    def __init__(self):
        self.selected_division: str = ""
        self.selected_unit: str = ""
        self.selected_segment: str = ""
        self.selected_tag: str = ""
        self.selected_total: bool = False
        self.request_string: str = ""
        self._navigation_stack = []

    def normalize_string(self, s: str) -> str:
        return s.lower().replace(" ", "")

    def form_request_string(self) -> str:
        if not self.selected_division:
            return ""

        request_parts = [
            self.selected_division,
            self.selected_unit if self.selected_unit else "",
            "масс" if self.selected_tag == "Массовые" else self.selected_tag if self.selected_tag else "",
            self.selected_segment if self.selected_segment else "",
            "итог" if self.selected_total else "",
            "масс" if hasattr(self, '_mass_tag') and self._mass_tag else ""
        ]

        self.request_string = self.normalize_string("".join(filter(None, request_parts)))
        return self.request_string

    def _save_state(self):
        self._navigation_stack.append({
            'division': self.selected_division,
            'unit': self.selected_unit,
            'segment': self.selected_segment,
            'tag': self.selected_tag,
            'total': self.selected_total
        })

    def go_back(self) -> bool:
        if not self._navigation_stack:
            return False

        prev_state = self._navigation_stack.pop()
        self.selected_division = prev_state['division']
        self.selected_unit = prev_state['unit']
        self.selected_segment = prev_state['segment']
        self.selected_tag = prev_state['tag']
        self.selected_total = prev_state['total']
        return True

    def reset_to_start(self):
        self._navigation_stack = []
        self.selected_division = ""
        self.selected_unit = ""
        self.selected_segment = ""
        self.selected_tag = ""
        self.selected_total = False
        self.request_string = ""

    def reset_selection(self) -> None:
        self._navigation_stack = []
        self.selected_division = ""
        self.selected_unit = ""
        self.selected_segment = ""
        self.selected_tag = ""
        self.selected_total = False
        self.request_string = ""

    def reset_to_start(self):
        self._navigation_stack = []
        self.selected_division = ""
        self.selected_unit = ""
        self.selected_segment = ""
        self.selected_tag = ""
        self.selected_total = False
        self.request_string = ""

    def get_all_divisions(self) -> List[str]:
        return list(STRUCTURE.keys())

    def select_division(self, division_name: str):
        self._save_state()
        self.selected_division = division_name
        self.selected_unit = ""
        self.selected_segment = ""
        self.selected_tag = ""
        self.selected_total = False

    def get_units_for_division(self, division_name: str = None) -> List[str]:
        division = division_name or self.selected_division
        if not division:
            return []

        if division not in STRUCTURE:
            raise ValueError(f"Дивизион '{division}' не найден в структуре")

        return list(STRUCTURE[division].keys())

    def get_units_with_options(self) -> List[str]:
        units = self.get_units_for_division()
        options = []
        if units:
            options.extend(units)
        options.append("Итог")
        if self._navigation_stack:
            options.append("Вернуться назад")
        return options

    def select_unit(self, unit_name: str) -> str:
        if unit_name == "Итог":
            self._save_state()
            self.selected_total = True
            return "total"
        elif unit_name == "Вернуться назад":
            return "back"

        self._save_state()
        self.selected_unit = unit_name
        self.selected_segment = ""
        self.selected_tag = ""
        self.selected_total = False
        return "continue"

    def get_segments_for_unit(self, unit_name: str = None) -> Union[List[str], Dict[str, Tuple[str]]]:
        unit = unit_name or self.selected_unit
        if not self.selected_division or not unit:
            return []

        division_data = STRUCTURE[self.selected_division]
        if unit not in division_data:
            return []

        segments = division_data[unit]

        if isinstance(segments, dict):
            if self.selected_tag:
                if self.selected_tag in segments:
                    return list(segments[self.selected_tag])
                else:
                    return []
            else:
                return list(segments.keys())

        elif isinstance(segments, (list, tuple)):
            return list(segments)

        return []

    def get_tags_with_options(self) -> List[str]:
        tags = self.get_segments_for_unit()
        options = []
        if tags:
            options.extend(tags)
        options.append("Итог")
        if self._navigation_stack:
            options.append("Вернуться назад")
        return options

    def select_tag(self, tag_name: str) -> str:
        if tag_name == "Вернуться назад":
            return "back" if self.go_back() else "continue"
        elif tag_name == "Итог":
            self._save_state()
            self.selected_total = True
            self.selected_segment = ""
            self.selected_tag = ""
            return "total"
        else:
            if not self.selected_unit:
                raise ValueError("Сначала выберите юнит")

            if self.selected_unit in ["РК", "СП"]:
                if tag_name not in ["Массовые", "Все"]:
                    raise ValueError(f"Недопустимый тег '{tag_name}' для юнита '{self.selected_unit}'")
            elif self.selected_unit == "КИБ":
                if tag_name not in ["Все", "ККСБ", "ММБ"]:
                    raise ValueError(f"Недопустимый тег '{tag_name}' для юнита 'КИБ'")
            else:
                raise ValueError(f"Юнит '{self.selected_unit}' не поддерживает теги")

            self._save_state()
            self.selected_tag = tag_name
            self.selected_total = False
            return "continue"

    def get_segments_with_options(self) -> List[str]:
        segments = self.get_segments_for_unit()
        options = []
        if segments:
            options.extend(segments)

        options.append("Итог")

        if self._navigation_stack:
            options.append("Вернуться назад")

        return options

    def select_segment(self, segment_name: str) -> str:
        if segment_name == "Вернуться назад":
            return "back" if self.go_back() else "continue"
        elif segment_name == "Итог":
            self._save_state()
            self.selected_total = True
            self.selected_segment = ""
            return "total"
        elif segment_name == "Масс":
            self._save_state()
            self.selected_total = False
            return "mass"
        else:
            available_segments = self.get_segments_for_unit()
            if segment_name not in available_segments:
                raise ValueError(f"Сегмент '{segment_name}' не найден в юните '{self.selected_unit}'")

            self._save_state()
            self.selected_segment = segment_name
            self.selected_total = False
            return "continue"

    def get_current_state(self) -> Dict[str, str]:
        return {
            "division": self.selected_division,
            "unit": self.selected_unit,
            "tag": self.selected_tag,
            "segment": self.selected_segment,
            "is_total": self.selected_total,
            "request_string": self.request_string,
            "can_go_back": bool(self._navigation_stack)
        }

    def process_excel_data(self, file_path: str, search_value: str = None) -> Optional[Tuple]:
        from excel_handler import get_data_from_table

        search_str = search_value or self.form_request_string()
        if not search_str:
            return None

        return get_data_from_table(file_path, search_str)