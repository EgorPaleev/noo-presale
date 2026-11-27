"""Build or verify an editable estimate and the proposal's commercial block."""

import argparse
import json
import tempfile
from copy import copy
from contextlib import ExitStack
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

START = "<!-- presale-commercial -->"
END = "<!-- /presale-commercial -->"
CENT = Decimal("0.01")


def text(value):
    if not isinstance(value, str) or not value.strip() or "\n" in value or "\r" in value:
        raise ValueError(f"Ожидалась непустая однострочная строка: {value!r}")
    return value


def number(value):
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise ValueError(f"Ожидалось число: {value!r}")
    amount = Decimal(str(value))
    if not amount.is_finite() or amount < 0 or amount.as_tuple().exponent < -2 or amount > 10**10:
        raise ValueError(f"Допустимо неотрицательное число с точностью до 2 знаков: {value}")
    return amount


def optional_number(value):
    return number(value) if value is not None else None


def load(path):
    data = json.loads(path.read_text(encoding="utf-8"), parse_float=Decimal)
    mode = data.get("mode", "fixed")
    if mode not in ("fixed", "range"):
        raise ValueError("mode должен быть fixed или range")
    currency = data.get("currency")
    vat = data.get("vat")
    duration = data.get("duration")
    for value in (currency, vat, duration):
        if value is not None:
            text(value)
    configured_rates = data.get("rates", {})
    if not isinstance(configured_rates, dict):
        raise ValueError("rates должен быть объектом со ставками по ролям")
    title = text(data.get("project", "Смета проекта"))
    deadlines = data.get("phase_durations", {})
    if not isinstance(deadlines, dict):
        raise ValueError("phase_durations должен быть объектом со сроками этапов")
    schedule = {text(phase): text(value) for phase, value in deadlines.items()}
    rates = {}
    rows = []
    seen = set()
    phases = {}
    groups = {}
    for work in data["work"]:
        work_id, phase, action, result, role = (text(work[key]) for key in ("id", "phase", "action", "result", "role"))
        if work_id in seen:
            raise ValueError(f"Повтор ID работы: {work_id}")
        seen.add(work_id)
        if phases.setdefault(phase.casefold(), phase) != phase:
            raise ValueError(f"Этапы различаются только регистром: {phase}")
        low = number(work["hours"])
        if mode == "fixed" and "max_hours" in work:
            raise ValueError(f"max_hours применяется только для range: {work_id}")
        high = number(work.get("max_hours", work["hours"]))
        if high < low:
            raise ValueError(f"max_hours меньше hours: {work_id}")
        rates[role] = optional_number(configured_rates.get(role))
        group = text(work.get("work_id", work_id))
        name = text(work.get("name", action))
        if groups.setdefault(group, (phase, name, action, result)) != (phase, name, action, result):
            raise ValueError(f"Работа {group} по ролям должна иметь одно описание и результат")
        rows.append((work_id, phase, action, result, role, low, high, group, name))
    if not rows:
        raise ValueError("Нужна хотя бы одна работа")
    unknown = set(schedule) - set(phases.values())
    if unknown:
        raise ValueError(f"Сроки указаны для неизвестных этапов: {sorted(unknown)}")
    expenses = []
    for item in data.get("expenses", []):
        expenses.append((text(item["name"]), text(item["period"]), optional_number(item.get("amount")), text(item["basis"])))
    return mode, currency, vat, duration, rates, rows, expenses, title, schedule


def display(amount):
    return f"{amount:,.2f}".replace(",", " ").rstrip("0").rstrip(".")


def interval(low, high, mode):
    return f"{display(low)}–{display(high)}" if mode == "range" else display(low)


def markdown(value):
    return "".join("\\" + c if c in r"\|*`_<>[]{}" else c for c in value)


def offer_block(estimate):
    mode, currency, vat, duration, rates, rows, expenses, *_ = estimate
    ready = currency is not None and vat is not None and all(rate is not None for rate in rates.values())
    hours = [sum(row[column] for row in rows) for column in (5, 6)]
    lines = [START, "## Коммерческие условия", "", f"- Статус цены — {'рассчитана' if ready else 'черновик'}",
             f"- Трудозатраты — {interval(*hours, mode)} ч",
             f"- Срок — {markdown(duration) if duration else 'уточняется'}",
             f"- НДС — {markdown(vat) if vat else 'уточняется'}"]
    if ready:
        amounts = [sum((row[column] * rates[row[4]]).quantize(CENT, rounding=ROUND_HALF_UP) for row in rows)
                   for column in (5, 6)]
        lines.append(f"- Стоимость работ — {interval(*amounts, mode)} {markdown(currency)}")
    else:
        lines.append("- Стоимость работ — уточняется")
    for name, period, amount, basis in expenses:
        price = f"{display(amount)} {markdown(currency)}" if amount is not None and currency else "уточняется"
        lines.append(f"- Допрасход — {markdown(name)} ({markdown(period)}), {price}. {markdown(basis)}")
    return "\n".join([*lines, END]) + "\n"


def workbook(estimate):
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    mode, currency, vat, duration, rates, rows, expenses, title, schedule = estimate
    book = Workbook()
    main = book.active
    main.title = "Смета"
    detail = book.create_sheet("Расчёт")
    params = book.create_sheet("Параметры")
    extra = book.create_sheet("Допрасходы")
    detail.sheet_state = "hidden"
    with (Path(__file__).resolve().parent.parent / "assets" / "estimate-style.xlsx").open("rb") as source:
        template = load_workbook(source)
    styles = template.active
    book.loaded_theme = template.loaded_theme
    ranged = mode == "range"
    number_format = "#,##0.00"

    def literal(sheet, row, column, value):
        if value is None:
            return
        cell = sheet.cell(row, column, value)
        if isinstance(value, str):
            cell.data_type = "s"

    def append(sheet, values):
        row = 1 if sheet.max_row == 1 and sheet["A1"].value is None else sheet.max_row + 1
        for column, value in enumerate(values, 1):
            literal(sheet, row, column, value)
        return row

    def style(row, source, start=2, end=6):
        main.row_dimensions[row].height = styles.row_dimensions[source].height
        source_end = 4 if source in (2, 3, 4) else 6
        for column in range(start, end + 1):
            origin = styles.cell(source, min(source_end, column))
            cell = main.cell(row, column)
            for attribute in ("font", "fill", "border", "alignment", "number_format"):
                setattr(cell, attribute, copy(getattr(origin, attribute)))

    def totals(cells, price=False):
        refs = ",".join(cells)
        value = f"ROUND(SUM({refs}),2)"
        return f'=IF(COUNT({refs})<{len(cells)},"",{value})' if price else f"={value}"

    def total_range(column, first, last, price=False):
        refs = f"{column}{first}:{column}{last}"
        value = f"ROUND(SUM({refs}),2)"
        return f'=IF(COUNT({refs})<ROWS({refs}),"",{value})' if price else f"={value}"

    append(params, ["Параметр", "Значение"])
    for label, value in (("Валюта", currency), ("НДС", vat), ("Общий срок", duration), ("Формат", mode)):
        append(params, [label, value])
    append(params, ["Роль", "Ставка за час"])
    rate_rows = {role: append(params, [role, rate]) for role, rate in rates.items()}

    append(detail, ["ID позиции", "Этап", "Роль", "Часы", "Часы max", "Ставка", "Стоимость", "Стоимость max"])
    groups = {}
    for work_id, phase, action, result, role, low, high, group, name in rows:
        i = append(detail, [work_id, phase, role, low, high if ranged else None])
        rate = f"'Параметры'!B{rate_rows[role]}"
        detail[f"F{i}"] = f'=IF(ISNUMBER({rate}),{rate},"")'
        detail[f"G{i}"] = f'=IF(ISNUMBER(F{i}),ROUND(D{i}*F{i},2),"")'
        if ranged:
            detail[f"H{i}"] = f'=IF(ISNUMBER(F{i}),ROUND(E{i}*F{i},2),"")'
        works = groups.setdefault(phase, {})
        works.setdefault(group, (name, action, result, []))[3].append(i)

    append(extra, ["Расход", "Период", "Сумма", "Основание"])
    for row in expenses:
        append(extra, list(row))

    last_column = 8 if ranged else 6
    main.sheet_view.showGridLines = False
    widths = [styles.column_dimensions[col].width for col in "ABCDEF"]
    if ranged:
        widths.extend((18, 18))
    for column, width in enumerate(widths, 1):
        main.column_dimensions[get_column_letter(column)].width = width

    style(1, 1, end=last_column)
    main.merge_cells(start_row=1, start_column=2, end_row=1, end_column=last_column)
    literal(main, 1, 2, title)
    summary_headers = ("Этапы", "Бюджет min", "Бюджет max", "Сроки", "Часы min", "Часы max") if ranged else (
        "Этапы", "Бюджет", "Сроки", "Трудозатраты, ч"
    )
    style(2, 2, end=1 + len(summary_headers))
    for col, label in enumerate(summary_headers, 2):
        literal(main, 2, col, label)

    grand = 3 + len(groups)
    stage_header = grand + 6
    style(grand, 4, end=1 + len(summary_headers))
    literal(main, grand, 2, "Итого")
    duration_col = "E" if ranged else "D"
    main[f"{duration_col}{grand}"] = '=IF(\'Параметры\'!B4="","уточняется",\'Параметры\'!B4)'
    style(grand + 2, 5, end=last_column)
    literal(main, grand + 2, 2, "Стоимость часа")
    rates_range = f"'Параметры'!B7:B{6 + len(rates)}"
    main[f"C{grand+2}"] = (
        f'=IF(COUNT({rates_range})<ROWS({rates_range}),"уточняется",'
        f'IF(MIN({rates_range})=MAX({rates_range}),MIN({rates_range}),"по ролям"))'
    )
    main[f"D{grand+2}"] = '=IF(\'Параметры\'!B2="","валюта уточняется",\'Параметры\'!B2)'
    main[f"E{grand+2}"] = '=IF(\'Параметры\'!B3="","НДС уточняется",\'Параметры\'!B3)'
    literal(main, grand + 3, 2, "Статус цены")
    main[f"C{grand+3}"] = (
        f'=IF(AND(\'Параметры\'!B2<>"",\'Параметры\'!B3<>"",ISNUMBER(C{grand})),'
        '"Рассчитана","Черновик")'
    )

    detail_headers = ["Этап / работа", "Что именно делаем", "Результат"]
    detail_headers += (["Часы min", "Часы max", "Стоимость min", "Стоимость max"] if ranged else
                       ["Трудозатраты, ч", "Стоимость"])
    style(stage_header, 6, end=last_column)
    for col, label in enumerate(detail_headers, 2):
        literal(main, stage_header, col, label)
    main.freeze_panes = f"B{stage_header + 1}"

    calc_cols = ("D", "E", "G", "H") if ranged else ("D", "G")
    detail_cols = ("E", "F", "G", "H") if ranged else ("E", "F")
    summary_cols = ("F", "G", "C", "D") if ranged else ("E", "C")
    price_from = 2 if ranged else 1
    row = stage_header + 1
    for summary_row, (phase, works) in enumerate(groups.items(), 3):
        style(summary_row, 3, end=1 + len(summary_headers))
        literal(main, summary_row, 2, phase)
        literal(main, summary_row, 5 if ranged else 4, schedule.get(phase, "уточняется"))
        style(row, 7, end=last_column)
        main.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_column)
        literal(main, row, 2, phase)
        row += 1
        first = row
        for work_id, (name, action, result, indices) in works.items():
            style(row, 8, end=last_column)
            literal(main, row, 2, f"{work_id} · {name}" if name != action else work_id)
            literal(main, row, 3, action)
            literal(main, row, 4, result)
            for index, (target, source) in enumerate(zip(detail_cols, calc_cols)):
                cells = [f"'Расчёт'!{source}{i}" for i in indices]
                main[f"{target}{row}"] = totals(cells, price=index >= price_from)
                main[f"{target}{row}"].number_format = number_format
            row += 1
        style(row, 9, end=last_column)
        literal(main, row, 2, "Итого по этапу")
        for index, (target, summary) in enumerate(zip(detail_cols, summary_cols)):
            price = index >= price_from
            main[f"{target}{row}"] = total_range(target, first, row - 1, price=price)
            main[f"{summary}{summary_row}"] = f'=IF(ISNUMBER({target}{row}),{target}{row},"")'
            for key in (f"{target}{row}", f"{summary}{summary_row}"):
                main[key].number_format = number_format
        row += 1

    for index, col in enumerate(summary_cols):
        main[f"{col}{grand}"] = total_range(col, 3, grand - 1, price=index >= price_from)
        main[f"{col}{grand}"].number_format = number_format
    main.print_area = f"B1:{get_column_letter(last_column)}{row - 1}"
    main.sheet_properties.pageSetUpPr.fitToPage = True
    main.page_setup.orientation = "landscape"
    main.page_setup.paperSize = main.PAPERSIZE_A4
    main.page_setup.fitToWidth = 1
    main.page_setup.fitToHeight = 0

    for sheet, widths, headers in (
        (params, [29, 27], (1, 6)),
        (extra, [27, 17, 15, 50], (1,)),
    ):
        sheet.freeze_panes = "A2"
        for i, width in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(i)].width = width
        for line in sheet.iter_rows():
            for cell in line:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                if cell.row in headers:
                    cell.fill = PatternFill("solid", fgColor="4285F4")
                    cell.font = Font(color="FFFFFF", bold=True)
                elif cell.data_type == "n" and cell.value is not None:
                    cell.number_format = number_format
    return book


def verify_book(expected, path):
    from openpyxl import load_workbook

    with path.open("rb") as file:
        actual = load_workbook(file, data_only=False)
    if expected.sheetnames != actual.sheetnames:
        raise ValueError("Изменился состав листов сметы")
    for sheet in expected:
        current = actual[sheet.title]
        if sheet.sheet_state != current.sheet_state or sheet.merged_cells != current.merged_cells:
            raise ValueError(f"Изменилась компоновка листа: {sheet.title}")
        if sheet.title == "Смета":
            widths = {key: dim.width for key, dim in sheet.column_dimensions.items()}
            actual_widths = {key: dim.width for key, dim in current.column_dimensions.items()}
            if widths != actual_widths or sheet.sheet_view.showGridLines != current.sheet_view.showGridLines:
                raise ValueError("Изменилась ширина колонок или сетка листа сметы")
        if current.max_row != sheet.max_row or current.max_column != sheet.max_column:
            raise ValueError(f"Изменился размер листа: {sheet.title}")
        for row in sheet:
            for cell in row:
                other = current[cell.coordinate]
                value = float(cell.value) if isinstance(cell.value, Decimal) else cell.value
                if value != other.value or (value is not None and cell.data_type != other.data_type):
                    raise ValueError(f"Смета расходится с расчётом: {sheet.title}!{cell.coordinate}")
                if sheet.title == "Смета" and any(
                    copy(getattr(cell, attribute)) != copy(getattr(other, attribute))
                    for attribute in ("font", "fill", "alignment", "number_format")
                ):
                    raise ValueError(f"Изменилось оформление сметы: {sheet.title}!{cell.coordinate}")


def verify_offer(expected, path):
    content = path.read_text(encoding="utf-8")
    if content.count(START) != 1 or content.count(END) != 1 or expected.rstrip("\n") not in content:
        raise ValueError("Коммерческий блок КП расходится с расчётом")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("build", "verify"))
    parser.add_argument("source", type=Path)
    parser.add_argument("--xlsx", type=Path)
    parser.add_argument("--offer", type=Path)
    args = parser.parse_args()
    if not (args.xlsx or args.offer):
        parser.error("Укажи --xlsx или --offer")
    paths = [path.resolve() for path in (args.source, args.xlsx, args.offer) if path]
    if len(set(paths)) != len(paths):
        parser.error("Вход и выходы должны иметь разные пути")
    try:
        estimate = load(args.source)
        with ExitStack() as stack:
            pending = []

            def output(path):
                if args.action == "verify":
                    return path
                directory = stack.enter_context(tempfile.TemporaryDirectory(dir=path.parent))
                temporary = Path(directory) / path.name
                pending.append((temporary, path))
                return temporary

            if args.xlsx:
                expected = workbook(estimate)
                path = output(args.xlsx)
                if args.action == "build":
                    expected.save(path)
                verify_book(expected, path)
            if args.offer:
                block = offer_block(estimate)
                path = output(args.offer)
                if args.action == "build":
                    path.write_text(block, encoding="utf-8")
                verify_offer(block, path)
            for temporary, destination in pending:
                temporary.replace(destination)
    except ModuleNotFoundError as error:
        parser.exit(1, f"{error}. Для XLSX установи openpyxl. Коммерческий блок можно собрать без него.\n")
    except (InvalidOperation, OSError, TypeError, ValueError, KeyError) as error:
        parser.exit(1, f"{error}\n")
    print("Проверка пройдена")


if __name__ == "__main__":
    main()
