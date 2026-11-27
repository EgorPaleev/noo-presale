"""Connect NooRules without rewriting project files or the other plugin's settings."""

import argparse
import re
from pathlib import Path

LEGACY = "noorules:task-entry"
SETUP = {
    "noo-setup": ("dev", "Dev", "noo-task", "noorules", "задач разработки"),
    "noo-presale-setup": ("presale", "Presale", "noo-presale-task", "noorules-presale", "пресейлов"),
}
FENCE = re.compile(r" {0,3}(`{3,}|~{3,})")
MARKER = re.compile(r"<!-- (/?)((?:noorules:)(?:(?:dev|presale):)?task-entry) -->")


def blocks(content):
    fence = None
    current = None
    for line in content.splitlines():
        boundary = FENCE.match(line)
        if boundary:
            token = boundary[1]
            if fence is None:
                fence = token
            elif token.startswith(fence) and not line[boundary.end():].strip():
                fence = None
            continue
        if fence:
            continue
        marker = MARKER.fullmatch(line.strip().lstrip("\ufeff"))
        if marker:
            closing, name = marker.groups()
            if not closing and current is None:
                current = (name, [])
            elif closing and current is not None and name == current[0]:
                yield current[0], "\n".join(current[1])
                current = None
            else:
                raise ValueError("Некорректный блок NooRules в проектных инструкциях")
        elif current is not None:
            current[1].append(line)
    if current is not None:
        raise ValueError("Незакрытый блок NooRules в проектных инструкциях")


def paths(project, entry_name, namespace):
    entry = project / entry_name
    rules = project / ".noorules" / namespace / "rules.md"
    for path in (entry, rules):
        if path.resolve() != path or path.is_symlink():
            raise ValueError(f"Путь подключения указывает на символьную ссылку: {path}")
        if path.exists() and not path.is_file():
            raise ValueError(f"Ожидался файл: {path}")
    return entry, rules


def connect(project, entry_name, setup):
    namespace, label, task, package, purpose = SETUP[setup]
    entry, rules = paths(project, entry_name, namespace)
    current = entry.read_bytes() if entry.exists() else b""
    content = current.decode("utf-8")
    marker = f"noorules:{namespace}:task-entry"
    own = [
        name for name, body in blocks(content)
        if name == marker or (name == LEGACY and f"`{task}`" in body)
    ]
    if len(own) > 1:
        raise ValueError(f"Несколько блоков NooRules {label}. Уточни действующую точку входа.")

    created = []
    if not rules.exists():
        template = Path(__file__).resolve().parent.parent / "references" / "rules-template.md"
        data = template.read_bytes()
        rules.parent.mkdir(parents=True, exist_ok=True)
        with rules.open("xb") as file:
            file.write(data)
        created.append(rules.relative_to(project).as_posix())

    if not own:
        newline = "\r\n" if "\r\n" in content else "\n"
        block = newline.join([
            f"<!-- {marker} -->",
            f"## NooRules {label}",
            "",
            f"Используй NooRules {label} для {purpose} в этом проекте через `{task}`.",
            f"Если вызов недоступен, открой `skills/{task}/SKILL.md` в установленном пакете `{package}`.",
            f"Правила плагина находятся в [`{rules.relative_to(project).as_posix()}`]({rules.relative_to(project).as_posix()}).",
            f"<!-- /{marker} -->",
            "",
        ]).encode("utf-8")
        with entry.open("ab") as file:
            if current:
                file.write(newline.encode() if content.endswith("\n") else (newline * 2).encode())
            file.write(block)
        created.append(entry.relative_to(project).as_posix())

    return created


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", type=Path, required=True)
    parser.add_argument("--entry", choices=("AGENTS.md", "CLAUDE.md"), required=True)
    args = parser.parse_args()
    project = args.project.resolve()
    if not project.is_dir():
        parser.error("Корень проекта должен существовать")
    try:
        setup = Path(__file__).resolve().parent.parent.name
        changed = connect(project, args.entry, setup)
    except (OSError, ValueError) as error:
        parser.exit(1, f"{error}\n")
    print("\n".join(changed) if changed else "Подключение уже настроено, изменений нет")


if __name__ == "__main__":
    main()
