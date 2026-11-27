"""Inspect a presale run or accept a changed external input."""

import argparse
import hashlib
import json
import re
import shutil
import tempfile
from pathlib import Path

IDENTIFIER = r"[a-z0-9][a-z0-9-]{0,63}"


def file_ref(project, relative):
    path = (project / relative).resolve()
    location = path.relative_to(project).as_posix()
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(65536), b""):
            digest.update(chunk)
    return {"path": location, "fingerprint": "sha256:" + digest.hexdigest()}


def refresh(run, role, updated):
    inputs = run["externalInputs"]
    if inputs.get(role) == updated:
        return []
    stages = run["steps"]
    if role not in inputs and not any(role in step.get("optionalInputs", []) for step in stages):
        raise ValueError(f"Неизвестная роль внешнего входа: {role}")
    inputs[role] = updated
    affected = set()
    changed_outputs = {role}
    while True:
        previous = len(affected)
        for step in stages:
            consumed = set(step["inputs"]) | set(step.get("optionalInputs", [])) | set(step.get("usedInputs", {}))
            if consumed & changed_outputs or affected.intersection(step.get("dependsOn", [])):
                affected.add(step["id"])
                changed_outputs.add(step["output"])
        if len(affected) == previous:
            break
    for step in stages:
        if step["id"] in affected:
            step["status"] = "stale" if "result" in step else "pending"
            step["reason"] = f"Обновлён вход: {role}"
            step.pop("activity", None)
    return [step["id"] for step in stages if step["id"] in affected]


def save(path, run):
    with tempfile.TemporaryDirectory(dir=path.parent) as directory:
        temporary = Path(directory) / "run.json"
        temporary.write_text(json.dumps(run, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        temporary.replace(path)


def check_demo(run, project):
    demo = next((step for step in run["steps"] if step["task"] == "presale-demo"), None)
    if demo is None:
        return
    evidence = demo.get("demoEvidence")
    if not evidence:
        raise ValueError("Для прототипа нет отчёта о визуальной приёмке")
    foundation = evidence["visualFoundation"]
    if foundation not in {"project", "default"}:
        raise ValueError("Укажи применённую визуальную основу прототипа")
    required = {"platforms.md", "review.md"}
    if foundation == "default":
        required.add("visual-profile.md")
    used = {
        Path(item["material"]).name for item in demo.get("appliedMaterials", {}).get("uses", [])
        if item.get("effect", "").strip()
    }
    if required - used:
        raise ValueError("Не зафиксировано применение справок: " + ", ".join(sorted(required - used)))
    views = evidence["views"]
    expected = sorted(demo.get("targetPlatforms", ["desktop", "mobile"]))
    if sorted(view["platform"] for view in views) != expected:
        raise ValueError("Нет визуальной проверки всех целевых платформ прототипа")
    entry = (project / demo["result"]["path"]).resolve()
    entry.relative_to(project)
    if not entry.is_file() or not entry.stat().st_size:
        raise ValueError("Нет рабочей точки входа прототипа")
    screenshots = set()
    for view in views:
        screenshot = (project / view["screenshot"]).resolve()
        screenshot.relative_to(project)
        with screenshot.open("rb") as file:
            header = file.read(24)
        if (header[:8] != b"\x89PNG\r\n\x1a\n" or header[12:16] != b"IHDR"
                or len(header) != 24 or screenshot in screenshots or not view["scenario"].strip()
                or int.from_bytes(header[16:20], "big") < view["width"]
                or int.from_bytes(header[20:24], "big") < view["height"]):
            raise ValueError(f"Нужен отдельный PNG-снимок и пройденный сценарий: {view['platform']}")
        screenshots.add(screenshot)


def close(path, run, project):
    if any(step["status"] != "completed" for step in run["steps"]):
        raise ValueError("В запуске остались незавершённые этапы")
    delivery = (project / run["deliveryPath"]).resolve()
    delivery.relative_to(project)
    if delivery == project or delivery.is_relative_to(path.parent.parent) or not delivery.is_dir():
        raise ValueError("Укажи существующую папку выдачи вне служебного каталога")
    if not run["delivered"]:
        raise ValueError("В запуске не записаны выданные результаты")
    for relative in run["delivered"]:
        file = (project / relative).resolve()
        if not file.is_relative_to(delivery) or not file.is_file():
            raise ValueError(f"Результат отсутствует в папке выдачи: {relative}")
    check_demo(run, project)
    shutil.rmtree(path.parent)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", type=Path, default=Path.cwd())
    parser.add_argument("run")
    action = parser.add_mutually_exclusive_group()
    action.add_argument("--set-input", nargs=2, metavar=("ROLE", "FILE"))
    action.add_argument("--prepare", action="store_true")
    action.add_argument("--check-demo", action="store_true")
    action.add_argument("--close", action="store_true")
    args = parser.parse_args()
    if not re.fullmatch(IDENTIFIER, args.run):
        parser.error("Некорректный ID запуска")
    project = args.project.resolve()
    runs = project / ".noorules" / "runs"
    directory = runs / args.run
    path = directory / "run.json"
    try:
        if path.resolve() != path:
            raise ValueError("Служебный путь запуска не должен проходить через символьную ссылку")
        if args.prepare:
            directory.mkdir(parents=True, exist_ok=True)
            ignore = runs / ".gitignore"
            if ignore.is_symlink():
                raise ValueError("Служебный .gitignore не должен быть символьной ссылкой")
            if not ignore.exists():
                ignore.write_text("*\n", encoding="utf-8")
            return
        run = json.loads(path.read_text(encoding="utf-8"))
        if args.set_input:
            role, relative = args.set_input
            if not re.fullmatch(IDENTIFIER, role):
                raise ValueError(f"Некорректная роль входа: {role}")
            affected = refresh(run, role, file_ref(project, relative))
            save(path, run)
            print(json.dumps({"affected": affected}, ensure_ascii=False))
        elif args.check_demo:
            check_demo(run, project)
            print("Доказательства проверки прототипа на месте")
        elif args.close:
            close(path, run, project)
            print("Служебные материалы удалены")
        else:
            for step in run["steps"]:
                activity = step.get("activity", {})
                executor = activity.get("executor") or step.get("executor")
                model = step.get("model", {}).get("used")
                message = step.get("reason") if step["status"] in ("waiting", "stale") else activity.get("message")
                print(" ".join(filter(None, (step["id"], step["status"], executor, model, message))))
    except (OSError, ValueError, KeyError) as error:
        parser.exit(1, f"{error}\n")


if __name__ == "__main__":
    main()
