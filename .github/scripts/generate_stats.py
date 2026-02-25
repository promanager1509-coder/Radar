#!/usr/bin/env python3
"""
generate_stats.py
Запускается GitHub Actions при каждом коммите.
Сканирует все JSON файлы в /data/developers/,
считает объекты, записывает /data/meta/stats.json
"""

import json
import os
import glob
from datetime import datetime, timezone

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DEVELOPERS_DIR = os.path.join(BASE_DIR, "data", "developers")
STATS_FILE = os.path.join(BASE_DIR, "data", "meta", "stats.json")
INDEX_FILE = os.path.join(BASE_DIR, "data", "meta", "index.json")

def load_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"⚠️  Ошибка чтения {path}: {e}")
        return None

def normalize_type(raw_type):
    """Нормализует тип помещения к стандарту"""
    t = (raw_type or "").lower().strip()
    if "офис" in t or "office" in t:
        return "Офис"
    if "габ" in t or "gab" in t or "готов" in t:
        return "ГАБ"
    if "пвз" in t or "pvz" in t or "пункт выдачи" in t:
        return "ПВЗ"
    if "премиум" in t or "premium" in t or "элит" in t:
        return "Премиум"
    return "ПСН"

def main():
    now = datetime.now(timezone.utc)
    seven_days_ago = now.timestamp() - (7 * 24 * 60 * 60)

    stats = {
        "version": "1.0",
        "description": "Автогенерируется скриптом GitHub Actions. Не редактировать вручную.",
        "generated": now.strftime("%Y-%m-%dT%H:%M:%S"),
        "total": 0,
        "sale": 0,
        "rent": 0,
        "added_last_7days": 0,
        "last_updated_developer": "",
        "last_updated_file": "",
        "last_updated_date": "",
        "by_category": {
            "ПСН": 0,
            "Офис": 0,
            "Аренда ПСН": 0,
            "ПВЗ": 0,
            "ГАБ": 0,
            "Премиум": 0
        },
        "by_developer": {}
    }

    latest_mtime = 0
    latest_file = ""
    latest_developer = ""

    # Сканируем все JSON файлы кроме шаблона
    pattern = os.path.join(DEVELOPERS_DIR, "**", "*.json")
    all_files = glob.glob(pattern, recursive=True)
    all_files = [f for f in all_files if "_template" not in f]

    if not all_files:
        print("⚠️  Файлы с данными не найдены в data/developers/")
    
    seen_ids = set()

    for filepath in sorted(all_files):
        data = load_json(filepath)
        if not data:
            continue

        units = data.get("units", [])
        developer = data.get("developer", os.path.basename(os.path.dirname(filepath)))
        deal = data.get("deal", "sale")
        is_rent = deal == "rent"

        if developer not in stats["by_developer"]:
            stats["by_developer"][developer] = 0

        file_mtime = os.path.getmtime(filepath)
        count_in_file = 0

        for unit in units:
            uid = unit.get("id", "")
            # Дедупликация по id
            if uid and uid in seen_ids:
                continue
            if uid:
                seen_ids.add(uid)

            area = float(unit.get("area", 0) or 0)
            if area <= 0:
                continue  # пропускаем объекты без площади

            count_in_file += 1
            stats["total"] += 1
            stats["by_developer"][developer] += 1

            # Продажа / аренда
            unit_deal = unit.get("deal", deal)
            unit_is_rent = unit_deal == "rent"
            if unit_is_rent:
                stats["rent"] += 1
            else:
                stats["sale"] += 1

            # По категории
            unit_type = normalize_type(unit.get("type", "ПСН"))
            if unit_is_rent and unit_type == "ПСН":
                stats["by_category"]["Аренда ПСН"] += 1
            elif unit_is_rent and unit_type == "ПВЗ":
                stats["by_category"]["ПВЗ"] += 1
            elif unit_type in stats["by_category"]:
                stats["by_category"][unit_type] += 1
            else:
                stats["by_category"]["ПСН"] += 1

        # Отслеживаем самый свежий файл
        if file_mtime > latest_mtime:
            latest_mtime = file_mtime
            latest_file = os.path.basename(filepath)
            latest_developer = developer

        # Объекты добавленные за 7 дней (по дате файла)
        if file_mtime > seven_days_ago:
            stats["added_last_7days"] += count_in_file

        print(f"✅ {os.path.basename(filepath)}: {count_in_file} объектов ({developer})")

    # Последнее обновление
    stats["last_updated_file"] = latest_file
    stats["last_updated_developer"] = latest_developer
    stats["last_updated_date"] = datetime.fromtimestamp(latest_mtime).strftime("%Y-%m-%d") if latest_mtime else ""

    # Записываем stats.json
    with open(STATS_FILE, "w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)

    print(f"\n📊 ИТОГО:")
    print(f"   Всего объектов: {stats['total']}")
    print(f"   Продажа: {stats['sale']}, Аренда: {stats['rent']}")
    print(f"   За 7 дней: +{stats['added_last_7days']}")
    print(f"   Последнее обновление: {stats['last_updated_developer']} ({stats['last_updated_date']})")
    print(f"\n✅ stats.json обновлён: {STATS_FILE}")

if __name__ == "__main__":
    main()
