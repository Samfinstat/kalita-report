#!/usr/bin/env python3
"""
КАЛИТА — Скрипт обновления данных в отчёте
Использование:
  python update_report.py --html index.html --spend Директ.xlsx --seo SEO.xlsx
"""

import argparse
import json
import re
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path


def parse_direct(filepath: str) -> dict:
    """Парсит экспорт из Яндекс.Директ (Мастер отчётов)."""
    df = pd.read_excel(filepath, header=None, skiprows=3)
    cols = ['Дата','Кампания','НомерКампании','Группа','НомерГруппы','НомерОбъявления',
            'Условие_показа','НомерУсловия','Корректировки','Устройство','Пол','Возраст',
            'Показы','Клики','CTR','Расход','CPC','Отказы','CPM','Глубина','CPA','Конверсии']
    df.columns = cols[:len(df.columns)]

    df = df[df['Дата'].notna() & (df['Дата'] != 'Дата')]
    df['Дата'] = pd.to_datetime(df['Дата'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['Дата'])

    for col in ['Показы','Клики','Расход','Конверсии']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df['Месяц'] = df['Дата'].dt.to_period('M')
    monthly = df.groupby('Месяц').agg(
        Показы=('Показы','sum'), Клики=('Клики','sum'),
        Расход=('Расход','sum'), Конверсии=('Конверсии','sum')
    ).reset_index()
    monthly['CTR'] = (monthly['Клики'] / monthly['Показы'] * 100).round(2)
    monthly['CPC'] = (monthly['Расход'] / monthly['Клики'].replace(0, np.nan)).round(2)

    MONTHS_RU = {
        1:'Янв', 2:'Фев', 3:'Мар', 4:'Апр', 5:'Май',
        6:'Июн', 7:'Июл', 8:'Авг', 9:'Сен', 10:'Окт', 11:'Ноя', 12:'Дек'
    }

    result = {
        'months': [],
        'spend':  [],
        'clicks': [],
        'imps':   [],
        'convs':  [],
        'ctr':    [],
        'cpc':    [],
    }
    for _, row in monthly.iterrows():
        p = row['Месяц']
        result['months'].append(f"{MONTHS_RU[p.month]} {p.year}")
        result['spend'].append(round(float(row['Расход']), 0))
        result['clicks'].append(int(row['Клики']))
        result['imps'].append(int(row['Показы']))
        result['convs'].append(int(row['Конверсии']))
        result['ctr'].append(round(float(row['CTR'] or 0), 2))
        result['cpc'].append(round(float(row['CPC'] or 0), 2))

    return result


def parse_seo(filepath: str) -> dict:
    """Парсит экспорт позиций из Яндекс.Вебмастер."""
    df = pd.read_excel(filepath)
    date_cols = [c for c in df.columns if hasattr(c, 'date')]
    for col in date_cols:
        df[col] = pd.to_numeric(df.iloc[:, df.columns.get_loc(col)].replace('--', np.nan), errors='coerce')

    latest = date_cols[0] if date_cols else None
    earliest = date_cols[-1] if len(date_cols) > 1 else None

    stats = {}
    if latest:
        vals = df[latest]
        stats = {
            'top3':    int((vals <= 3).sum()),
            'top10':   int((vals <= 10).sum()),
            'top30':   int((vals <= 30).sum()),
            'top100':  int((vals <= 100).sum()),
            'out':     int(vals.isna().sum()),
            'total':   len(df),
        }
    return stats


def update_html(html_path: str, direct: dict, seo: dict, password: str = None):
    """Встраивает обновлённые данные в HTML-файл."""
    content = Path(html_path).read_text(encoding='utf-8')

    def replace_js_var(name: str, value, text: str) -> str:
        val_str = json.dumps(value, ensure_ascii=False)
        pattern = rf'(const {name}\s*=\s*)(\[.*?\]|\'.*?\'|".*?");'
        replacement = rf'\g<1>{val_str};'
        return re.sub(pattern, replacement, text, flags=re.DOTALL)

    content = replace_js_var('months', direct['months'], content)
    content = replace_js_var('spend',  direct['spend'],  content)
    content = replace_js_var('convs',  direct['convs'],  content)
    content = replace_js_var('ctrs',   direct['ctr'],    content)
    content = replace_js_var('cpcs',   direct['cpc'],    content)

    # Update date in topbar
    today = datetime.today().strftime('%d.%m.%Y')
    content = re.sub(r'Отчёт обновлён: [\d\.]+', f'Отчёт обновлён: {today}', content)

    if password:
        content = re.sub(r"const PASSWORD = '.*?';", f"const PASSWORD = '{password}';", content)

    Path(html_path).write_text(content, encoding='utf-8')
    print(f"✅ Отчёт обновлён: {html_path}")
    print(f"   Месяцев в данных: {len(direct['months'])}")
    print(f"   Последний месяц: {direct['months'][-1]}")
    if seo:
        print(f"   SEO ТОП-10: {seo.get('top10', '?')} ключей")


def main():
    parser = argparse.ArgumentParser(description='Обновление отчёта КАЛИТА')
    parser.add_argument('--html',     default='index.html',     help='Путь к HTML-отчёту')
    parser.add_argument('--spend',    required=False,           help='Файл экспорта Яндекс.Директ (.xlsx)')
    parser.add_argument('--seo',      required=False,           help='Файл позиций Яндекс.Вебмастер (.xlsx)')
    parser.add_argument('--password', required=False,           help='Новый пароль (опционально)')
    args = parser.parse_args()

    direct = {}
    seo    = {}

    if args.spend:
        print(f"📊 Читаю данные Директ: {args.spend}")
        direct = parse_direct(args.spend)

    if args.seo:
        print(f"🔍 Читаю SEO-позиции: {args.seo}")
        seo = parse_seo(args.seo)

    if direct or seo:
        update_html(args.html, direct, seo, args.password)
    else:
        print("⚠️  Укажите хотя бы один файл данных: --spend или --seo")


if __name__ == '__main__':
    main()
