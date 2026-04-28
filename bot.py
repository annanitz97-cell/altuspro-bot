import telebot
import requests
from datetime import datetime, timedelta
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

WEBHOOK = 'https://altuspro.bitrix24.ru/rest/33/1r23rn74ww8yq892/'
BOT_TOKEN = '8612831715:AAE1OVngy867YStfZhEZUiyqcqbEAt_8ZA0'

bot = telebot.TeleBot(BOT_TOKEN)

WON_STAGES = ['WON', 'UC_ZWS97R', 'EXECUTING']
LOST_STAGES = ['LOSE', 'APOLOGY']
MISSED_STAGES = ['NEW', '1']
EXPECTED_STAGES = ['PREPAYMENT_INVOICE']

STAGE_NAMES = {
    'NEW': 'Потенциальная потребность',
    '1': 'Потребность подтверждена',
    'UC_FZ1R5C': 'ДЛЯ АНИ!!',
    'PREPARATION': 'ТКП направлено',
    '3': 'ТКП перекупам',
    '2': 'Работа с возражениями',
    'PREPAYMENT_INVOICE': 'Счет на предоплату',
    'EXECUTING': 'В работе - счет оплачен',
    'UC_ZWS97R': 'Отгружен БЕЗ ДОКУМЕНТОВ',
    'UC_018IHX': 'ПРОРАБОТАТЬ',
    '6': 'Напомнить',
    '5': 'нет бюджета',
    'UC_IOD5R7': 'Документооборот',
    'UC_H0H6EG': 'Комплексная закупка',
    '4': 'Тендер',
    'UC_WLPIEC': 'Заказ до 10к/Озон',
    'UC_3W24DF': 'Не рассматривают аналоги',
    'WON': 'Сделка успешна',
    'LOSE': 'Сделка провалена',
    'APOLOGY': 'Спам'
}

def bx(method, params=None):
    url = WEBHOOK + method + '.json'
    r = requests.get(url, params=params or {})
    return r.json().get('result', [])

def get_users():
    users = bx('user.get', {'select[]': ['ID', 'NAME', 'LAST_NAME']})
    return {u['ID']: (u.get('NAME','') + ' ' + u.get('LAST_NAME','')).strip() for u in users}

def get_deals():
    deals = []
    start = 0
    users = get_users()
    while True:
        params = {
            'select[]': ['ID','TITLE','STAGE_ID','ASSIGNED_BY_ID','OPPORTUNITY','DATE_CREATE','CLOSEDATE'],
            'order[DATE_CREATE]': 'DESC',
            'start': start
        }
        result = bx('crm.deal.list', params)
        if not result:
            break
        for d in result:
            d['MANAGER'] = users.get(d.get('ASSIGNED_BY_ID',''), 'Не назначен')
        deals.extend(result)
        if len(result) < 50 or len(deals) > 400:
            break
        start += 50
    return deals

def parse_date(s):
    for fmt in ['%d.%m.%Y', '%d.%m.%y', '%d.%m']:
        try:
            d = datetime.strptime(s.strip(), fmt)
            if fmt == '%d.%m':
                d = d.replace(year=datetime.now().year)
            return d
        except:
            pass
    return None

def parse_period(text):
    text = text.lower().strip()
    now = datetime.now()
    if 'неделя' in text or 'неделю' in text or 'неделе' in text:
        return now - timedelta(days=7), now
    if 'месяц' in text or 'месяца' in text or 'месяце' in text:
        return now - timedelta(days=30), now
    if 'квартал' in text:
        return now - timedelta(days=90), now
    if 'год' in text:
        return now - timedelta(days=365), now
    if 'сегодня' in text:
        return now.replace(hour=0,minute=0,second=0), now
    if 'вчера' in text:
        y = now - timedelta(days=1)
        return y.replace(hour=0,minute=0,second=0), y.replace(hour=23,minute=59,second=59)
    range_match = re.search(r'(\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)\s*[-–]\s*(\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)', text)
    if range_match:
        d1 = parse_date(range_match.group(1))
        d2 = parse_date(range_match.group(2))
        if d1 and d2:
            return d1, d2.replace(hour=23,minute=59,second=59)
    return None, None

def get_deal_date(d):
    date_str = d.get('DATE_CREATE','')
    if not date_str:
        return None
    try:
        return datetime.fromisoformat(date_str.replace('+03:00','').replace('T',' ').split('.')[0])
    except:
        return None

def filter_by_period(deals, date_from, date_to):
    if not date_from:
        return deals
    result = []
    for d in deals:
        dt = get_deal_date(d)
        if dt and date_from <= dt <= date_to:
            result.append(d)
    return result

def days_since(date_str):
    if not date_str:
        return 999
    try:
        d = datetime.fromisoformat(date_str.replace('+03:00','').replace('T',' ').split('.')[0])
        return (datetime.now() - d).days
    except:
        return 999

def format_money(v):
    try:
        n = float(v or 0)
        if n > 0:
            return f"{int(n):,}".replace(',', ' ') + ' ₽'
    except:
        pass
    return '—'

def format_period(date_from, date_to):
    if not date_from:
        return 'все время'
    return f"{date_from.strftime('%d.%m')}–{date_to.strftime('%d.%m.%Y')}"

def find_manager(name, deals):
    name_lower = name.lower().strip()
    if not name_lower:
        return None
    for d in deals:
        m = d.get('MANAGER','')
        if name_lower in m.lower():
            return m
    return None

def sum_revenue(deals_list):
    return sum(float(d.get('OPPORTUNITY',0) or 0) for d in deals_list)

def manager_report(manager_name, deals, date_from=None, date_to=None):
    filtered = [d for d in deals if manager_name.lower() in d.get('MANAGER','').lower()]
    if date_from:
        filtered = filter_by_period(filtered, date_from, date_to)

    won_final = [d for d in filtered if d.get('STAGE_ID') == 'WON']
    won_shipped = [d for d in filtered if d.get('STAGE_ID') == 'UC_ZWS97R']
    won_paid = [d for d in filtered if d.get('STAGE_ID') == 'EXECUTING']
    won = won_final + won_shipped + won_paid
    lost = [d for d in filtered if d.get('STAGE_ID') in LOST_STAGES]
    expected = [d for d in filtered if d.get('STAGE_ID') in EXPECTED_STAGES]
    active = [d for d in filtered if d.get('STAGE_ID') not in WON_STAGES + LOST_STAGES]
    missed = [d for d in active if d.get('STAGE_ID') in MISSED_STAGES and days_since(d.get('DATE_CREATE','')) >= 3]

    total_revenue = sum_revenue(won)
    expected_sum = sum_revenue(expected)
    conv = round(len(won) / max(len(filtered), 1) * 100)

    stages = {}
    for d in active:
        s = STAGE_NAMES.get(d.get('STAGE_ID',''), d.get('STAGE_ID','—'))
        stages[s] = stages.get(s, 0) + 1

    full_name = filtered[0].get('MANAGER', manager_name) if filtered else manager_name
    period_str = format_period(date_from, date_to)

    text = f"👤 *{full_name}*\n"
    text += f"📅 Период: {period_str}\n\n"
    text += f"📊 Всего сделок: {len(filtered)}\n\n"
    text += f"💰 *Оплаты:*\n"
    text += f"  ✅ Сделка успешна: {len(won_final)} ({format_money(sum_revenue(won_final))})\n"
    text += f"  📦 Отгружен без документов: {len(won_shipped)} ({format_money(sum_revenue(won_shipped))})\n"
    text += f"  💳 Счет оплачен: {len(won_paid)} ({format_money(sum_revenue(won_paid))})\n"
    text += f"  📈 Итого оплат: {len(won)} ({format_money(total_revenue)})\n\n"

    if expected:
        text += f"⏳ *Ожидаем оплату: {len(expected)}* ({format_money(expected_sum)})\n\n"

    text += f"❌ Провалов: {len(lost)}\n"
    text += f"🔄 В работе: {len(active)}\n"
    text += f"📈 Конверсия: {conv}%\n"

    if stages:
        text += f"\n📋 *По стадиям:*\n"
        for stage, count in sorted(stages.items(), key=lambda x: x[1], reverse=True)[:6]:
            text += f"  • {stage}: {count}\n"

    if missed:
        text += f"\n🔴 *Пропущенных: {len(missed)}*\n"
        for d in missed[:5]:
            days = days_since(d.get('DATE_CREATE',''))
            text += f"  • {d.get('TITLE','—')} ({days} дн.)\n"

    return text

def create_excel_report(deals, date_from, date_to):
    wb = openpyxl.Workbook()
    
    # Цвета
    BLUE = "1E3A5F"
    LIGHT_BLUE = "2E86AB"
    GREEN = "27AE60"
    ORANGE = "F39C12"
    RED = "E74C3C"
    GRAY = "95A5A6"
    LIGHT_GRAY = "ECF0F1"
    WHITE = "FFFFFF"
    YELLOW = "FFF9C4"
    GREEN_LIGHT = "E8F5E9"
    RED_LIGHT = "FFEBEE"

    def style_cell(cell, bold=False, color=None, bg=None, size=11, align='left', wrap=False):
        cell.font = Font(bold=bold, size=size, color=color or "000000", name='Calibri')
        if bg:
            cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
        cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)

    def border_cell(cell, style='thin'):
        thin = Side(style=style, color="CCCCCC")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ======= ЛИСТ 1: СВОДКА =======
    ws = wb.active
    ws.title = "Сводка по менеджерам"
    ws.sheet_view.showGridLines = False

    # Заголовок
    ws.merge_cells('A1:J1')
    ws['A1'] = f'ОТЧЕТ AltusPro CRM — {date_from.strftime("%d.%m.%Y")} – {date_to.strftime("%d.%m.%Y")}'
    style_cell(ws['A1'], bold=True, color=WHITE, bg=BLUE, size=14, align='center')
    ws.row_dimensions[1].height = 35

    ws.merge_cells('A2:J2')
    ws['A2'] = f'Сформирован: {datetime.now().strftime("%d.%m.%Y %H:%M")}'
    style_cell(ws['A2'], color="666666", bg=LIGHT_GRAY, size=10, align='center')
    ws.row_dimensions[2].height = 20

    ws.row_dimensions[3].height = 10

    # Заголовки таблицы
    headers = ['Менеджер', 'Всего сделок', 'Сделка успешна', 'Отгружен б/д', 'Счет оплачен',
               'Итого оплат (шт)', 'Итого оплат (руб)', 'Ожидаем оплату', 'В работе', 'Провалы', 'Конверсия']
    
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=h)
        style_cell(cell, bold=True, color=WHITE, bg=LIGHT_BLUE, size=10, align='center', wrap=True)
        border_cell(cell)
    ws.row_dimensions[4].height = 40

    # Данные по менеджерам
    managers = {}
    for d in filter_by_period(deals, date_from, date_to):
        m = d.get('MANAGER', 'Не назначен')
        if m not in managers:
            managers[m] = {'total':0,'won_f':0,'won_s':0,'won_p':0,'lost':0,'expected':0,'active':0,
                          'rev_f':0,'rev_s':0,'rev_p':0,'rev_e':0}
        sid = d.get('STAGE_ID','')
        opp = float(d.get('OPPORTUNITY',0) or 0)
        managers[m]['total'] += 1
        if sid == 'WON':
            managers[m]['won_f'] += 1; managers[m]['rev_f'] += opp
        elif sid == 'UC_ZWS97R':
            managers[m]['won_s'] += 1; managers[m]['rev_s'] += opp
        elif sid == 'EXECUTING':
            managers[m]['won_p'] += 1; managers[m]['rev_p'] += opp
        elif sid in LOST_STAGES:
            managers[m]['lost'] += 1
        elif sid in EXPECTED_STAGES:
            managers[m]['expected'] += 1; managers[m]['rev_e'] += opp; managers[m]['active'] += 1
        else:
            managers[m]['active'] += 1

    row = 5
    totals = {'total':0,'won_f':0,'won_s':0,'won_p':0,'lost':0,'expected':0,'active':0,
              'rev_f':0,'rev_s':0,'rev_p':0,'rev_e':0}

    for name, s in sorted(managers.items(), key=lambda x: x[1]['rev_f']+x[1]['rev_s']+x[1]['rev_p'], reverse=True):
        won_total = s['won_f'] + s['won_s'] + s['won_p']
        rev_total = s['rev_f'] + s['rev_s'] + s['rev_p']
        conv = round(won_total / max(s['total'], 1) * 100)
        bg = WHITE if row % 2 == 0 else LIGHT_GRAY

        values = [name, s['total'], s['won_f'], s['won_s'], s['won_p'],
                  won_total, rev_total, s['rev_e'], s['active'], s['lost'], f"{conv}%"]

        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=val)
            if col == 7:
                style_cell(cell, bg=GREEN_LIGHT if rev_total > 0 else bg, align='right')
                cell.number_format = '#,##0 "₽"'
            elif col == 8:
                style_cell(cell, bg=YELLOW if s['rev_e'] > 0 else bg, align='right')
                cell.number_format = '#,##0 "₽"'
            elif col == 11:
                style_cell(cell, bold=True, color=GREEN if conv >= 50 else (ORANGE if conv >= 25 else RED), bg=bg, align='center')
            elif col == 1:
                style_cell(cell, bold=True, bg=bg)
            else:
                style_cell(cell, bg=bg, align='center')
            border_cell(cell)

        for k in totals:
            totals[k] += s[k]
        row += 1

    # Итоговая строка
    won_total_all = totals['won_f'] + totals['won_s'] + totals['won_p']
    rev_total_all = totals['rev_f'] + totals['rev_s'] + totals['rev_p']
    conv_all = round(won_total_all / max(totals['total'], 1) * 100)

    total_values = ['ИТОГО', totals['total'], totals['won_f'], totals['won_s'], totals['won_p'],
                    won_total_all, rev_total_all, totals['rev_e'], totals['active'], totals['lost'], f"{conv_all}%"]

    for col, val in enumerate(total_values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        style_cell(cell, bold=True, color=WHITE, bg=BLUE, align='center' if col > 1 else 'left')
        if col == 7:
            cell.number_format = '#,##0 "₽"'
        if col == 8:
            cell.number_format = '#,##0 "₽"'
        border_cell(cell)
    ws.row_dimensions[row].height = 25

    # Ширина колонок
    col_widths = [22, 12, 14, 14, 13, 14, 16, 16, 11, 10, 11]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ======= ЛИСТ 2: ДЕТАЛИЗАЦИЯ =======
    ws2 = wb.create_sheet("Детализация сделок")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells('A1:G1')
    ws2['A1'] = f'ДЕТАЛИЗАЦИЯ СДЕЛОК — {date_from.strftime("%d.%m.%Y")} – {date_to.strftime("%d.%m.%Y")}'
    style_cell(ws2['A1'], bold=True, color=WHITE, bg=BLUE, size=13, align='center')
    ws2.row_dimensions[1].height = 30

    headers2 = ['Менеджер', 'Название сделки', 'Статус', 'Сумма', 'Дата создания', 'Дней в работе', 'Источник']
    for col, h in enumerate(headers2, 1):
        cell = ws2.cell(row=2, column=col, value=h)
        style_cell(cell, bold=True, color=WHITE, bg=LIGHT_BLUE, size=10, align='center')
        border_cell(cell)
    ws2.row_dimensions[2].height = 25

    filtered_deals = filter_by_period(deals, date_from, date_to)
    filtered_deals.sort(key=lambda d: (d.get('MANAGER',''), d.get('STAGE_ID','')))

    row2 = 3
    for d in filtered_deals:
        sid = d.get('STAGE_ID','')
        stage_name = STAGE_NAMES.get(sid, sid)
        days = days_since(d.get('DATE_CREATE',''))
        opp = float(d.get('OPPORTUNITY',0) or 0)

        if sid in ['WON']:
            bg = GREEN_LIGHT
        elif sid in ['UC_ZWS97R', 'EXECUTING']:
            bg = "#E3F2FD"
        elif sid in LOST_STAGES:
            bg = RED_LIGHT
        elif sid in EXPECTED_STAGES:
            bg = YELLOW
        elif days >= 7:
            bg = "#FFF3E0"
        else:
            bg = WHITE if row2 % 2 == 0 else LIGHT_GRAY

        values2 = [
            d.get('MANAGER','—'),
            d.get('TITLE','—'),
            stage_name,
            opp if opp > 0 else None,
            d.get('DATE_CREATE','')[:10] if d.get('DATE_CREATE') else '—',
            days if days < 999 else '—',
            d.get('SOURCE_ID','—') or '—'
        ]

        for col, val in enumerate(values2, 1):
            cell = ws2.cell(row=row2, column=col, value=val)
            style_cell(cell, bg=bg, wrap=(col==2))
            if col == 4 and val:
                cell.number_format = '#,##0 "₽"'
                style_cell(cell, bg=bg, align='right')
            border_cell(cell)
        row2 += 1

    col_widths2 = [22, 40, 22, 14, 13, 13, 12]
    for i, w in enumerate(col_widths2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # ======= ЛИСТ 3: ОЖИДАЕМЫЕ ОПЛАТЫ =======
    ws3 = wb.create_sheet("Ожидаемые оплаты")
    ws3.sheet_view.showGridLines = False

    ws3.merge_cells('A1:E1')
    ws3['A1'] = f'ОЖИДАЕМЫЕ ОПЛАТЫ (Счет на предоплату)'
    style_cell(ws3['A1'], bold=True, color=WHITE, bg="E67E22", size=13, align='center')
    ws3.row_dimensions[1].height = 30

    headers3 = ['Менеджер', 'Название сделки', 'Сумма', 'Дата создания', 'Дней ожидания']
    for col, h in enumerate(headers3, 1):
        cell = ws3.cell(row=2, column=col, value=h)
        style_cell(cell, bold=True, color=WHITE, bg=LIGHT_BLUE, size=10, align='center')
        border_cell(cell)

    expected_deals = [d for d in filter_by_period(deals, date_from, date_to) if d.get('STAGE_ID') in EXPECTED_STAGES]
    row3 = 3
    total_expected = 0
    for d in sorted(expected_deals, key=lambda x: x.get('MANAGER','')):
        opp = float(d.get('OPPORTUNITY',0) or 0)
        total_expected += opp
        days = days_since(d.get('DATE_CREATE',''))
        bg = "#FFF3E0" if days > 7 else YELLOW

        vals3 = [d.get('MANAGER','—'), d.get('TITLE','—'), opp if opp > 0 else None,
                 d.get('DATE_CREATE','')[:10] if d.get('DATE_CREATE') else '—', days if days < 999 else '—']
        for col, val in enumerate(vals3, 1):
            cell = ws3.cell(row=row3, column=col, value=val)
            style_cell(cell, bg=bg)
            if col == 3 and val:
                cell.number_format = '#,##0 "₽"'
                style_cell(cell, bg=bg, align='right')
            border_cell(cell)
        row3 += 1

    # Итог
    ws3.cell(row=row3, column=1, value='ИТОГО').font = Font(bold=True)
    total_cell = ws3.cell(row=row3, column=3, value=total_expected)
    style_cell(total_cell, bold=True, color=WHITE, bg="E67E22", align='right')
    total_cell.number_format = '#,##0 "₽"'

    col_widths3 = [22, 40, 16, 14, 14]
    for i, w in enumerate(col_widths3, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # Сохраняем в буфер
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

@bot.message_handler(commands=['report'])
def handle_report(message):
    now = datetime.now()
    date_from = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    date_to = now

    bot.reply_to(message, f"⏳ Формирую отчет за {date_from.strftime('%d.%m')}–{date_to.strftime('%d.%m.%Y')}...")
    
    deals = get_deals()
    excel_buffer = create_excel_report(deals, date_from, date_to)
    
    filename = f"AltusPro_отчет_{date_from.strftime('%d.%m')}-{date_to.strftime('%d.%m.%Y')}.xlsx"
    bot.send_document(message.chat.id, excel_buffer, visible_file_name=filename,
                      caption=f"📊 Отчет AltusPro CRM\n📅 {date_from.strftime('%d.%m')}–{date_to.strftime('%d.%m.%Y')}\n3 листа: Сводка | Детализация | Ожидаемые оплаты")

@bot.message_handler(commands=['start', 'help'])
def handle_start(message):
    bot.reply_to(message, """👋 Привет! Я CRM бот AltusPro.

📊 *Команды:*
`/all` — все менеджеры за месяц
`/all неделя` — за неделю
`/all 01.04-28.04` — за период
`/missed` — пропущенные заявки
`/revenue` — выручка
`/expected` — ожидаемые оплаты
`/attention` — требуют внимания
`/report` — Excel отчет с 1-го числа

👤 *По менеджеру:*
`Александра` — за месяц
`Федоткин неделя` — за неделю
`Разумовская 01.04-28.04` — за период""", parse_mode='Markdown')

@bot.message_handler(commands=['all'])
def handle_all(message):
    args = message.text.replace('/all', '').strip()
    date_from, date_to = parse_period(args) if args else (datetime.now() - timedelta(days=30), datetime.now())
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    filtered = filter_by_period(deals, date_from, date_to)

    managers = {}
    for d in filtered:
        m = d.get('MANAGER', 'Не назначен')
        if m not in managers:
            managers[m] = {'won_f':0,'won_s':0,'won_p':0,'active':0,'lost':0,'expected':0,'revenue':0,'expected_sum':0}
        sid = d.get('STAGE_ID')
        opp = float(d.get('OPPORTUNITY',0) or 0)
        if sid == 'WON':
            managers[m]['won_f'] += 1; managers[m]['revenue'] += opp
        elif sid == 'UC_ZWS97R':
            managers[m]['won_s'] += 1; managers[m]['revenue'] += opp
        elif sid == 'EXECUTING':
            managers[m]['won_p'] += 1; managers[m]['revenue'] += opp
        elif sid in LOST_STAGES:
            managers[m]['lost'] += 1
        elif sid in EXPECTED_STAGES:
            managers[m]['expected'] += 1; managers[m]['expected_sum'] += opp; managers[m]['active'] += 1
        else:
            managers[m]['active'] += 1

    period_str = format_period(date_from, date_to)
    text = f"📊 *Все менеджеры* | {period_str}\nВсего: {len(filtered)} сделок\n\n"

    for name, s in sorted(managers.items(), key=lambda x: x[1]['revenue'], reverse=True):
        won_total = s['won_f'] + s['won_s'] + s['won_p']
        total = won_total + s['active'] + s['lost']
        conv = round(won_total / max(total, 1) * 100)
        text += f"👤 *{name}*\n"
        text += f"  ✅{s['won_f']} 📦{s['won_s']} 💳{s['won_p']} | 💰{format_money(s['revenue'])}\n"
        if s['expected'] > 0:
            text += f"  ⏳ Ожидаем: {s['expected']} ({format_money(s['expected_sum'])})\n"
        text += f"  🔄{s['active']} ❌{s['lost']} 📈{conv}%\n\n"

    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['expected'])
def handle_expected(message):
    args = message.text.replace('/expected', '').strip()
    date_from, date_to = parse_period(args) if args else (datetime.now() - timedelta(days=30), datetime.now())
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    filtered = filter_by_period(deals, date_from, date_to)
    expected = [d for d in filtered if d.get('STAGE_ID') in EXPECTED_STAGES]

    if not expected:
        bot.reply_to(message, "✅ Нет сделок в ожидании оплаты!")
        return

    total = sum_revenue(expected)
    period_str = format_period(date_from, date_to)
    by_manager = {}
    for d in expected:
        m = d.get('MANAGER','—')
        if m not in by_manager:
            by_manager[m] = []
        by_manager[m].append(d)

    text = f"⏳ *Ожидаемые оплаты* | {period_str}\nИтого: *{format_money(total)}* ({len(expected)} сделок)\n\n"
    for name, mgr_deals in sorted(by_manager.items()):
        mgr_sum = sum_revenue(mgr_deals)
        text += f"👤 *{name}* — {format_money(mgr_sum)}\n"
        for d in mgr_deals[:5]:
            text += f"  • {d.get('TITLE','—')} — {format_money(d.get('OPPORTUNITY',0))}\n"
        text += "\n"

    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['missed'])
def handle_missed(message):
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    missed = [d for d in deals if d.get('STAGE_ID') in MISSED_STAGES and days_since(d.get('DATE_CREATE','')) >= 3]

    if not missed:
        bot.reply_to(message, "🎉 Пропущенных заявок нет!")
        return

    text = f"🔴 *Пропущенные ({len(missed)})*\n\n"
    for d in missed[:20]:
        days = days_since(d.get('DATE_CREATE',''))
        emoji = "🔴" if days >= 7 else "🟡"
        text += f"{emoji} {d.get('TITLE','—')}\n  👤 {d.get('MANAGER','—')} | {days} дн.\n\n"

    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['revenue'])
def handle_revenue(message):
    args = message.text.replace('/revenue', '').strip()
    date_from, date_to = parse_period(args) if args else (datetime.now() - timedelta(days=30), datetime.now())
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    filtered = filter_by_period(deals, date_from, date_to)
    won = [d for d in filtered if d.get('STAGE_ID') in WON_STAGES]
    total = sum_revenue(won)

    by_manager = {}
    for d in won:
        m = d.get('MANAGER','—')
        by_manager[m] = by_manager.get(m, 0) + float(d.get('OPPORTUNITY',0) or 0)

    period_str = format_period(date_from, date_to)
    text = f"💰 *Выручка* | {period_str}\nИтого: *{format_money(total)}*\n\n"
    for name, rev in sorted(by_manager.items(), key=lambda x: x[1], reverse=True):
        text += f"👤 {name}: {format_money(rev)}\n"

    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['attention'])
def handle_attention(message):
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    attention = [d for d in deals if d.get('STAGE_ID') not in WON_STAGES + LOST_STAGES and days_since(d.get('DATE_CREATE','')) >= 7]

    if not attention:
        bot.reply_to(message, "✅ Все сделки в норме!")
        return

    text = f"⚠️ *Требуют внимания ({len(attention)})*\n_(висят 7+ дней)_\n\n"
    for d in attention[:15]:
        days = days_since(d.get('DATE_CREATE',''))
        stage = STAGE_NAMES.get(d.get('STAGE_ID',''), d.get('STAGE_ID',''))
        text += f"• {d.get('TITLE','—')}\n  👤 {d.get('MANAGER','—')} | {stage} | {days} дн.\n\n"

    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(func=lambda m: True)
def handle_text(message):
    text = message.text.strip()
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()

    date_from, date_to = parse_period(text)
    name_part = re.sub(r'(неделя|неделю|неделе|месяц|месяца|квартал|год|сегодня|вчера|\d{1,2}\.\d{1,2}(?:\.\d{2,4})?[-–]\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)', '', text, flags=re.IGNORECASE).strip()

    if not date_from:
        date_from = datetime.now() - timedelta(days=30)
        date_to = datetime.now()

    manager_name = find_manager(name_part, deals)

    if not manager_name:
        bot.reply_to(message, f"❓ Менеджер '{name_part}' не найден.\n\nКоманды:\n`/all` — все менеджеры\n`/missed` — пропущенные\n`/revenue` — выручка\n`/expected` — ожидаемые оплаты\n`/report` — Excel отчет\n`/attention` — требуют внимания", parse_mode='Markdown')
        return

    report = manager_report(manager_name, deals, date_from, date_to)
    bot.reply_to(message, report, parse_mode='Markdown')

if __name__ == '__main__':
    print("Бот запущен!")
    bot.infinity_polling()
