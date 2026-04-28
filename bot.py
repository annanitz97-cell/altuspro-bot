import telebot
import requests
from datetime import datetime, timedelta
import re

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
    'PREPAYMENT_INVOICE': 'Счёт на предоплату',
    'EXECUTING': 'В работе - счёт оплачен',
    'UC_ZWS97R': 'Отгружен БЕЗ ДОКУМЕНТОВ',
    'UC_018IHX': 'ПРОРАБОТАТЬ',
    '6': 'Напомнить',
    '5': 'нет бюджета',
    'UC_IOD5R7': 'Документооборот',
    'UC_H0H6EG': 'Комплексная закупка',
    '4': 'Тендер',
    'UC_WLPIEC': 'Заказ до 10к/Озон',
    'UC_3W24DF': 'Не рассматривают аналоги',
    'WON': 'Сделка успешна ✅',
    'LOSE': 'Сделка провалена ❌',
    'APOLOGY': 'Спам ❌'
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
        return 'всё время'
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

def manager_report(manager_name, deals, date_from=None, date_to=None):
    filtered = [d for d in deals if manager_name.lower() in d.get('MANAGER','').lower()]
    if date_from:
        filtered = filter_by_period(filtered, date_from, date_to)

    won = [d for d in filtered if d.get('STAGE_ID') in WON_STAGES]
    lost = [d for d in filtered if d.get('STAGE_ID') in LOST_STAGES]
    expected = [d for d in filtered if d.get('STAGE_ID') in EXPECTED_STAGES]
    active = [d for d in filtered if d.get('STAGE_ID') not in WON_STAGES + LOST_STAGES]
    missed = [d for d in active if d.get('STAGE_ID') in MISSED_STAGES and days_since(d.get('DATE_CREATE','')) >= 3]
    
    revenue = sum(float(d.get('OPPORTUNITY',0) or 0) for d in won)
    expected_sum = sum(float(d.get('OPPORTUNITY',0) or 0) for d in expected)
    conv = round(len(won) / max(len(filtered), 1) * 100)

    stages = {}
    for d in active:
        s = STAGE_NAMES.get(d.get('STAGE_ID',''), d.get('STAGE_ID','—'))
        stages[s] = stages.get(s, 0) + 1

    full_name = filtered[0].get('MANAGER', manager_name) if filtered else manager_name
    period_str = format_period(date_from, date_to)

    text = f"👤 *{full_name}*\n"
    text += f"📅 Период: {period_str}\n\n"
    text += f"📊 Всего сделок: {len(filtered)}\n"
    text += f"💰 Оплачено: {len(won)} ({format_money(revenue)})\n"
    text += f"⏳ Ожидаем оплату: {len(expected)} ({format_money(expected_sum)})\n"
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

👤 *По менеджеру:*
`Александра` — за месяц
`Александра неделя` — за неделю
`Федоткин 01.04-28.04` — за период""", parse_mode='Markdown')

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
            managers[m] = {'won': 0, 'active': 0, 'lost': 0, 'expected': 0, 'revenue': 0, 'expected_sum': 0}
        if d.get('STAGE_ID') in WON_STAGES:
            managers[m]['won'] += 1
            managers[m]['revenue'] += float(d.get('OPPORTUNITY', 0) or 0)
        elif d.get('STAGE_ID') in LOST_STAGES:
            managers[m]['lost'] += 1
        elif d.get('STAGE_ID') in EXPECTED_STAGES:
            managers[m]['expected'] += 1
            managers[m]['expected_sum'] += float(d.get('OPPORTUNITY', 0) or 0)
            managers[m]['active'] += 1
        else:
            managers[m]['active'] += 1

    period_str = format_period(date_from, date_to)
    text = f"📊 *Все менеджеры* | {period_str}\n"
    text += f"Всего сделок: {len(filtered)}\n\n"

    for name, s in sorted(managers.items(), key=lambda x: x[1]['revenue'], reverse=True):
        total = s['won'] + s['active'] + s['lost']
        conv = round(s['won'] / max(total, 1) * 100)
        text += f"👤 *{name}*\n"
        text += f"  💰 {s['won']} ({format_money(s['revenue'])})"
        if s['expected'] > 0:
            text += f" | ⏳ {s['expected']} ({format_money(s['expected_sum'])})"
        text += f"\n  🔄 {s['active']} | ❌ {s['lost']} | 📈 {conv}%\n\n"

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

    total = sum(float(d.get('OPPORTUNITY',0) or 0) for d in expected)
    period_str = format_period(date_from, date_to)

    by_manager = {}
    for d in expected:
        m = d.get('MANAGER','—')
        if m not in by_manager:
            by_manager[m] = []
        by_manager[m].append(d)

    text = f"⏳ *Ожидаемые оплаты* | {period_str}\n"
    text += f"Итого: *{format_money(total)}* ({len(expected)} сделок)\n\n"

    for name, manager_deals in sorted(by_manager.items()):
        mgr_sum = sum(float(d.get('OPPORTUNITY',0) or 0) for d in manager_deals)
        text += f"👤 *{name}* — {format_money(mgr_sum)}\n"
        for d in manager_deals[:5]:
            text += f"  • {d.get('TITLE','—')} — {format_money(d.get('OPPORTUNITY',0))}\n"
        text += "\n"

    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['missed'])
def handle_missed(message):
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    missed = [d for d in deals
              if d.get('STAGE_ID') in MISSED_STAGES
              and days_since(d.get('DATE_CREATE','')) >= 3]

    if not missed:
        bot.reply_to(message, "🎉 Пропущенных заявок нет!")
        return

    text = f"🔴 *Пропущенные ({len(missed)})*\n\n"
    for d in missed[:20]:
        days = days_since(d.get('DATE_CREATE',''))
        emoji = "🔴" if days >= 7 else "🟡"
        text += f"{emoji} {d.get('TITLE','—')}\n"
        text += f"  👤 {d.get('MANAGER','—')} | {days} дн.\n\n"

    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['revenue'])
def handle_revenue(message):
    args = message.text.replace('/revenue', '').strip()
    date_from, date_to = parse_period(args) if args else (datetime.now() - timedelta(days=30), datetime.now())
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    filtered = filter_by_period(deals, date_from, date_to)
    won = [d for d in filtered if d.get('STAGE_ID') in WON_STAGES]
    total = sum(float(d.get('OPPORTUNITY',0) or 0) for d in won)

    by_manager = {}
    for d in won:
        m = d.get('MANAGER','—')
        by_manager[m] = by_manager.get(m, 0) + float(d.get('OPPORTUNITY',0) or 0)

    period_str = format_period(date_from, date_to)
    text = f"💰 *Выручка* | {period_str}\n"
    text += f"Итого: *{format_money(total)}*\n\n"
    for name, rev in sorted(by_manager.items(), key=lambda x: x[1], reverse=True):
        text += f"👤 {name}: {format_money(rev)}\n"

    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['attention'])
def handle_attention(message):
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    attention = [d for d in deals
                 if d.get('STAGE_ID') not in WON_STAGES + LOST_STAGES
                 and days_since(d.get('DATE_CREATE','')) >= 7]

    if not attention:
        bot.reply_to(message, "✅ Все сделки в норме!")
        return

    text = f"⚠️ *Требуют внимания ({len(attention)})*\n_(висят 7+ дней)_\n\n"
    for d in attention[:15]:
        days = days_since(d.get('DATE_CREATE',''))
        stage = STAGE_NAMES.get(d.get('STAGE_ID',''), d.get('STAGE_ID',''))
        text += f"• {d.get('TITLE','—')}\n"
        text += f"  👤 {d.get('MANAGER','—')} | {stage} | {days} дн.\n\n"

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
        bot.reply_to(message, f"❓ Менеджер '{name_part}' не найден.\n\nКоманды:\n`/all` — все менеджеры\n`/missed` — пропущенные\n`/revenue` — выручка\n`/expected` — ожидаемые оплаты\n`/attention` — требуют внимания", parse_mode='Markdown')
        return

    report = manager_report(manager_name, deals, date_from, date_to)
    bot.reply_to(message, report, parse_mode='Markdown')

if __name__ == '__main__':
    print("Бот запущен!")
    bot.infinity_polling()
