import telebot
import requests
from datetime import datetime, timedelta

WEBHOOK = 'https://altuspro.bitrix24.ru/rest/33/1r23rn74ww8yq892/'
BOT_TOKEN = '8612831715:AAE1OVngy867YStfZhEZUiyqcqbEAt_8ZA0'

bot = telebot.TeleBot(BOT_TOKEN)

WON_STAGES = ['WON', 'UC_ZWS97R']
LOST_STAGES = ['LOSE', 'APOLOGY']

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

def days_since(date_str):
    if not date_str:
        return 999
    try:
        d = datetime.fromisoformat(date_str.replace('+03:00',''))
        return (datetime.now() - d).days
    except:
        return 999

def format_money(v):
    try:
        n = float(v or 0)
        if n > 0:
            return f"{n:,.0f} ₽".replace(',', ' ')
    except:
        pass
    return '—'

def get_manager_stats(name, deals, period_days=None):
    now = datetime.now()
    
    if period_days:
        cutoff = now - timedelta(days=period_days)
        filtered = [d for d in deals if d.get('MANAGER','').lower() == name.lower() 
                   and days_since(d.get('DATE_CREATE','')) <= period_days]
    else:
        filtered = [d for d in deals if d.get('MANAGER','').lower() == name.lower()]
    
    won = [d for d in filtered if d.get('STAGE_ID') in WON_STAGES]
    lost = [d for d in filtered if d.get('STAGE_ID') in LOST_STAGES]
    active = [d for d in filtered if d.get('STAGE_ID') not in WON_STAGES + LOST_STAGES]
    missed = [d for d in active if d.get('STAGE_ID') in ['NEW','1'] and days_since(d.get('DATE_CREATE','')) >= 3]
    revenue = sum(float(d.get('OPPORTUNITY',0) or 0) for d in won)
    
    # По стадиям
    stages = {}
    for d in active:
        s = STAGE_NAMES.get(d.get('STAGE_ID',''), d.get('STAGE_ID',''))
        stages[s] = stages.get(s, 0) + 1
    
    return {
        'total': len(filtered),
        'won': len(won),
        'lost': len(lost),
        'active': len(active),
        'missed': len(missed),
        'revenue': revenue,
        'stages': stages
    }

def find_manager(name, users_dict):
    name_lower = name.lower()
    for uid, full_name in users_dict.items():
        if name_lower in full_name.lower():
            return full_name
    return None

@bot.message_handler(commands=['start', 'help'])
def handle_start(message):
    bot.reply_to(message, """👋 Привет! Я CRM бот AltusPro.

Что я умею:
📊 /manager Имя — сводка по менеджеру
📋 /all — сводка по всем менеджерам  
🔴 /missed — пропущенные заявки
💰 /revenue — выручка за месяц
⚠️ /attention — сделки требующие внимания

Или просто напиши имя менеджера, например:
"Александра" или "Федоткин"
""")

@bot.message_handler(commands=['all'])
def handle_all(message):
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    users = get_users()
    
    managers = {}
    for d in deals:
        m = d.get('MANAGER', 'Не назначен')
        if m not in managers:
            managers[m] = {'active': 0, 'won': 0, 'revenue': 0}
        if d.get('STAGE_ID') in WON_STAGES:
            managers[m]['won'] += 1
            managers[m]['revenue'] += float(d.get('OPPORTUNITY', 0) or 0)
        elif d.get('STAGE_ID') not in LOST_STAGES:
            managers[m]['active'] += 1
    
    text = "📊 *Сводка по всем менеджерам*\n\n"
    for name, s in sorted(managers.items(), key=lambda x: x[1]['revenue'], reverse=True):
        conv = round(s['won'] / max(s['won'] + s['active'], 1) * 100)
        text += f"👤 *{name}*\n"
        text += f"  Активных: {s['active']} | Успешных: {s['won']}\n"
        text += f"  Выручка: {format_money(s['revenue'])}\n"
        text += f"  Конверсия: {conv}%\n\n"
    
    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['missed'])
def handle_missed(message):
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    missed = [d for d in deals 
              if d.get('STAGE_ID') in ['NEW', '1'] 
              and days_since(d.get('DATE_CREATE','')) >= 3
              and d.get('STAGE_ID') not in WON_STAGES + LOST_STAGES]
    
    if not missed:
        bot.reply_to(message, "🎉 Пропущенных заявок нет!")
        return
    
    text = f"🔴 *Пропущенные заявки ({len(missed)})*\n\n"
    for d in missed[:20]:
        days = days_since(d.get('DATE_CREATE',''))
        emoji = "🔴" if days >= 7 else "🟡"
        text += f"{emoji} {d.get('TITLE','—')}\n"
        text += f"  👤 {d.get('MANAGER','—')} | {days} дн.\n\n"
    
    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['revenue'])
def handle_revenue(message):
    bot.reply_to(message, "⏳ Загружаю данные...")
    deals = get_deals()
    now = datetime.now()
    
    month_deals = [d for d in deals 
                   if d.get('STAGE_ID') in WON_STAGES
                   and days_since(d.get('DATE_CREATE','')) <= 30]
    
    total = sum(float(d.get('OPPORTUNITY',0) or 0) for d in month_deals)
    
    by_manager = {}
    for d in month_deals:
        m = d.get('MANAGER','—')
        by_manager[m] = by_manager.get(m, 0) + float(d.get('OPPORTUNITY',0) or 0)
    
    text = f"💰 *Выручка за последние 30 дней*\n"
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
    
    text = f"⚠️ *Сделки требующие внимания ({len(attention)})*\n_(висят 7+ дней)_\n\n"
    for d in attention[:15]:
        days = days_since(d.get('DATE_CREATE',''))
        stage = STAGE_NAMES.get(d.get('STAGE_ID',''), d.get('STAGE_ID',''))
        text += f"• {d.get('TITLE','—')}\n"
        text += f"  👤 {d.get('MANAGER','—')} | {stage} | {days} дн.\n\n"
    
    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(func=lambda m: True)
def handle_text(message):
    text = message.text.strip()
    users = get_users()
    manager_name = find_manager(text, users)
    
    if not manager_name:
        bot.reply_to(message, f"❓ Менеджер '{text}' не найден.\n\nДоступные команды:\n/all — все менеджеры\n/missed — пропущенные\n/revenue — выручка\n/attention — требуют внимания")
        return
    
    bot.reply_to(message, f"⏳ Загружаю данные по {manager_name}...")
    deals = get_deals()
    
    week = get_manager_stats(manager_name, deals, 7)
    month = get_manager_stats(manager_name, deals, 30)
    
    text = f"👤 *{manager_name}*\n\n"
    
    text += f"📅 *За неделю:*\n"
    text += f"  Сделок: {week['total']} | Успешных: {week['won']}\n"
    text += f"  Выручка: {format_money(week['revenue'])}\n\n"
    
    text += f"📆 *За месяц:*\n"
    text += f"  Сделок: {month['total']} | Успешных: {month['won']}\n"
    text += f"  Выручка: {format_money(month['revenue'])}\n"
    text += f"  Конверсия: {round(month['won']/max(month['total'],1)*100)}%\n\n"
    
    text += f"📋 *Сейчас в работе: {month['active']}*\n"
    if month['stages']:
        for stage, count in sorted(month['stages'].items(), key=lambda x: x[1], reverse=True)[:5]:
            text += f"  • {stage}: {count}\n"
    
    if month['missed'] > 0:
        text += f"\n🔴 *Пропущенных: {month['missed']}*\n"
    
    bot.reply_to(message, text, parse_mode='Markdown')

if __name__ == '__main__':
    print("Бот запущен!")
    bot.infinity_polling()
