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
    'EXECUTING': 'В работе - счёт
