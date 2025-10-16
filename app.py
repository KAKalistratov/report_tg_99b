import os
import re
from dotenv import load_dotenv
import telebot
import gspread
from oauth2client.service_account import ServiceAccountCredentials

load_dotenv()
TG_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
SHEET_ID = os.getenv("GOOGLE_SHEETS_ID")

# Названия листов
SHEET_VK_TG = "ВК + ТГ"
SHEET_AIRAT = "Отчет для Айрата ВК + ТГ"

# Авторизация в Google Sheets
scope = ["https://www.googleapis.com/auth/spreadsheets.readonly",
         "https://www.googleapis.com/auth/drive.readonly"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
gc = gspread.authorize(creds)
bot = telebot.TeleBot(TG_TOKEN, parse_mode="Markdown")

def fmt_money(val):
    if not val:
        return "0"
    return str(val).replace("₽", "").replace(" ", "").replace("\xa0", "").strip()

def fmt_percent(val):
    if not val:
        return "0%"
    val = val.replace(",", ".").replace("%", "").strip()
    try:
        f = float(val)
        return f"{f:.2f}%".replace(".", ",")
    except:
        return val

def report_vk_tg():
    ws = gc.open_by_key(SHEET_ID).worksheet(SHEET_VK_TG)
    vals = ws.get_all_values()

    # достаём дату
    date = vals[0][1] if len(vals) > 0 and len(vals[0]) > 1 else "?"

    # маленький помощник: найти значение по ключу в колонке А
    def get(a):
        for row in vals:
            if len(row) >= 2 and a.strip().lower() in row[0].strip().lower():
                return row[1], row[2] if len(row) > 2 else ""
        return "", ""

    parts = [f"*ОТЧЕТ {date}*\n"]

    # --- ДП ВК ---
    parts.append("*ДП ВК*")
    parts.append(f"План выручки на новых: {fmt_money(get('План выручки на новых')[0])}")
    parts.append(f"Сумма на новых за вчера: {fmt_money(get('Сумма на новых за вчера')[0])}")
    parts.append(f"Сумма на новых за месяц: {fmt_money(get('Сумма на новых за месяц')[0])} ({fmt_percent(get('Сумма на новых за месяц')[1])})")
    parts.append(f"Дельта с прошлым месяцем: {fmt_percent(get('Дельта с прошлым месяцем')[1])}\n")

    parts.append(f"План выручки на продлениях: {fmt_money(get('План выручки на продлениях')[0])}")
    parts.append(f"Сумма на продлениях за вчера: {fmt_money(get('Сумма на продлениях за вчера')[0])}")
    parts.append(f"Сумма на продлениях за месяц: {fmt_money(get('Сумма на продлениях за месяц')[0])} ({fmt_percent(get('Сумма на продлениях за месяц')[1])})")
    parts.append(f"Дельта с прошлым месяцем: {fmt_percent(get('Дельта с прошлым месяцем')[1])}\n")

    # --- ТГ ---
    parts.append("*ТГ*")
    # просто читаем вторые блоки по порядку, потому что на листе они идут ниже
    parts.append(f"План выручки на новых: {fmt_money(vals[15][1])}")
    parts.append(f"Сумма на новых за вчера: {fmt_money(vals[16][1])}")
    parts.append(f"Сумма на новых за месяц: {fmt_money(vals[17][1])} ({fmt_percent(vals[17][2])})")
    parts.append(f"Дельта с прошлым месяцем: {fmt_percent(vals[18][1])}\n")

    parts.append(f"План выручки на продлениях: {fmt_money(vals[20][1])}")
    parts.append(f"Сумма на продлениях за вчера: {fmt_money(vals[21][1])}")
    parts.append(f"Сумма на продлениях за месяц: {fmt_money(vals[22][1])} ({fmt_percent(vals[22][2])})")
    parts.append(f"Дельта с прошлым месяцем: {fmt_percent(vals[23][1])}\n")

    # --- ИТОГ ---
    parts.append("*ИТОГ*")
    parts.append(f"Итого план выручки на новых: {fmt_money(vals[26][1])}")
    parts.append(f"Итого факт выручки на новых: {fmt_money(vals[27][1])} ({fmt_percent(vals[27][2])})")
    parts.append(f"Итого план выручки на продлениях: {fmt_money(vals[28][1])}")
    parts.append(f"Итого факт выручки на продлениях: {fmt_money(vals[29][1])} ({fmt_percent(vals[29][2])})")

    return "\n".join(parts)


def report_airat():
    ws = gc.open_by_key(SHEET_ID).worksheet(SHEET_AIRAT)
    vals = ws.get_all_values()
    date = vals[0][1] if len(vals) > 0 and len(vals[0]) > 1 else "?"

    def get(a):
        for row in vals:
            if len(row) >= 2 and a.strip().lower() in row[0].strip().lower():
                return row[1]
        return ""

    parts = [f"*ОТЧЕТ ДП ЗА {date}*\n"]
    parts.append(f"Прогноз выручки на сегодня: {fmt_money(get('Прогноз выручки на сегодня'))}")
    parts.append(f"Выручка на новых за вчера: {fmt_money(get('Выручка на новых за вчера'))}")
    parts.append(f"Итого выручки на новых: {fmt_money(get('Итого выручки на новых'))}")
    parts.append(f"Прогнозы выручки на сентябрь: {fmt_money(get('Прогнозы выручки на сентябрь'))}")
    parts.append(f"План выручки на новых: {fmt_money(get('План выручки на новых'))}\n")

    parts.append(f"Получено Ц лидов за сентябрь: {fmt_money(get('Получено Ц лидов за сентябрь'))}")
    parts.append(f"Получено Ц лидов за вчера: {fmt_money(get('Получено Ц лидов за вчера'))}")
    parts.append(f"Конверсия по Ц лидам за сентябрь: {fmt_percent(get('Конверсия по Ц лидам за сентябрь'))}")
    parts.append(f"Конверсия по Ц лидам за вчера: {fmt_percent(get('Конверсия по Ц лидам за вчера'))}\n")

    parts.append(f"Плановая численность на вчера: {fmt_money(get('Плановая численность на вчера'))}")
    parts.append(f"Фактическая численность на вчера: {fmt_money(get('Фактическая численность на вчера'))}")

    return "\n".join(parts)


@bot.message_handler(commands=['get_date'])
def send_reports(message):
    try:
        vk_tg_report = report_vk_tg()
        airat_report = report_airat()

        bot.send_message(message.chat.id, vk_tg_report)
        bot.send_message(message.chat.id, airat_report)
    except Exception as e:
        bot.send_message(message.chat.id, f"Ошибка: {e}")


@bot.message_handler(commands=['start', 'help'])
def help_cmd(message):
    bot.reply_to(message, "Команда: /get_date — присылает два отчета: ВК+ТГ и Отчет для Айрата ВК+ТГ.")


if __name__ == "__main__":
    bot.infinity_polling(skip_pending=True)