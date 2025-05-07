import os
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes, CallbackQueryHandler
import requests
from io import BytesIO
from datetime import datetime, timedelta

# Глобальная переменная для хранения данных
data_cache = None
last_update_time = None


def get_updated_data():
    global data_cache, last_update_time

    # Обновляем данные если прошло больше 15 минут или данные еще не загружены
    if data_cache is None or (datetime.now() - last_update_time) > timedelta(minutes=15):
        try:
            file_url = 'https://docs.google.com/spreadsheets/d/1zBAff5u-tirVwRNt7KoclhdkOaQ_S_p8/export?format=xlsx'
            response = requests.get(file_url)
            response.raise_for_status()  # Проверяем успешность запроса

            excel_data = BytesIO(response.content)
            data_cache = pd.read_excel(excel_data,
                                       sheet_name=1,
                                       engine="openpyxl")
            last_update_time = datetime.now()

        except Exception as e:
            print(f"Ошибка при обновлении данных: {e}")
            if data_cache is None:
                raise
    return data_cache


def get_item_status(item_number):
    try:
        data = get_updated_data()
        # Более надежный поиск по первому столбцу
        mask = data.iloc[:, 0].astype(str).str.contains(f"{item_number}-цб",
                                                        na=False)

        if not mask.any():
            return 'Товар не найден'

        item_status = data.loc[mask, data.columns[1]].values[0]

        if item_status == "отгружено":
            item_date = data.loc[mask, data.columns[4]].values[0]
            formatted_date = item_date.strftime("%d.%m.%Y")
            return f"Продукция по счету {item_number} отгружена: {formatted_date}"
        elif item_status == "готово":
            return (
                f"Готово:\nПродукция по счету {item_number} готова.\n"
                "Для оформления документов подъезжайте в офис:\n"
                "ул. Б.Хмельницкого, 128 – с доверенностью или печатью организации.\n"
                "Затем на склад: ул. 10 лет Октября, 219 к2А")
        elif item_status == "производство":
            return "Производство:\nЗаказ передан в производство. По готовности сообщим."
        elif item_status == "оплачено":
            return f"Оплачено:\nОплата по счету {item_number} поступила. Заказ передан в производство."
        else:
            return f"Неизвестный статус: {item_status}"

    except Exception as e:
        print(f"Ошибка в get_item_status: {str(e)}")
        return f"Произошла ошибка при обработке запроса: {str(e)}"


def get_menu_keyboard():
    keyboard = [
        [
            InlineKeyboardButton("Статус заказа",
                                 callback_data='enter_item_number')
        ],
        [
            InlineKeyboardButton("Контактная информация",
                                 callback_data='contact_info')
        ],
    ]
    return InlineKeyboardMarkup(keyboard)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        if update.message:
            await update.message.reply_text('Выберите действие:',
                                            reply_markup=get_menu_keyboard())
        else:
            query = update.callback_query
            await query.answer()
            await query.edit_message_text(text='Выберите действие:',
                                          reply_markup=get_menu_keyboard())
    except Exception as e:
        print(f"Ошибка в start: {e}")


async def button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        query = update.callback_query
        await query.answer()

        if query.data == 'enter_item_number':
            await query.edit_message_text(text="Введите номер счета :")
        elif query.data == 'contact_info':
            keyboard = [[
                InlineKeyboardButton("Назад", callback_data='back_to_menu')
            ]]
            reply_markup = InlineKeyboardMarkup(keyboard)
            await query.edit_message_text(text=(
                "Наша контактная информация:\nАдрес главного офиса:\n"
                "644001 г. Омск, ул. Б. Хмельницкого, 128\nТелефоны для связи:\n"
                "+7 (3812) 36-39-43\n+7 (3812) 36-40-88\n+7 (3812) 36-41-07\n"
                "Email: tender@lakcolor.ru"),
                reply_markup=reply_markup)
        elif query.data == 'back_to_menu':
            await start(update, context)
    except Exception as e:
        print(f"Ошибка в button: {e}")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        item_number = update.message.text.strip()

        # Проверяем, что введены только цифры
        if not item_number.isdigit():
            await update.message.reply_text(
                "Пожалуйста, введите только цифры номера счета.")
            return

        status = get_item_status(item_number)

        keyboard = [[
            InlineKeyboardButton("Вернуться в меню",
                                 callback_data='back_to_menu')
        ]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text(status, reply_markup=reply_markup)

    except Exception as e:
        print(f"Ошибка в handle_message: {e}")
        await update.message.reply_text(
            "Произошла ошибка при обработке запроса. Пожалуйста, попробуйте позже."
        )


async def force_update(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global data_cache, last_update_time
    try:
        data_cache = None
        get_updated_data()
        await update.message.reply_text("Данные успешно обновлены!")
    except Exception as e:
        print(f"Ошибка в force_update: {e}")
        await update.message.reply_text(f"Ошибка при обновлении данных: {e}")


def main():
    try:
        # Инициализация данных при старте
        get_updated_data()

        application = ApplicationBuilder().token(
            "BOT_TOKEN").build()

        application.add_handler(CommandHandler('start', start))
        application.add_handler(CommandHandler('update', force_update))
        application.add_handler(CallbackQueryHandler(button))
        application.add_handler(
            MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

        print("Бот запущен...")
        application.run_polling()

    except Exception as e:
        print(f"Ошибка в main: {e}")


if __name__ == '__main__':
    main()
