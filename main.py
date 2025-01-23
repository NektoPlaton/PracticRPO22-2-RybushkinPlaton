import telebot
from telebot import types # - для кнопок
import webbrowser
import pandas as pd
from io import BytesIO
#import xlrd


bot = telebot.TeleBot("7656313726:AAGkZmmVN_f9PX0bjmt8lesOyWYToJE3uFM")

@bot.message_handler(commands=['start'])
def start(message): # - кнопки у самого пользователя под клавиатурой
    markup = types.ReplyKeyboardMarkup()

    siteBtn = types.KeyboardButton('Перейти в JournalTop')
    helpBtn = types.KeyboardButton('Помощь')
    markup.row(siteBtn, helpBtn)

    lessonsCountBtn = types.KeyboardButton('Подсчёт кол-ва проведенных пар для групп')
    lowMarkStudentsBtn = types.KeyboardButton('Ученики с низкими оценками')
    markup.row(lessonsCountBtn, lowMarkStudentsBtn)

    bot.send_message(message.chat.id, f'Здравствуйте, {message.from_user.first_name} '
                                      f'\nВы можете начать работу с ботом, нажмите на кнопку с нужной вам функцией', reply_markup=markup)

    bot.register_next_step_handler(message, click_button)

def click_button(message): # - обработка клавиатурных кнопок
    if message.text == 'Перейти в JournalTop':
        webbrowser.open('https://journal.top-academy.ru/ru/main/schedule/page/index')
    if message.text == 'Помощь':
        bot.send_message(message.chat.id, 'По вопросам бота обращаться к @NektoDetox')
    if message.text == 'Подсчёт кол-ва проведенных пар для групп':
        bot.send_message(message.chat.id, 'Отправьте xlsx файл(файл Excel):')

        bot.register_next_step_handler(message)
    if message.text == 'Ученики с низкими оценками':
        bot.send_message(message.chat.id, 'Отправьте xlsx файл(файл Excel):')

        bot.register_next_step_handler(message)

    bot.register_next_step_handler(message, click_button)



