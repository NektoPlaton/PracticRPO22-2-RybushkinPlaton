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

        bot.register_next_step_handler(message, count_lessons)
    if message.text == 'Ученики с низкими оценками':
        bot.send_message(message.chat.id, 'Отправьте xlsx файл(файл Excel):')

        bot.register_next_step_handler(message, find_students_with_low_scores)

    bot.register_next_step_handler(message, click_button)

def count_lessons(message):
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    df = pd.read_excel(BytesIO(downloaded_file))

    first_column_value = df.iloc[1, 0]

    subject_count = {}

    for column in df.columns[1:]:
        for value in df[column]:
            if isinstance(value, str):
                if "предмет:" in value.lower():

                    subject = value.lower().split("предмет:")[1].split("\n")[
                        0].strip()
                    if subject in subject_count:
                        subject_count[subject] += 1
                    else:
                        subject_count[subject] = 1

    result_message = "Количество предметов:\n"
    for subject, count in subject_count.items():
        result_message += f'Предмет: {subject}, Количество: {count}\n'

    bot.send_message(message.chat.id, f'Количество пар для группы {first_column_value}: \n{result_message}')


    #bot.register_next_step_handler(message, lambda m: count_lessons(m, df))
    del downloaded_file
    del df

