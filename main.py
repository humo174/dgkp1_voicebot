import subprocess
import sys
import zipfile
import time
import datetime
from creds import oauth_token, catalog_id, bot_api, allow_users, admin_id


def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


checkImports = False
while not checkImports:
    try:
        import telebot
        import openpyxl
        from telebot import types
        from speechkit import Session, SpeechSynthesis
        import pandas as pd
        import xlsxwriter
        from thefuzz import fuzz as f
        checkImports = True
    except ImportError:
        install('pytelegrambotapi')
        install('speechkit')
        install('openpyxl')
        install('pandas')
        install('xlsxwriter')
        install('thefuzz')
        install('python-Levenshtein')
        install('lxml')

bot = telebot.TeleBot(bot_api)


def update_notice():
    try:
        with open('updatedone', 'r', encoding='utf-8') as done:
            done.close()
    except FileNotFoundError:
        with open('update.md', 'r', encoding='utf-8') as update_open:
            notice = update_open.read()
            for user in allow_users:
                bot.send_message(user, f'{notice}', parse_mode='html')

        with open('updatedone', 'w+', encoding='utf-8') as done:
            done.close()


def convert_file(chatid):
    @bot.message_handler(content_types=['document'])
    def lets_convert(message):
        if message.text != 'Отмена' and message.text is not None:
            kbc = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
            kbc1 = types.KeyboardButton(text='Отмена')
            kbc.add(kbc1)
            msg = bot.send_message(message.chat.id, f'Я жду файл, либо жми "Отмена"', reply_markup=kbc)
            bot.register_next_step_handler(msg, lets_convert)
        elif message.text == 'Отмена':
            kb = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
            kb1 = types.KeyboardButton(text='📢 Озвучить')
            kb2 = types.KeyboardButton(text='📄 Сформировать файл')
            kb.add(kb1, kb2)
            bot.send_message(message.chat.id, f'Действие отменено', reply_markup=kb)
        else:
            file_name = f'{message.chat.id}.xls'
            file_info = bot.get_file(message.document.file_id)
            print(file_info.file_path[-4::])
            if file_info.file_path[-4::] == '.xls':
                downloaded_file = bot.download_file(file_info.file_path)
                with open(file_name, 'wb') as new_file:
                    new_file.write(downloaded_file)
                bot.send_message(message.chat.id, f'Обрабатываю файл')

                try:
                    df = pd.read_html(f'{message.chat.id}.xls')
                    df5 = pd.DataFrame()
                    df6 = pd.DataFrame()
                    for j in range(len(df)):
                        df2 = pd.DataFrame(df[j])
                        df4 = pd.DataFrame(columns=['ФИО'])
                        df2 = df2.drop(columns=[0, 2, 3, 4, 5, 6, 8, 9, 10], axis=1)
                        df2 = df2.rename(columns={1: "ФИО", 7: "Номер"})
                        df3 = pd.DataFrame(columns=['ФИО', "Номер"])
                        df2 = df2[(df2.ФИО != "ФИО пациента")]
                        df4 = df2[df2['Номер'].isnull()].reset_index()
                        df2 = df2.dropna().reset_index()

                        for i in range(len(df2)):
                            if f.WRatio('Врач: ', df2['ФИО'][i]) < 70:
                                for c in range(0, (df2['Номер'][i].count(',') + 1)):
                                    if df2['Номер'][i].count(',') >= 1:
                                        df3 = df3.append({'ФИО': df2['ФИО'][i], 'Номер': (
                                            ((((df2['Номер'][i].split(",")[c]).replace("(", "")).replace(
                                                ")", "")).replace("-", "")).replace(" ", "")[2::])}, ignore_index=True)
                                    else:
                                        df3 = df3.append({'ФИО': df2['ФИО'][i], 'Номер': (
                                            ((((df2['Номер'][i].split(",")[0]).replace("(", "")).replace(
                                                ")", "")).replace("-", "")).replace(" ", "")[2::])}, ignore_index=True)

                        df5 = df5.append(df3, ignore_index=True)
                        df6 = df6.append(df4, ignore_index=True)

                    workbook = xlsxwriter.Workbook(f'{message.chat.id}-converted.xlsx')
                    workbook1 = xlsxwriter.Workbook(f'{message.chat.id}-declined.xlsx')
                    worksheet = workbook.add_worksheet('Лист1')
                    worksheet1 = workbook1.add_worksheet('Лист1')
                    worksheet.write('A1', 'Номер')
                    worksheet.write('B1', 'Комментарий')
                    worksheet1.write('A1', 'ФИО')
                    for i in range(len(df6)):
                        worksheet1.write('A' + str(i + 2), str(df6['ФИО'][i]))
                    for i in range(len(df5)):
                        worksheet.write('A' + str(i + 2), str(df5['Номер'][i]))
                        worksheet.write('B' + str(i + 2), str(df5['ФИО'][i]))
                    workbook.close()
                    workbook1.close()

                    time.sleep(4)
                    kb = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                    kb1 = types.KeyboardButton(text='📢 Озвучить')
                    kb2 = types.KeyboardButton(text='📄 Сформировать файл')
                    kb.add(kb1, kb2)
                    with open(f'{message.chat.id}-converted.xlsx', 'rb') as converted_file:
                        bot.send_document(message.chat.id,
                                          caption='Готовый файл для автообзвона\n'
                                                  'Просто импортируй его в ВАТС "Ростелеком"',
                                          document=converted_file, parse_mode='html',
                                          visible_file_name=f'Автообзвон {datetime.datetime.now().strftime("%d-%m-%Y")}'
                                                            f'.xlsx')

                    with open(f'{message.chat.id}-declined.xlsx', 'rb') as declined_file:
                        bot.send_document(message.chat.id, caption="Отклоненные пациенты без номера телефона",
                                          document=declined_file,
                                          visible_file_name=f'Отклонено {datetime.datetime.now().strftime("%d-%m-%Y")}'
                                                            f'.xlsx', reply_markup=kb)

                except zipfile.BadZipfile:
                    time.sleep(2)
                    conv_file = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                    conv_file_1 = types.KeyboardButton(text='Отмена')
                    conv_file.add(conv_file_1)
                    msg = bot.send_message(message.chat.id, f'Ошибка при обработке файла\nЗагруженный файл не '
                                                            f'соответствует '
                                                            f'скелету.\n\nОжидаю правильный файл',
                                           reply_markup=conv_file)
                    bot.register_next_step_handler(msg, lets_convert)

            else:
                conv_file = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                conv_file_1 = types.KeyboardButton(text='Отмена')
                conv_file.add(conv_file_1)
                msg = bot.send_message(message.chat.id, f'Ошибка при обработке файла\nЗагруженный файл не '
                                                        f'соответствует '
                                                        f'скелету.\n\nОжидаю правильный файл',
                                       reply_markup=conv_file)
                bot.register_next_step_handler(msg, lets_convert)

    conv_file = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    conv_file_1 = types.KeyboardButton(text='Отмена')
    conv_file.add(conv_file_1)
    msg = bot.send_message(chatid, f'1) Сформируй файл в РМИС "БАРС" в меню: Отчеты → Регистратура → График записей.\n'
                                   f'2) Отправь его мне.', reply_markup=conv_file)
    bot.register_next_step_handler(msg, lets_convert)


def lets_rock():
    @bot.message_handler(content_types=['text'])
    def message_main(message):
        def ozvuch(message):
            if message.text == 'Отмена':
                kb = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                kb1 = types.KeyboardButton(text='📢 Озвучить')
                kb2 = types.KeyboardButton(text='📄 Сформировать файл')
                kb.add(kb1, kb2)
                bot.send_message(message.chat.id, f'Запрос отменен', reply_markup=kb)

            else:
                text_to_sound = message.text
                kbcl = types.ReplyKeyboardRemove()
                try:
                    bot.send_message(message.chat.id, f'Дождитесь формирования аудиофайла', reply_markup=kbcl)
                    session = Session.from_yandex_passport_oauth_token(oauth_token, catalog_id)
                    synthesizeaudio = SpeechSynthesis(session)
                    synthesizeaudio.synthesize(
                        str(f'{message.chat.id}-out.wav'), text=f'{text_to_sound}',
                        voice='oksana', sampleRateHertz='16000'
                    )
                    kb = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                    kb1 = types.KeyboardButton(text='📢 Озвучить')
                    kb2 = types.KeyboardButton(text='📄 Сформировать файл')
                    kb.add(kb1, kb2)
                    audio = open(f'{message.chat.id}-out.wav', 'rb')
                    time.sleep(4)
                    bot.send_message(message.chat.id, f'Аудиофайл сформирован успешно', reply_markup=kb)
                    bot.send_audio(message.chat.id, audio)
                    audio.close()
                except Exception as audiofail:
                    error_audio = open(r'error_audio.log', 'a+')
                    error_audio.write(f'{datetime.datetime.now()} | AudioFail: {audiofail}\n\n\n')
                    error_audio.close()
                    bot.send_message(message.chat.id, f'Системная ошибка. Запись сохранена в лог. Обратитесь к '
                                                      f'разработчику')

        if message.text == '📢 Озвучить' and message.chat.id in allow_users:
            kbc = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
            kbc1 = types.KeyboardButton(text='Отмена')
            kbc.add(kbc1)
            msg = bot.send_message(message.chat.id, f'Введите текст для озвучки', reply_markup=kbc)
            bot.register_next_step_handler(msg, ozvuch)

        if message.text == '📄 Сформировать файл':
            convert_file(message.chat.id)


def mainbody():
    @bot.message_handler(commands=['start'])
    def start_command(message):
        if message.chat.id in allow_users:
            kb = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
            kb1 = types.KeyboardButton(text='📢 Озвучить')
            kb2 = types.KeyboardButton(text='📄 Сформировать файл')
            kb.add(kb1, kb2)
            bot.send_message(message.chat.id, f'Давай начнем работу!', reply_markup=kb)
            lets_rock()
        else:
            bot.send_message(message.chat.id, f'У вас нет прав на пользование данным ботом, '
                                              f'перешлите это сообщение администратору\n\n'
                                              f'<code>{message.chat.id}</code>',
                             parse_mode='html')

    @bot.message_handler(commands=['add'])
    def add_command(message):
        if message.chat.id == admin_id:
            if message.text == '/add':
                pass
            elif message.text != '/add' and message.text.split(' ')[0] == '/add':
                newallow = []
                try:
                    with open('creds.py', 'r+', encoding='utf-8') as fo:
                        for line in fo.readlines():
                            if line.find('allow_users =') == -1:
                                newallow.append(line)
                            if line.find('allow_users =') != -1:
                                oldstring = f'{line[:-1:]}'
                                newallow.append(oldstring.replace(")", f", {int(message.text.split(' ')[1])})\n"))
                                bot.send_message(message.chat.id, f'Добавлен пользователь с ID = '
                                                                  f'{int(message.text.split(" ")[1])}')

                    with open('creds.py', 'r+', encoding='utf-8') as fo:
                        for line in newallow:
                            fo.write(line)
                except ValueError:
                    bot.send_message(message.chat.id, f'Неверный формат ID. Должен быть целым числом')
        else:
            pass

    bot.polling(non_stop=True)


update_notice()

nkb = types.ReplyKeyboardRemove()
bot.send_message(admin_id, f'Бот перезапущен', reply_markup=nkb)

while 1 == 1:
    try:
        mainbody()
    except Exception as exc:
        f = open(r'error_connection.log', 'a+')
        f.write(f'{datetime.datetime.now()} | ErrorConnection: {exc}\n')
        f.close()
