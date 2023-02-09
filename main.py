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

        checkImports = True
    except ImportError:
        install('pytelegrambotapi')
        install('speechkit')
        install('openpyxl')
        install('pandas')
        install('xlsxwriter')

bot = telebot.TeleBot(bot_api)


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
            file_name = f'{message.chat.id}.xlsx'
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)
            with open(file_name, 'wb') as new_file:
                new_file.write(downloaded_file)
            bot.send_message(message.chat.id, f'Обрабатываю файл')

            try:
                wb = openpyxl.load_workbook(f'{message.chat.id}.xlsx')
                sheet = wb.active
                rows = sheet.max_row

                wb_convert = openpyxl.Workbook()
                wb_convert.create_sheet(title='Лист1', index=0)
                del wb_convert['Sheet']
                sheet_convert = wb_convert['Лист1']
                sheet_convert['A1'] = 'Номер'
                sheet_convert['B1'] = 'Комментарий'

                wb_decline = openpyxl.Workbook()
                wb_decline.create_sheet(title='Лист1', index=0)
                del wb_decline['Sheet']
                sheet_decline = wb_decline['Лист1']
                sheet_decline['A1'] = 'fio'

                for i in range(1, rows + 1):
                    fio = sheet.cell(row=i, column=2)
                    phone = sheet.cell(row=i, column=8)
                    if fio.value is not None and fio.value != 'ФИО пациента' and phone.value is not None:
                        a = (
                            f'{((((phone.value.split(",")[0]).replace("(", "")).replace(")", "")).replace("-", "")).replace(" ", "")[2::]}',
                            f'{fio.value}')
                        sheet_convert.append(a)
                    if fio.value is not None and fio.value != 'ФИО пациента' and phone.value is None:
                        b = (f'{fio.value}',)
                        sheet_decline.append(b)

                wb_decline.save(f'{message.chat.id}-declined.xlsx')
                wb_convert.save(f'{message.chat.id}-converted.xlsx')

                df = pd.read_excel(f'{message.chat.id}-converted.xlsx', sheet_name='Лист1')
                workbook = xlsxwriter.Workbook(f'{message.chat.id}-converted.xlsx')
                worksheet = workbook.add_worksheet('Лист1')
                worksheet.write('A1', 'Номер')
                worksheet.write('B1', 'Комментарий')
                for i in range(len(df)):
                    worksheet.write('A' + str(i + 2), str(df['Номер'][i]))
                    worksheet.write('B' + str(i + 2), str(df['Комментарий'][i]))

                workbook.close()
                time.sleep(4)
                kb = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                kb1 = types.KeyboardButton(text='📢 Озвучить')
                kb2 = types.KeyboardButton(text='📄 Сформировать файл')
                kb.add(kb1, kb2)
                with open(f'{message.chat.id}-converted.xlsx', 'rb') as converted_file:
                    bot.send_document(message.chat.id,
                                      caption='Готовый файл для автообзвона\n'
                                              'Необходимо открыть и просто <b>СОХРАНИТЬ</b> файл еще раз, '
                                              'чтобы он загрузился в ВАТС "Ростелеком"',
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
                                                        f'скелету. Убедитесь, что сформированный файл пересохранен как '
                                                        f'"Книга Excel (*.xlsx)"\n\nОжидаю правильный файл',
                                       reply_markup=conv_file)
                bot.register_next_step_handler(msg, lets_convert)

    conv_file = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    conv_file_1 = types.KeyboardButton(text='Отмена')
    conv_file.add(conv_file_1)
    msg = bot.send_message(chatid, f'1. Сформируй файл в РМИС "БАРС" в меню: Отчеты → Регистратура → График записей\n'
                                   f'2. Пересохрани скачанный файл в правильный формат .xlsx '
                                   f'(Открыть файл → Сохранить как...)\n'
                                   f'3. Пришли мне правильно сохраненный файл', reply_markup=conv_file)
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

    bot.polling(non_stop=True)


bot.send_message(admin_id, f'Бот перезапущен')
while 1 == 1:
    try:
        mainbody()
    except Exception as exc:
        f = open(r'error_connection.log', 'a+')
        f.write(f'{datetime.datetime.now()} | ErrorConnection: {exc}\n')
        f.close()
