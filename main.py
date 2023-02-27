import json
import subprocess
import sys
import time
import datetime


def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


checkImports = False
while not checkImports:
    try:
        import telebot
        import openpyxl
        import paramiko
        from telebot import types
        from speechkit import Session, SpeechSynthesis
        import pandas as pd
        import xlsxwriter

        checkImports = True
    except ImportError:
        install('pytelegrambotapi')
        install('speechkit')
        install('openpyxl')
        install('paramiko')
        install('pandas')
        install('xlsxwriter')
        install('lxml')


def creds():
    with open(r'creds.json', 'r', encoding='utf-8') as read:
        return json.load(read)


def update_notice():
    try:
        with open('updatedone', 'r', encoding='utf-8') as done:
            done.close()
    except FileNotFoundError:
        with open('update.md', 'r', encoding='utf-8') as update_open:
            notice = update_open.read()
            for user in creds()['allow_users']:
                bot.send_message(user, f'{notice}', parse_mode='html')

        with open('updatedone', 'w+', encoding='utf-8') as done:
            done.close()


def add_allow_user(message):
    with open(r'creds.json', 'r', encoding='utf-8') as read:
        data = json.load(read)

    try:
        bool(data['allow_users'][f'{int(message.text.split(" ")[1])}'])
        bot.send_message(message.chat.id, f'Пользователь с ID = {message.text.split(" ")[1]} уже существует')

    except KeyError:
        data['allow_users'][f'{int(message.text.split(" ")[1])}'] = message.text.split(" ")[2]

        with open(r'creds.json', 'w', encoding='utf-8') as write:
            json.dump(data, write)

        bot.send_message(message.chat.id, f'Добавлен пользователь с ID = {message.text.split(" ")[1]}, '
                                          f'{message.text.split(" ")[2]}')

    except ValueError:
        bot.send_message(message.chat.id, f'Неверный формат ID')

    except IndexError:
        bot.send_message(message.chat.id, f'Index Error')


def show_allow_users(message):
    with open(r'creds.json', 'r', encoding='utf-8') as read:
        data = json.load(read)

    users = ''
    for i in data['allow_users']:
        users += f'<code>{i}</code>, {data["allow_users"].get(i)}\n'

    bot.send_message(message.chat.id, users, parse_mode='html')


def del_allow_user(message):
    with open(r'creds.json', 'r', encoding='utf-8') as read:
        data = json.load(read)

    try:
        del data['allow_users'][f'{int(message.text.split(" ")[1])}']

        with open(r'creds.json', 'w', encoding='utf-8') as write:
            json.dump(data, write)

        bot.send_message(message.chat.id, f'Удален пользователь с ID = {int(message.text.split(" ")[1])}')
    except (ValueError, KeyError):
        bot.send_message(message.chat.id, f'Пользователь с таким ID не найден')


def reboot_bot(message):
    try:
        host = f'{creds()["host"]["host"]}'
        user = f'{creds()["host"]["user"]}'
        secret = f'{creds()["host"]["secret"]}'
        port = 22

        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(hostname=host, username=user, password=secret, port=port)
        stdin, stdout, stderr = client.exec_command(
            'sudo docker compose -f /opt/dgkp1_voicebot/docker-compose.yaml restart')
        data = stdout.read() + stderr.read()
        client.close()
    except Exception as exc:
        bot.send_message(message.chat.id, f'Ошибка при перезагрузке бота')
        f = open(r'ssh_connection.log', 'a+')
        f.write(f'{datetime.datetime.now()} | ErrorSSH: {exc}\n')
        f.close()


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
            if file_info.file_path[-4::] == '.xls':
                downloaded_file = bot.download_file(file_info.file_path)
                with open(file_name, 'wb') as new_file:
                    new_file.write(downloaded_file)

                try:
                    df = pd.read_html(f'{message.chat.id}.xls')
                    bot.send_message(message.chat.id, f'Обрабатываю файл')
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
                        df2 = df2.drop(labels=[0], axis=0)
                        df2 = df2.dropna().reset_index()
                        for i in range(len(df2)):
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

                    time.sleep(2)
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

                except (ImportError, ValueError):
                    conv_file = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                    conv_file_1 = types.KeyboardButton(text='Отмена')
                    conv_file.add(conv_file_1)
                    msg = bot.send_message(message.chat.id, f'Ошибка при обработке файла\nЗагруженный файл '
                                                            f'не соответствует структуре. Убедитесь, что загружаемый '
                                                            f'файл является исходным из РМИС "БАРС"\n\n'
                                                            f'Ожидаю правильный файл',
                                           reply_markup=conv_file)
                    bot.register_next_step_handler(msg, lets_convert)

            else:
                conv_file = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                conv_file_1 = types.KeyboardButton(text='Отмена')
                conv_file.add(conv_file_1)
                msg = bot.send_message(message.chat.id, f'Ошибка при обработке файла\nЗагруженный файл не '
                                                        f'является файлом .xls\n\nОжидаю правильный файл',
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
                    session = Session.from_yandex_passport_oauth_token(creds()['oauth_token'],
                                                                       creds()['catalog_id'])
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
                    error_audio.write(f'{datetime.datetime.now()} | AudioFail: {audiofail}\n')
                    error_audio.close()
                    bot.send_message(message.chat.id, f'Системная ошибка. Запись сохранена в лог. Обратитесь к '
                                                      f'администратору')

        if message.text == '📢 Озвучить':
            try:
                bool(creds()['allow_users'][f'{message.chat.id}'])
                kbc = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                kbc1 = types.KeyboardButton(text='Отмена')
                kbc.add(kbc1)
                msg = bot.send_message(message.chat.id, f'Введите текст для озвучки', reply_markup=kbc)
                bot.register_next_step_handler(msg, ozvuch)
            except KeyError:
                bot.send_message(message.chat.id, f'У вас нет прав на пользование данным ботом, '
                                                  f'перешлите это сообщение администратору\n\n'
                                                  f'<code>{message.chat.id}</code>',
                                 parse_mode='html')

        if message.text == '📄 Сформировать файл':
            try:
                bool(creds()['allow_users'][f'{message.chat.id}'])
                convert_file(message.chat.id)
            except KeyError:
                bot.send_message(message.chat.id, f'У вас нет прав на пользование данным ботом, '
                                                  f'перешлите это сообщение администратору\n\n'
                                                  f'<code>{message.chat.id}</code>',
                                 parse_mode='html')


bot = telebot.TeleBot(creds()['bot_api'])


def mainbody():
    @bot.message_handler(commands=['start'])
    def start_command(message):
        try:
            bool(creds()['allow_users'][f'{message.chat.id}'])
            kb = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
            kb1 = types.KeyboardButton(text='📢 Озвучить')
            kb2 = types.KeyboardButton(text='📄 Сформировать файл')
            kb.add(kb1, kb2)
            bot.send_message(message.chat.id, f'Давай начнем работу!', reply_markup=kb)
            lets_rock()
        except KeyError:
            bot.send_message(message.chat.id, f'У вас нет прав на пользование данным ботом, '
                                              f'перешлите это сообщение администратору\n\n'
                                              f'<code>{message.chat.id}</code>',
                             parse_mode='html')

    @bot.message_handler(commands=['add'])
    def add_command(message):
        if message.chat.id == creds()['admin_id']:
            if message.text == '/add':
                pass
            elif message.text != '/add' and message.text.split(' ')[0] == '/add' and message.text.split(' ')[2]:
                add_allow_user(message)

        else:
            pass

    @bot.message_handler(commands=['del'])
    def del_command(message):
        if message.chat.id == creds()['admin_id']:
            if message.text == '/del':
                pass
            elif message.text != '/del' and message.text.split(' ')[0] == '/del':
                del_allow_user(message)

        else:
            pass

    @bot.message_handler(commands=['sau'])
    def sau_command(message):
        if message.chat.id == creds()['admin_id']:
            show_allow_users(message)

    @bot.message_handler(commands=['reboot'])
    def reboot_command(message):
        if message.chat.id == creds()['admin_id']:
            reboot_bot(message)

    bot.polling(non_stop=True)


update_notice()

rkb = types.ReplyKeyboardRemove()
bot.send_message(creds()['admin_id'], f'Бот перезапущен', reply_markup=rkb)

# while 1 == 1:
# try:
mainbody()
# except Exception as exc:
#     f = open(r'error_connection.log', 'a+')
#     f.write(f'{datetime.datetime.now()} | ErrorConnection: {exc}\n')
#     f.close()
