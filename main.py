from speechkit import Session, SpeechSynthesis
import telebot
from telebot import types
from creds import oauth_token, catalog_id, bot_api

allow_users = (1134241949, 571876631)
bot = telebot.TeleBot(bot_api)

try:
    while 1 == 1:
        @bot.message_handler(commands=['start'])
        def start_command(message):
            if message.chat.id in allow_users:
                kb = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                kb1 = types.KeyboardButton(text='Озвучить')
                kb.add(kb1)
                bot.send_message(message.chat.id, f'Давай начнем работу!', reply_markup=kb)
            else:
                bot.send_message(message.chat.id, f'У вас нет прав на пользование данным ботом')


        @bot.message_handler(content_types=['text'])
        def message_main(message):
            def ozvuch(mes_to_sound):
                if mes_to_sound.text == 'Отмена':
                    kb = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                    kb1 = types.KeyboardButton(text='Озвучить')
                    kb.add(kb1)
                    bot.send_message(message.chat.id, f'Запрос отменен', reply_markup=kb)

                else:
                    text_to_sound = mes_to_sound.text
                    session = Session.from_yandex_passport_oauth_token(oauth_token, catalog_id)
                    synthesizeaudio = SpeechSynthesis(session)
                    synthesizeaudio.synthesize(
                        str(f'out.wav'), text=f'{text_to_sound}',
                        voice='oksana', sampleRateHertz='16000'
                    )
                    bot.send_message(message.chat.id, f'Вот готовый файл')
                    audio = open(r'./out.wav', 'rb')
                    bot.send_audio(message.chat.id, audio)
                    audio.close()

            if message.text == 'Озвучить' and message.chat.id in allow_users:
                kbc = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
                kbc1 = types.KeyboardButton(text='Отмена')
                kbc.add(kbc1)
                msg = bot.send_message(message.chat.id, f'Введите текст для озвучки', reply_markup=kbc)
                bot.register_next_step_handler(msg, ozvuch)


        bot.polling(none_stop=True)
except Exception as exc:
    print(f'Rebooted service because {exc}')
