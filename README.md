<h1> Бот синтеза речи для автооповещений </h1>
<h2> Необходимые файлы в корне </h2>
<h3> creds.py </h3>
<code>oauth_token = "YandexSpeechKitOAuthToken"
catalog_id = "CloudYandexCatalogId"
bot_api = 'TelegramBotApi'
allow_users = (NumericIDUser1, NumericIDUser2)</code>
<h2> Логирование </h2>
Логи из-за ошибок создания и отправки аудиофайла будут сыпаться в <b>error_audio.log</b></br>
Логи из-за ошибок соединения с серварами Telegram будут сыпаться в error_connection.log
