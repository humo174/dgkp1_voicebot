<h1> Бот синтеза речи для автооповещений </h1>
<h2> Необходимые файлы в корне </h2>
<h3> creds.json </h3>

```
{
    'bot_api': 'YOUR_BOT_API',
    'oauth_token': "YOUR_OAUTH_TOKEN",
    'catalog_id': "YOUR_CATALOG_ID",
    'admin_id': YOUR_ID,
    'allow_users': {
        'USER_ID1': 'NAME_USER1',
        'USER_ID2': 'NAME_USER2'
    }
}
```

<h2> Логирование </h2>
Логи из-за ошибок создания и отправки аудиофайла будут сыпаться в <b>error_audio.log</b></br>
Логи из-за ошибок соединения с серварами Telegram будут сыпаться в <b>error_connection.log</b>
