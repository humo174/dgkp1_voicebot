import json

data = {
    'bot_api': 'YOUR_BOT_API',
    'oauth_token': "YOUR_OAUTH_TOKEN",
    'catalog_id': "YOUR_CATALOG_ID",
    'admin_id': YOUR_ID,
    'allow_users': {
        'USER_ID1': 'NAME_USER1',
        'USER_ID2': 'NAME_USER2'
    }
}

with open('creds.json', 'w', encoding='utf-8') as jo:
    json.dump(data, jo)
