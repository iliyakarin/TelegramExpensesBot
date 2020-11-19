# TelegramExpensesBot
This is a personal Telegram bot that helps me manage my family budget.
I've decided to use webhooks instead of polling because of low intensity of overall operations on this application from users.
I don't see any reason to constantly generate polls to web hosting if we are using it few times a day or even less often.
But webhooks have some limitations, like we need to use web hosting on the internet with a dns name or static IP and valid ssl certificate provided by a trusted ssl authority provider.

Plans to do:
1. Creation set of google spreadsheet documents for whole year with cell formatting.
   Thinking of different schemas, create it from a template file or generate it from scratch asking users for parameters.
   
2. Refactor code for better performance.

3. Add a full deployment manual.


Description (draft):

To use this application you need to:

1. Create a .env file in the same folder where __init__.py file exists.

2. Register telegram bot using @BotFather bot.
   You will receive an API token for your bot.
   It looks similar to: '3332224444:RTYuiURTHjiUYTgHJUYTFghGFFRFGt'
   **Very sensitive information, keep it safe !!!**
   You need to put this line in the .env file for variable TG_API_TOKEN=''. Example: TG_API_TOKEN='3332224444:RTYuiURTHjiUYTgHJUYTFghGFFRFGt'
   
3. Register service account in your google account settings.
   And save service_account.json file (this file is a key file to use your google spreadsheets under service account without password)!
   **Very sensitive information, keep it safe !!!**
   
   Add path to it into the .env file for variable SERVICE_ACCOUNT="".
   Example: 
   windows: SERVICE_ACCOUNT="C:\Users\Iliya\secret\service_account.json" 
   or 
   linux : '/home/iliya/ExpensesBot/service_account.json'


