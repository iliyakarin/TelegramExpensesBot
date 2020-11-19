from flask import Flask
import flask
from flask_sslify import SSLify
import telebot
import gspread
from datetime import datetime
import logging
from time import sleep
import os
from dotenv import load_dotenv

load_dotenv()
TG_API_TOKEN = os.getenv('TG_API_TOKEN')
SERVICE_ACCOUNT = os.getenv('SERVICE_ACCOUNT')
URL = 'https://api.telegram.org/bot' + TG_API_TOKEN + '/'
# noinspection SpellCheckingInspection
datem = datetime(datetime.today().year, datetime.today().month, 1)
app = Flask(__name__)
# noinspection SpellCheckingInspection
sslify = SSLify(app)
gc = gspread.service_account(filename=SERVICE_ACCOUNT)
bot = telebot.TeleBot(TG_API_TOKEN)


def next_available_row(worksheet):
    """
    This function accepts worksheet as an argument and returns empty row number as result.
    Function also adds 5 rows if last row is not empty.
    Worksheet should be whole element from worksheet list returned from Spreadsheet.
    """
    i = worksheet.col_count
    result_list = []
    while i > 0:
        result = len(list(filter(None, worksheet.col_values(i)))) + 1
        result_list.append(result)
        i = i - 1
    result_list.sort()
    empty_row = result_list[-1]
    last_row = worksheet.get_all_records()[-1]
    if last_row:
        worksheet.resize(int(empty_row) + 5)
        while worksheet.row_values(empty_row):
            empty_row = empty_row + 1
        else:
            return empty_row
    else:
        while worksheet.row_values(empty_row):
            empty_row = empty_row + 1
        else:
            return empty_row


@bot.message_handler(func=lambda message: True, regexp='^(\w+(:))+\w+$')
def add_current_month_expense_by_default(message):
    """
    This function takes as input initial message that should be in expense:price format. Example bread:50
    And asks user to choose category (worksheet) for this expense from current month document.
    Finally this function sends these two inputs to next step add_current_month_expense_by_default_category function.
    """
    if message.from_user.id == 395147397:
        try:
            input_list = message.text.split(':')
            sh = gc.open(datem.today().strftime("%Y.%m") + " Family budget")
            worksheet_list = sh.worksheets()
            worksheet_list = worksheet_list[:len(worksheet_list) - 2]
            category_list = []
            for i in worksheet_list:
                category_list.append(str(worksheet_list.index(i) + 1) + ") " + str(i.title))
            bot.send_message(message.chat.id, "\n".join([s for s in category_list]))
            msg = bot.reply_to(message, 'Please choose the expense category number from message above: ')
            bot.register_next_step_handler(msg, add_current_month_expense_by_default_category,
                                           input_list, worksheet_list)
        except gspread.SpreadsheetNotFound as fe:
            if any(str(message.text) in s for s in ['exit', 'start', 'help']):
                bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                                  'command /start or /help')
                return
            else:
                bot.reply_to(message, 'ERROR!\nFile Copy of '
                                      '' + datem.today().strftime("%Y.%m") + ' Family budget FOR TEST not found!')
                bot.send_message(message.chat.id, 'Please ensure that document for current month exists: ')
                logging.error(str(fe))
                return
        except Exception as e:
            if any(str(message.text) in s for s in ['exit', 'start', 'help']):
                bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                                  'command /start or /help')
                return
            else:
                bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
                bot.send_message(message.chat.id, 'Please check if you have access and internet connection is stable: ')
                logging.error(str(e))
                return
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


def add_current_month_expense_by_default_category(message, input_list, worksheet_list):
    """
    This function takes variables from previous step add_current_month_expense_by_default and adds
    expense from input_list[0] as expense
    price from input_list[1] as price
    and category number (worksheet) from message.text
    And finally format cells as expected and fills up information to empty row.
    """
    category_num = int(message.text) - 1
    try:
        if worksheet_list[category_num] in worksheet_list and int(message.text) > 0:
            bot.send_message(message.chat.id, 'You have chosen: ' + str(message.text) + ") " +
                             (worksheet_list[category_num]).title)
            sh = gc.open(datem.today().strftime("%Y.%m") + " Family budget")
            today_date = datem.today().strftime("%d.%m.%Y")
            worksheet_list = sh.worksheets()
            worksheet = worksheet_list[category_num]
            empty_string = next_available_row(worksheet)
            worksheet.format('A' + str(empty_string), {"horizontalAlignment": "LEFT",
                                                       "textFormat": {
                                                           "fontSize": 12
                                                       }
                                                       })
            worksheet.format('B' + str(empty_string), {"horizontalAlignment": "CENTER",
                                                       "numberFormat": {
                                                           "type": "DATE",
                                                           "pattern": "dd.mm.yyyy"},
                                                       "textFormat": {
                                                           "fontSize": 12
                                                       }
                                                       })
            worksheet.format('C' + str(empty_string), {"horizontalAlignment": "RIGHT",
                                                       "numberFormat": {
                                                           "type": "CURRENCY",
                                                           "pattern": "#,##0.00 руб."},
                                                       "textFormat": {
                                                           "fontSize": 12
                                                       }
                                                       })
            worksheet.update('A' + str(empty_string), input_list[0])
            worksheet.update('B' + str(empty_string), today_date, value_input_option='USER_ENTERED')
            worksheet.update('C' + str(empty_string), int(input_list[1]))
            read_row_str = (', '.join(map(str, worksheet.get(
                'A' + str(empty_string) + ":" + 'C' + str(empty_string))
                                          ))).strip('[]')
            bot.send_message(
                message.chat.id, read_row_str + ""
                                                " has been added to "
                                                "" + worksheet.title + ""
                                                " worksheet into " + datem.today().strftime("%Y.%m") + " file.")
        else:
            bot.send_message(message.chat.id, 'Category with number: ' + str(message.text) + ' '
                                                                                             'not found! Please retry')
            msg = bot.reply_to(message, 'Please choose the expense category number from message above: ')
            bot.register_next_step_handler(msg, add_current_month_expense_by_default_category,
                                           input_list, worksheet_list)
    except Exception as e:
        if any(str(message.text) in s for s in ['exit', 'start', 'help']):
            bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                              'command /start or /help')
            return
        else:
            bot.reply_to(message, 'ERROR!\nCategory ' + str(message.text) + ' not found!\nTry once more!')
            msg = bot.reply_to(message, 'Please choose the expense category number from message above: ')
            bot.register_next_step_handler(msg, add_current_month_expense_by_default_category,
                                           input_list, worksheet_list)
            logging.error(str(e))


@bot.message_handler(commands=['start', 'Start', 'help', 'Help'])
def handle_start_help(message):
    """
    Shows start menu with all supported commands.
    """
    if message.from_user.id == 395147397:
        # noinspection SpellCheckingInspection
        bot.send_message(message.chat.id, "All available commands:\n"
                                          "/start or /help shows help menu\n\n"
                                          "/CurrentMonthBalance or /CMB\n shows current month balance\n\n"
                                          "/DefinedMonthBalance or /DMB\n shows defined month balance\n\n\n"
                                          "/CurrentMonthExpenseByCategory\n retrieves expenses for current month for "
                                          "every category\n\n"
                                          "/ExactMonthExpenseByCategory\n retrieves expenses for defined month for "
                                          "every "
                                          "category\n\n\n"
                                          "/AddExpenseToCurrentMonth or /AECM\n adds expense to current month\n\n"
                                          "/AddExpenseToDefinedMonth or /AEDM\n adds expense to defined month\n\n\n"
                                          "/FormatDefinedFile or /FDF\n restores correct formating for whole document"
                                          " defined by date, takes up to 5 minutes, do not use frequently\n"
                         )
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['CurrentMonthBalance', 'currentmonthbalance', 'cmb', 'CMB'])
def current_month_balance(message):
    """
    Sending 3 messages back to chat with balance statistics for current month.
    income, expenses, and balance
    executes with command /CurrentMonthBalance
    """
    if message.from_user.id == 395147397:
        sh = gc.open(datem.today().strftime("%Y.%m") + " Family budget")
        worksheet = sh.worksheet("Balance")
        bot.send_message(message.chat.id, "Current month income is: " + worksheet.acell('B18').value)
        bot.send_message(message.chat.id, "Current month expenses are: " + worksheet.acell('D18').value)
        bot.send_message(message.chat.id, "Current month balance is: " + worksheet.acell('F1').value)
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


def count_category_rows_expenses(month):
    """
    Counts number of rows on balance sheet for all expenses.
    And returning list of all expenses from balance sheet.
    """
    sh = gc.open(month + " Family budget")
    worksheet = sh.worksheet("Balance")
    values_list = worksheet.col_values(3)
    values_list.pop(0)
    return values_list


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['AddExpenseToCurrentMonth', 'addexpensetocurrentmonth', 'AECM', 'aecm'])
def add_current_month_expense(message):
    """
    Adds expense for current month for defined category.
    executes with command /AddExpenseToCurrentMonth or /addexpensetocurrentmonth or /AECM or /aecm
    This function is 1/3 step process of adding expense to current month.
    Consists of 3 step process
    add_current_month_expense - checks if file for curent month exists and asks for category input.
    add_current_month_expense_input_cat -
    """
    if message.from_user.id == 395147397:
        sh = gc.open(datem.today().strftime("%Y.%m") + " Family budget")
        worksheet_list = sh.worksheets()
        worksheet_list = worksheet_list[:len(worksheet_list) - 2]
        category_list = []
        for i in worksheet_list:
            category_list.append(str(worksheet_list.index(i) + 1) + ") " + str(i.title))
        bot.send_message(message.chat.id, "\n".join([s for s in category_list]))
        msg = bot.reply_to(message, 'Please choose the expense category number from message above: ')
        bot.register_next_step_handler(msg, add_current_month_expense_input_cat, worksheet_list)
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


def add_current_month_expense_input_cat(message, worksheet_list):
    """
    This function is 2/3 step process of adding expense to current month
    Accepts 2 arguments message and worksheet_list
    Message is category number
    worksheet_list is list of all worksheets from spreadsheet of current month
    """
    category_num = int(message.text) - 1
    try:
        if worksheet_list[category_num] in worksheet_list and int(message.text) > 0:
            bot.send_message(message.chat.id, 'You have chosen: ' + str(message.text) + ") " +
                             (worksheet_list[category_num]).title)
            msg = bot.reply_to(message, 'Please enter tag and price as example "Bread: 50": ')
            bot.register_next_step_handler(msg, add_current_month_expense_input_string, category_num)
        else:
            bot.send_message(message.chat.id, 'Category with number: ' + str(message.text) + ' '
                                                                                             'not found! Please retry')
            msg = bot.reply_to(message, 'Please choose the expense category number from message above: ')
            bot.register_next_step_handler(msg, add_current_month_expense_input_cat, worksheet_list)
    except Exception as e:
        bot.reply_to(message, 'ERROR!\nCategory ' + str(message.text) + ' not found!\nTry once more!')
        msg = bot.reply_to(message, 'Please choose the expense category number from message above: ')
        bot.register_next_step_handler(msg, add_current_month_expense_input_cat, worksheet_list)
        logging.error(str(e))


def add_current_month_expense_input_string(message, category_num):
    """
    This function is 3/3 step process of adding expense to current month
    expense from input_list[0] as expense
    price from input_list[1] as price
    and category number (worksheet) from message.text
    And finally format cells as expected and fills up information to empty row
    """
    input_list = message.text.split(':')
    if len(input_list) == 2:
        sh = gc.open(datem.today().strftime("%Y.%m") + " Family budget")
        today_date = datem.today().strftime("%d.%m.%Y")
        worksheet_list = sh.worksheets()
        worksheet = worksheet_list[category_num]
        empty_string = next_available_row(worksheet)
        worksheet.format('A' + str(empty_string), {"horizontalAlignment": "LEFT",
                                                   "textFormat": {
                                                       "fontSize": 12
                                                   }
                                                   })
        worksheet.format('B' + str(empty_string), {"horizontalAlignment": "CENTER",
                                                   "numberFormat": {
                                                       "type": "DATE",
                                                       "pattern": "dd.mm.yyyy"},
                                                   "textFormat": {
                                                       "fontSize": 12
                                                   }
                                                   })
        worksheet.format('C' + str(empty_string), {"horizontalAlignment": "RIGHT",
                                                   "numberFormat": {
                                                       "type": "CURRENCY",
                                                       "pattern": "#,##0.00 руб."},
                                                   "textFormat": {
                                                       "fontSize": 12
                                                   }
                                                   })
        worksheet.update('A' + str(empty_string), input_list[0])
        worksheet.update('B' + str(empty_string), today_date, value_input_option='USER_ENTERED')
        worksheet.update('C' + str(empty_string), int(input_list[1]))
        read_row_str = (', '.join(map(str, worksheet.get(
            'A' + str(empty_string) + ":" + 'C' + str(empty_string))
                                      ))).strip('[]')
        bot.send_message(
            message.chat.id, read_row_str + ""
                                            " has been added to "
                                            "" + worksheet.title + ""
                                            " worksheet into " + datem.today().strftime("%Y.%m") + " file.")
    else:
        bot.send_message(message.chat.id,
                         "You have entered *** " + message.text + " *** - wrong input format (Expense:price) or not 2 "
                                                                  "parameters in the input")
        sh = gc.open(datem.today().strftime("%Y.%m") + " Family budget")
        worksheet_list = sh.worksheets()
        msg = bot.reply_to(message, "Please enter the expense for "
                                    "" + str(datem.today().strftime("%Y.%m")) + ""
                                    " and " + str(worksheet_list[category_num].title) + ""
                                    " category in format expense:price")
        bot.register_next_step_handler(msg, add_current_month_expense_input_string, category_num)


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['AddExpenseToDefinedMonth', 'addexpensetodefinedmonth', 'AEDM' 'aedm'])
def add_defined_month_expense(message):
    """
    Adds expense for defined month and for defined category.
    This function is 1/3 step process of adding expense to defined month
    executes with command /AddExpenseToDefinedMonth or /addexpensetodefinedmonth or /AEDM or /aedm
    Asks user to input year and month of document which will be used to add expense
    """
    if message.from_user.id == 395147397:
        msg = bot.reply_to(message, 'Enter year and month in format YYYY.MM: ')
        bot.register_next_step_handler(msg, defined_month_expense_date)
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


def defined_month_expense_date(message):
    """
    This function is 2/3 step process of adding expense to defined month
    """
    month = message.text
    try:
        bot.send_message(message.chat.id, 'You have entered: ' + month)
        sh = gc.open(month + " Family budget")
        worksheet_list = sh.worksheets()
        worksheet_list = worksheet_list[:len(worksheet_list) - 2]
        category_list = []
        for i in worksheet_list:
            category_list.append(str(worksheet_list.index(i) + 1) + ") " + str(i.title))
        bot.send_message(message.chat.id, "\n".join([s for s in category_list]))
        msg = bot.reply_to(message, 'Please choose the expense category from message above: ')
        bot.register_next_step_handler(msg, add_defined_month_expense_category, month, worksheet_list)
    except gspread.SpreadsheetNotFound as fe:
        bot.reply_to(message, 'ERROR!\nFile ' + str(message.text) + ' not found!\nTry once more!')
        msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
        bot.register_next_step_handler(msg, defined_month_expense_date)
        logging.error(str(fe))
    except Exception as e:
        bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
        msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
        bot.register_next_step_handler(msg, defined_month_expense_date)
        logging.error(str(e))


def add_defined_month_expense_category(message, month, worksheet_list):
    try:
        category_num = int(message.text) - 1
        if worksheet_list[category_num] in worksheet_list and int(message.text) > 0:
            category = worksheet_list[category_num].title
            msg = bot.reply_to(message, "Please enter the expense for "
                                        "" + str(month) + " and " + str(category) + ""
                                        " category in format expense:date:price")
            bot.register_next_step_handler(msg, add_defined_month_expense_input, month, category_num)
        else:
            bot.send_message(message.chat.id, 'Category with number: '
                                              '' + str(message.text) + ' not found! Please retry')
            message = bot.reply_to(message, 'Please choose the expense category number from message above: ')
            bot.register_next_step_handler(message, add_defined_month_expense_category, worksheet_list, month)

    except gspread.WorksheetNotFound as fe:
        category_num = int(message.text) - 1
        category = worksheet_list[category_num].title
        bot.reply_to(message, 'ERROR!\nWorksheet ' + str(category) + ' not found!\nTry once more!')
        msg = bot.reply_to(message, 'Please choose existing category from message above: ')
        bot.register_next_step_handler(msg, add_defined_month_expense_category, month, worksheet_list)
        logging.error(str(fe))
    except Exception as e:
        bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
        msg = bot.reply_to(message, 'Please choose existing category from message above: ')
        bot.register_next_step_handler(msg, add_defined_month_expense_category, month, worksheet_list)
        logging.error(str(e))


def add_defined_month_expense_input(message, month, category_num):
    expense = message.text
    input_list = expense.split(':')
    try:
        if len(input_list) == 3:
            sh = gc.open(month + " Family budget")
            worksheet_list = sh.worksheets()
            worksheet = worksheet_list[category_num]
            empty_string = next_available_row(worksheet)
            worksheet.format('A' + str(empty_string), {"horizontalAlignment": "LEFT",
                                                       "textFormat": {
                                                           "fontSize": 12
                                                       }
                                                       })
            worksheet.format('B' + str(empty_string), {"horizontalAlignment": "CENTER",
                                                       "numberFormat": {
                                                           "type": "DATE",
                                                           "pattern": "dd.mm.yyyy"},
                                                       "textFormat": {
                                                           "fontSize": 12
                                                       }
                                                       })
            worksheet.format('C' + str(empty_string), {"horizontalAlignment": "RIGHT",
                                                       "numberFormat": {
                                                           "type": "CURRENCY",
                                                           "pattern": "#,##0.00 руб."},
                                                       "textFormat": {
                                                           "fontSize": 12
                                                       }
                                                       })
            worksheet.update('A' + str(empty_string), input_list[0])
            worksheet.update('B' + str(empty_string), input_list[1], value_input_option='USER_ENTERED')
            worksheet.update('C' + str(empty_string), int(input_list[2]))
            read_row_str = (', '.join(map(str, worksheet.get(
                'A' + str(empty_string) + ":" + 'C' + str(empty_string))
                                          ))).strip('[]')
            bot.send_message(message.chat.id, read_row_str + " has been added to "
                                                             "" + worksheet.title + " "
                                                             "worksheet into " + str(month) + " file.")
        else:
            bot.send_message(message.chat.id,
                             "You have entered *** "
                             "" + str(expense) + "*** - wrong input format (Expense:date:price) or "
                             "not 3 parameters in the input")
            sh = gc.open(month + " Family budget")
            worksheet_list = sh.worksheets()
            msg = bot.reply_to(message, "Please enter the expense for "
                                        "" + str(month) + ""
                                        " and "
                                        "" + worksheet_list[category_num] + ""
                                        " category in format expense:date:price")
            bot.register_next_step_handler(msg, add_defined_month_expense_input, month, category_num)
    except gspread.SpreadsheetNotFound as fe:
        if any(str(message.text) in s for s in ['exit', 'start', 'help']):
            bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                              'command /start or /help')
            return
        else:
            bot.reply_to(message, 'ERROR!\nFile ' + str(message.text) + ' not found!\nTry once more!')
            bot.send_message(message.chat.id, 'Please ensure that document for current month exists: ')
            logging.error(str(fe))
            return
    except Exception as e:
        bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
        bot.send_message(message.chat.id, 'Please check if you have access and internet connection is stable: ')
        logging.error(str(e))
        return


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['CurrentMonthExpenseByCategory', 'currentmonthexpensebycategory'])
def current_month_expense_by_category(message):
    """
    Returns message with balance statistics for current month expenses by every category.
    income, expenses, and balance
    executes with command /CurrentMonthExpenseByCategory
    """
    if message.from_user.id == 395147397:
        try:
            sh = gc.open(datem.today().strftime("%Y.%m") + " Family budget")
            month = datem.today().strftime("%Y.%m")
            worksheet = sh.worksheet("Balance")
            values_list = count_category_rows_expenses(month)
            if len(values_list) >= 1:
                for i in values_list:
                    index = values_list.index(i) + 2
                    bot.send_message(message.chat.id,
                                     "Current month expense for " + i + " is: " + worksheet.acell(
                                         'D' + str(index)).value + "\n")
            elif len(values_list) < 1:
                bot.send_message(message.chat.id, "No categories found!")
            else:
                bot.send_message(message.chat.id, "Something went wrong!")
        except gspread.SpreadsheetNotFound as fe:
            if any(str(message.text) in s for s in ['exit', 'start', 'help']):
                bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                                  'command /start or /help')
                return
            else:
                bot.reply_to(message, 'ERROR!\nFile ' + str(message.text) + ' not found!')
                bot.send_message(message.chat.id, 'Please ensure that document for current month exists: ')
                logging.error(str(fe))
                return
        except Exception as e:
            bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
            bot.send_message(message.chat.id, 'Please check if you have access and internet connection is stable: ')
            logging.error(str(e))
            return
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['DefinedMonthBalance', 'definedmonthbalance', 'DMB', 'dmb'])
def exact_month_balance(message):
    """
    Sending 3 messages back with balance statistics for defined month.
    income, expenses, and balance
    This function is 1/2 step process of returning to chat balance statistics for defined month.
    executes with command /DefinedMonthBalance or /definedmonthbalance or /DMB or /dmb
    """
    if message.from_user.id == 395147397:
        msg = bot.reply_to(message, 'Enter year and month in format YYYY.MM: ')
        bot.register_next_step_handler(msg, exact_month_balance_input)
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


def exact_month_balance_input(message):
    """
    This function is 2/2 step process of returning to chat balance statistics for defined month.
    """
    try:
        sh = gc.open(message.text + " Family budget")
        bot.send_message(message.chat.id, 'You have entered: ' + message.text)
        worksheet = sh.worksheet("Balance")
        bot.send_message(message.chat.id, message.text + " month income is: " + worksheet.acell('B18').value)
        bot.send_message(message.chat.id, message.text + " month expenses are: " + worksheet.acell('D18').value)
        bot.send_message(message.chat.id, message.text + " month balance is: " + worksheet.acell('F1').value)
    except gspread.SpreadsheetNotFound as fe:
        if any(str(message.text) in s for s in ['exit', 'start', 'help']):
            bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                              'command /start or /help')
            return
        else:
            bot.reply_to(message, 'ERROR!\nFile ' + str(message.text) + ' not found!\nTry once more!')
            msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
            bot.register_next_step_handler(msg, exact_month_balance_input)
            logging.error(str(fe))
    except Exception as e:
        bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
        msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
        bot.register_next_step_handler(msg, exact_month_balance_input)
        logging.error(str(e))


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['ExactMonthExpenseByCategory', 'exactmonthexpensebycategory', 'EMEC', 'emec'])
def exact_month_expense_by_category(message):
    """
    This function is 1/2 step process of returning to chat balance statistics for defined month for each category.
    executes with command /ExactMonthExpenseByCategory or /exactmonthexpensebycategory or /EMEC or /emec

    """
    if message.from_user.id == 395147397:
        msg = bot.reply_to(message, 'Enter year and month in format YYYY.MM: ')
        bot.register_next_step_handler(msg, exact_month_expense_by_category_input)
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


def exact_month_expense_by_category_input(message):
    """
    This function is 2/2 step process of returning to chat balance statistics for defined month for each category.
    This function accepts message as an argument in format year.month (YYYY.MM). Example: 2020.08
    """
    month = message.text
    try:
        sh = gc.open(month + " Family budget")
        bot.send_message(message.chat.id, 'You have entered: ' + month)
        worksheet = sh.worksheet("Balance")
        values_list = count_category_rows_expenses(month)
        if len(values_list) >= 1:
            for i in values_list:
                index = values_list.index(i) + 2
                bot.send_message(message.chat.id,
                                 month + " month expense for " + i + " is: " + worksheet.acell(
                                     'D' + str(index)).value + "\n")
        elif len(values_list) < 1:
            bot.send_message(message.chat.id, "No categories found!")
        else:
            bot.send_message(message.chat.id, "Something went wrong!")
    except gspread.SpreadsheetNotFound as fe:
        if any(str(message.text) in s for s in ['exit', 'start', 'help']):
            bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                              'command /start or /help')
            return
        else:
            bot.reply_to(message, 'ERROR!\nFile ' + str(message.text) + ' not found!\nTry once more!')
            msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
            bot.register_next_step_handler(msg, exact_month_expense_by_category_input)
            logging.error(str(fe))
    except Exception as e:
        bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
        msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
        bot.register_next_step_handler(msg, exact_month_expense_by_category_input)
        logging.error(str(e))


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['FormatDefinedFile', 'formatdefinedfile', 'fdf', 'FDF'])
def format_defined_file(message):
    """
    This function is 1/2 step process of formatting defined file to reference cells format.
    """
    if message.from_user.id == 395147397:
        msg = bot.reply_to(message, 'Enter year and month in format YYYY.MM to update cell format file wide: ')
        bot.register_next_step_handler(msg, format_defined_file_input)
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


def format_defined_file_input(message):
    """
    This function is 2/2 step process of formatting defined file to reference cells format.
    This function accept 1 argument in format of document date YYYY.MM Example: 2020.08
    and reformat whole file sheet by sheet.
    Google API has a Quota quota group 'WriteGroup' and limit 'Write requests per user per 100 seconds
    That's why timeout is introduced in for cycle
    This operation should normally take ~5 minutes
    """
    month = message.text
    try:
        bot.send_message(message.chat.id, 'You have entered: ' + month + '\nPlease be patient, '
                                          'operation could take a while. Execution time depends on number of cells'
                                          ' and lists in the document, you will get message on completion.')
        sh = gc.open(month + " Family budget")
        worksheet_list = sh.worksheets()
        balance_sheet = worksheet_list[-1]
        worksheet_list = worksheet_list[:len(worksheet_list) - 1]
        balance_sheet.format('A1:F1', {"horizontalAlignment": "LEFT",
                                       "textFormat": {
                                           "fontSize": 14,
                                           "fontFamily": 'Arial'
                                       }
                                       })
        balance_sheet.format('A2:A' + str(balance_sheet.row_count), {"horizontalAlignment": "LEFT",
                                                                     "textFormat": {
                                                                         "fontSize": 12,
                                                                         "fontFamily": 'Arial'
                                                                     }
                                                                     })
        balance_sheet.format('C2:C' + str(balance_sheet.row_count), {"horizontalAlignment": "LEFT",
                                                                     "textFormat": {
                                                                         "fontSize": 12,
                                                                         "fontFamily": 'Arial'
                                                                     }
                                                                     })
        balance_sheet.format('B2:B' + str(balance_sheet.row_count), {"horizontalAlignment": "RIGHT",
                                                                     "numberFormat": {
                                                                         "type": "CURRENCY",
                                                                         "pattern": "#,##0.00 руб."},
                                                                     "textFormat": {
                                                                         "fontSize": 12,
                                                                         "fontFamily": 'Arial'
                                                                     }
                                                                     })
        balance_sheet.format('D2:D' + str(balance_sheet.row_count), {"horizontalAlignment": "RIGHT",
                                                                     "numberFormat": {
                                                                         "type": "CURRENCY",
                                                                         "pattern": "#,##0.00 руб."},
                                                                     "textFormat": {
                                                                         "fontSize": 12,
                                                                         "fontFamily": 'Arial'
                                                                     }
                                                                     })
        for item in worksheet_list:
            sleep(30)
            item.format('A1', {"horizontalAlignment": "LEFT",
                               "textFormat": {
                                   "fontSize": 14,
                                   "fontFamily": 'Arial'
                               }
                               })
            item.format('B1:C1', {"horizontalAlignment": "CENTER",
                                  "textFormat": {
                                      "fontSize": 14,
                                      "fontFamily": 'Arial'
                                  }
                                  })
            item.format('D1', {"horizontalAlignment": "RIGHT",
                               "numberFormat": {
                                   "type": "CURRENCY",
                                   "pattern": "#,##0.00 руб."},
                               "textFormat": {
                                   "fontSize": 14,
                                   "fontFamily": 'Arial'
                               }
                               })
            item.format('A2:A' + str(item.row_count), {"horizontalAlignment": "LEFT",
                                                       "textFormat": {
                                                           "fontSize": 12,
                                                           "fontFamily": 'Arial'
                                                       }
                                                       })
            item.format('B2:B' + str(item.row_count), {"horizontalAlignment": "CENTER",
                                                       "numberFormat": {
                                                           "type": "DATE",
                                                           "pattern": "dd.mm.yyyy"},
                                                       "textFormat": {
                                                           "fontSize": 12,
                                                           "fontFamily": 'Arial'
                                                       }
                                                       })
            item.format('C2:C' + str(item.row_count), {"horizontalAlignment": "RIGHT",
                                                       "numberFormat": {
                                                           "type": "CURRENCY",
                                                           "pattern": "#,##0.00 руб."},
                                                       "textFormat": {
                                                           "fontSize": 12,
                                                           "fontFamily": 'Arial'
                                                       }
                                                       })
        bot.send_message(message.chat.id, "Document re-formatting completed successfully!!!")
    except gspread.SpreadsheetNotFound as fe:
        if any(str(message.text) in s for s in ['exit', 'start', 'help']):
            bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                              'command /start or /help')
            return
        else:
            bot.reply_to(message, 'ERROR!\nFile ' + str(message.text) + ' not found!\nTry once more!')
            msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
            bot.register_next_step_handler(msg, format_defined_file_input)
            logging.error(str(fe))
    except Exception as e:
        bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
        msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
        bot.register_next_step_handler(msg, format_defined_file_input)
        logging.error(str(e))


@app.route('/', methods=['POST'])
def webhook():
    if flask.request.headers.get('content-type') == 'application/json':
        json_string = flask.request.get_data().decode('utf-8')
        update = telebot.types.Update.de_json(json_string)
        bot.process_new_updates([update])
        return ''
    else:
        flask.abort(403)


# Enable saving next step handlers to file "./.handlers-saves/step.save".
# Delay=2 means that after any change in next step handlers (e.g. calling register_next_step_handler())
# saving will happen after delay 2 seconds.
bot.enable_save_next_step_handlers(delay=2, filename="./step.save")

# Load next_step_handlers from save file (default "./.handlers-saves/step.save")
# WARNING It will work only if enable_save_next_step_handlers was called!
bot.load_next_step_handlers(filename="./step.save")

if __name__ == '__main__':
    app.run()
