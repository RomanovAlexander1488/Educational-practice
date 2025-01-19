import openpyxl
import telebot

API_TOKEN = '7635646554:AAHaH6Yi8ET705pT4QIaBjukWFJy1_VLKCs'
bot = telebot.TeleBot(API_TOKEN)
#преобразование в число
def safe_int(value):
    try:
        return float(value) if value else 0
    except ValueError:
        return 0

#для подсчета % выполнения дз
def calculate_homework_percentage(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    teacher_stats = {}

    for row in sheet.iter_rows(min_row=3, values_only=True):
        teacher = row[1]
        total_given = safe_int(row[3])
        total_planned = safe_int(row[6])

        if teacher is None or teacher == '':
            continue

        if teacher not in teacher_stats:
            teacher_stats[teacher] = {'given': 0, 'planned': 0}

        teacher_stats[teacher]['given'] += total_given
        teacher_stats[teacher]['planned'] += total_planned

#подсчет % результата
    results = {}
    for teacher, stats in teacher_stats.items():
        if stats['planned'] > 0:
            percentage = (stats['given'] / stats['planned']) * 100
        else:
            percentage = 0
        results[teacher] = percentage

    return results

def send_report(message, file_path):
    results = calculate_homework_percentage(file_path)

    report = "Отчет по выполнению дз:\n"
    if not results:
        report += "Нет данных по преподавателю или данные неправельные."
    else:
        for teacher, percentage in results.items():
            report += f"{teacher}: {percentage:.2f}% выполнено\n"

    bot.send_message(message.chat.id, report)

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.send_message(message.chat.id, "Йоу! Скинь мне Excel файл с данными по дз.")

#для получения файла
@bot.message_handler(content_types=['document'])
def handle_document(message):

    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    #сохранение файла
    with open("homework_data.xlsx", "wb") as f:
        f.write(downloaded_file)

    #отчёт
    send_report(message, "homework_data.xlsx")

bot.polling()