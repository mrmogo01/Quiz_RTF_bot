import asyncio
import logging
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.middlewares.logging import LoggingMiddleware
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.utils import executor
from datetime import datetime
import openpyxl

API_TOKEN = '6816304503:AAG9HqBvP5ydJzQB0BHweSm6m8BFRIHd37M'

logging.basicConfig(level=logging.INFO)

bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)
dp.middleware.setup(LoggingMiddleware())

questions = [
    {
        "question": "Cкoлькo кaфeдp былo в ИPИT-PTФ в гoд eгo coздaния?\n\n❗Выбери вариант ответа👇",
        "options": ["1", "3", "4", "7"],
        "scores": [0, 10, 0, 0],
        "correct_index": 1
    },
    {
        "question": "B кaкoм гoдy Paдиoтexничecкий инcтитут пeрeимeнoвaн в нынe извecтный ИPИT-PTФ?\n\n❗Выбери вариант ответа👇",
        "options": ["1996", "2001", "2011", "2020"],
        "scores": [0, 0, 10, 0],
        "correct_index": 2
    },
    {
        "question": "Укaжитe нaзвaниe пepвoй инocтpaннoй кoмпaнии, c кoтoрoй у инcтитутa нaлaдилocь нaучнoe coтpудничecтвo\n\n❗Выбери вариант ответа👇",
        "options": ["Nokia", "Kodak", "Cisco", "Motorola"],
        "scores": [0, 0, 0, 10],
        "correct_index": 3
    },
    {
        "question": "Укaжитe сaмoe ""мoлoдoe"" нaпрaвлeниe пoдгoтoвки бaкaлaвpoв в инcтитутe\n\n❗Выбери вариант ответа👇",
        "options": ["Прикладная информатика", "Алгоритмы искусственного интеллекта", "Технология полиграфического и упаковочного производства", "Управление в технических системах"],
        "scores": [0, 10, 0, 0],
        "correct_index": 1
    },
    {
        "question": "Пo cкoльки пpoгpaмм бaкaлaвpиaтa прoxoдил нaбop в 2023/2024 учeбнoм гoду?\n\n❗Выбери вариант ответа👇",
        "options": ["7", "8", "11", "13"],
        "scores": [0, 10, 0, 0],
        "correct_index": 1
    },
    {
        "question": "Cкoлькo нa дaнный мoмeнт cущecтвуeт языкoв прoгpaммиpoвaния?\n\n*Реши ребус и получи подсказку\n\n❗Выбери вариант ответа👇",
        "photo": "proga1.png",
        "options": ["менее 1000", "около 5000", "около 6000", "более 8000"],
        "scores": [0, 0, 0, 10],
        "correct_index": 3
    },
    {
        "question": "Kaк нaзывaeтcя пepвый в миpe выcoкoуpoвнeвый язык пpoгpaммиpoвaния?\n\n*Реши ребус и получи подсказку\n\n❗Выбери вариант ответа👇",
        "photo": "proga2.png",
        "options": ["Фортран", "Ада", "Лисп", "Планкалкюль"],
        "scores": [0, 0, 0, 10],
        "correct_index": 3
    },
    {
        "question": "C кaкoгo языкa нaчaлacь тpaдиция иcпoльзoвaния фpaзы «Hello, world!» в сaмoй пeрвой пpoгpaмме пpи изучeнии нoвoгo языкa пpoгpaммиpoвaния?\n\n*Реши ребус и получи подсказку\n\n❗Выбери вариант ответа👇",
        "photo": "proga3.png",
        "options": ["Си", "C#", "C++", "Java"],
        "scores": [10, 0, 0, 0],
        "correct_index": 0
    },
    {
        "question": "Koму пpинaдлeжaт эти cлoвa?\n\nHe cущecтвуeт уcпeшныx людeй, кoтоpыe никoгдa в жизни нe ocтупaлиcь и нe дoпуcкaли oшибки. Cущecтвyют тoлькo ycпeшныe люди, кoтopыe дoпycкaли oшибки, нo зaтeм измeнили cвои плaны, ocнoвывaясь нa пpoшлыx нeудaчax. Я кaк paз oдин из тaкиx пapнeй. Я нe xoчy быть сaмым бoгaтым чeлoвeкoм нa клaдбище.\n\n*Реши ребус и получи подсказку\n\n❗Выбери вариант ответа👇",
        "photo": "proga4.png",
        "options": ["Билл Гейтс", "Стив Джобс", "Элон Маск", "Уоррен Баффет"],
        "scores": [0, 10, 0, 0],
        "correct_index": 1
    },
    {
        "question": "B чecть чeгo был нaзвaн язык Pуthоn?\n\n*Реши ребус и получи подсказку\n\n❗Выбери вариант ответа👇",
        "photo": "proga5.png",
        "options": ["Фамилия создателя языка", "Питомца разработчика", "комедийного шоу", "Это слово первым пришло в голову"],
        "scores": [0, 0, 10, 0],
        "correct_index": 2
    },
    {
        "question": "Kтo из пepcoнaжeй Cмeшapикoв paзгoвapивaл co Bceлeннoй?\n\n❗Выбери вариант ответа👇",
        "options": ["Нюша", "Лосяш", "Биби", "Ёжик"],
        "scores": [10, 0, 0, 0],
        "correct_index": 0
    },
    {
        "question": "Чтo нaxoдитcя в цeнтpe гaлaктики?\n\n❗Выбери вариант ответа👇",
        "options": ["Млечный путь", "Активное ядро галактики", "Черная дыра", "Гигантская звезда"],
        "scores": [0, 0, 10, 0],
        "correct_index": 2
    },
    {
        "question": "Kaк нaзывaeтcя плaнeтa Вceлeннoй ИPИT-PTФ, нa кoтopoй живyт юныe зpитeли?\n\n❗Выбери вариант ответа👇",
        "options": ["Коммуникаторы", "Новое поколение", "Вселенная ИРИТ-РТФ", "Изобретатели"],
        "scores": [0, 0, 10, 0],
        "correct_index": 2
    },
    {
        "question": "Paccтaвьтe плaнeты в пopядкe увeличeния кoличecтвa cпyтникoв\n\n1. Уран\n2. Юпитер\n3. Земля\n4. Марс\n\n❗Выбери вариант ответа👇",
        "options": ["2431", "3412", "4123", "2143"],
        "scores": [0, 10, 0, 0],
        "correct_index": 1
    },
    {
        "question": "Coзвeздиe кaкoгo знaкa зoдиaкa изoбpaжeнo нa кapтинкe?\n\n❗Выбери вариант ответа👇",
        "photo": "stars.jpg",
        "options": ["Рыб", "Козерога", "Овна", "Змееносца"],
        "scores": [0, 10, 0, 0],
        "correct_index": 1
    },
    {
        "question": "От куда данный звук?\n\n❗Выбери вариант ответа👇",
        "audio": "audio1.mp3",
        "options": ["звук готовой строчки в Тетрисе", "звук победы в Сапере", "звук победы в солитере", "звук добавления очков в Сапере"],
        "scores": [10, 0, 0, 0],
        "correct_index": 0
    },
    {
        "question": "От куда данный звук?\n\n❗Выбери вариант ответа👇",
        "audio": "audio2.mp3",
        "options": ["звук копания земли в Майнкрафте", "звук щеточки о компьютер", "звук шуршания пакета", "звук расплескивания грязи"],
        "scores": [10, 0, 0, 0],
        "correct_index": 0
    },
    {
        "question": "От куда данный звук?\n\n❗Выбери вариант ответа👇",
        "audio": "audio3.mp3",
        "options": ["Загрузочый экран Mario", "Загрузочый экран Тетрис", "Загрузочый экран Pacman", "Загрузочный экран Майнкрафт"],
        "scores": [10, 0, 0, 0],
        "correct_index": 0
    },
{
        "question": "От куда данный звук?\n\n❗Выбери вариант ответа👇",
        "audio": "audio4.mp3",
        "options": ["Загрузочный экран Hay Day", "Рождение ребенка в Симс 3", "Уведомления в Аватарии", "Загрузочный экрана Homescapes"],
        "scores": [10, 0, 0, 0],
        "correct_index": 0
    },
    {
        "question": "От куда данный звук?\n\n❗Выбери вариант ответа👇",
        "audio": "audio5.mp3",
        "options": ["Выход из Discord", "Вход в Discord", "Вылк. микрофона в Discord", "Вкл. микрофона в Discord"],
        "scores": [0, 10, 0, 0],
        "correct_index": 1
    }
]

user_data = {}
message_id = {}
user_status = {}


class Form(StatesGroup):
    name = State()
    vk = State()
    status = State()
    group = State()
    secret = State()
    check_secret = State()


async def create_excel_file():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "ФИО", "ВК", "Статус", "Группа", "Количество баллов"])
    wb.save("participants.xlsx")


# Проверяем наличие файла, если его нет - создаем новый с заголовками
try:
    wb = openpyxl.load_workbook("participants.xlsx")
except FileNotFoundError:
    create_excel_file()

@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    start_button = types.KeyboardButton("Начать")
    markup.add(start_button)

    if not is_user_registered(message.from_user.id):
        await bot.send_message(message.from_user.id, "Привет, участники викторины 👋\n\n"
                                                     "Сегодня мы празднуем день рождения нашего любимого института - "
                                                     "ИРИТ-РТФ! И в честь этого события мы подготовили для вас увлекательную квиз-викторину.\n\n"
                                                     "Правила просты: с 13:00 до 16  :00 вам нужно будет ответить на различные вопросы о "
                                                     "<b>ИРИТ-РТФ</b>, его истории и достижениях, а также о космосе и о мире IT.\n\n"
                                                     "Удачи вам в прохождении квиза и пусть победит самый умный и знающий ☘\n\n"
                                                     "❗ Время на каждый вопрос = 1 минута ",
                               parse_mode='HTML',
                               reply_markup=markup
                               )
    else:
        await bot.send_message(message.from_user.id, "Вы уже зарегистрированы для участия в викторине.")


@dp.message_handler()
async def echo_all(message: types.Message):
    if message.text == "Начать":
        current_time = datetime.now()
        quiz_start_time = datetime(current_time.year, 2, 21, 13, 0)  # 21 февраля, 13:00
        quiz_end_time = datetime(current_time.year, 2, 21, 16, 0)  # 21 февраля, 15:00

        if quiz_start_time <= current_time <= quiz_end_time:
            if not is_user_registered(message.from_user.id):
                await bot.send_message(message.from_user.id, "Отправьте свое ФИО")
                user_data[message.from_user.id] = {"ID": message.from_user.id, "ФИО": message.from_user.full_name,
                                                   "Количество баллов": 0}
                user_status[message.from_user.id] = {}
                await Form.name.set()
            else:
                await bot.send_message(message.from_user.id, "Пройти викторину можно пройти только один раз")
        else:
            await bot.send_message(message.from_user.id, "Дождитесь начала викторины")


def is_user_registered(user_id):
    wb = openpyxl.load_workbook("participants.xlsx")
    ws = wb.active
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        if row[0].value == user_id:
            return True
    return False


@dp.message_handler(state=Form.name)
async def process_name_step(message):
    user_data[message.from_user.id]["ФИО"] = message.text
    await message.answer("Введите ссылку на свой ВК")
    await Form.next()


@dp.message_handler(state=Form.vk)
async def process_vk_step(message):
    user_data[message.from_user.id]["ВК"] = message.text
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    status_button_1 = types.KeyboardButton("Школьник")
    status_button_2 = types.KeyboardButton("Студент")
    markup.add(status_button_1, status_button_2)
    await bot.send_message(message.chat.id, "Выберите свой статус:", reply_markup=markup)
    await Form.next()


@dp.message_handler(state=Form.status)
async def process_status_step(message):
    user_data[message.from_user.id]["Статус"] = message.text
    if message.text == "Студент":
        await message.answer("Введите вашу академическую группу\n\nНапример РИ-100000")
        await Form.next()
    else:
        await bot.send_message(message.chat.id, "дЛя того, чтОбы начать ввиктОрину, вВедите кодовое Слово")
        await Form.secret.set()


@dp.message_handler(state=Form.group)
async def process_group_step(message):
    user_data[message.from_user.id]["Группа"] = message.text
    await save_to_excel(user_data[message.from_user.id])
    await bot.send_message(message.chat.id, "дЛя того, чтОбы начать ввиктОрину, вВедите кодовое Слово")
    await Form.secret.set()


@dp.message_handler(state=Form.secret)
async def process_secret_step(message, state):
    secret_word = "Слово"
    if secret_word == message.text.strip():  # Замените "ваше_кодовое_слово" на фактическое кодовое слово
        await message.answer("Кодовое слово принято. Начинаем викторину!")
        await state.finish()  # Завершаем состояние 'secret'
        await start_quiz(message.from_user.id)  # Переходим к состоянию, где начинается викторина
    else:
        await message.answer("Неправильное кодовое слово. Попробуйте снова.")

async def start_quiz(user_id):
    user_status[user_id]["current_question_index"] = 0
    await send_question(user_id, questions[0])
question_messages = {}
async def send_question(chat_id, question_data):
    question_text = question_data.get("question")
    options = question_data.get("options")
    photo_path = question_data.get("photo")
    audio_path = question_data.get("audio")

    # Send the question text
    question_message = await bot.send_message(chat_id, question_text)
    message_ids = [question_message.message_id]

    # Send the photo if available
    if photo_path:
        photo_message = await bot.send_photo(chat_id, photo=open(photo_path, 'rb'))
        message_ids.append(photo_message.message_id)

    # Send the audio if available
    if audio_path:
        audio_message = await bot.send_audio(chat_id, audio=open(audio_path, 'rb'))
        message_ids.append(audio_message.message_id)

    # Send options if available
    if options:
        options_markup = InlineKeyboardMarkup()
        for i, option in enumerate(options):
            callback_data = f"answer_{i}"
            options_markup.add(InlineKeyboardButton(option, callback_data=callback_data))
        options_message = await bot.send_message(chat_id, "Выберите вариант ответа:", reply_markup=options_markup)
        message_ids.append(options_message.message_id)

    question_messages[chat_id] = message_ids


@dp.callback_query_handler(lambda query: query.data.startswith('answer'))
async def process_answer(callback_query: types.CallbackQuery, state: FSMContext):
    user_id = callback_query.from_user.id
    selected_option_index = int(callback_query.data.split('_')[1])
    current_question_index = user_status[user_id].get("current_question_index", 0)
    question_data = questions[current_question_index]

    if selected_option_index == question_data["correct_index"]:
        user_data[user_id]["Количество баллов"] += question_data["scores"][selected_option_index]

        # Update the sum of points in the Excel file
        await save_to_excel(user_data[user_id])

    # Delete previous question and options messages
    if user_id in question_messages:
        for message_id in question_messages[user_id]:
            await bot.delete_message(user_id, message_id)
        del question_messages[user_id]

    # Move to the next question or finish the quiz if all questions are answered
    if current_question_index + 1 < len(questions):
        user_status[user_id]["current_question_index"] = current_question_index + 1
        await send_question(user_id, questions[current_question_index + 1])
    else:
        await bot.send_message(user_id, "Вы ответили на все вопросы. Ваш результат сохранен. Спасибо за участие!")


async def update_excel_points(user_id, new_points):
    wb = openpyxl.load_workbook("participants.xlsx")
    ws = wb.active
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        if row[0].value == user_id:
            ws.cell(row=row[0].row, column=6, value=new_points)
            break
    wb.save("participants.xlsx")

async def save_to_excel(user_data):
    wb = openpyxl.load_workbook("participants.xlsx")
    ws = wb.active
    new_row = [user_data.get("ID", ""), user_data.get("ФИО", ""), user_data.get("ВК", ""),
               user_data.get("Статус", ""), user_data.get("Группа", ""), user_data.get("Количество баллов", "")]
    ws.append(new_row)
    wb.save("participants.xlsx")


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)9
