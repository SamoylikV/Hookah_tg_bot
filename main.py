# !venv/bin/python
import logging
import json
from openpyxl import Workbook

from aiogram.contrib.middlewares.logging import LoggingMiddleware
from aiogram.dispatcher import FSMContext
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.utils import executor

json_file = open("credentials.json", "r")
data = json.load(json_file)
json_file.close()
token = data["token"]
admin_password = data["admin_password"]
moderator_password = data["moderator_password"]
bot = Bot(token=token)
dp = Dispatcher(bot, storage=MemoryStorage())
dp.middleware.setup(LoggingMiddleware())
logging.basicConfig(level=logging.INFO)
admin_keyboard = types.ReplyKeyboardMarkup()
admin_keyboard.add(types.InlineKeyboardButton(text="Добавить товар", callback_data="add_item"))
admin_keyboard.add(types.InlineKeyboardButton(text="Удалить товар", callback_data="delete_item"))
admin_keyboard.add(types.InlineKeyboardButton(text="Показать товары", callback_data="show_items"))
admin_keyboard.add(types.InlineKeyboardButton(text="Высрать exel", callback_data="export_to_excel"))

moderator_keyboard = types.ReplyKeyboardMarkup()
moderator_keyboard.add(types.InlineKeyboardButton(text="Написать остаток товаров", callback_data="add_mod_amount"))


class AdminState(StatesGroup):
    admin = State()
    add_item = State()
    add_amount = State()
    delete_item = State()
    show_items = State()


class ModeratorState(StatesGroup):
    moderator = State()
    add_mod_amount = State()
    edit_mod_amount = State()
    mod_amount_done = State()


@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    if message.text == "/start " + moderator_password:
        await message.answer("Дарова модератор", reply_markup=moderator_keyboard)
        await ModeratorState.moderator.set()
    elif message.text == "/start " + admin_password:
        await message.answer("Бонсуар, админ", reply_markup=admin_keyboard)
        await AdminState.admin.set()
    else:
        await message.answer("403\nAccess to the requested resource is forbidden")


@dp.message_handler(state=AdminState.admin)
async def admin(message: types.Message, state: FSMContext):
    if message.text == "Добавить товар":
        await message.answer("Введите название товара")
        await AdminState.add_item.set()
    elif message.text == "Удалить товар":
        await message.answer("Введите название товара")
        await AdminState.delete_item.set()
    elif message.text == "Показать товары":
        await message.answer("товары:")
        with open("items.json", "r") as file:
            items = json.load(file)
        print(items)
        for item in items:
            await message.answer(item + " " + items[item][0] + " " + items[item][1])
        await AdminState.admin.set()
    elif message.text == "Высрать exel":
        await message.answer("Высрал exel")
        wb = Workbook()
        ws = wb.active
        with open("items.json", "r") as file:
            items = json.load(file)
        for item in items:
            ws.append([item, items[item][0], items[item][1]])
        wb.save("items.xlsx")
        await bot.send_document(message.from_user.id, open("items.xlsx", "rb"))
        await AdminState.admin.set()
    else:
        await message.answer("Неверная команда")
        await AdminState.admin.set()


@dp.message_handler(state=ModeratorState.moderator)
async def moderator(message: types.Message, state: FSMContext):
    if message.text == "Написать остаток товаров":
        items_keyboard = types.ReplyKeyboardMarkup()
        items_keyboard.add(types.InlineKeyboardButton(text="↩️Назад ↩️", callback_data="back"))
        with open("items.json", "r") as file:
            items = json.load(file)
        for item in items:
            items_keyboard.add(types.InlineKeyboardButton(text=item, callback_data=item))
        await message.answer("Выберите название товара", reply_markup=items_keyboard)
        await ModeratorState.add_mod_amount.set()

    else:
        await message.answer("Неверная команда")
        await ModeratorState.moderator.set()


# moderator choose item and enters its price
@dp.message_handler(state=ModeratorState.add_mod_amount)
async def add_mod_amount(message: types.Message, state: FSMContext):
    await state.update_data(item=message.text)
    if message.text != "↩️Назад ↩️":
        await message.answer("Введите остаток")
        await ModeratorState.edit_mod_amount.set()
    else:
        await message.answer("Назад", reply_markup=moderator_keyboard)
        await ModeratorState.moderator.set()


# moderator enters price
@dp.message_handler(state=ModeratorState.edit_mod_amount)
async def edit_mod_amount(message: types.Message, state: FSMContext):
    try:
        if message.text != "↩️Назад ↩️":
            data = await state.get_data()
            item = data.get("item")
            with open("items.json", "r") as file:
                items = json.load(file)
            items[item][1] = message.text
            with open("items.json", "w") as file:
                json.dump(items, file)
            await message.answer("Остаток изменен")
            await ModeratorState.add_mod_amount.set()
        else:
            await message.answer("Меню", reply_markup=moderator_keyboard)
            await ModeratorState.moderator.set()
    except Exception as e:
        if message.text == "↩️Назад ↩️":
            await ModeratorState.moderator.set()
        else:
            await message.answer("Неверная команда")
            await ModeratorState.moderator.set()



@dp.message_handler(state=AdminState.add_item)
async def add_item(message: types.Message, state: FSMContext):
    await message.answer("Введите количество")
    await AdminState.add_amount.set()
    await state.update_data(item_name=message.text)


@dp.message_handler(state=AdminState.add_amount)
async def add_amount(message: types.Message, state: FSMContext):
    data = await state.get_data()
    item_name = data.get("item_name")
    item_info = [message.text, "NULL"]
    with open("items.json", "r") as file:
        items = json.load(file)
    items[item_name] = item_info
    with open("items.json", "w") as file:
        json.dump(items, file)
    await message.answer("Товар добавлена")
    await AdminState.admin.set()


@dp.message_handler(state=AdminState.delete_item)
async def delete_item(message: types.Message, state: FSMContext):
    with open("items.json", "r") as file:
        items = json.load(file)
    items.pop(message.text)
    with open("items.json", "w") as file:
        json.dump(items, file)
    await message.answer("Товар удалена")
    await AdminState.admin.set()


if __name__ == '__main__':
    executor.start_polling(dp)
