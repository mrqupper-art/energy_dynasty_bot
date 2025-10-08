from aiogram import Bot, Dispatcher, types
from aiogram.utils import executor
import pandas as pd
import os
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.contrib.fsm_storage.memory import MemoryStorage

TOKEN = "7994739525:AAGJWzvn-oLVUOWheeh4RPbf_-mk4PD6zqQ"

bot = Bot(token=TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

FILE_NAME = "orders.xlsx"
if not os.path.exists(FILE_NAME):
    df = pd.DataFrame(columns=["Имя", "Телефон", "Адрес", "Описание"])
    df.to_excel(FILE_NAME, index=False)

class OrderState(StatesGroup):
    name = State()
    phone = State()
    address = State()
    description = State()

@dp.message_handler(commands=["start"])
async def start(message: types.Message):
    await message.answer(
        "👋 Привет! Я бот компании *ENERGY DYNASTY* ⚡\n"
        "Мы выполняем электромонтажные работы для квартир и домов.\n\n"
        "Чтобы оставить заявку, напиши команду: /order",
        parse_mode="Markdown"
    )

@dp.message_handler(commands=["order"])
async def order_start(message: types.Message):
    await message.answer("Введите своё имя:")
    await OrderState.name.set()

@dp.message_handler(state=OrderState.name)
async def process_name(message: types.Message, state: FSMContext):
    await state.update_data(name=message.text)
    await message.answer("Введите номер телефона 📞:")
    await OrderState.next()

@dp.message_handler(state=OrderState.phone)
async def process_phone(message: types.Message, state: FSMContext):
    await state.update_data(phone=message.text)
    await message.answer("Введите адрес 🏠:")
    await OrderState.next()

@dp.message_handler(state=OrderState.address)
async def process_address(message: types.Message, state: FSMContext):
    await state.update_data(address=message.text)
    await message.answer("Опишите, какие работы нужно выполнить ⚙️:")
    await OrderState.next()

@dp.message_handler(state=OrderState.description)
async def process_description(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    user_data["description"] = message.text

    # Сохраняем в Excel
    df = pd.read_excel(FILE_NAME)
    df = pd.concat([df, pd.DataFrame([user_data])], ignore_index=True)
    df.to_excel(FILE_NAME, index=False)

    await message.answer("✅ Спасибо! Ваша заявка принята. Мы свяжемся с вами в ближайшее время.")
    await state.finish()

if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)
