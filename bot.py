import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.types import KeyboardButton, ReplyKeyboardMarkup
from aiogram.enums import ParseMode
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
import openpyxl
from datetime import datetime
import os

TOKEN = os.getenv("BOT_TOKEN")

bot = Bot(token=TOKEN, parse_mode=ParseMode.HTML)
dp = Dispatcher(storage=MemoryStorage())

class Form(StatesGroup):
    waiting_for_contact = State()
    waiting_for_age = State()
    waiting_for_days = State()
    waiting_for_time = State()

@dp.message(commands=['start'])
async def cmd_start(message: types.Message, state: FSMContext):
    keyboard = [
        [KeyboardButton(text="📍 START POINT"), KeyboardButton(text="🤖 ROBOTICS")],
        [KeyboardButton(text="🧪 SCIENCE LAB"), KeyboardButton(text="✈️ FLIGHT ACADEMY")],
        [KeyboardButton(text="🔧 ENGINEERING LAB"), KeyboardButton(text="💻 CODING ROOM")],
        [KeyboardButton(text="🕶️ VR ROOM"), KeyboardButton(text="🤖 VEX V5- IQ ROOM")],
        [KeyboardButton(text="🚀 CHALLENGE LAB")]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)
    await message.answer("Quyidagi kurslardan birini tanlang:", reply_markup=reply_markup)

@dp.message(lambda msg: msg.text.startswith("📍") or msg.text.startswith("🤖") or msg.text.startswith("🧪"))
async def course_chosen(message: types.Message, state: FSMContext):
    await state.update_data(course=message.text)
    contact_btn = KeyboardButton("Kontaktni yuborish", request_contact=True)
    markup = ReplyKeyboardMarkup(resize_keyboard=True, keyboard=[[contact_btn]])
    await message.answer("Iltimos, kontakt raqamingizni yuboring:", reply_markup=markup)
    await state.set_state(Form.waiting_for_contact)

@dp.message(Form.waiting_for_contact, content_types=types.ContentType.CONTACT)
async def process_contact(message: types.Message, state: FSMContext):
    await state.update_data(phone=message.contact.phone_number)
    await message.answer("Yoshingizni kiriting:")
    await state.set_state(Form.waiting_for_age)

@dp.message(Form.waiting_for_age)
async def process_age(message: types.Message, state: FSMContext):
    await state.update_data(age=message.text)
    await message.answer("Hafta kunlaridan sizga qulaylarini kiriting (masalan: Dushanba, Payshanba):")
    await state.set_state(Form.waiting_for_days)

@dp.message(Form.waiting_for_days)
async def process_days(message: types.Message, state: FSMContext):
    await state.update_data(days=message.text)
    await message.answer("Sizga qulay vaqtni kiriting (masalan: 14:00 - 15:00):")
    await state.set_state(Form.waiting_for_time)

@dp.message(Form.waiting_for_time)
async def process_time(message: types.Message, state: FSMContext):
    data = await state.update_data(time=message.text)
    user_data = await state.get_data()
    save_to_excel(user_data)
    await message.answer("Ma'lumotlaringiz qabul qilindi! ✅", reply_markup=types.ReplyKeyboardRemove())
    await state.clear()

def save_to_excel(data):
    filename = "users.xlsx"
    if not os.path.exists(filename):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Sana", "Kurs", "Telefon", "Yosh", "Kunlar", "Vaqt"])
    else:
        wb = openpyxl.load_workbook(filename)
        ws = wb.active

    ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), data["course"], data["phone"], data["age"], data["days"], data["time"]])
    wb.save(filename)

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
