import asyncio
from aiogram import Bot, Dispatcher, types, F
from aiogram.types import (
    Message, CallbackQuery, ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton, InputFile
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.state import StatesGroup, State
from aiogram.filters import Command, CommandStart
from config import TOKEN, ADMIN_ID
import openpyxl
from datetime import datetime
import os

bot = Bot(token=TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# Holatlar (FSM)
class Form(StatesGroup):
    waiting_for_contact = State()
    waiting_for_age = State()
    waiting_for_day = State()
    waiting_for_time = State()

# Kurslar ro‘yxati
courses = [
    "🚀 START POINT", "🤖 ROBOTICS", "🧪 CHALLENGE LAB",
    "✈️ FLIGHT ACADEMY", "🔬 SCIENCE LAB", "🔧 ENGINEERING LAB",
    "💻 CODING ROOM", "🕶️ VR ROOM", "⚙️ VEX V5 – IQ ROOM"
]

user_data = {}

# Kurslar klaviaturasi
def course_keyboard():
    kb = InlineKeyboardMarkup(row_width=3)
    for c in courses:
        kb.add(InlineKeyboardButton(text=c, callback_data=c))
    return kb

# START
@dp.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await message.answer("Quyidagi kurslardan birini tanlang:", reply_markup=course_keyboard())

# Kurs tanlandi
@dp.callback_query(F.data.in_(courses))
async def course_chosen(callback: CallbackQuery, state: FSMContext):
    await state.update_data(course=callback.data)
    contact_kb = ReplyKeyboardMarkup(
        resize_keyboard=True, one_time_keyboard=True,
        keyboard=[[KeyboardButton(text="📞 Kontakt yuborish", request_contact=True)]]
    )
    await callback.message.answer("📞 Telefon raqamingizni yuboring:", reply_markup=contact_kb)
    await state.set_state(Form.waiting_for_contact)

# Kontakt qabul qilish
@dp.message(F.contact, Form.waiting_for_contact)
async def contact_received(message: Message, state: FSMContext):
    await state.update_data(phone=message.contact.phone_number, name=message.from_user.full_name)
    await message.answer("🧒 Yoshingizni kiriting:")
    await state.set_state(Form.waiting_for_age)

# Yoshni olish
@dp.message(Form.waiting_for_age)
async def age_received(message: Message, state: FSMContext):
    try:
        age = int(message.text)
        if age < 4 or age > 99:
            raise ValueError
        await state.update_data(age=age)

        # Kunlarni tanlash
        days_kb = InlineKeyboardMarkup(row_width=3).add(
            *[InlineKeyboardButton(text=day, callback_data=day) for day in
              ["Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba", "Yakshanba"]]
        )
        await message.answer("📆 Qaysi kun qatnashasiz?", reply_markup=days_kb)
        await state.set_state(Form.waiting_for_day)
    except:
        await message.answer("❗ Iltimos, yoshni to‘g‘ri raqam bilan kiriting.")

# Kunni tanlash
@dp.callback_query(Form.waiting_for_day)
async def day_chosen(callback: CallbackQuery, state: FSMContext):
    await state.update_data(day=callback.data)
    times_kb = InlineKeyboardMarkup(row_width=3).add(
        *[InlineKeyboardButton(text=t, callback_data=t) for t in ["09:00", "11:00", "14:00", "16:00"]]
    )
    await callback.message.answer("🕒 Qaysi vaqtda qatnashasiz?", reply_markup=times_kb)
    await state.set_state(Form.waiting_for_time)

# Soatni tanlash va faylga yozish
@dp.callback_query(Form.waiting_for_time)
async def time_chosen(callback: CallbackQuery, state: FSMContext):
    await state.update_data(time=callback.data)
    data = await state.get_data()
    data["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Excelga yozish
    path = "data.xlsx"
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["ID", "Ism", "Telefon", "Yosh", "Kurs", "Kun", "Soat", "Vaqti"])

    ws.append([
        callback.from_user.id,
        data["name"],
        data["phone"],
        data["age"],
        data["course"],
        data["day"],
        data["time"],
        data["timestamp"]
    ])
    wb.save(path)

    await callback.message.answer("✅ Muvaffaqiyatli ro‘yxatdan o‘tdingiz! Rahmat!")
    await state.clear()

# Admin eksport
@dp.message(Command("export"))
async def export_excel(message: Message):
    if message.from_user.id == ADMIN_ID:
        if os.path.exists("data.xlsx"):
            await message.answer_document(InputFile("data.xlsx"))
        else:
            await message.answer("📂 Fayl hali yaratilmagan.")
    else:
        await message.answer("🚫 Bu buyruq faqat admin uchun!")

# Run
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
