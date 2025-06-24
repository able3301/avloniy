import logging
import openpyxl
import os
import re
from telegram import Update, KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, filters,
    ConversationHandler, ContextTypes
)

# Logni sozlash
logging.basicConfig(level=logging.INFO)

# Token va admin ID
BOT_TOKEN = "7586148058:AAEa8tfucoM5fBaYXwUQNpmBflkkdgaFFcY"
ADMIN_ID = 1350513135
FILE_NAME = "users_data.xlsx"

# Bosqichlar
COURSE, PHONE, AGE, DAY, TIME = range(5)

# Kurslar
COURSES = [
    "🟢 START POINT", "🤖 ROBOTICS", "⚙️ CHALLENGE LAB",
    "✈️ FLIGHT ACADEMY", "🧪 SCINECE LAB", "🏗️ ENGINEERING LAB",
    "💻 CODING ROOM", "🎮 VR ROOM", "🔧 VEX V5- IQ ROOM"
]

ALLOWED_DAYS = ["Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba", "Yakshanba"]

def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kun", "Vaqt"])
        wb.save(FILE_NAME)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    init_excel()
    keyboard = [[KeyboardButton(course)] for course in COURSES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("Quyidagi kurslardan birini tanlang:", reply_markup=reply_markup)
    return COURSE

async def course_chosen(update: Update, context: ContextTypes.DEFAULT_TYPE):
    course = update.message.text
    if course not in COURSES:
        await update.message.reply_text("Iltimos, menyudan kursni tanlang.")
        return COURSE
    context.user_data["course"] = course
    button = KeyboardButton("📞 Telefon raqamni yuborish", request_contact=True)
    reply_markup = ReplyKeyboardMarkup([[button]], resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("Telefon raqamingizni yuboring:", reply_markup=reply_markup)
    return PHONE

async def phone_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    contact = update.message.contact
    if not contact:
        await update.message.reply_text("Faqat kontakt tugmasidan foydalaning.")
        return PHONE
    context.user_data["phone"] = contact.phone_number
    context.user_data["name"] = update.effective_user.full_name
    await update.message.reply_text("Yoshingizni kiriting:", reply_markup=ReplyKeyboardRemove())
    return AGE

async def age_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    age = update.message.text.strip()
    if not age.isdigit():
        await update.message.reply_text("Faqat raqam kiriting.")
        return AGE
    context.user_data["age"] = age
    await update.message.reply_text("Qaysi kunlari qatnasha olasiz? (Dushanba, Seshanba, ...)")
    return DAY

async def day_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    input_days = update.message.text.strip().capitalize()
    days_list = [d.strip().capitalize() for d in input_days.split(",") if d.strip()]
    
    if not days_list:
        await update.message.reply_text("Iltimos, hech bo‘lmasa 1 ta kun kiriting.")
        return DAY

    if any(day not in ALLOWED_DAYS for day in days_list):
        await update.message.reply_text("Faqat quyidagi kunlardan foydalaning:\n" + ", ".join(ALLOWED_DAYS))
        return DAY

    if len(days_list) > 7:
        await update.message.reply_text("Eng ko‘pi bilan 7 ta kun kiriting.")
        return DAY

    context.user_data["day"] = ", ".join(days_list)
    await update.message.reply_text("Qaysi soatlarda qatnasha olasiz? (Masalan: 14:00 - 16:00)")
    return TIME


async def time_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    time_text = update.message.text.strip()
    pattern = r"^(?:[01]\d|2[0-3]):[0-5]\d\s*-\s*(?:[01]\d|2[0-3]):[0-5]\d$"
    if not re.match(pattern, time_text):
        await update.message.reply_text("Iltimos, soatni quyidagicha kiriting: 14:00 - 16:00")
        return TIME
    context.user_data["time"] = time_text

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active
    ws.append([
        context.user_data["name"],
        context.user_data["phone"],
        context.user_data["age"],
        context.user_data["course"],
        context.user_data["day"],
        context.user_data["time"]
    ])
    wb.save(FILE_NAME)
    await update.message.reply_text("Ma’lumotlaringiz saqlandi. Rahmat!")
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Bekor qilindi.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# === ADMIN BUYRUQLARI ===

async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id == ADMIN_ID:
        await update.message.reply_document(open(FILE_NAME, "rb"))
    else:
        await update.message.reply_text("Siz admin emassiz!")

async def list_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("Siz admin emassiz!")
        return
    if not os.path.exists(FILE_NAME):
        await update.message.reply_text("Hozircha foydalanuvchilar yo‘q.")
        return
    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))[1:]
    if not rows:
        await update.message.reply_text("Hech kim ro‘yxatdan o‘tmagan.")
        return
    text = "\n\n".join([f"{i+1}. {r[0]} | {r[1]} | {r[2]} yosh | {r[3]} | {r[4]} | {r[5]}" for i, r in enumerate(rows)])
    await update.message.reply_text(text)

async def clear_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("Siz admin emassiz!")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kun", "Vaqt"])
    wb.save(FILE_NAME)
    await update.message.reply_text("Barcha ma’lumotlar o‘chirildi!")

# === ===

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logging.error(f"Xato: {context.error}")

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            COURSE: [MessageHandler(filters.TEXT & ~filters.COMMAND, course_chosen)],
            PHONE: [MessageHandler(filters.CONTACT, phone_received)],
            AGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, age_received)],
            DAY: [MessageHandler(filters.TEXT & ~filters.COMMAND, day_received)],
            TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, time_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv_handler)

    # Admin buyruqlar
    app.add_handler(CommandHandler("export", export))
    app.add_handler(CommandHandler("list", list_users))
    app.add_handler(CommandHandler("clear", clear_data))

    app.add_error_handler(error_handler)
    app.run_polling()

if __name__ == "__main__":
    main()
