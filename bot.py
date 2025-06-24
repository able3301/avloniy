import logging
from telegram import Update, KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, filters,
    ContextTypes, ConversationHandler
)
import openpyxl
import os

# Loglar
logging.basicConfig(level=logging.INFO)

# Bosqichlar
COURSE, PHONE, AGE, DAY, TIME = range(5)

# Fayl nomi
FILE_NAME = "users_data.xlsx"

# Admin ID (raqam shaklida yoziladi)
ADMIN_ID = 1350513135

# Kurslar
COURSES = [
    "🟢 START POINT", "🤖 ROBOTICS", "⚙️ CHALLENGE LAB",
    "✈️ FLIGHT ACADEMY", "🧪 SCINECE LAB", "🏗️ ENGINEERING LAB",
    "💻 CODING ROOM", "🎮 VR ROOM", "🔧 VEX V5- IQ ROOM"
]

# Excel fayl tayyorlash
def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kun", "Vaqt"])
        wb.save(FILE_NAME)

# /start komandasi
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    init_excel()
    keyboard = [[KeyboardButton(course)] for course in COURSES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("Quyidagi kurslardan birini tanlang:", reply_markup=reply_markup)
    return COURSE

# Kurs tanlandi
async def course_chosen(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text not in COURSES:
        await update.message.reply_text("Iltimos, menyudan kursni tanlang.")
        return COURSE
    context.user_data["course"] = text
    contact_button = KeyboardButton("📞 Telefon raqamni yuborish", request_contact=True)
    reply_markup = ReplyKeyboardMarkup([[contact_button]], resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("Iltimos, telefon raqamingizni yuboring:", reply_markup=reply_markup)
    return PHONE

# Telefon raqam qabul qilindi
async def phone_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    contact = update.message.contact
    if not contact:
        await update.message.reply_text("Faqat kontakt tugmasidan foydalaning.")
        return PHONE
    context.user_data["phone"] = contact.phone_number
    context.user_data["name"] = update.effective_user.full_name
    await update.message.reply_text("Yoshingizni kiriting:", reply_markup=ReplyKeyboardRemove())
    return AGE

# Yosh qabul qilindi
async def age_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    age = update.message.text
    if not age.isdigit():
        await update.message.reply_text("Faqat raqam kiriting. Iltimos, yoshingizni qaytadan kiriting:")
        return AGE
    context.user_data["age"] = age
    await update.message.reply_text("Qaysi kunlari qatnasha olasiz?")
    return DAY

# Kun qabul qilindi
async def day_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["day"] = update.message.text
    await update.message.reply_text("Qaysi soatlarda qatnasha olasiz?")
    return TIME

# Soat qabul qilindi va Excelga yoziladi
async def time_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["time"] = update.message.text
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
    await update.message.reply_text("Ma'lumotlar saqlandi. Rahmat!", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# /cancel komandasi
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Bekor qilindi.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# /export komandasi - faqat admin
async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id == ADMIN_ID:
        if os.path.exists(FILE_NAME):
            await update.message.reply_document(open(FILE_NAME, "rb"))
        else:
            await update.message.reply_text("Fayl mavjud emas.")
    else:
        await update.message.reply_text("Siz admin emassiz!")

# /clear komandasi - Excelni tozalash
async def clear_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id == ADMIN_ID:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kun", "Vaqt"])
        wb.save(FILE_NAME)
        await update.message.reply_text("Barcha ma’lumotlar tozalandi.")
    else:
        await update.message.reply_text("Siz admin emassiz!")

# /users komandasi - ro'yxatni matn sifatida ko'rsatish
async def show_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id == ADMIN_ID:
        if not os.path.exists(FILE_NAME):
            await update.message.reply_text("Fayl mavjud emas.")
            return
        wb = openpyxl.load_workbook(FILE_NAME)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        if not rows:
            await update.message.reply_text("Hozircha foydalanuvchi yo‘q.")
            return
        message = "📋 Ro'yxatdan o'tgan foydalanuvchilar:\n\n"
        for i, row in enumerate(rows, 1):
            message += f"{i}. {row[0]} | {row[1]} | {row[2]} yosh | {row[3]} | {row[4]} | {row[5]}\n"
        await update.message.reply_text(message)
    else:
        await update.message.reply_text("Siz admin emassiz!")

# Xatolarni ushlash
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    print(f"Xato: {context.error}")

# Asosiy funksiya
def main():
    app = ApplicationBuilder().token("YOUR_BOT_TOKEN").build()

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
    app.add_handler(CommandHandler("export", export_excel))
    app.add_handler(CommandHandler("clear", clear_excel))
    app.add_handler(CommandHandler("users", show_users))
    app.add_error_handler(error_handler)
    app.run_polling()

if __name__ == "__main__":
    main()
