import logging
from telegram import Update, KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)
import openpyxl
import os

# Log sozlamalari
logging.basicConfig(level=logging.INFO)

# Bosqichlar
COURSE, PHONE, AGE, DAY, TIME = range(5)

# Excel fayl nomi
FILE_NAME = "users_data.xlsx"

# Kurslar ro'yxati + yopish
COURSES = [
    "🟢 START POINT", "🤖 ROBOTICS", "⚙️ CHALLENGE LAB",
    "✈️ FLIGHT ACADEMY", "🧪 SCIENCE LAB", "🏗️ ENGINEERING LAB",
    "💻 CODING ROOM", "🎮 VR ROOM", "🔧 VEX V5- IQ ROOM",
    "❌ Yopish"
]

# KUNLAR VA VAQT variantlari
DAYS = ["Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba", "Yakshanba"]
TIMES = ["09:00-11:00", "11:00-13:00", "14:00-16:00", "16:00-18:00"]

# Admin ID
ADMIN_ID = 1350513135  # O'zingizga moslashtiring

# Excel faylini yaratish
def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kun", "Vaqt"])
        wb.save(FILE_NAME)

# /start komutasi
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    init_excel()
    matn = (
        "Assalomu alaykum! 👋\n\n"
        "Sizga \"Abdulla Avloniy nomidagi pedogogik mahorat milliy instituti\" STEAM markazi tomonidan tashkil etilgan innovatsion kurslar bo‘yicha ro‘yxatdan o‘tish uchun bir nechta savollar beriladi.\n\n"
        "\uD83D\uDCCC Bu markaz zamonaviy laboratoriyalar, ilg‘or texnologiyalar va amaliy loyihalar asosida ta’lim beradi. Har bir yo‘nalish o‘quvchilarning bilim olishiga, ixtirochilik salohiyatini oshirishga qaratilgan.\n\n"
        "Iltimos, quyidagi yo‘nalishlardan birini tanlang:"
    )
    keyboard = [[KeyboardButton(course)] for course in COURSES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text(matn, reply_markup=reply_markup)
    return COURSE

# Kurs tanlandi
async def course_chosen(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text == "❌ Yopish":
        await update.message.reply_text("✅ Ro‘yxatdan o‘tish yakunlandi. Botdan foydalanishingiz mumkin.")
        return ConversationHandler.END
    if text not in COURSES:
        await update.message.reply_text("Iltimos, menyudan kursni tanlang.")
        return COURSE
    context.user_data["course"] = text
    contact_button = KeyboardButton("📞 Telefon raqamni yuborish", request_contact=True)
    reply_markup = ReplyKeyboardMarkup([[contact_button]], resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("📱 Iltimos, telefon raqamingizni yuboring:", reply_markup=reply_markup)
    return PHONE

# Telefon qabul qilindi
async def phone_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    contact = update.message.contact
    if not contact:
        await update.message.reply_text("Faqat kontakt tugmasidan foydalaning.")
        return PHONE
    context.user_data["phone"] = contact.phone_number
    context.user_data["name"] = update.effective_user.full_name
    await update.message.reply_text("🎂 Yoshingizni kiriting:", reply_markup=ReplyKeyboardRemove())
    return AGE

# Yosh qabul qilindi
async def age_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    age = update.message.text
    if not age.isdigit():
        await update.message.reply_text("Faqat raqam kiriting. Iltimos, yoshingizni qaytadan kiriting:")
        return AGE
    context.user_data["age"] = age
    keyboard = [[KeyboardButton(day)] for day in DAYS]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("📅 Qaysi kunlari qatnasha olasiz?", reply_markup=reply_markup)
    return DAY

# Kun qabul qilindi
async def day_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day = update.message.text
    if day not in DAYS:
        await update.message.reply_text("Faqat menyudan kunni tanlang!")
        return DAY
    context.user_data["day"] = day
    keyboard = [[KeyboardButton(t)] for t in TIMES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("🕒 Qaysi soatlarda qatnasha olasiz?", reply_markup=reply_markup)
    return TIME

# Soat qabul qilindi
async def time_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    time = update.message.text
    if time not in TIMES:
        await update.message.reply_text("Faqat menyudan vaqtni tanlang!")
        return TIME
    context.user_data["time"] = time

    # Excelga yozish
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

    # Yana kurs tanlash
    keyboard = [[KeyboardButton(course)] for course in COURSES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text(
        "✅ Ma'lumotlar saqlandi. Agar boshqa kursga ham qatnashmoqchi bo‘lsangiz, tanlang.\n\nAgar ro‘yxatdan o‘tishni yakunlamoqchi bo‘lsangiz, '❌ Yopish' tugmasini bosing:",
        reply_markup=reply_markup
    )
    return COURSE

# /export - Excel faylini admin uchun
async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id == ADMIN_ID:
        await update.message.reply_document(open(FILE_NAME, "rb"))
    else:
        await update.message.reply_text("❌ Siz admin emassiz!")

# /list - foydalanuvchilarni ko'rish (admin)
async def list_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Siz admin emassiz!")
        return
    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))[1:]  # sarlavhasiz
    if not rows:
        await update.message.reply_text("Foydalanuvchilar yo‘q.")
        return
    msg = "\n".join([f"{r[0]} - {r[3]} - {r[4]} {r[5]}" for r in rows])
    await update.message.reply_text(f"📋 Ro'yxat:\n{msg}")

# /clear - ma’lumotlarni tozalash (admin)
async def clear_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Siz admin emassiz!")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kun", "Vaqt"])
    wb.save(FILE_NAME)
    await update.message.reply_text("✅ Barcha ma’lumotlar tozalandi.")

# Bekor qilish
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ Bekor qilindi.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# Xatolik
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    print(f"Xato: {context.error}")

# Ishga tushirish
def main():
    app = ApplicationBuilder().token("7586148058:AAEa8tfucoM5fBaYXwUQNpmBflkkdgaFFcY").build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            COURSE: [MessageHandler(filters.TEXT & ~filters.COMMAND, course_chosen)],
            PHONE: [MessageHandler(filters.CONTACT, phone_received)],
            AGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, age_received)],
            DAY: [MessageHandler(filters.TEXT & ~filters.COMMAND, day_received)],
            TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, time_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)]
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("export", export))
    app.add_handler(CommandHandler("list", list_users))
    app.add_handler(CommandHandler("clear", clear_data))
    app.add_error_handler(error_handler)
    app.run_polling()

if __name__ == "__main__":
    main()
