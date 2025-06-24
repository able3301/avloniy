import logging
import os
import openpyxl
from telegram import (
    Update, KeyboardButton, ReplyKeyboardMarkup,
    ReplyKeyboardRemove, ReplyKeyboardMarkup
)
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)

# Logger
logging.basicConfig(level=logging.INFO)

# Fayl nomi
FILE_NAME = "users.xlsx"

# Admin Telegram ID
ADMIN_ID = 1350513135  # o'zingizning admin ID'ingizni kiriting

# Bosqichlar
COURSE, PHONE, AGE, DAY, TIME, CONFIRM_DAY = range(6)

# Kurslar ro'yxati
COURSES = [
    "🟢 START POINT", "🤖 ROBOTICS", "⚙️ CHALLENGE LAB",
    "✈️ FLIGHT ACADEMY", "🧪 SCINECE LAB", "🏗️ ENGINEERING LAB",
    "💻 CODING ROOM", "🎮 VR ROOM", "🔧 VEX V5- IQ ROOM", "❌ Yopish"
]

# Kunlar
DAYS = ["Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba", "Yakshanba"]

# Soatlar
TIMES = ["09:00-11:00", "11:00-13:00", "14:00-16:00", "16:00-18:00"]

# Excel tayyorlash
def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kunlar", "Vaqt"])
        wb.save(FILE_NAME)

# /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    init_excel()
    
    welcome_text = (
        "👋 *Assalomu alaykum!*\n\n"
        "Sizga _\"Abdulla Avloniy nomidagi Pedagogik Mahorat Milliy Instituti\"_ STEAM markazi tomonidan tashkil etilgan "
        "*innovatsion kurslar* bo‘yicha ro‘yxatdan o‘tish uchun bir nechta savollar beriladi.\n\n"
        "📌 *Bu markaz* zamonaviy laboratoriyalar, ilg‘or texnologiyalar va amaliy loyihalar asosida ta’lim beradi. "
        "Har bir yo‘nalish o‘quvchilarning bilim olishiga, *ixtirochilik salohiyatini* oshirishga qaratilgan.\n\n"
        "🧭 *Iltimos, quyidagi yo‘nalishlardan birini tanlang:*"
    )
    
    keyboard = [[KeyboardButton(course)] for course in COURSES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    
    await update.message.reply_text(welcome_text, parse_mode="Markdown", reply_markup=reply_markup)
    return COURSE

# Kurs tanlash
async def course_chosen(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text == "❌ Yopish":
        await update.message.reply_text("✅ Ro‘yxatdan o‘tish yakunlandi. Botdan foydalanishingiz mumkin.")
        return ConversationHandler.END

    if text not in COURSES:
        await update.message.reply_text("Iltimos, menyudan kursni tanlang.")
        return COURSE

    context.user_data["course"] = text
    # Telefonni so‘rash
    contact_button = KeyboardButton("📞 Telefon raqamni yuborish", request_contact=True)
    reply_markup = ReplyKeyboardMarkup([[contact_button]], resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("📱 Iltimos, telefon raqamingizni yuboring:", reply_markup=reply_markup)
    return PHONE


# Telefon qabul qilish
async def phone_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    contact = update.message.contact
    if not contact:
        await update.message.reply_text("Iltimos, telefon raqam tugmasidan foydalaning.")
        return PHONE
    context.user_data["phone"] = contact.phone_number
    context.user_data["name"] = update.effective_user.full_name
    await update.message.reply_text("Yoshingizni kiriting:", reply_markup=ReplyKeyboardRemove())
    return AGE

# Yosh
async def age_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    age = update.message.text
    if not age.isdigit():
        await update.message.reply_text("Faqat raqam kiriting. Yoshingizni qayta kiriting:")
        return AGE
    context.user_data["age"] = age
    context.user_data["days"] = []
    return await send_day_keyboard(update, context)

# Kunlar uchun klaviatura
async def send_day_keyboard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    selected = context.user_data.get("days", [])
    keyboard = []
    for day in DAYS:
        if day in selected:
            keyboard.append([KeyboardButton(f"✅ {day}")])
        else:
            keyboard.append([KeyboardButton(day)])
    keyboard.append([KeyboardButton("✅ Tayyor")])
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("Qaysi kunlarda qatnasha olasiz? Bir nechta tanlang:", reply_markup=reply_markup)
    return CONFIRM_DAY

# Kun tanlash bosqichi
async def day_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.replace("✅", "").strip()
    if text == "Tayyor":
        if not context.user_data["days"]:
            await update.message.reply_text("Iltimos, kamida bitta kun tanlang.")
            return await send_day_keyboard(update, context)
        return await ask_time(update, context)
    if text in DAYS:
        selected = context.user_data["days"]
        if text in selected:
            selected.remove(text)
        else:
            if len(selected) < 7:
                selected.append(text)
            else:
                await update.message.reply_text("Eng ko‘pi bilan 7 ta kun tanlashingiz mumkin.")
        return await send_day_keyboard(update, context)
    else:
        await update.message.reply_text("Iltimos, menyudan tanlang.")
        return await send_day_keyboard(update, context)

# Soat tanlash
async def ask_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [[KeyboardButton(time)] for time in TIMES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("Qaysi soatlarda qatnasha olasiz?", reply_markup=reply_markup)
    return TIME

# Soat qabul qilish va saqlash
async def time_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["time"] = update.message.text

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

    # Yana boshqa kurs tanlash yoki tugatish menyusi
    keyboard = [[KeyboardButton(course)] for course in COURSES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text(
        "✅ Ma'lumotlar saqlandi. Agar boshqa kursga ham qatnashmoqchi bo‘lsangiz, tanlang.\n\nAgar ro‘yxatdan o‘tishni yakunlamoqchi bo‘lsangiz, '❌ Yopish' tugmasini bosing:",
        reply_markup=reply_markup
    )

    return COURSE


# Bekor qilish
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Bekor qilindi.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# Admin: Excel fayl jo'natish
async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("Siz admin emassiz.")
        return
    await update.message.reply_document(open(FILE_NAME, "rb"))

# Admin: Ro‘yxatni matn ko‘rinishda olish
async def list_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("Siz admin emassiz.")
        return
    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active
    text = ""
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        text += f"{i}. {row[0]} - {row[3]} - {row[4]} - {row[5]}\n"
    await update.message.reply_text(text or "Foydalanuvchilar mavjud emas.")

# Admin: Tozalash
async def clear_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("Siz admin emassiz.")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kunlar", "Vaqt"])
    wb.save(FILE_NAME)
    await update.message.reply_text("Barcha ma’lumotlar o‘chirildi.")

# Botni ishga tushurish
def main():
    app = ApplicationBuilder().token("7586148058:AAEa8tfucoM5fBaYXwUQNpmBflkkdgaFFcY").build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            COURSE: [MessageHandler(filters.TEXT & ~filters.COMMAND, course_chosen)],
            PHONE: [MessageHandler(filters.CONTACT, phone_received)],
            AGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, age_received)],
            CONFIRM_DAY: [MessageHandler(filters.TEXT & ~filters.COMMAND, day_selection)],
            TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, time_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv_handler)
    app.add_handler(CommandHandler("file", export))
    app.add_handler(CommandHandler("list", list_users))
    app.add_handler(CommandHandler("clear", clear_data))
    app.run_polling()

if __name__ == "__main__":
    main()
