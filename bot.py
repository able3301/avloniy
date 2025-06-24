import logging
from telegram import Update, KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, filters,
    ContextTypes, ConversationHandler
)
import openpyxl
import os

# Log sozlamalari
logging.basicConfig(level=logging.INFO)

# Bosqichlar
COURSE, PHONE, AGE, DAY, TIME = range(5)

# Excel fayl nomi
FILE_NAME = "users_data.xlsx"

# Kurslar ro'yxati
COURSES = [
    "🟢 START POINT", "🤖 ROBOTICS", "⚙️ CHALLENGE LAB",
    "✈️ FLIGHT ACADEMY", "🧪 SCIENCE LAB", "🏗️ ENGINEERING LAB",
    "💻 CODING ROOM", "🎮 VR ROOM", "🔧 VEX V5- IQ ROOM",
    "❌ Kurs tanlashni yakunlash"
]

# Kun va vaqt tugmalari
DAYS = ["Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba", "Yakshanba"]
TIMES = ["9:00 - 11:00", "11:00 - 13:00", "14:00 - 16:00", "16:00 - 18:00"]

# Admin ID
ADMIN_ID = "1350513135"  # o'zingizniki bilan almashtiring

# Excel yaratish
def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kun", "Vaqt"])
        wb.save(FILE_NAME)

# /start komandasi
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    init_excel()
    greeting = (
        "Assalomu alaykum! 👋\n\n"
        "Sizga \"Abdulla Avloniy nomidagi pedogogik mahorat milliy instituti\" STEAM markazi tomonidan tashkil etilgan innovatsion kurslar bo‘yicha ro‘yxatdan o‘tish uchun bir nechta savollar beriladi.\n\n"
        "📌 Bu markaz zamonaviy laboratoriyalar, ilg‘or texnologiyalar va amaliy loyihalar asosida ta’lim beradi. Har bir yo‘nalish o‘quvchilarning bilim olishiga, ixtirochilik salohiyatini oshirishga qaratilgan.\n\n"
        "Iltimos, quyidagi yo‘nalishlardan birini tanlang:"
    )
    keyboard = [[KeyboardButton(course)] for course in COURSES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text(greeting, reply_markup=reply_markup)
    return COURSE

# Kurs tanlandi
async def course_chosen(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text == "❌ Kurs tanlashni yakunlash":
        await update.message.reply_text("Ro'yxatdan o'tish yakunlandi. Rahmat!", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END
    if text not in COURSES:
        await update.message.reply_text("Iltimos, menyudan kursni tanlang.")
        return COURSE
    context.user_data["course"] = text
    contact_button = KeyboardButton("📞 Telefon raqamni yuborish", request_contact=True)
    reply_markup = ReplyKeyboardMarkup([[contact_button]], resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("Iltimos, telefon raqamingizni yuboring:", reply_markup=reply_markup)
    return PHONE

# Telefon
async def phone_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    contact = update.message.contact
    if not contact:
        await update.message.reply_text("Faqat kontakt tugmasidan foydalaning.")
        return PHONE
    context.user_data["phone"] = contact.phone_number
    context.user_data["name"] = update.effective_user.full_name
    await update.message.reply_text("Yoshingizni kiriting:", reply_markup=ReplyKeyboardRemove())
    return AGE

# Yosh
async def age_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    age = update.message.text
    if not age.isdigit():
        await update.message.reply_text("Faqat raqam kiriting. Iltimos, yoshingizni qaytadan kiriting:")
        return AGE
    context.user_data["age"] = age
    day_buttons = [[KeyboardButton(day)] for day in DAYS]
    reply_markup = ReplyKeyboardMarkup(day_buttons, resize_keyboard=True)
    await update.message.reply_text("Qaysi kunlari qatnasha olasiz?", reply_markup=reply_markup)
    return DAY

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


# Vaqt
async def time_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    time = update.message.text
    if time not in TIMES:
        await update.message.reply_text("Faqat tugmadan tanlang.")
        return TIME
    context.user_data["time"] = time
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
    await update.message.reply_text("Ma'lumotlar saqlandi! Yana bir kurs tanlashni istaysizmi?", reply_markup=ReplyKeyboardMarkup(
        [[KeyboardButton(course)] for course in COURSES], resize_keyboard=True
    ))
    return COURSE

# Bekor qilish
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Bekor qilindi.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# Admin /export
async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.effective_user.id) == ADMIN_ID:
        await update.message.reply_document(open(FILE_NAME, "rb"))
    else:
        await update.message.reply_text("Siz admin emassiz!")

# Admin /list
async def list_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.effective_user.id) == ADMIN_ID:
        wb = openpyxl.load_workbook(FILE_NAME)
        ws = wb.active
        users = ""
        for row in ws.iter_rows(min_row=2, values_only=True):
            users += ", ".join(map(str, row)) + "\n"
        await update.message.reply_text(users or "Hozircha foydalanuvchilar mavjud emas.")
    else:
        await update.message.reply_text("Siz admin emassiz!")

# Admin /clear
async def clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.effective_user.id) == ADMIN_ID:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kun", "Vaqt"])
        wb.save(FILE_NAME)
        await update.message.reply_text("Barcha ma'lumotlar o'chirildi.")
    else:
        await update.message.reply_text("Siz admin emassiz!")

# Xatolar
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logging.error(msg="Xatolik yuz berdi:", exc_info=context.error)

# Botni ishga tushurish
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
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("export", export))
    app.add_handler(CommandHandler("list", list_users))
    app.add_handler(CommandHandler("clear", clear))
    app.add_error_handler(error_handler)

    app.run_polling()

if __name__ == "__main__":
    main()
