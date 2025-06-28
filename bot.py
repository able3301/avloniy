import logging
import os
import openpyxl
from telegram import (
    Update, KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove
)
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)

logging.basicConfig(level=logging.INFO)
FILE_NAME = "users.xlsx"
ADMIN_ID = 1350513135
COURSE, PHONE, AGE, DAY, TIME, CONFIRM_DAY = range(6)

COURSES = [
    "🟢 START POINT", "🤖 ROBOTICS", "⚙️ CHALLENGE LAB",
    "✈️ FLIGHT ACADEMY", "🧪 SCINECE LAB", "🏗️ ENGINEERING LAB",
    "💻 CODING ROOM", "🎮 VR ROOM", "🔧 VEX V5- IQ ROOM"
]

DAYS = ["Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba", "Yakshanba"]
TIMES = ["09:00-11:00", "11:00-13:00", "14:00-16:00", "16:00-18:00"]

def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kunlar", "Vaqt"])
        wb.save(FILE_NAME)

async def clean_previous_messages(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    old_msg_ids = context.user_data.get("prev_messages", [])
    for msg_id in old_msg_ids:
        try:
            await context.bot.delete_message(chat_id, msg_id)
        except:
            pass
    context.user_data["prev_messages"] = [update.message.message_id]

async def send_reply(update: Update, context: ContextTypes.DEFAULT_TYPE, text, reply_markup=None):
    msg = await update.message.reply_text(text, reply_markup=reply_markup, parse_mode="Markdown")
    context.user_data["prev_messages"].append(msg.message_id)
    return msg

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    init_excel()
    await clean_previous_messages(update, context)
    welcome = (
        "👋 *Assalomu alaykum!*

"
        "Sizga _\"Abdulla Avloniy nomidagi Pedagogik Mahorat Milliy Instituti\"_ STEAM markazi tomonidan tashkil etilgan "
        "*innovatsion kurslar* bo‘yicha ro‘yxatdan o‘tish uchun bir nechta savollar beriladi.\n\n"
        "📌 *Bu markaz* zamonaviy laboratoriyalar, ilg‘or texnologiyalar va amaliy loyihalar asosida ta’lim beradi. "
        "Har bir yo‘nalish o‘quvchilarning bilim olishiga, *ixtirochilik salohiyatini* oshirishga qaratilgan.\n\n"
        "🗭 *Iltimos, quyidagi yo‘nalishlardan birini tanlang:*"
    )
    keyboard = [[KeyboardButton(course)] for course in COURSES]
    markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await send_reply(update, context, welcome, markup)
    return COURSE

async def course_chosen(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await clean_previous_messages(update, context)
    text = update.message.text
    if text not in COURSES:
        await send_reply(update, context, "Iltimos, menyudan kursni tanlang.")
        return COURSE
    context.user_data["course"] = text
    btn = KeyboardButton("📞 Telefon raqamni yuborish", request_contact=True)
    markup = ReplyKeyboardMarkup([[btn]], resize_keyboard=True, one_time_keyboard=True)
    await send_reply(update, context, "Telefon raqamingizni yuboring:", markup)
    return PHONE

async def phone_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await clean_previous_messages(update, context)
    contact = update.message.contact
    if not contact:
        await send_reply(update, context, "Iltimos, tugmadan foydalaning.")
        return PHONE
    context.user_data.update({
        "phone": contact.phone_number,
        "name": update.effective_user.full_name
    })
    await send_reply(update, context, "Yoshingizni kiriting:", ReplyKeyboardRemove())
    return AGE

async def age_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await clean_previous_messages(update, context)
    age = update.message.text
    if not age.isdigit():
        await send_reply(update, context, "Faqat raqam kiriting. Qayta urinib ko'ring:")
        return AGE
    context.user_data["age"] = age
    context.user_data["days"] = []
    return await send_day_keyboard(update, context)

async def send_day_keyboard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    selected = context.user_data.get("days", [])
    keyboard = [[KeyboardButton(f"✅ {d}" if d in selected else d)] for d in DAYS]
    keyboard.append([KeyboardButton("✅ Tayyor")])
    markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await send_reply(update, context, "Qaysi kunlarda qatnasha olasiz?", markup)
    return CONFIRM_DAY

async def day_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await clean_previous_messages(update, context)
    text = update.message.text.replace("✅", "").strip()
    if text == "Tayyor":
        if not context.user_data["days"]:
            await send_reply(update, context, "Kamida bitta kun tanlang.")
            return await send_day_keyboard(update, context)
        return await ask_time(update, context)
    if text in DAYS:
        selected = context.user_data["days"]
        selected.remove(text) if text in selected else selected.append(text)
        return await send_day_keyboard(update, context)
    await send_reply(update, context, "Iltimos, menyudan tanlang.")
    return await send_day_keyboard(update, context)

async def ask_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await clean_previous_messages(update, context)
    keyboard = [[KeyboardButton(t)] for t in TIMES]
    markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
    await send_reply(update, context, "Qaysi soatlarda qatnasha olasiz?", markup)
    return TIME

async def time_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await clean_previous_messages(update, context)
    time = update.message.text
    if time not in TIMES:
        await send_reply(update, context, "Soatlardan birini tanlang.")
        return TIME
    context.user_data["time"] = time
    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active
    ws.append([
        context.user_data["name"],
        context.user_data["phone"],
        context.user_data["age"],
        context.user_data["course"],
        ", ".join(context.user_data["days"]),
        context.user_data["time"]
    ])
    wb.save(FILE_NAME)
    await send_reply(update, context, "Ma'lumotlar saqlandi. Yana /start buyrug'ini yuborishingiz mumkin.", ReplyKeyboardRemove())
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await clean_previous_messages(update, context)
    await send_reply(update, context, "Bekor qilindi.", ReplyKeyboardRemove())
    return ConversationHandler.END

async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await send_reply(update, context, "Siz admin emassiz.")
        return
    await update.message.reply_document(open(FILE_NAME, "rb"))

async def list_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await send_reply(update, context, "Siz admin emassiz.")
        return
    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active
    text = "\n".join([f"{i}. {r[0]} - {r[3]} - {r[4]} - {r[5]}" for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), 1)])
    await send_reply(update, context, text or "Foydalanuvchilar yo'q.")

async def clear_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await send_reply(update, context, "Siz admin emassiz.")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Ism", "Telefon", "Yosh", "Kurs", "Kunlar", "Vaqt"])
    wb.save(FILE_NAME)
    await send_reply(update, context, "Ma'lumotlar o'chirildi.")

def main():
    app = ApplicationBuilder().token("7586148058:AAEa8tfucoM5fBaYXwUQNpmBflkkdgaFFcY").build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            COURSE: [MessageHandler(filters.TEXT & ~filters.COMMAND, course_chosen)],
            PHONE: [MessageHandler(filters.CONTACT, phone_received)],
            AGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, age_received)],
            CONFIRM_DAY: [MessageHandler(filters.TEXT & ~filters.COMMAND, day_selection)],
            TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, time_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)]
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("file", export))
    app.add_handler(CommandHandler("list", list_users))
    app.add_handler(CommandHandler("clear", clear_data))
    app.run_polling()

if __name__ == "__main__":
    main()
