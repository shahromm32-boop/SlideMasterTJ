import telebot
from telebot import types
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO

TOKEN = "8718259141:AAHAGxLvfWq6wvrugpQyFuEfz4qoynRbV9I"

# Маълумоти пардохтро худат пур кун
CARD_NUMBER = "4439200028187457"
CARD_OWNER = "SHAHROM MURODOV"


# Telegram ID-и худатро баъдтар мегузорем
ADMIN_ID = 0

bot = telebot.TeleBot(TOKEN)

orders = {}


@bot.message_handler(commands=["start"])
def start(message):
    bot.send_message(
        message.chat.id,
        "Салом! 👋\n"
        "Ба SlideMasterTJ хуш омадӣ!\n\n"
        "Мавзӯи презентатсияро навис:"
    )
    bot.register_next_step_handler(message, get_topic)


def get_topic(message):
    topic = message.text.strip()

    if len(topic) < 3:
        msg = bot.send_message(
            message.chat.id,
            "Мавзӯъро дуруст навис:"
        )
        bot.register_next_step_handler(msg, get_topic)
        return

    msg = bot.send_message(
        message.chat.id,
        "Чанд слайд мехоҳӣ?\n"
        "5, 10, 15 ё 20 навис."
    )
    bot.register_next_step_handler(
        msg, get_count, topic
    )


def get_count(message, topic):
    try:
        count = int(message.text.strip())
        if count not in [5, 10, 15, 20]:
            raise ValueError
    except ValueError:
        msg = bot.send_message(
            message.chat.id,
            "Танҳо 5, 10, 15 ё 20 навис."
        )
        bot.register_next_step_handler(
            msg, get_count, topic
        )
        return

    order_id = str(message.from_user.id)
    price = count * 2

    orders[order_id] = {
        "topic": topic,
        "count": count,
        "price": price,
        "status": "waiting_payment"
    }

    bot.send_message(
        message.chat.id,
        f"📚 Мавзӯъ: {topic}\n"
        f"📄 Слайдҳо: {count}\n"
        f"💰 Нарх: {price} сомонӣ\n\n"
        f"🏦 Бонк: Эсхата\n"
        f"💳 Корт: {CARD_NUMBER}\n"
        f"👤 Қабулкунанда: {CARD_OWNER}\n\n"
        "Пас аз пардохт расидро ҳамчун акс фирист.\n"
        "Файл танҳо баъди санҷиши пардохт дода мешавад."
    )


@bot.message_handler(content_types=["photo"])
def receive_receipt(message):
    user_id = str(message.from_user.id)

    if user_id not in orders:
        bot.reply_to(
            message,
            "Аввал /start-ро пахш карда фармоиш деҳ."
        )
        return

    order = orders[user_id]

    if order["status"] != "waiting_payment":
        bot.reply_to(
            message,
            "Расид аллакай қабул шудааст."
        )
        return

    order["status"] = "receipt_sent"

    bot.reply_to(
        message,
        "📸 Расид қабул шуд!\n"
        "Пас аз санҷиши пардохт файлат фиристода мешавад."
    )

    if ADMIN_ID == 0:
        bot.send_message(
            message.chat.id,
            "⚠️ Администратор ҳанӯз танзим нашудааст."
        )
        return

    keyboard = types.InlineKeyboardMarkup()
    keyboard.add(
        types.InlineKeyboardButton(
            "✅ Тасдиқ ва фиристодани слайдҳо",
            callback_data=f"approve_{user_id}"
        )
    )

    bot.send_photo(
        ADMIN_ID,
        message.photo[-1].file_id,
        caption=(
            f"Фармоиши нав\n"
            f"ID: {user_id}\n"
            f"Мавзӯъ: {order['topic']}\n"
            f"Слайдҳо: {order['count']}\n"
            f"Нарх: {order['price']} сомонӣ\n\n"
            "Аввал дар барномаи бонкӣ маблағро санҷ!"
        ),
        reply_markup=keyboard
    )


def make_presentation(topic, count):
    prs = Presentation()
    prs.slide_width = Inches(12)
    prs.slide_height = Inches(6.75)

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = topic
    slide.placeholders[1].text = "SlideMasterTJ"

    for i in range(2, count + 1):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"{topic} — қисми {i - 1}"

        box = slide.placeholders[1].text_frame
        box.clear()
        p = box.paragraphs[0]
        p.text = (
            f"Мавзӯъ: {topic}\n\n"
            f"Қисми {i - 1}\n"
            "Ин слайдро бо маълумоти зарурӣ пур кунед."
        )
        p.font.size = Pt(24)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    output.name = "SlideMasterTJ.pptx"
    return output


@bot.callback_query_handler(
    func=lambda call: call.data.startswith("approve_")
)
def approve_payment(call):
    if call.from_user.id != ADMIN_ID:
        bot.answer_callback_query(
            call.id,
            "Ин тугма танҳо барои администратор аст.",
            show_alert=True
        )
        return

    user_id = call.data.split("_", 1)[1]

    if user_id not in orders:
        bot.answer_callback_query(call.id, "Фармоиш ёфт нашуд.")
        return

    order = orders[user_id]

    if order["status"] != "receipt_sent":
        bot.answer_callback_query(
            call.id,
            "Ин фармоиш барои тасдиқ омода нест."
        )
        return

    try:
        file_data = make_presentation(
            order["topic"], order["count"]
        )

        bot.send_document(
            int(user_id),
            file_data,
            caption=(
                "✅ Пардохт тасдиқ шуд!\n"
                "Презентатсияи ту омода аст."
            )
        )

        order["status"] = "completed"
        bot.answer_callback_query(
            call.id,
            "Презентатсия фиристода шуд!"
        )
        bot.send_message(
            ADMIN_ID,
            f"Фармоиши {user_id} анҷом ёфт."
        )

    except Exception as e:
        print("Хато:", e)
        bot.answer_callback_query(
            call.id,
            "Ҳангоми сохтани файл хато шуд."
        )


print("SlideMasterTJ фаъол аст!")
bot.infinity_polling()
