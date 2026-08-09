import re
import os
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes
from docx import Document

TOKEN = "8876238881:AAEVbcBHKdpsFRIHxj_P5me6NLEc0JXA2lU"
TEMPLATE = "template.docx"
FINAL_FILE = "all_cards.docx"

def fill_template(cat, ser, pin, date):
    # إذا لم يوجد الملف النهائي، ننشئه لأول مرة من القالب
    if not os.path.exists(FINAL_FILE):
        doc = Document(TEMPLATE) if os.path.exists(TEMPLATE) else Document()
    else:
        # إذا كان موجوداً، نفتح الملف الحالي ونضيف صفحة جديدة ننسخ فيها القالب
        doc = Document(FINAL_FILE)
        doc.add_page_break()
        if os.path.exists(TEMPLATE):
            temp_doc = Document(TEMPLATE)
            for element in temp_doc.element.body:
                doc.element.body.append(element)

    # استبدال النصوص في آخر جزء (الكارت) تمت إضافته
    for paragraph in doc.paragraphs:
        if "[CATEGORY]" in paragraph.text: paragraph.text = paragraph.text.replace("[CATEGORY]", cat)
        if "[SERIAL]" in paragraph.text: paragraph.text = paragraph.text.replace("[SERIAL]", ser)
        if "[PIN]" in paragraph.text: paragraph.text = paragraph.text.replace("[PIN]", pin)
        if "[DATE]" in paragraph.text: paragraph.text = paragraph.text.replace("[DATE]", date)
    
    doc.save(FINAL_FILE)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()

    # أمر التصفير (تفريغ الملف)
    if text.lower() == "clear":
        if os.path.exists(TEMPLATE):
            doc = Document(TEMPLATE)
            doc.save(FINAL_FILE)
            await update.message.reply_text("🧹 تم تصفير الملف، أنت الآن على لوح أبيض!")
        else:
            await update.message.reply_text("⚠️ ملف القالب (template.docx) غير موجود!")
        return

    # أمر الحصول على الملف
    if text.lower() == "file":
        if os.path.exists(FINAL_FILE):
            await update.message.reply_document(document=open(FINAL_FILE, 'rb'))
        else:
            await update.message.reply_text("⚠️ الملف فارغ.")
        return

    # نمط استقبال الكارت (تأكد من إرسال البيانات بنفس الترتيب)
    pattern = r"(.*?)\n.*?Ser:\s*(.*?)\n.*?PIN:\s*(.*?)\n.*?Date:\s*(.*)"
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)

    if match:
        fill_template(match.group(1).strip(), match.group(2).strip(), match.group(3).strip(), match.group(4).strip())
        await update.message.reply_text(f"🖨️ تمت إضافة كارت {match.group(1).strip()} بنجاح!")
    else:
        await update.message.reply_text("❌ صيغة غير صحيحة. يرجى الإرسال كالتالي:\n15K\nSer: ...\nPIN: ...\nDate: ...")

app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
app.run_polling()
