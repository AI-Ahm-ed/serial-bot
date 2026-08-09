import os
import logging
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, MessageHandler, filters
from docx import Document

# إعداد السجلات للتتبع
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

TOKEN = "8876238881:AAEVbcBHKdpsFRIHxj_P5me6NLEc0JXA2lU"
cards_database = []

def fill_and_append_card(master_doc, template_path, category, serial, pin, exp, is_first):
    """
    تقوم بفتح نسخة من القالب، استبدال البيانات فيها، ثم إضافتها للمستند الرئيسي 
    مع الحفاظ على التصميم واللوجو والجداول تماماً.
    """
    if not is_first:
        master_doc.add_page_break()

    # نفتح نسخة جديدة من القالب لكل كارت لضمان بقاء التصميم واللوجو والصور
    temp_doc = Document(template_path)
    
    replacements = {
        "[CATEGORY]": category,
        "[SERIAL]": serial,
        "[PIN]": pin,
        "[EXP]": exp
    }

    # استبدال النصوص في الفقرات
    for p in temp_doc.paragraphs:
        for key, value in replacements.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, value)

    # استبدال النصوص في الجداول
    for table in temp_doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replacements.items():
                    if key in cell.text:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, value)

    # إذا لم يكن الكارت الأول، ننقل محتويات النسخة إلى المستند الرئيسي
    if is_first:
        # ننسخ محتويات أول قالب للمستند الرئيسي مباشرة
        master_doc._body._element.clear()
        for element in temp_doc._body._element:
            master_doc._body._element.append(element)
    else:
        # نضيف عناصر الكارت الجديد للمستند الرئيسي
        for element in temp_doc._body._element:
            master_doc._body._element.append(element)

def create_combined_word_document(cards_list):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, "template.docx")
    
    if not os.path.exists(template_path):
        logger.error(f"ملف القالب غير موجود في المسار: {template_path}")
        return None

    try:
        # نبدأ بمستند فارغ لنبني فيه الكارتات بالتصميم الأصلي
        master_doc = Document(template_path)
    except Exception as e:
        logger.error(f"فشل في فتح ملف القالب: {e}")
        return None

    for index, card in enumerate(cards_list):
        fill_and_append_card(
            master_doc=master_doc,
            template_path=template_path,
            category=card['category'],
            serial=card['serial'],
            pin=card['pin'],
            exp=card['exp'],
            is_first=(index == 0)
        )

    output_filename = "all_cards_combined.docx"
    output_path = os.path.join(base_dir, output_filename)
    master_doc.save(output_path)
    return output_path

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global cards_database
    text = update.message.text.strip()
    
    if text.lower() == "clear":
        count = len(cards_database)
        cards_database = []
        await update.message.reply_text(f"🗑️ تم مسح وتفريغ القائمة بنجاح! (تم حذف {count} كارت من الذاكرة).")
        return

    if text.lower() == "file":
        if not cards_database:
            await update.message.reply_text("⚠️ لم تقم بإرسال أي كارت بعد! أرسل الكارتات أولاً ثم اكتب file.")
            return
        
        await update.message.reply_text(f"⏳ جاري دمج {len(cards_database)} كارت في ملف واحد...")
        
        file_path = create_combined_word_document(cards_database)
        
        if file_path and os.path.exists(file_path):
            with open(file_path, 'rb') as doc_file:
                await update.message.reply_document(
                    document=doc_file,
                    caption=f"✅ تم إرسال الملف المجمع ويحتوي على ({len(cards_database)}) كارت بنفس التصميم واللوجو.\n(ملاحظة: الكارتات لا تزال محفوظة، يمكنك طلب الملف مرة أخرى أو كتابة clear للتفريغ)."
                )
            os.remove(file_path)
        else:
            await update.message.reply_text("❌ حدث خطأ أثناء دمج المستندات، تأكد من وجود ملف `template.docx`.")
        return

    lines = text.split('\n')
    
    if len(lines) < 4:
        await update.message.reply_text(
            "⚠️ الصيغة غير صحيحة!\n"
            "الرجاء إرسال الرسالة بنفس التنسيق:\n\n"
            "15K\n"
            "Ser: 26041900225094\n"
            "PIN: 2421896018520315\n"
            "Exp: 2028-08-31\n\n"
            "الأوامر المتاحة:\n"
            "- **file**: لاستلام الملف المجمع.\n"
            "- **clear**: لتفريغ القائمة وحذف الكارتات المخزنة."
        )
        return

    category = lines[0].strip()
    serial = ""
    pin = ""
    exp = ""

    for line in lines[1:]:
        if "Ser:" in line:
            serial = line.replace("Ser:", "").strip()
        elif "PIN:" in line:
            pin = line.replace("PIN:", "").strip()
        elif "Exp:" in line:
            exp = line.replace("Exp:", "").strip()

    if not serial or not pin or not exp:
        await update.message.reply_text("⚠️ لم يتم العثور على جميع البيانات (Ser, PIN, Exp). تأكد من كتابة الكلمات بشكل صحيح.")
        return

    cards_database.append({
        'category': category,
        'serial': serial,
        'pin': pin,
        'exp': exp
    })

    await update.message.reply_text(f"📥 تم حفظ الكارت (Ser: {serial})\nالعدد الإجمالي المخزن حالياً: {len(cards_database)}")

def main():
    application = ApplicationBuilder().token(TOKEN).build()
    application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    print("Bot is running...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    main()
