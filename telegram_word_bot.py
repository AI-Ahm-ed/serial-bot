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

# التوكن الخاص بك
TOKEN = "8876238881:AAEVbcBHKdpsFRIHxj_P5me6NLEc0JXA2lU"

# قائمة مؤقتة لتخزين الكارتات المرسلة
cards_database = []

def fill_card_data(doc, category, serial, pin, exp):
    """
    دالة مساعدة لتعبئة الحقول داخل المستند
    """
    for paragraph in doc.paragraphs:
        if any(k in paragraph.text for k in ["[CATEGORY]", "[SERIAL]", "[PIN]", "[EXP]"]):
            for run in paragraph.runs:
                run.text = (run.text
                            .replace("[CATEGORY]", category)
                            .replace("[SERIAL]", serial)
                            .replace("[PIN]", pin)
                            .replace("[EXP]", exp))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if any(k in cell.text for k in ["[CATEGORY]", "[SERIAL]", "[PIN]", "[EXP]"]):
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.text = (run.text
                                        .replace("[CATEGORY]", category)
                                        .replace("[SERIAL]", serial)
                                        .replace("[PIN]", pin)
                                        .replace("[EXP]", exp))

def create_combined_word_document(cards_list):
    """
    دمج قالب الوورد في ملف واحد بحيث يأخذ كل كارت صفحة جديدة بشكل صحيح ومنظم
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, "template.docx")
    
    if not os.path.exists(template_path):
        logger.error(f"ملف القالب غير موجود في المسار: {template_path}")
        return None

    try:
        # نبدأ بإنشاء المستند الرئيسي باستخدام أول كارت
        master_doc = Document(template_path)
    except Exception as e:
        logger.error(f"فشل في فتح ملف القالب: {e}")
        return None

    for index, card in enumerate(cards_list):
        category = card['category']
        serial = card['serial']
        pin = card['pin']
        exp = card['exp']

        if index == 0:
            # تعبئة الكارت الأول في المستند الرئيسي
            fill_card_data(master_doc, category, serial, pin, exp)
        else:
            # للكارتات اللاحقة، نضيف فاصل صفحات ثم نقرأ نسخة جديدة ونعبئها ونضيفها للمستند الرئيسي
            master_doc.add_page_break()
            temp_doc = Document(template_path)
            fill_card_data(temp_doc, category, serial, pin, exp)
            
            # نسخ العناصر من المستند المؤقت إلى المستند الرئيسي بشكل نظيف
            for element in temp_doc.element.body:
                master_doc.element.body.append(element)

    output_filename = "all_cards_combined.docx"
    output_path = os.path.join(base_dir, output_filename)
    master_doc.save(output_path)
    return output_path

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global cards_database
    text = update.message.text.strip()
    
    # 1. أمر تفريغ القائمة
    if text.lower() == "clear":
        count = len(cards_database)
        cards_database = []
        await update.message.reply_text(f"🗑️ تم مسح وتفريغ القائمة بنجاح! (تم حذف {count} كارت من الذاكرة).")
        return

    # 2. أمر طلب الملف المجمع
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
                    caption=f"✅ تم إرسال الملف المجمع ويحتوي على ({len(cards_database)}) كارت مرتبة في صفحات منفصلة.\n(ملاحظة: الكارتات لا تزال محفوظة، يمكنك طلب الملف مرة أخرى أو كتابة clear للتفريغ)."
                )
            os.remove(file_path)
        else:
            await update.message.reply_text("❌ حدث خطأ أثناء دمج المستندات، تأكد من وجود ملف `template.docx`.")
        return

    # 3. معالجة وتخزين الكارت الوارد
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
    application.run_polling()

if __name__ == '__main__':
    main()
