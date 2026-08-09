import os
import logging
import re
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

def generate_word_document(category, serial, pin, exp):
    """
    تعبئة قالب الوورد بالبيانات المطلوبة واستبدال الحقول بناءً على النموذج
    """
    doc_path = "template.docx"  # تأكد من رفع ملف القالب بنفس هذا الاسم
    
    try:
        doc = Document(doc_path)
    except Exception as e:
        logger.error(f"فشل في فتح ملف القالب: {e}")
        return None

    # استبدال الحقول داخل المستند (النصوص والجداول)
    for paragraph in doc.paragraphs:
        if "[CATEGORY]" in paragraph.text or "[SERIAL]" in paragraph.text or "[PIN]" in paragraph.text or "[EXP]" in paragraph.text:
            for run in paragraph.runs:
                run.text = (run.text
                            .replace("[CATEGORY]", category)
                            .replace("[SERIAL]", serial)
                            .replace("[PIN]", pin)
                            .replace("[EXP]", exp))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "[CATEGORY]" in cell.text or "[SERIAL]" in cell.text or "[PIN]" in cell.text or "[EXP]" in cell.text:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.text = (run.text
                                        .replace("[CATEGORY]", category)
                                        .replace("[SERIAL]", serial)
                                        .replace("[PIN]", pin)
                                        .replace("[EXP]", exp))

    output_filename = f"card_{serial}.docx"
    doc.save(output_filename)
    return output_filename

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    معالجة الرسائل الواردة بالصيغة متعددة الأسطر:
    15K 
    Ser: 26041900225094 
    PIN: 2421896018520315 
    Exp: 2028-08-31
    """
    text = update.message.text.strip()
    lines = text.split('\n')
    
    if len(lines) < 4:
        await update.message.reply_text(
            "⚠️ الصيغة غير صحيحة!\n"
            "الرجاء إرسال الرسالة بنفس التنسيق:\n\n"
            "15K\n"
            "Ser: 26041900225094\n"
            "PIN: 2421896018520315\n"
            "Exp: 2028-08-31"
        )
        return

    # السطر الأول يمثل الـ Category
    category = lines[0].strip()
    
    serial = ""
    pin = ""
    exp = ""

    # استخراج البيانات باستخدام البحث عن المفاتيح بغض النظر عن ترتيب الأسطر الباقية
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

    # توليد ملف الوورد
    file_path = generate_word_document(category, serial, pin, exp)
    
    if file_path and os.path.exists(file_path):
        with open(file_path, 'rb') as doc_file:
            await update.message.reply_document(
                document=doc_file,
                caption=f"✅ تم توليد الكارت بنجاح:\nSer: {serial}"
            )
        # حذف الملف من السيرفر بعد الإرسال
        os.remove(file_path)
    else:
        await update.message.reply_text("❌ حدث خطأ أثناء توليد المستند، تأكد من وجود ملف `template.docx` في المشروع.")

def main():
    # بناء تطبيق البوت
    application = ApplicationBuilder().token(TOKEN).build()

    # استقبال الرسائل النصية
    application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    print("Bot is running...")
    application.run_polling()

if __name__ == '__main__':
    main()
