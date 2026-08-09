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

# قراءة الـ Token من متغيرات البيئة أو وضعه مباشرة هنا
TOKEN = os.getenv("TELEGRAM_TOKEN", "YOUR_BOT_TOKEN_HERE")

def generate_word_document(category, serial, pin, exp):
    """
    تعبئة قالب الوورد بالبيانات المطلوبة واستبدال الحقول بناءً على النموذج الجديد
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
    معالجة الرسائل الواردة، توقع رسالة بالشكل:
    Category | Serial | Pin | Exp
    """
    text = update.message.text.strip()
    
    # إذا أرسل المستخدم كلمة file أو ما شابه
    if text.lower() == "file":
        await update.message.reply_text("الرجاء إرسال بيانات الكارت بهذا الترتيب:\nCategory | Serial | Pin | Exp")
        return

    # تقسيم النص بناءً على الفاصل "|"
    parts = [p.strip() for p in text.split('|')]
    
    if len(parts) != 4:
        await update.message.reply_text(
            "⚠️ الصيغة غير صحيحة!\n"
            "الرجاء إرسال البيانات بالترتيب التالي فاصلاً بينها بـ |\n\n"
            "مثال:\n"
            "15K | 26041900225094 | 2421896018520315 | 2028-08-31"
        )
        return

    category, serial, pin, exp = parts

    # توليد ملف الوورد
    file_path = generate_word_document(category, serial, pin, exp)
    
    if file_path and os.path.exists(file_path):
        with open(file_path, 'rb') as doc_file:
            await update.message.reply_document(
                document=doc_file,
                caption=f"✅ تم توليد كارت الـ Exp بنجاح:\nSer: {serial}"
            )
        # حذف الملف من السيرفر بعد الإرسال للحفاظ على المساحة
        os.remove(file_path)
    else:
        await update.message.reply_text("❌ حدث خطأ أثناء توليد المستند، تأكد من وجود ملف `template.docx`.")

def main():
    # بناء تطبيق البوت
    application = ApplicationBuilder().token(TOKEN).build()

    # استقبال أي رسالة نصية ومعالجتها
    application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    # تشغيل البوت
    print("Bot is running...")
    application.run_polling()

if __name__ == '__main__':
    main()
