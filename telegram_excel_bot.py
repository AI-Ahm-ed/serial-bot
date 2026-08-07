import re
import os
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

# ضع الـ Token الذي حصلت عليه من BotFather هنا
TOKEN = "8876238881:AAEVbcBHKdpsFRIHxj_P5me6NLEc0JXA2lU"
EXCEL_FILE = "messages_data.xlsx"

# دالة لحفظ البيانات في Excel
def save_to_excel(serial, pin, date_val):
    new_row = {"Serial Number": serial, "PIN": pin, "Date": date_val}
    
    # إذا كان الملف موجوداً مسبقاً، نضيف البيانات إليه، وإلا ننشئ ملفاً جديداً
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    else:
        df = pd.DataFrame([new_row])
        
    df.to_excel(EXCEL_FILE, index=False)

# دالة استقبال الرسائل وتفكيكها
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    
    # التعبير النمطي لالتقاط الحقول الثلاثة بغض النظر عن المسافات أو حالة الأحرف
    pattern = r"serial\s*numb:\s*(.*?)\nPIN:\s*(.*?)\nDATE:\s*(.*)"
    match = re.search(pattern, text, re.IGNORECASE)
    
    if match:
        serial = match.group(1).strip()
        pin = match.group(2).strip()
        date_val = match.group(3).strip()
        
        # حفظ في Excel
        save_to_excel(serial, pin, date_val)
        
        # إرسال تأكيد للمُرسل
        await update.message.reply_text(
            f"✅ تم حفظ البيانات بنجاح!\n\n🔹 Serial: {serial}\n🔹 PIN: {pin}\n🔹 Date: {date_val}"
        )
    else:
        await update.message.reply_text(
            "⚠️ لم يتم التعرف على نمط الرسالة.\nتأكد أن الرسالة تحتوي على:\nserial numb:\nPIN:\nDATE:"
        )

if __name__ == "__main__":
    print("🤖 البوت يعمل الان ويستقبل الرسائل...")
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
    app.run_polling()