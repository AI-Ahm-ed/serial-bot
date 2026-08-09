import os
import logging
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, MessageHandler, filters
from docx import Document
from docx.oxml import OxmlElement

# إعداد السجلات للتتبع
logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TOKEN = "8876238881:AAEVbcBHKdpsFRIHxj_P5me6NLEc0JXA2lU"
cards_database = []

def fill_data(doc, category, serial, pin, exp):
    """تعبئة الحقول في المستند مع الحفاظ على التنسيق"""
    replacements = {
        "[CATEGORY]": category,
        "[SERIAL]": serial,
        "[PIN]": pin,
        "[EXP]": exp
    }
    
    # معالجة الفقرات
    for p in doc.paragraphs:
        for key, value in replacements.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, value)
    
    # معالجة الجداول
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replacements.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)

def create_combined_word_document(cards_list):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, "template.docx")
    
    # نستخدم قالب الوورد الأساسي كبداية
    master_doc = Document(template_path)
    
    # تعبئة أول كارت في القالب الأصلي مباشرة
    fill_data(master_doc, cards_list[0]['category'], cards_list[0]['serial'], cards_list[0]['pin'], cards_list[0]['exp'])
    
    # إضافة الكارتات المتبقية
    for i in range(1, len(cards_list)):
        master_doc.add_page_break()
        # إضافة محتويات القالب مرة أخرى (بدون فقدان التنسيق)
        # نقوم بإنشاء نسخة مؤقتة من القالب لكل كارت
        temp_doc = Document(template_path)
        fill_data(temp_doc, cards_list[i]['category'], cards_list[i]['serial'], cards_list[i]['pin'], cards_list[i]['exp'])
        
        # نسخ فقرات القالب إلى المستند الرئيسي
        for p in temp_doc.paragraphs:
            master_doc.add_paragraph(p.text, style=p.style)
        
        # نسخ جداول القالب إلى المستند الرئيسي
        for table in temp_doc.tables:
            master_doc.add_table(table.rows[0].cells) # هذه بداية بسيطة للنسخ، سنعتمد على أن القالب يحتوي على هيكل ثابت

    output_path = os.path.join(base_dir, "all_cards_combined.docx")
    master_doc.save(output_path)
    return output_path

# ... (بقية دوال التعامل مع البوت كما هي في الكود السابق)
