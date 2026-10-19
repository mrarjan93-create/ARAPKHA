import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# ضبط إعدادات الصفحة
st.set_page_config(page_title="منصة تقديم البحوث الأكاديمية", layout="centered")

st.title("📄 نموذج تقديم وإعداد البحث الأكاديمي")
st.write("يرجى ملء الخانات التالية ليتم إخراج بحثك وفق القالب المعتمد للمجلة تلقائياً.")

# إنشاء استمارة إدخال البيانات
with st.form("journal_form"):
    st.subheader("1. المعلومات الأساسية")
    title = st.text_input("عنوان البحث بالعربية")
    authors = st.text_area("أسماء الباحثين والانتساب (مثال: د. أحمد علي - جامعة بغداد)")
    
    st.subheader("2. الملخص والكلمات المفتاحية")
    abstract = st.text_area("الملخص (Abstract)", height=150)
    keywords = st.text_input("الكلمات المفتاحية (تفصل بينها فاصلة)")

    st.subheader("3. متون البحث")
    introduction = st.text_area("1. المقدمة", height=200)
    methodology = st.text_area("2. منهجية البحث والجانب التطبيقي", height=200)
    results = st.text_area("3. النتائج والمناقشة", height=200)
    conclusions = st.text_area("4. الاستنتاجات والتوصيات", height=150)
    
    st.subheader("4. المراجع والمصادر")
    references = st.text_area("قائمة المراجع (وفق الترتيب المطلوب)", height=200)
    
    submitted = st.form_submit_button("توليد ملف Word المنسق")

# عند ضغط الباحث على زر التوليد
if submitted:
    if not title or not authors:
        st.error("يرجى تعبئة عنوان البحث وأسماء الباحثين على الأقل!")
    else:
        # إنشاء مستند Word جديد وتطبيق التنسيق الأكاديمي
        doc = Document()

        # ضبط الهوامش (مثلاً 2.5 سم من كل جانب)
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # دالة مساعدة لإضافة العناوين بتنسيق معين
        def add_heading(text, level=1):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if level != 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            run.font.name = 'Traditional Arabic' # اسم الخط المعتمد لمجلتك
            run.font.size = Pt(16 if level == 1 else 18)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 51, 102) # لون أزرق غامق للعناوين
            return p

        # دالة مساعدة لإضافة النصوص العادية
        def add_body_paragraph(text):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # محاذاة من الجانبين
            run = p.add_run(text)
            run.font.name = 'Traditional Arabic'
            run.font.size = Pt(14)
            p.paragraph_format.line_spacing = 1.15
            return p

        # 1. إضافة عنوان البحث
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_title = p_title.add_run(title)
        r_title.font.name = 'Traditional Arabic'
        r_title.font.size = Pt(22)
        r_title.font.bold = True

        # 2. إضافة الباحثين
        p_auth = doc.add_paragraph()
        p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_auth = p_auth.add_run(authors)
        r_auth.font.name = 'Traditional Arabic'
        r_auth.font.size = Pt(12)
        r_auth.font.italic = True

        doc.add_paragraph() # مسافة فارغة

        # 3. الملخص والكلمات المفتاحية
        add_heading("الملخص")
        add_body_paragraph(abstract)
        
        p_key = doc.add_paragraph()
        r_key_lbl = p_key.add_run("الكلمات المفتاحية: ")
        r_key_lbl.font.bold = True
        r_key_lbl.font.name = 'Traditional Arabic'
        p_key.add_run(keywords).font.name = 'Traditional Arabic'

        doc.add_paragraph()

        # 4. أجزاء البحث
        if introduction:
            add_heading("1. المقدمة")
            add_body_paragraph(introduction)

        if methodology:
            add_heading("2. منهجية البحث")
            add_body_paragraph(methodology)

        if results:
            add_heading("3. النتائج والمناقشة")
            add_body_paragraph(results)

        if conclusions:
            add_heading("4. الاستنتاجات والتوصيات")
            add_body_paragraph(conclusions)

        if references:
            add_heading("المراجع والمصادر")
            add_body_paragraph(references)

        # حفظ المستند في الذاكرة لتنزيله
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("تم توليد ملف البحث بنجاح وفق قالب المجلة!")
        st.download_button(
            label="📥 تحميل ملف البحث بصيغة Word (.docx)",
            data=buffer,
            file_name=f"Research_{title[:15]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
