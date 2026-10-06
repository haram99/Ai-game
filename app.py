# -*- coding: utf-8 -*-
"""
لعبة "رحلة الذكاء الاصطناعي" لطالبات المرحلة الثانوية (نسخة الويب)
التشغيل محلياً: streamlit run app.py
"""

import math
import sqlite3
import streamlit as st
import pandas as pd
import plotly.express as px

# إعداد الصفحة
st.set_page_config(page_title="رحلة الذكاء الاصطناعي 🤖", page_icon="🤖", layout="centered")

# ------------------------- إعداد قاعدة البيانات -------------------------
def init_db():
    conn = sqlite3.connect("ai_game_players.db")
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS players (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT,
                    phone TEXT,
                    score INTEGER,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
                )''')
    conn.commit()
    conn.close()

init_db()

def save_player_data(name, phone, score):
    conn = sqlite3.connect("ai_game_players.db")
    c = conn.cursor()
    c.execute("INSERT INTO players (name, phone, score) VALUES (?, ?, ?)", (name, phone, score))
    conn.commit()
    conn.close()

# ------------------------- بنك الأسئلة -------------------------
LEVEL1 = [
    ("ما الذي يميّز الذكاء الاصطناعي عن البرنامج التقليدي؟",
     ["يتعلّم من البيانات ويحسّن أداءه", "يعمل بدون كهرباء", "لا يحتاج إلى بيانات أبدًا"], 0,
     "الذكاء الاصطناعي يتعلّم الأنماط من البيانات بدل أن نكتب له كل قاعدة يدويًا."),
    ("أي مما يلي مثال على تطبيق للذكاء الاصطناعي؟",
     ["الآلة الحاسبة البسيطة", "التعرّف على الوجه لفتح الجوال", "مصباح كهربائي"], 1,
     "التعرّف على الوجه يتعلّم من صور كثيرة ليميّز الوجوه."),
    ("ما أهم مادة خام لتدريب نموذج ذكاء اصطناعي؟",
     ["البيانات", "الألوان", "الصوت العالي"], 0,
     "جودة البيانات وكميتها تحدّد جودة النموذج: «مدخلات سيئة = نتائج سيئة»."),
    ("عندما يقترح يوتيوب فيديو يعجبك، فهو يستخدم:",
     ["الحظ", "التعلّم من سلوكك وسلوك المستخدمين (أنظمة التوصية)", "اختيار موظف لكل شخص"], 1,
     "أنظمة التوصية تحلّل ما شاهدته لتتوقّع ما تحبّينه."),
    ("ما معنى «تدريب النموذج»؟",
     ["إعطاؤه أمثلة ليتعلّم منها", "إطفاؤه وتشغيله", "رسم شكل له"], 0,
     "ندرّب النموذج بأمثلة معنونة (مثلاً: هذه تفاحة، وهذه موزة) فيكتشف الفروق."),
]

LEVEL3 = [
    ("درّبوا نظامًا للتعرّف على الوجوه بصور أشخاص من بلد واحد فقط. ماذا سيحدث غالبًا؟",
     ["سيعمل بدقة متساوية مع الجميع", "سيخطئ أكثر مع وجوه لم يرها كثيرًا", "سيصبح أذكى من البشر"], 1,
     "هذا يسمّى «تحيّز البيانات»: النموذج يتقن ما رآه كثيرًا ويضعف فيما رآه قليلًا."),
    ("لتقليل التحيّز في النموذج، الأفضل أن:",
     ["نجمع بيانات متنوعة ومتوازنة", "نحذف نصف البيانات عشوائيًا", "نستخدم صورة واحدة فقط"], 0,
     "التنوّع والتوازن في البيانات يجعلان النموذج أعدل وأدق."),
    ("نظام يرفض طلبات التوظيف اعتمادًا على بيانات قديمة فيها تمييز. ما المشكلة؟",
     ["النموذج يتعلّم التمييز نفسه من البيانات", "الحاسوب بطيء", "الشاشة صغيرة"], 0,
     "النموذج يعكس ما في بياناته؛ لذلك يجب مراجعة البيانات وعدم الاعتماد الأعمى على النتائج."),
    ("هل يجب أن يراجع إنسان القرارات المهمة التي يتخذها الذكاء الاصطناعي؟",
     ["لا، الآلة لا تخطئ أبدًا", "نعم، للمراجعة والمسؤولية", "فقط في أيام العطل"], 1,
     "الإنسان مسؤول عن القرار النهائي، خاصة في الصحة والتعليم والعمل."),
]

TEST_SET = [
    (4.5, 6.5, "تفاحة"), (7.5, 3.0, "موزة"), (3.0, 9.0, "تفاحة"),
    (8.5, 1.5, "موزة"), (6.0, 4.0, "موزة"), (5.0, 6.0, "تفاحة"),
]

def knn_predict(train, point, k=1):
    dists = sorted((math.dist((x, y), point), label) for x, y, label in train)
    top = [lbl for _, lbl in dists[:k]]
    return max(set(top), key=top.count)

# ------------------------- إدارة جلسة اللعبة -------------------------
if "step" not in st.session_state:
    st.session_state.step = "register"
    st.session_state.name = ""
    st.session_state.phone = ""
    st.session_state.score = 0
    st.session_state.train_data = []

# 1. شاشة التسجيل
if st.session_state.step == "register":
    st.title("🤖 رحلة الذكاء الاصطناعي")
    st.markdown("أهلاً بكِ في رحلتنا الممتعة! الرجاء إدخال بياناتكِ للبدء:")
    
    with st.form("reg_form"):
        name = st.text_input("اسم الطالبة الثلاثي:")
        phone = st.text_input("رقم الجوال (للتواصل أو حفظ النتيجة):")
        submitted = st.form_submit_button("ابدأ الرحلة 🚀")
        
        if submitted:
            if name.strip() and phone.strip():
                st.session_state.name = name
                st.session_state.phone = phone
                st.session_state.step = "home"
                st.rerun()
            else:
                st.error("الرجاء إدخال الاسم ورقم الجوال بشكل صحيح.")

# 2. القائمة الرئيسية
elif st.session_state.step == "home":
    st.title(f"مرحباً بكِ، {st.session_state.name} 🌟")
    st.success(f"مجموع نقاطكِ الحالي: {st.session_state.score}")
    
    st.markdown("---")
    choice = st.radio("اختاري المستوى:", [
        "المستوى 1: ما هو الذكاء الاصطناعي؟",
        "المستوى 2: درّبي الآلة 🍎🍌",
        "المستوى 3: اكتشفي التحيّز ⚖️",
        "إنهاء وحفظ النتيجة 🏁"
    ])
    
    if st.button("انتقلي للمستوى المختار"):
        if "المستوى 1" in choice:
            st.session_state.step = "level1"
            st.session_state.q_idx = 0
            st.session_state.correct_count = 0
        elif "المستوى 2" in choice:
            st.session_state.step = "level2"
        elif "المستوى 3" in choice:
            st.session_state.step = "level3"
            st.session_state.q_idx = 0
            st.session_state.correct_count = 0
        elif "إنهاء" in choice:
            save_player_data(st.session_state.name, st.session_state.phone, st.session_state.score)
            st.session_state.step = "finished"
        st.rerun()

# 3. المستوى الأول والثالث (نظام الأسئلة)
elif st.session_state.step in ["level1", "level3"]:
    lvl_title = "المستوى 1: ما هو الذكاء الاصطناعي؟" if st.session_state.step == "level1" else "المستوى 3: اكتشفي التحيّز ⚖️"
    questions = LEVEL1 if st.session_state.step == "level1" else LEVEL3
    
    st.subheader(lvl_title)
    idx = st.session_state.q_idx
    
    if idx < len(questions):
        q, options, ans, explain = questions[idx]
        st.write(f"**سؤال {idx + 1} من {len(questions)}:**")
        st.markdown(f"### {q}")
        
        user_choice = st.radio("اختاري الإجابة المناسبة:", options, key=f"q_{idx}")
        
        if st.button("تأكيد الإجابة"):
            selected_idx = options.index(user_choice)
            if selected_idx == ans:
                st.success("✅ إجابة صحيحة! " + explain)
                st.session_state.score += 10
                st.session_state.correct_count += 1
            else:
                st.error("❌ إجابة خاطئة. " + explain)
            
            if st.button("السؤال التالي ◀"):
                st.session_state.q_idx += 1
                st.rerun()
    else:
        st.balloons()
        st.success(f"🎉 أنهيتِ المستوى بنجاح! أجبتِ بشكل صحيح على {st.session_state.correct_count} من {len(questions)}")
        if st.button("العودة للقائمة الرئيسية"):
            st.session_state.step = "home"
            st.rerun()

# 4. المستوى الثاني (تدريب الآلة)
elif st.session_state.step == "level2":
    st.subheader("🍎🍌 المستوى 2: درّبي الآلة على التفريق بين التفاحة والموزة")
    st.write("حرّكي المنزلقين لوصف الفاكهة (الطول والاستدارة) ثم أضيفي أمثلة للآلة لتتعلم منها.")
    
    length = st.slider("الطول (قصير ← طويل)", 0.0, 10.0, 5.0, 0.5)
    roundness = st.slider("الاستدارة (مستقيم ← مستدير)", 0.0, 10.0, 5.0, 0.5)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("أضيفها كتفاحة 🍎"):
            st.session_state.train_data.append((length, roundness, "تفاحة"))
            st.success("تمت إضافة تفاحة بنجاح!")
    with col2:
        if st.button("أضيفها كموزة 🍌"):
            st.session_state.train_data.append((length, roundness, "موزة"))
            st.success("تمت إضافة موزة بنجاح!")
    with col3:
        if st.button("مسح الأمثلة"):
            st.session_state.train_data = []
            st.info("تم مسح جميع الأمثلة.")
            
    st.write(f"عدد أمثلة التدريب الحالية: **{len(st.session_state.train_data)}**")
    
    if st.button("اختبري الآلة ✅"):
        labels_set = {lbl for _, _, lbl in st.session_state.train_data}
        if len(labels_set) < 2 or len(st.session_state.train_data) < 4:
            st.error("⚠️ أضيفي 4 أمثلة على الأقل، ومن النوعين معًا (تفاحة وموزة).")
        else:
            preds = [((x, y, t), knn_predict(st.session_state.train_data, (x, y))) for x, y, t in TEST_SET]
            right = sum(1 for (_, _, t), p in preds if t == p)
            st.write(f"**دقة الآلة:** {right} من {len(TEST_SET)}")
            if right == len(TEST_SET):
                st.balloons()
                st.success("🏆 ممتاز! الآلة تعلّمت من أمثلتك بدقة كاملة.")
                st.session_state.score += 30
            else:
                st.warning("جربي إضافة أمثلة أكثر وأكثر تنوعاً قرب الحدود بين النوعين.")
                
    if st.button("العودة للقائمة الرئيسية"):
        st.session_state.step = "home"
        st.rerun()

# 5. شاشة النهاية
elif st.session_state.step == "finished":
    st.title("🏆 شكراً لكِ لمشاركتكِ!")
    st.markdown(f"**الاسم:** {st.session_state.name}")
    st.markdown(f"**رقم الجوال:** {st.session_state.phone}")
    st.markdown(f"**النقاط النهائية:** {st.session_state.score} نقطة 🌟")
    st.success("تم حفظ نتيجتكِ وبياناتكِ بنجاح في قاعدة البيانات.")
    
    if st.button("العب مرة أخرى"):
        st.session_state.step = "register"
        st.session_state.score = 0
        st.session_state.train_data = []
        st.rerun()
