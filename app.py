# -*- coding: utf-8 -*-
"""
رحلة الذكاء الاصطناعي — أكاديمية نمو (نسخة Streamlit + Google Sheets)

- تسجيل اسم الطالبة ورقم الجوال
- ثلاثة مستويات: أسئلة، تدريب آلة (KNN)، اكتشاف التحيّز
- حفظ بيانات اللاعبات ونتائجهن مباشرة في Google Sheets
- لوحة مشرفة محمية بكلمة مرور لعرض البيانات وتنزيلها بصيغة CSV

التشغيل:
    streamlit run app.py

المكتبات المطلوبة في requirements.txt:
    streamlit
    pandas
    altair
    Pillow
    gspread
    google-auth

إعداد Streamlit Secrets:

GOOGLE_SHEET_ID = "معرف Google Sheet"
GOOGLE_WORKSHEET_NAME = "اللاعبات"

[google_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email = "..."
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."

ADMIN_PASSWORD = "كلمة_مرور_المشرفة"
"""

import hmac
import math
import re
import uuid
from datetime import datetime
from pathlib import Path

import altair as alt
import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from PIL import Image


# ============================== الإعدادات ==============================

BASE = Path(__file__).parent
LOGO = BASE / "logo.jpg"

AR_DIGITS = str.maketrans(
    "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹",
    "0123456789" * 2,
)

HEADERS = [
    "المعرّف",
    "تاريخ التسجيل",
    "الاسم",
    "رقم الجوال",
    "درجة المستوى 1",
    "نقاط المستوى 2",
    "درجة المستوى 3",
    "مجموع النقاط",
    "آخر تحديث",
]

LEVEL_COL = {
    1: "درجة المستوى 1",
    2: "نقاط المستوى 2",
    3: "درجة المستوى 3",
}

FRUIT_COLORS = ["#e03b3b", "#f2c200"]

st.set_page_config(
    page_title="رحلة الذكاء الاصطناعي | أكاديمية نمو",
    page_icon=Image.open(LOGO) if LOGO.exists() else "🤖",
    layout="centered",
)


# ============================== التنسيق ==============================

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');

.stApp,
.stApp p,
.stApp label,
.stApp button,
.stApp input,
.stApp h1,
.stApp h2,
.stApp h3,
.stApp h4,
.stApp li,
.stApp span.stMarkdown {
    font-family: 'Tajawal', sans-serif;
}

.stApp {
    direction: rtl;
    text-align: right;
}

[data-testid="stMarkdownContainer"],
[data-testid="stCaptionContainer"],
[data-testid="stAlert"],
label,
.stRadio {
    direction: rtl;
    text-align: right;
}

.block-container {
    max-width: 760px;
    padding-top: 1.5rem;
}

div.stButton > button,
div[data-testid="stFormSubmitButton"] > button {
    width: 100%;
    border-radius: 12px;
    font-weight: 700;
    padding: .6rem 1rem;
}

input[aria-label^="رقم الجوال"] {
    direction: ltr;
    text-align: center;
    letter-spacing: 2px;
}
</style>
""",
    unsafe_allow_html=True,
)


# ============================== المحتوى ==============================

LEVEL1 = [
    (
        "ما الذي يميّز الذكاء الاصطناعي عن البرنامج التقليدي؟",
        [
            "يتعلّم من البيانات ويحسّن أداءه",
            "يعمل بدون كهرباء",
            "لا يحتاج إلى بيانات أبدًا",
        ],
        0,
        "الذكاء الاصطناعي يتعلّم الأنماط من البيانات بدل أن نكتب له كل قاعدة يدويًا.",
    ),
    (
        "أي مما يلي مثال على تطبيق للذكاء الاصطناعي؟",
        [
            "الآلة الحاسبة البسيطة",
            "التعرّف على الوجه لفتح الجوال",
            "مصباح كهربائي",
        ],
        1,
        "التعرّف على الوجه يتعلّم من صور كثيرة ليميّز الوجوه.",
    ),
    (
        "ما أهم مادة خام لتدريب نموذج ذكاء اصطناعي؟",
        [
            "البيانات",
            "الألوان",
            "الصوت العالي",
        ],
        0,
        "جودة البيانات وكميتها تحدّد جودة النموذج: «مدخلات سيئة = نتائج سيئة».",
    ),
    (
        "عندما يقترح يوتيوب فيديو يعجبك، فهو يستخدم:",
        [
            "الحظ",
            "التعلّم من سلوكك وسلوك المستخدمين (أنظمة التوصية)",
            "اختيار موظف لكل شخص",
        ],
        1,
        "أنظمة التوصية تحلّل ما شاهدته لتتوقّع ما تحبّينه.",
    ),
    (
        "ما معنى «تدريب النموذج»؟",
        [
            "إعطاؤه أمثلة ليتعلّم منها",
            "إطفاؤه وتشغيله",
            "رسم شكل له",
        ],
        0,
        "ندرّب النموذج بأمثلة معنونة (مثلاً: هذه تفاحة، وهذه موزة) فيكتشف الفروق.",
    ),
]

LEVEL3 = [
    (
        "درّبوا نظامًا للتعرّف على الوجوه بصور أشخاص من بلد واحد فقط. ماذا سيحدث غالبًا؟",
        [
            "سيعمل بدقة متساوية مع الجميع",
            "سيخطئ أكثر مع وجوه لم يرها كثيرًا",
            "سيصبح أذكى من البشر",
        ],
        1,
        "هذا يسمّى «تحيّز البيانات»: النموذج يتقن ما رآه كثيرًا ويضعف فيما رآه قليلًا.",
    ),
    (
        "لتقليل التحيّز في النموذج، الأفضل أن:",
        [
            "نجمع بيانات متنوعة ومتوازنة",
            "نحذف نصف البيانات عشوائيًا",
            "نستخدم صورة واحدة فقط",
        ],
        0,
        "التنوّع والتوازن في البيانات يجعلان النموذج أعدل وأدق.",
    ),
    (
        "نظام يرفض طلبات التوظيف اعتمادًا على بيانات قديمة فيها تمييز. ما المشكلة؟",
        [
            "النموذج يتعلّم التمييز نفسه من البيانات",
            "الحاسوب بطيء",
            "الشاشة صغيرة",
        ],
        0,
        "النموذج يعكس ما في بياناته؛ لذلك يجب مراجعة البيانات وعدم الاعتماد الأعمى على النتائج.",
    ),
    (
        "هل يجب أن يراجع إنسان القرارات المهمة التي يتخذها الذكاء الاصطناعي؟",
        [
            "لا، الآلة لا تخطئ أبدًا",
            "نعم، للمراجعة والمسؤولية",
            "فقط في أيام العطل",
        ],
        1,
        "الإنسان مسؤول عن القرار النهائي، خاصة في الصحة والتعليم والعمل.",
    ),
]

# (الطول، الاستدارة، الصنف الحقيقي)
TEST_SET = [
    (4.5, 6.5, "تفاحة"),
    (7.5, 3.0, "موزة"),
    (3.0, 9.0, "تفاحة"),
    (8.5, 1.5, "موزة"),
    (6.0, 4.0, "موزة"),
    (5.0, 6.0, "تفاحة"),
]

LEVELS = [
    (1, "المستوى 1: ما هو الذكاء الاصطناعي؟", "quiz1"),
    (2, "المستوى 2: درّبي الآلة 🍎🍌", "train"),
    (3, "المستوى 3: اكتشفي التحيّز ⚖️", "quiz3"),
]


# ============================== Google Sheets ==============================

@st.cache_resource
def get_google_sheet():
    """
    إنشاء اتصال واحد مع Google Sheets لكل عملية Streamlit.
    يجب أن تكون بيانات Service Account موجودة في st.secrets.
    """

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    if "google_service_account" not in st.secrets:
        raise RuntimeError(
            "لم يتم العثور على google_service_account في Streamlit Secrets."
        )

    if "GOOGLE_SHEET_ID" not in st.secrets:
        raise RuntimeError(
            "لم يتم العثور على GOOGLE_SHEET_ID في Streamlit Secrets."
        )

    credentials = Credentials.from_service_account_info(
        dict(st.secrets["google_service_account"]),
        scopes=scopes,
    )

    client = gspread.authorize(credentials)

    spreadsheet = client.open_by_key(
        str(st.secrets["GOOGLE_SHEET_ID"])
    )

    worksheet_name = str(
        st.secrets.get("GOOGLE_WORKSHEET_NAME", "اللاعبات")
    )

    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=worksheet_name,
            rows=1000,
            cols=len(HEADERS),
        )

    # تجهيز الصف الأول عند إنشاء ورقة جديدة أو إذا كانت فارغة.
    current_headers = worksheet.row_values(1)

    if current_headers != HEADERS:
        worksheet.update(
            "A1:I1",
            [HEADERS],
            value_input_option="USER_ENTERED",
        )

    return worksheet


def save_player(pid: str, fields: dict) -> None:
    """
    إنشاء صف للاعبة أو تحديث صفها في Google Sheets حسب المعرّف.
    """

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    worksheet = get_google_sheet()

    # قراءة عمود المعرّفات للعثور على صف اللاعبة.
    ids = worksheet.col_values(1)

    row = None

    for index, value in enumerate(ids[1:], start=2):
        if str(value).strip() == str(pid).strip():
            row = index
            break

    # إذا لم توجد اللاعبة، ننشئ صفًا جديدًا.
    if row is None:
        worksheet.append_row(
            [
                pid,
                now,
                "",
                "",
                "",
                "",
                "",
                0,
                now,
            ],
            value_input_option="USER_ENTERED",
        )
        row = len(ids) + 1

    # تحديث الحقول المطلوبة.
    for col_name, value in fields.items():
        if col_name not in HEADERS:
            continue

        col = HEADERS.index(col_name) + 1

        worksheet.update_cell(
            row,
            col,
            "" if value is None else value,
        )

    # تحديث آخر تحديث.
    last_update_col = HEADERS.index("آخر تحديث") + 1

    worksheet.update_cell(
        row,
        last_update_col,
        now,
    )


def persist() -> None:
    """
    حفظ نتائج المستويات الحالية في Google Sheets.
    """

    s = st.session_state

    fields = {
        LEVEL_COL[n]: s.scores[n]
        for n in (1, 2, 3)
    }

    fields["مجموع النقاط"] = total()

    try:
        save_player(s.pid, fields)
        s.save_error = False

    except Exception as e:
        s.save_error = True
        print(f"Google Sheets save error: {e}")


# ============================== الحالة ==============================

def init_state():
    defaults = {
        "stage": "login",
        "name": "",
        "phone": "",
        "pid": "",
        "scores": {1: None, 2: None, 3: None},
        "qi": 0,
        "qc": 0,
        "picked": None,
        "train": [],
        "test": None,
        "save_error": False,
    }

    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def total() -> int:
    return sum(
        value or 0
        for value in st.session_state.scores.values()
    )


def go(stage: str):
    st.session_state.stage = stage
    st.rerun()


def start_level(stage: str):
    s = st.session_state

    s.qi = 0
    s.qc = 0
    s.picked = None
    s.train = []
    s.test = None

    go(stage)


def knn_predict(train, point, k=1):
    distances = sorted(
        (
            math.dist((x, y), point),
            label,
        )
        for x, y, label in train
    )

    top = [
        label
        for _, label in distances[:k]
    ]

    return max(
        set(top),
        key=top.count,
    )


# ============================== الشاشات ==============================

def header_bar():
    c1, c2 = st.columns([1, 5])

    if LOGO.exists():
        c1.image(str(LOGO), width=70)

    c2.markdown(
        f"**أكاديمية نمو | Numo Academy**  \n"
        f"👩‍🎓 {st.session_state.name}"
    )

    st.divider()


def screen_login():
    if LOGO.exists():
        _, mid, _ = st.columns([1, 1, 1])
        mid.image(LOGO, use_container_width=True)

    st.markdown(
        "<h2 style='text-align:center'>"
        "مرحبًا بكِ في رحلة الذكاء الاصطناعي 🤖"
        "</h2>",
        unsafe_allow_html=True,
    )

    with st.form("login"):
        name = st.text_input("اسم الطالبة")

        phone = st.text_input(
            "رقم الجوال (مثال: 0500000000)",
            max_chars=10,
        )

        consent = st.checkbox(
            "أوافق على حفظ اسمي ورقم جوالي لدى أكاديمية نمو "
            "لأغراض المتابعة والتواصل."
        )

        submitted = st.form_submit_button(
            "ابدئي الرحلة ◀",
            type="primary",
        )

    if not submitted:
        return

    name = " ".join(name.split())
    phone = phone.strip().translate(AR_DIGITS)

    if len(name) < 2 or not all(
        ch.isalpha() or ch == " "
        for ch in name
    ):
        st.error(
            "⚠️ فضلًا اكتبي اسمك بالحروف فقط (حرفان على الأقل)."
        )
        return

    if not re.fullmatch(r"05\d{8}", phone):
        st.error(
            "⚠️ رقم الجوال يجب أن يكون بالصيغة "
            "0500000000 (10 أرقام تبدأ بـ 05)."
        )
        return

    if not consent:
        st.error(
            "⚠️ فضلًا وافقي على حفظ البيانات للمتابعة."
        )
        return

    s = st.session_state

    s.name = name
    s.phone = phone
    s.pid = uuid.uuid4().hex[:10]

    try:
        save_player(
            s.pid,
            {
                "الاسم": name,
                "رقم الجوال": phone,
            },
        )

        s.save_error = False

    except Exception as e:
        s.save_error = True
        print(f"Google Sheets save error: {e}")

        st.error(
            "تعذر حفظ بياناتك في Google Sheets. "
            "تأكدي من إعداد الاتصال ثم حاولي مرة أخرى."
        )
        return

    go("home")


def screen_home():
    s = st.session_state

    st.markdown(
        f"### أهلًا بكِ يا {s.name} 👋"
    )

    st.metric(
        "مجموع نقاطك",
        total(),
    )

    st.write(
        "ثلاثة مستويات لتتعلّمي أساسيات الذكاء الاصطناعي بالمرح!"
    )

    for number, title, stage in LEVELS:
        done = s.scores[number] is not None

        if st.button(
            ("✅ " if done else "") + title,
            key=f"lvl{number}",
        ):
            start_level(stage)

    if all(
        value is not None
        for value in s.scores.values()
    ):
        st.balloons()

        st.success(
            f"🏆 أنهيتِ كل المستويات يا {s.name}! "
            f"مجموع نقاطك {total()}."
        )

    st.divider()

    if st.button("🚪 إنهاء الجلسة"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]

        st.rerun()

    if s.save_error:
        st.warning(
            "تعذّر حفظ النتيجة في Google Sheets، "
            "أبلغي المشرفة."
        )


def screen_quiz(n: int):
    s = st.session_state
    questions = LEVEL1 if n == 1 else LEVEL3

    if s.qi >= len(questions):
        st.markdown(
            f"## 🎉 أحسنتِ يا {s.name}!"
        )

        st.write(
            f"إجاباتك الصحيحة: **{s.qc} من {len(questions)}**  |  "
            f"نقاط المستوى: **{s.scores[n]}**"
        )

        if st.button(
            "العودة للقائمة",
            type="primary",
        ):
            go("home")

        return

    question, options, answer, explanation = questions[s.qi]

    st.caption(
        f"المستوى {n} — سؤال {s.qi + 1} من {len(questions)}"
    )

    st.markdown(
        f"#### {question}"
    )

    for j, option in enumerate(options):
        if st.button(
            option,
            key=f"q{n}_{s.qi}_{j}",
            disabled=s.picked is not None,
        ):
            s.picked = j

            if j == answer:
                s.qc += 1

            st.rerun()

    if s.picked is not None:
        (
            st.success
            if s.picked == answer
            else st.error
        )(
            (
                "✅ إجابة صحيحة! "
                if s.picked == answer
                else "❌ ليست صحيحة. "
            )
            + explanation
        )

        last = s.qi == len(questions) - 1

        if st.button(
            "إنهاء المستوى ✔" if last else "التالي ◀",
            type="primary",
            key=f"next{n}_{s.qi}",
        ):
            s.qi += 1
            s.picked = None

            if s.qi >= len(questions):
                new_score = s.qc * 10

                s.scores[n] = max(
                    s.scores[n] or 0,
                    new_score,
                )

                persist()

            st.rerun()

    if st.button(
        "◀ القائمة",
        key=f"back{n}",
    ):
        go("home")


def draw_chart():
    s = st.session_state

    if not s.train:
        st.info(
            "أضيفي أمثلة لتظهر هنا على الرسم."
        )
        return

    color = alt.Color(
        "النوع:N",
        legend=alt.Legend(title=None),
        scale=alt.Scale(
            domain=["تفاحة", "موزة"],
            range=FRUIT_COLORS,
        ),
    )

    x = alt.X(
        "الطول:Q",
        scale=alt.Scale(domain=[0, 10]),
    )

    y = alt.Y(
        "الاستدارة:Q",
        scale=alt.Scale(domain=[0, 10]),
    )

    df = pd.DataFrame(
        s.train,
        columns=["الطول", "الاستدارة", "النوع"],
    )

    chart = (
        alt.Chart(df)
        .mark_circle(
            size=230,
            opacity=0.9,
            stroke="black",
            strokeWidth=1,
        )
        .encode(
            x=x,
            y=y,
            color=color,
        )
    )

    if s.test and "preds" in s.test:
        test_df = pd.DataFrame(
            s.test["preds"],
            columns=[
                "الطول",
                "الاستدارة",
                "النوع",
                "الحقيقي",
                "النتيجة",
            ],
        )

        squares = (
            alt.Chart(test_df)
            .mark_square(
                size=300,
                strokeWidth=4,
            )
            .encode(
                x=x,
                y=y,
                color=color,
                stroke=alt.Stroke(
                    "النتيجة:N",
                    legend=alt.Legend(
                        title="اختبار الآلة"
                    ),
                    scale=alt.Scale(
                        domain=["صحيح", "خطأ"],
                        range=["#2e9e5b", "#000000"],
                    ),
                ),
                tooltip=[
                    "الطول",
                    "الاستدارة",
                    "النوع",
                    "الحقيقي",
                    "النتيجة",
                ],
            )
        )

        chart = chart + squares

    st.altair_chart(
        chart.properties(height=320),
        use_container_width=True,
    )


def screen_train():
    s = st.session_state

    st.markdown(
        "#### 🍎🍌 درّبي الآلة على التفريق بين التفاحة والموزة"
    )

    st.write(
        "حرّكي المنزلقين لوصف الفاكهة ثم اختاري اسمها. "
        "كلما أعطيتِ الآلة أمثلة أكثر وأكثر تنوّعًا، تعلّمت أفضل!"
    )

    c1, c2 = st.columns(2)

    length = c1.slider(
        "الطول (قصير ← طويل)",
        0.0,
        10.0,
        5.0,
        0.5,
    )

    roundness = c2.slider(
        "الاستدارة (مستقيم ← مستدير)",
        0.0,
        10.0,
        5.0,
        0.5,
    )

    b1, b2, b3, b4 = st.columns(4)

    if b1.button("تفاحة 🍎"):
        s.train.append(
            (length, roundness, "تفاحة")
        )
        s.test = None

    if b2.button("موزة 🍌"):
        s.train.append(
            (length, roundness, "موزة")
        )
        s.test = None

    if b3.button("🗑️ مسح"):
        s.train = []
        s.test = None

    if b4.button(
        "اختبري ✅",
        type="primary",
    ):
        labels = {
            label
            for *_, label in s.train
        }

        if len(labels) < 2 or len(s.train) < 4:
            s.test = {
                "error": (
                    "⚠️ أضيفي 4 أمثلة على الأقل، "
                    "ومن النوعين معًا (تفاحة وموزة)."
                )
            }

        else:
            preds = []

            for x, y, true_label in TEST_SET:
                prediction = knn_predict(
                    s.train,
                    (x, y),
                )

                preds.append(
                    (
                        x,
                        y,
                        prediction,
                        true_label,
                        "صحيح"
                        if prediction == true_label
                        else "خطأ",
                    )
                )

            right = sum(
                1
                for *_, result in preds
                if result == "صحيح"
            )

            s.test = {
                "preds": preds,
                "right": right,
                "counts": {
                    label: sum(
                        1
                        for *_, train_label in s.train
                        if train_label == label
                    )
                    for label in labels
                },
            }

            # حتى 30 نقطة
            s.scores[2] = max(
                s.scores[2] or 0,
                right * 5,
            )

            persist()

    draw_chart()

    st.caption(
        f"أمثلة التدريب: {len(s.train)}"
    )

    test = s.test

    if test:
        if "error" in test:
            st.error(test["error"])

        else:
            number_of_tests = len(TEST_SET)

            message = (
                f"دقة الآلة: **{test['right']} من {number_of_tests}** "
                "(المربعات = اختبار جديد، الإطار الأخضر = تنبؤ صحيح)"
            )

            if test["right"] == number_of_tests:
                st.success(
                    "🏆 ممتاز! الآلة تعلّمت من أمثلتك. "
                    + message
                )

            else:
                st.warning(
                    message
                    + "  \nجرّبي إضافة أمثلة أكثر وأكثر "
                    "تنوّعًا قرب الحدود بين النوعين."
                )

            counts = test["counts"]

            if counts and max(counts.values()) >= 3 * min(counts.values()):
                st.info(
                    "⚖️ لاحظي: أمثلتك غير متوازنة بين النوعين، "
                    "وهذا قد يسبّب تحيّزًا في النموذج!"
                )

    if st.button(
        "◀ القائمة",
        key="back2",
    ):
        go("home")


# ============================== لوحة المشرفة ==============================

def admin_password():
    try:
        if "ADMIN_PASSWORD" in st.secrets:
            return str(st.secrets["ADMIN_PASSWORD"])
    except Exception:
        pass

    return None


def admin_panel():
    password = admin_password()

    # إذا لم يتم إعداد كلمة المرور، لا نظهر لوحة المشرفة.
    if not password:
        return

    with st.sidebar.expander("🔒 لوحة المشرفة"):
        entered = st.text_input(
            "كلمة المرور",
            type="password",
            key="admin_pw",
        )

        if not entered:
            return

        if not hmac.compare_digest(
            entered.encode(),
            password.encode(),
        ):
            st.error("كلمة المرور غير صحيحة.")
            return

        try:
            worksheet = get_google_sheet()
            records = worksheet.get_all_records()

            if not records:
                st.info("لا توجد بيانات بعد.")
                return

            data = pd.DataFrame(records)

            st.dataframe(
                data,
                use_container_width=True,
                hide_index=True,
            )

            csv_data = data.to_csv(
                index=False,
                encoding="utf-8-sig",
            )

            st.download_button(
                "⬇️ تنزيل بيانات اللاعبات CSV",
                data=csv_data,
                file_name="players.csv",
                mime="text/csv",
            )

        except Exception as e:
            st.error(
                "تعذر الاتصال بـ Google Sheets."
            )
            st.caption(
                "تحققي من GOOGLE_SHEET_ID، وبيانات Service Account، "
                "ومن مشاركة Google Sheet مع بريد Service Account."
            )
            print(f"Admin Google Sheets error: {e}")


# ============================== التشغيل ==============================

def main():
    init_state()
    admin_panel()

    stage = st.session_state.stage

    if stage != "login":
        header_bar()

    if stage == "login":
        screen_login()

    elif stage == "home":
        screen_home()

    elif stage == "quiz1":
        screen_quiz(1)

    elif stage == "quiz3":
        screen_quiz(3)

    elif stage == "train":
        screen_train()


if __name__ == "__main__":
    main()
