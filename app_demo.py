
import sqlite3
import pandas as pd
import streamlit as st
from io import BytesIO
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import arabic_reshaper
from bidi.algorithm import get_display

import os as _os
_HERE = _os.path.dirname(_os.path.abspath(__file__))
DB_PATH = _os.path.join(_HERE, "inventory_demo.db")



def create_excel_export():
    """ساخت گزارش کامل اکسل از دیتابیس آزمایشی."""

    from datetime import datetime
    from io import BytesIO
    import sqlite3

    output = BytesIO()

    connection = sqlite3.connect(
        f"file:{DB_PATH}?mode=ro",
        uri=True,
        timeout=10
    )

    try:
        products = pd.read_sql_query("""
            SELECT
                id AS "شناسه محصول",
                name AS "نام محصول",
                purchase_price AS "قیمت خرید",
                selling_price AS "قیمت فروش",
                quantity AS "موجودی فعلی",
                min_stock AS "حداقل موجودی"
            FROM products
            ORDER BY name
        """, connection)

        sales = pd.read_sql_query("""
            SELECT
                s.id AS "شناسه فروش",
                p.name AS "نام محصول",
                s.quantity AS "تعداد فروش",
                s.unit_price AS "قیمت واحد فروش",
                s.total_amount AS "مبلغ کل فروش",
                s.purchase_price_at_sale AS "بهای خرید هنگام فروش",
                s.sale_date AS "تاریخ فروش"
            FROM sales s
            LEFT JOIN products p ON p.id = s.product_id
            ORDER BY s.sale_date DESC
        """, connection)

        purchases = pd.read_sql_query("""
            SELECT
                pu.id AS "شناسه خرید",
                p.name AS "نام محصول",
                pu.quantity AS "تعداد خرید",
                pu.unit_price AS "قیمت واحد خرید",
                pu.total_amount AS "مبلغ کل خرید",
                pu.purchase_date AS "تاریخ خرید"
            FROM purchases pu
            LEFT JOIN products p ON p.id = pu.product_id
            ORDER BY pu.purchase_date DESC
        """, connection)

        profit = pd.read_sql_query("""
            SELECT
                p.name AS "نام محصول",
                COALESCE(SUM(s.quantity), 0) AS "تعداد فروخته‌شده",
                COALESCE(SUM(s.total_amount), 0) AS "درآمد فروش",
                COALESCE(
                    SUM(s.quantity * s.purchase_price_at_sale), 0
                ) AS "بهای تمام‌شده",
                COALESCE(
                    SUM(
                        s.total_amount -
                        (s.quantity * s.purchase_price_at_sale)
                    ), 0
                ) AS "سود ناخالص"
            FROM products p
            LEFT JOIN sales s ON s.product_id = p.id
            GROUP BY p.id, p.name
            ORDER BY "سود ناخالص" DESC
        """, connection)

        inventory = pd.read_sql_query("""
            SELECT
                name AS "نام محصول",
                quantity AS "موجودی فعلی",
                min_stock AS "حداقل موجودی",
                purchase_price AS "قیمت خرید واحد",
                selling_price AS "قیمت فروش واحد",
                quantity * purchase_price AS "ارزش موجودی بر اساس خرید",
                quantity * selling_price AS "ارزش موجودی بر اساس فروش",
                CASE
                    WHEN quantity <= 0 THEN 'ناموجود'
                    WHEN quantity <= min_stock THEN 'موجودی کم'
                    ELSE 'موجودی مناسب'
                END AS "وضعیت موجودی"
            FROM products
            ORDER BY quantity ASC
        """, connection)

    finally:
        connection.close()

    sheets = {
        "محصولات": products,
        "فروش‌ها": sales,
        "خریدها": purchases,
        "تحلیل سود": profit,
        "خلاصه موجودی": inventory
    }

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for sheet_name, dataframe in sheets.items():

            dataframe.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

            worksheet = writer.sheets[sheet_name]

            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            worksheet.sheet_view.rightToLeft = True

            for cell in worksheet[1]:
                cell.fill = __import__("openpyxl").styles.PatternFill(
                    fill_type="solid",
                    fgColor="17345E"
                )
                cell.font = __import__("openpyxl").styles.Font(
                    name="Arial",
                    bold=True,
                    color="FFFFFF",
                    size=11
                )
                cell.alignment = __import__("openpyxl").styles.Alignment(
                    horizontal="center",
                    vertical="center"
                )

            worksheet.row_dimensions[1].height = 30

            for column_cells in worksheet.columns:
                column_letter = column_cells[0].column_letter

                max_length = max(
                    len(str(cell.value or ""))
                    for cell in column_cells
                )

                worksheet.column_dimensions[column_letter].width = min(
                    max(max_length + 4, 14),
                    38
                )

                for cell in column_cells[1:]:
                    cell.alignment = __import__(
                        "openpyxl"
                    ).styles.Alignment(
                        horizontal="center",
                        vertical="center"
                    )

            # قالب‌بندی ستون‌های مالی
            for row in worksheet.iter_rows(min_row=2):
                for cell in row:
                    if any(
                        word in str(worksheet.cell(1, cell.column).value)
                        for word in [
                            "قیمت",
                            "مبلغ",
                            "درآمد",
                            "بهای تمام‌شده",
                            "سود",
                            "ارزش موجودی"
                        ]
                    ):
                        cell.number_format = '#,##0.00'

    output.seek(0)
    return output


def get_db_connection(*args, **kwargs):
    """ایجاد اتصال دیتابیس با فعال بودن کلیدهای خارجی."""
    connection = sqlite3.connect(*args, **kwargs)
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def create_inventory_movements_table():
    """ایجاد جدول تاریخچه تغییرات موجودی."""

    with get_db_connection(DB_PATH) as connection:

        connection.execute("""
            CREATE TABLE IF NOT EXISTS inventory_movements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER NOT NULL,
                movement_type TEXT NOT NULL,
                quantity_change INTEGER NOT NULL,
                previous_quantity INTEGER NOT NULL,
                new_quantity INTEGER NOT NULL,
                movement_date TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                description TEXT,
                source_type TEXT,
                source_id INTEGER,
                FOREIGN KEY (product_id)
                    REFERENCES products(id)
                    ON DELETE RESTRICT,
                UNIQUE (source_type, source_id)
            )
        """)

        connection.execute("""
            CREATE INDEX IF NOT EXISTS
            idx_inventory_movements_product_date
            ON inventory_movements(product_id, movement_date)
        """)


create_inventory_movements_table()



st.set_page_config(
    page_title="Smart Inventory Manager",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Vazirmatn', sans-serif;
}

.stApp {
    background-color: #f4f7fb;
}

.main .block-container {
    padding: 2rem 2.5rem;
    max-width: 1500px;
    direction: rtl;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #10213f 0%, #172f55 100%);
}

section[data-testid="stSidebar"] * {
    color: #ffffff;
}

section[data-testid="stSidebar"] .stRadio label {
    background-color: rgba(255,255,255,0.05);
    padding: 10px 14px;
    border-radius: 10px;
}

.hero {
    background: linear-gradient(120deg, #10213f, #2456a6);
    padding: 30px;
    border-radius: 20px;
    color: white;
    margin-bottom: 25px;
    box-shadow: 0 8px 24px rgba(20, 48, 90, 0.12);
}

.hero h1 {
    color: white;
    font-size: 30px;
    margin: 0 0 10px 0;
}

.hero p {
    color: #dce8ff;
    font-size: 15px;
    margin: 0;
}

.section-title {
    color: #17345e;
    font-size: 22px;
    font-weight: 700;
    margin: 25px 0 15px 0;
}

div[data-testid="stMetric"] {
    background: white;
    border: 1px solid #e5ebf4;
    padding: 20px;
    border-radius: 16px;
    box-shadow: 0 4px 14px rgba(25, 50, 90, 0.04);
}

div[data-testid="stMetricLabel"] {
    color: #17345e !important;
    font-size: 14px;
    font-weight: 600;
}

div[data-testid="stMetricValue"] {
    color: #16345d;
    font-weight: 800;
}

div[data-testid="stDataFrame"] {
    background: white;
    border-radius: 14px;
    overflow: hidden;
}

.stAlert {
    border-radius: 12px;
}

#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

header[data-testid="stHeader"] {
    background: rgba(244,247,251,0.9);
}

/* قابل مشاهده کردن دکمه باز و بسته کردن منوی کناری */
button[data-testid="stSidebarCollapseButton"],
button[data-testid="stSidebarCollapsedControl"],
[data-testid="stSidebarCollapsedControl"] button {
    visibility: visible !important;
    opacity: 1 !important;
    color: #17345e !important;
    background-color: #ffffff !important;
}

button[data-testid="stSidebarCollapseButton"] svg,
button[data-testid="stSidebarCollapsedControl"] svg,
[data-testid="stSidebarCollapsedControl"] button svg {
    visibility: visible !important;
    opacity: 1 !important;
    color: #17345e !important;
    fill: #17345e !important;
    stroke: #17345e !important;
}


/* اصلاح رنگ آیکون منوی اصلی Streamlit */
header[data-testid="stHeader"] button {
    color: #17345e !important;
    opacity: 1 !important;
}

header[data-testid="stHeader"] button svg {
    color: #17345e !important;
    opacity: 1 !important;
    fill: #17345e !important;
    stroke: #17345e !important;
}

header[data-testid="stHeader"] button svg path {
    stroke: #17345e !important;
    fill: #17345e !important;
}

header[data-testid="stHeader"] [data-testid="stSidebarCollapsedControl"] {
    color: #17345e !important;
    opacity: 1 !important;
}


/* خوانا کردن عنوان‌ها و فیلدهای بخش اصلی */
.main .block-container {
    color: #17345e !important;
}

.main .block-container h1,
.main .block-container h2,
.main .block-container h3,
.main .block-container h4,
.main .block-container p,
.main .block-container label,
.main .block-container [data-testid="stWidgetLabel"],
.main .block-container [data-testid="stMarkdownContainer"] {
    color: #17345e !important;
    opacity: 1 !important;
}

.main .block-container input,
.main .block-container textarea {
    color: #17345e !important;
    background-color: #ffffff !important;
    opacity: 1 !important;
}

.main .block-container [data-baseweb="select"] * {
    color: #17345e !important;
    opacity: 1 !important;
}

</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=5)
def load_products():
    with get_db_connection(DB_PATH) as connection:
        return pd.read_sql_query("""
            SELECT id, name, purchase_price, selling_price,
                   quantity, min_stock
            FROM products
            ORDER BY name
        """, connection)


@st.cache_data(ttl=5)
def load_sales_summary():
    with get_db_connection(DB_PATH) as connection:
        return pd.read_sql_query("""
            SELECT
                COALESCE(SUM(total_amount), 0) AS revenue,
                COALESCE(
                    SUM(total_amount - purchase_price_at_sale * quantity),
                    0
                ) AS profit,
                COALESCE(SUM(quantity), 0) AS units_sold
            FROM sales
        """, connection).iloc[0]


st.markdown(
    "<div style='direction:rtl;text-align:center;background:#fff4d6;"
    "color:#7a5a00;border:1px solid #f0d58a;border-radius:10px;"
    "padding:8px 12px;margin-bottom:12px;font-size:14px;'>"
    "🧪 نسخه نمایشی؛ داده‌ها نمونه هستند و با اطلاعات واقعی ارتباطی ندارند"
    "</div>",
    unsafe_allow_html=True
)


def reset_demo_database():
    """بازسازی دیتابیس دمو با داده نمونه (فقط برای inventory_demo.db)."""
    import os, sqlite3
    if not DB_PATH.endswith("inventory_demo.db"):
        raise RuntimeError("بازنشانی فقط برای دیتابیس دمو مجاز است.")
    tmp = DB_PATH + ".tmp"
    if os.path.exists(tmp):
        os.remove(tmp)
    con = sqlite3.connect(tmp)
    try:
        con.execute("PRAGMA foreign_keys = ON")
        con.executescript("""
CREATE TABLE products (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, purchase_price REAL NOT NULL, selling_price REAL NOT NULL, quantity INTEGER NOT NULL DEFAULT 0, min_stock INTEGER NOT NULL DEFAULT 5);
CREATE TABLE sales (id INTEGER PRIMARY KEY AUTOINCREMENT, product_id INTEGER NOT NULL, quantity INTEGER NOT NULL CHECK(quantity > 0), unit_price REAL NOT NULL, total_amount REAL NOT NULL, sale_date TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, purchase_price_at_sale REAL, FOREIGN KEY (product_id) REFERENCES products(id));
CREATE TABLE purchases (id INTEGER PRIMARY KEY AUTOINCREMENT, product_id INTEGER NOT NULL, quantity INTEGER NOT NULL CHECK (quantity > 0), unit_price REAL NOT NULL CHECK (unit_price >= 0), total_amount REAL NOT NULL CHECK (total_amount >= 0), purchase_date TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY (product_id) REFERENCES products(id));
CREATE TABLE inventory_movements (id INTEGER PRIMARY KEY AUTOINCREMENT, product_id INTEGER NOT NULL, movement_type TEXT NOT NULL, quantity_change INTEGER NOT NULL, previous_quantity INTEGER NOT NULL, new_quantity INTEGER NOT NULL, movement_date TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, description TEXT, source_type TEXT, source_id INTEGER, FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE RESTRICT, UNIQUE (source_type, source_id));
CREATE TABLE customers (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, phone TEXT DEFAULT '', email TEXT DEFAULT '', address TEXT DEFAULT '', notes TEXT DEFAULT '', created_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE suppliers (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, phone TEXT DEFAULT '', email TEXT DEFAULT '', address TEXT DEFAULT '', notes TEXT DEFAULT '', created_at TEXT DEFAULT CURRENT_TIMESTAMP);
""")
        items = [
            ("برنج ایرانی ۱۰ کیلویی", 850000, 950000, 5),
            ("روغن مایع ۱.۸ لیتری", 320000, 365000, 8),
            ("چای ممتاز ۵۰۰ گرمی", 210000, 245000, 6),
            ("شکر ۱ کیلویی", 48000, 55000, 10),
            ("ماکارونی ۵۰۰ گرمی", 32000, 38000, 10),
            ("رب گوجه ۸۰۰ گرمی", 95000, 110000, 5),
        ]
        con.executemany("INSERT INTO products(name,purchase_price,selling_price,quantity,min_stock) VALUES(?,?,?,0,?)",
                        [(n, b, s, m) for n, b, s, m in items])

        def move(pid, mtype, change, desc, st_, sid, d):
            prev = con.execute("SELECT quantity FROM products WHERE id=?", (pid,)).fetchone()[0]
            con.execute("UPDATE products SET quantity=? WHERE id=?", (prev + change, pid))
            con.execute("INSERT INTO inventory_movements(product_id,movement_type,quantity_change,previous_quantity,new_quantity,movement_date,description,source_type,source_id) VALUES(?,?,?,?,?,?,?,?,?)",
                        (pid, mtype, change, prev, prev + change, d, desc, st_, sid))

        for pid, q, d in [(1,30,"2026-10-01 09:00:00"),(2,40,"2026-10-01 09:10:00"),(3,25,"2026-10-02 10:00:00"),
                          (4,60,"2026-10-02 10:15:00"),(5,50,"2026-10-03 11:00:00"),(6,20,"2026-10-03 11:20:00")]:
            p = items[pid-1][1]
            cur = con.execute("INSERT INTO purchases(product_id,quantity,unit_price,total_amount,purchase_date) VALUES(?,?,?,?,?)", (pid, q, p, q*p, d))
            move(pid, "purchase", q, "خرید نمونه", "purchase", cur.lastrowid, d)

        for pid, q, d in [(1,10,"2026-10-10 12:00:00"),(2,15,"2026-10-11 13:00:00"),(3,8,"2026-10-12 14:00:00"),
                          (4,52,"2026-10-13 15:00:00"),(5,20,"2026-10-14 16:00:00"),(1,12,"2026-10-20 12:30:00"),
                          (6,6,"2026-10-22 17:00:00"),(2,10,"2026-10-25 18:00:00")]:
            b, s = items[pid-1][1], items[pid-1][2]
            cur = con.execute("INSERT INTO sales(product_id,quantity,unit_price,total_amount,sale_date,purchase_price_at_sale) VALUES(?,?,?,?,?,?)", (pid, q, s, q*s, d, b))
            move(pid, "sale", -q, "فروش نمونه", "sale", cur.lastrowid, d)

        con.executemany("INSERT INTO customers(name,phone,email,address,notes) VALUES(?,?,?,?,?)", [
            ("علی رضایی","09120000001","","تهران، نمونه","مشتری نمونه"),
            ("مریم احمدی","09120000002","","اصفهان، نمونه","مشتری نمونه"),
            ("فروشگاه نمونه آفتاب","09120000003","","شیراز، نمونه","مشتری نمونه")])
        con.executemany("INSERT INTO suppliers(name,phone,email,address,notes) VALUES(?,?,?,?,?)", [
            ("شرکت پخش نمونه الف","02100000001","","تهران، نمونه","تأمین‌کننده نمونه"),
            ("بازرگانی نمونه ب","03100000002","","اصفهان، نمونه","تأمین‌کننده نمونه")])
        con.commit()
    except Exception:
        con.close()
        if os.path.exists(tmp):
            os.remove(tmp)
        raise
    con.close()
    os.replace(tmp, DB_PATH)


def _demo_db_ready():
    import sqlite3
    if not _os.path.exists(DB_PATH):
        return False
    try:
        con = sqlite3.connect(DB_PATH)
        ok = con.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='products'"
        ).fetchone()
        con.close()
        return bool(ok)
    except Exception:
        return False


if not _demo_db_ready():
    reset_demo_database()

with st.sidebar:
    st.markdown(
        "<h2 style='text-align:center;'>📦 Smart Inventory</h2>",
        unsafe_allow_html=True
    )
    st.markdown(
        "<p style='text-align:center;color:#c8d8f2;'>مدیریت هوشمند فروشگاه</p>",
        unsafe_allow_html=True
    )
    st.divider()

    page = st.radio(
        "منوی اصلی",
        ["🏡 صفحه اصلی", "🏠 داشبورد", "📦 محصولات", "🛒 ثبت فروش", "📥 ثبت خرید", "🧾 فاکتور فروش", "📊 گزارش‌های فروش", "👥 مشتریان", "🏢 تأمین‌کنندگان"],
        label_visibility="visible",
        key="main_navigation"
    )

    st.divider()
    st.caption("نسخه اولیه مدیریتی")

    st.divider()
    confirm_reset = st.checkbox("مطمئن هستم", key="demo_reset_confirm")
    if st.button("🔄 بازنشانی داده‌های دمو", disabled=not confirm_reset, key="demo_reset_btn"):
        try:
            reset_demo_database()
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"بازنشانی انجام نشد: {e}")


try:
    products = load_products()
    sales = load_sales_summary()

    low_stock = products[
        products["quantity"] <= products["min_stock"]
    ]

    if page == "🏡 صفحه اصلی":
        st.markdown("""
        <div style="
            background: linear-gradient(120deg, #10213f, #2864a8);
            padding: 34px 30px;
            border-radius: 20px;
            color: white;
            text-align: center;
            margin-bottom: 25px;
        ">
            <div style="font-size: 48px;">📦</div>
            <h1 style="color:white; margin: 8px 0;">
                Smart Inventory Manager
            </h1>
            <h3 style="color:#e1ecff; font-weight:500;">
                سیستم هوشمند مدیریت موجودی و فروش
            </h3>
            <p style="color:#e1ecff; font-size:16px;">
                مدیریت محصولات، بررسی موجودی و مشاهده عملکرد فروشگاه
                در یک محیط ساده و منظم.
            </p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(
            '<div style="font-family: Vazirmatn, sans-serif; '
            'color: #17345e; font-size: 22px; font-weight: 700; '
            'margin: 25px 0 15px 0;">'
            'خلاصه وضعیت فروشگاه'
            '</div>',
            unsafe_allow_html=True
        )

        home_col1, home_col2, home_col3, home_col4 = st.columns(4)

        with home_col1:
            st.markdown(
                '<div style="font-family: Vazirmatn, sans-serif; '
                'color: #17345e; font-size: 15px; font-weight: 600; '
                'margin-bottom: 6px;">📦 تعداد محصولات</div>',
                unsafe_allow_html=True
            )
            st.metric(
                "تعداد محصولات",
                f"{len(products):,}",
                label_visibility="collapsed"
            )

        with home_col2:
            st.markdown(
                '<div style="font-family: Vazirmatn, sans-serif; '
                'color: #17345e; font-size: 15px; font-weight: 600; '
                'margin-bottom: 6px;">🗃️ مجموع موجودی</div>',
                unsafe_allow_html=True
            )
            st.metric(
                "مجموع موجودی",
                f"{int(products['quantity'].sum()):,}",
                label_visibility="collapsed"
            )

        with home_col3:
            st.markdown(
                '<div style="font-family: Vazirmatn, sans-serif; '
                'color: #17345e; font-size: 15px; font-weight: 600; '
                'margin-bottom: 6px;">💰 مجموع فروش (افغانی)</div>',
                unsafe_allow_html=True
            )
            st.metric(
                "مجموع فروش (افغانی)",
                f"{float(sales['revenue']):,.2f}",
                label_visibility="collapsed"
            )

        with home_col4:
            st.markdown(
                '<div style="font-family: Vazirmatn, sans-serif; '
                'color: #17345e; font-size: 15px; font-weight: 600; '
                'margin-bottom: 6px;">📈 سود ناخالص (افغانی)</div>',
                unsafe_allow_html=True
            )
            st.metric(
                "سود ناخالص (افغانی)",
                f"{float(sales['profit']):,.2f}",
                label_visibility="collapsed"
            )

        st.markdown(
            '<div style="'
            'color:#10213f !important; '
            '-webkit-text-fill-color:#10213f !important; '
            'opacity:1 !important; '
            'font-family:Vazirmatn,sans-serif; '
            'font-size:22px; font-weight:800; '
            'margin:25px 0 15px 0; '
            'text-align:right;">'
            'ورود به بخش‌های سیستم'
            '</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="section-title">'
            'ورود سریع به بخش‌های سیستم'
            '</div>',
            unsafe_allow_html=True
        )

        shortcuts = [
            (
                "📊", "داشبورد مدیریتی",
                "مشاهده وضعیت انبار و عملکرد فروشگاه",
                "🏠 داشبورد", "#e5efff", "#17345e",
                "ورود به داشبورد", "home_dashboard_button"
            ),
            (
                "📦", "مدیریت محصولات",
                "مشاهده و مدیریت محصولات فروشگاه",
                "📦 محصولات", "#e4f6ed", "#17633d",
                "مدیریت محصولات", "home_products_button"
            ),
            (
                "🛒", "ثبت فروش",
                "ثبت فروش جدید و به‌روزرسانی موجودی",
                "🛒 ثبت فروش", "#fff0df", "#914d0c",
                "رفتن به ثبت فروش", "home_sales_button"
            ),
            (
                "📥", "ثبت خرید",
                "ثبت خریدهای جدید فروشگاه",
                "📥 ثبت خرید", "#e8e7ff", "#40358a",
                "رفتن به ثبت خرید", "home_purchases_button"
            ),
            (
                "🧾", "فاکتور فروش",
                "دسترسی به بخش فاکتورها",
                "🧾 فاکتور فروش", "#fce7f3", "#9d174d",
                "رفتن به فاکتورها", "home_invoice_button"
            ),
            (
                "📈", "گزارش‌های فروش",
                "بررسی میزان فروش و سود فروشگاه",
                "📊 گزارش‌های فروش", "#e0f2fe", "#075985",
                "مشاهده گزارش‌ها", "home_reports_button"
            )
        ]

        for row_start in range(0, len(shortcuts), 2):
            card_columns = st.columns(2)

            for card_col, item in zip(
                card_columns,
                shortcuts[row_start:row_start + 2]
            ):
                (
                    icon, title, description, target,
                    background, title_color, button_text, button_key
                ) = item

                with card_col:
                    st.markdown(
                        f"""
                        <div class="home-shortcut-card" style="
                            background:{background};
                            padding:20px 14px;
                            border-radius:16px;
                            text-align:center;
                            margin:8px 0 10px 0;
                            min-height:155px;
                            border:1px solid #b8c5d8;
                        ">
                            <div style="
                                font-size:34px;
                                opacity:1;
                            ">{icon}</div>
                            <div class="section-title" style="
                                color:#17345e;
                                font-family:Vazirmatn,sans-serif;
                                font-size:22px;
                                font-weight:700;
                                opacity:1;
                                margin:8px 0;
                            ">{title}</div>
                            <div class="section-title" style="
                                font-family:Vazirmatn,sans-serif;
                                font-size:16px;
                                font-weight:600;
                                line-height:1.9;
                                margin:8px 0;
                                text-align:center;
                            ">{description}</div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    st.button(
                        button_text,
                        key=button_key,
                        use_container_width=True,
                        on_click=lambda target=target: (
                            st.session_state.update(
                                main_navigation=target
                            )
                        )
                    )

    else:
        st.markdown("""
        <div class="hero">
            <h1>داشبورد مدیریت فروشگاه</h1>
            <p>نمای کلی موجودی، فروش و عملکرد کسب‌وکار شما</p>
        </div>
        """, unsafe_allow_html=True)

    if page == "🏠 داشبورد":
        st.markdown(
            '<div class="section-title">خلاصه عملکرد</div>',
            unsafe_allow_html=True
        )

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown(
                '<div style="font-family: Vazirmatn, sans-serif; '
                'color: #17345e; font-size: 15px; font-weight: 600; '
                'margin-bottom: 6px;">📦 تعداد محصولات</div>',
                unsafe_allow_html=True
            )
            st.metric(
                "تعداد محصولات",
                f"{len(products):,}",
                label_visibility="collapsed"
            )

        with col2:
            st.markdown(
                '<div style="font-family: Vazirmatn, sans-serif; '
                'color: #17345e; font-size: 15px; font-weight: 600; '
                'margin-bottom: 6px;">🗃️ مجموع موجودی</div>',
                unsafe_allow_html=True
            )
            st.metric(
                "مجموع موجودی",
                f"{int(products['quantity'].sum()):,}",
                label_visibility="collapsed"
            )

        with col3:
            st.markdown(
                '<div style="font-family: Vazirmatn, sans-serif; '
                'color: #17345e; font-size: 15px; font-weight: 600; '
                'margin-bottom: 6px;">💰 مجموع فروش</div>',
                unsafe_allow_html=True
            )
            st.metric(
                "مجموع فروش",
                f"{float(sales['revenue']):,.2f}",
                label_visibility="collapsed"
            )

        with col4:
            st.markdown(
                '<div style="font-family: Vazirmatn, sans-serif; '
                'color: #17345e; font-size: 15px; font-weight: 600; '
                'margin-bottom: 6px;">📈 سود ناخالص</div>',
                unsafe_allow_html=True
            )
            st.metric(
                "سود ناخالص",
                f"{float(sales['profit']):,.2f}",
                label_visibility="collapsed"
            )

        st.markdown(
            '<div class="section-title">وضعیت انبار</div>',
            unsafe_allow_html=True
        )

        left, right = st.columns([3, 2])

        with left:
            if not products.empty:
                chart_data = products.set_index("name")["quantity"]
                st.bar_chart(
                    chart_data,
                    color="#3478d4",
                    height=330
                )
            else:
                st.info("هنوز محصولی ثبت نشده است.")

        with right:
            st.markdown(
                '<div style="font-family: Vazirmatn, sans-serif; '
                'color: #17345e; font-size: 15px; font-weight: 600; '
                'margin-bottom: 6px;">⚠️ محصولات با موجودی کم</div>',
                unsafe_allow_html=True
            )
            st.metric(
                "محصولات با موجودی کم",
                len(low_stock),
                label_visibility="collapsed"
            )

            if not low_stock.empty:
                st.warning("برخی محصولات به حداقل موجودی رسیده‌اند.")

                st.dataframe(
                    low_stock[["name", "quantity", "min_stock"]]
                    .rename(columns={
                        "name": "نام محصول",
                        "quantity": "موجودی",
                        "min_stock": "حداقل موجودی"
                    }),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.success("موجودی محصولات در حداقل تعیین‌شده قرار دارد.")

        st.markdown(
            '<div class="section-title">آخرین نمای محصولات</div>',
            unsafe_allow_html=True
        )

        st.dataframe(
            products.rename(columns={
                "id": "شناسه",
                "name": "نام محصول",
                "purchase_price": "قیمت خرید هر واحد (افغانی ؋)",
                "selling_price": "قیمت فروش هر واحد (افغانی ؋)",
                "quantity": "موجودی",
                "min_stock": "حداقل موجودی"
            }),
            use_container_width=True,
            hide_index=True
        )

    elif page == "📦 محصولات":
        st.markdown(
            '<div class="section-title">مدیریت محصولات</div>',
            unsafe_allow_html=True
        )


        st.markdown(
            '<div style="font-family: Vazirmatn, sans-serif; '
            'color: #17345e; font-size: 22px; font-weight: 700; '
            'margin: 25px 0 15px 0;">'
            '➕ ثبت محصول جدید'
            '</div>',
            unsafe_allow_html=True
        )

        with st.form("add_product_form"):
            st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">نام محصول</div>', unsafe_allow_html=True)
            new_name = st.text_input(
                "نام محصول",
                max_chars=150,
                label_visibility="collapsed",
            )

            col_add1, col_add2 = st.columns(2)

            with col_add1:
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">قیمت خرید هر واحد (افغانی)</div>', unsafe_allow_html=True)
                new_purchase = st.number_input(
                    "قیمت خرید هر واحد (افغانی ؋)",
                    min_value=0.0,
                    value=0.0,
                    step=1.0,
                    format="%.2f",
                    label_visibility="collapsed",
                )

                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">تعداد موجودی اولیه (عدد)</div>', unsafe_allow_html=True)
                new_quantity = st.number_input(
                    "تعداد موجودی اولیه (عدد)",
                    min_value=0,
                    value=0,
                    step=1,
                    label_visibility="collapsed",
                )

            with col_add2:
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">قیمت فروش هر واحد (افغانی)</div>', unsafe_allow_html=True)
                new_selling = st.number_input(
                    "قیمت فروش هر واحد (افغانی ؋)",
                    min_value=0.0,
                    value=0.0,
                    step=1.0,
                    format="%.2f",
                    label_visibility="collapsed",
                )

                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">حداقل موجودی برای هشدار (عدد)</div>', unsafe_allow_html=True)
                new_min_stock = st.number_input(
                    "حداقل موجودی برای هشدار (عدد)",
                    min_value=0,
                    value=5,
                    step=1,
                    label_visibility="collapsed",
                )

            add_submitted = st.form_submit_button(
                "ثبت محصول",
                use_container_width=True
            )

        if add_submitted:
            product_name = new_name.strip()

            if not product_name:
                st.error("نام محصول را وارد کنید.")

            elif len(product_name) > 150:
                st.error("نام محصول بیش از حد طولانی است.")

            else:
                try:
                    with get_db_connection(DB_PATH) as connection:
                        existing = connection.execute(
                            """
                            SELECT id
                            FROM products
                            WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))
                            LIMIT 1
                            """,
                            (product_name,)
                        ).fetchone()

                        if existing:
                            st.error(
                                "محصولی با این نام قبلاً ثبت شده است."
                            )

                        else:
                            connection.execute(
                                """
                                INSERT INTO products (
                                    name,
                                    purchase_price,
                                    selling_price,
                                    quantity,
                                    min_stock
                                )
                                VALUES (?, ?, ?, ?, ?)
                                """,
                                (
                                    product_name,
                                    float(new_purchase),
                                    float(new_selling),
                                    int(new_quantity),
                                    int(new_min_stock)
                                )
                            )

                            connection.commit()

                            load_products.clear()
                            st.session_state[
                                "product_added_message"
                            ] = "محصول با موفقیت ثبت شد."
                            st.rerun()

                except sqlite3.Error as error:
                    st.error("ثبت محصول ناموفق بود.")
                    st.code(str(error))

        if st.session_state.get("product_added_message"):
            st.success(
                st.session_state.pop("product_added_message")
            )

        st.divider()

        st.markdown(
            '<div style="font-family: Vazirmatn, sans-serif; '
            'color: #17345e; font-size: 15px; font-weight: 600; '
            'margin-bottom: 6px;">🔍 جست‌وجوی محصولات</div>',
            unsafe_allow_html=True
        )

        search = st.text_input(
            "جست‌وجوی محصولات",
            placeholder="نام محصول را وارد کنید...",
            label_visibility="collapsed"
        )

        filtered_products = products[
            products["name"].str.contains(
                search,
                case=False,
                na=False
            )
        ]

        st.dataframe(
            filtered_products.rename(columns={
                "id": "شناسه",
                "name": "نام محصول",
                "purchase_price": "قیمت خرید",
                "selling_price": "قیمت فروش",
                "quantity": "موجودی",
                "min_stock": "حداقل موجودی"
            }),
            use_container_width=True,
            hide_index=True
        )

        st.caption(f"تعداد نتایج: {len(filtered_products)}")

        # --- فرم ویرایش محصول ---

        if st.session_state.get("product_updated_message"):
            st.success(
                st.session_state.pop("product_updated_message")
            )

        st.divider()
        st.markdown(
            '<div style="font-family: Vazirmatn, sans-serif; '
            'color: #17345e; font-size: 22px; font-weight: 700; '
            'margin: 25px 0 15px 0;">'
            '✏️ ویرایش محصول'
            '</div>',
            unsafe_allow_html=True
        )

        if products.empty:
            st.info("برای ویرایش، ابتدا باید محصولی ثبت کنید.")

        else:
            product_options = products["id"].astype(int).tolist()

            selected_product_id = st.selectbox(
                "محصول موردنظر را انتخاب کنید",
                options=product_options,
                format_func=lambda product_id: (
                    products.loc[
                        products["id"] == product_id, "name"
                    ].iloc[0]
                ),
                key="edit_product_selector"
            )

            selected_product = products.loc[
                products["id"] == selected_product_id
            ].iloc[0]

            with st.form("edit_product_form"):
                st.markdown(
                    '<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 2px;">'
                    '📝 نام محصول</div>',
                    unsafe_allow_html=True
                )

                edit_name = st.text_input(
                    "نام محصول",
                    value=str(selected_product["name"]),
                    max_chars=150,
                    label_visibility="collapsed"
                )

                edit_col1, edit_col2 = st.columns(2)

                with edit_col1:
                    st.markdown(
                        '<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 2px;">'
                        '💰 قیمت خرید هر واحد</div>',
                        unsafe_allow_html=True
                    )

                    edit_purchase = st.number_input(
                        "قیمت خرید هر واحد",
                        min_value=0.0,
                        value=float(selected_product["purchase_price"]),
                        step=1.0,
                        format="%.2f",
                        label_visibility="collapsed"
                    )

                    st.markdown(
                        '<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 2px;">'
                        '📦 تعداد موجودی انبار</div>',
                        unsafe_allow_html=True
                    )

                    edit_quantity = st.number_input(
                        "تعداد موجودی انبار",
                        min_value=0,
                        value=int(selected_product["quantity"]),
                        step=1,
                        label_visibility="collapsed"
                    )

                with edit_col2:
                    st.markdown(
                        '<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 2px;">'
                        '💵 قیمت فروش هر واحد</div>',
                        unsafe_allow_html=True
                    )

                    edit_selling = st.number_input(
                        "قیمت فروش هر واحد",
                        min_value=0.0,
                        value=float(selected_product["selling_price"]),
                        step=1.0,
                        format="%.2f",
                        label_visibility="collapsed"
                    )

                    st.markdown(
                        '<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 2px;">'
                        '⚠️ حداقل موجودی برای هشدار</div>',
                        unsafe_allow_html=True
                    )

                    edit_min_stock = st.number_input(
                        "حداقل موجودی برای هشدار",
                        min_value=0,
                        value=int(selected_product["min_stock"]),
                        step=1,
                        label_visibility="collapsed"
                    )

                edit_submitted = st.form_submit_button(
                    "💾 ذخیره تغییرات",
                    use_container_width=True
                )

            if edit_submitted:
                cleaned_name = edit_name.strip()

                if not cleaned_name:
                    st.error("نام محصول را وارد کنید.")

                elif len(cleaned_name) > 150:
                    st.error("نام محصول بیش از حد طولانی است.")

                else:
                    try:
                        duplicate = None
                        updated_rows = 0

                        with get_db_connection(
                            DB_PATH,
                            timeout=10
                        ) as connection:

                            connection.execute("BEGIN IMMEDIATE")

                            current_product = connection.execute(
                                """
                                SELECT quantity
                                FROM products
                                WHERE id = ?
                                """,
                                (int(selected_product_id),)
                            ).fetchone()

                            if current_product is None:
                                raise ValueError(
                                    "محصول انتخاب‌شده دیگر وجود ندارد."
                                )

                            previous_stock = int(current_product[0])
                            new_stock = int(edit_quantity)
                            quantity_change = new_stock - previous_stock

                            duplicate = connection.execute(
                                """
                                SELECT id
                                FROM products
                                WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))
                                  AND id != ?
                                LIMIT 1
                                """,
                                (
                                    cleaned_name,
                                    int(selected_product_id)
                                )
                            ).fetchone()

                            if duplicate is None:

                                cursor = connection.execute(
                                    """
                                    UPDATE products
                                    SET name = ?,
                                        purchase_price = ?,
                                        selling_price = ?,
                                        quantity = ?,
                                        min_stock = ?
                                    WHERE id = ?
                                    """,
                                    (
                                        cleaned_name,
                                        float(edit_purchase),
                                        float(edit_selling),
                                        new_stock,
                                        int(edit_min_stock),
                                        int(selected_product_id)
                                    )
                                )

                                updated_rows = cursor.rowcount

                                if updated_rows == 1 and quantity_change != 0:

                                    connection.execute(
                                        """
                                        INSERT INTO inventory_movements (
                                            product_id,
                                            movement_type,
                                            quantity_change,
                                            previous_quantity,
                                            new_quantity,
                                            movement_date,
                                            description,
                                            source_type,
                                            source_id
                                        )
                                        VALUES (
                                            ?, 'adjustment', ?, ?, ?,
                                            CURRENT_TIMESTAMP, NULL,
                                            'adjustment', NULL
                                        )
                                        """,
                                        (
                                            int(selected_product_id),
                                            quantity_change,
                                            previous_stock,
                                            new_stock
                                        )
                                    )

                            if duplicate is None and updated_rows == 1:
                                connection.commit()
                            else:
                                connection.rollback()

                        # این قسمت پس از پایان بلوک اتصال دیتابیس اجرا می‌شود.
                        if duplicate:
                            st.error(
                                "محصول دیگری با این نام وجود دارد."
                            )

                        elif updated_rows == 1:
                            load_products.clear()

                            st.session_state[
                                "product_updated_message"
                            ] = "تغییرات محصول با موفقیت ذخیره شد."

                            st.rerun()

                        else:
                            st.error(
                                "تغییری در محصول ذخیره نشد. "
                                "شناسه محصول را بررسی کنید."
                            )

                    except sqlite3.Error as error:
                        st.error("ذخیره تغییرات ناموفق بود.")
                        st.code(str(error))






    elif page == "📥 ثبت خرید":
        st.markdown(
            """
            <h2 style="
                color: #17345e;
                opacity: 1;
                font-size: 26px;
                font-weight: 700;
                margin-bottom: 20px;
            ">📥 ثبت خرید کالا</h2>
            """,
            unsafe_allow_html=True
        )

        if "purchase_success_message" in st.session_state:
            st.success(
                st.session_state.pop("purchase_success_message")
            )

        purchase_products = load_products()

        if purchase_products.empty:
            st.info("ابتدا یک محصول در بخش محصولات ثبت کنید.")
        else:
            product_labels = {
                int(row["id"]):
                    f'{row["name"]} — موجودی فعلی: {int(row["quantity"])}'
                for _, row in purchase_products.iterrows()
            }

            st.markdown(
                '<p style="color:#17345e;opacity:1;font-weight:700;'
                'font-size:16px;">انتخاب محصول</p>',
                unsafe_allow_html=True
            )

            selected_purchase_product_id = st.selectbox(
                "محصول موردنظر را انتخاب کنید",
                options=list(product_labels.keys()),
                format_func=lambda product_id: product_labels[product_id],
                key="purchase_product_id",
                label_visibility="collapsed"
            )

            selected_product = purchase_products.loc[
                purchase_products["id"] == selected_purchase_product_id
            ].iloc[0]

            st.markdown(
                f"""
                <p style="
                    color: #17345e;
                    opacity: 1;
                    font-size: 16px;
                    font-weight: 600;
                    margin-top: 12px;
                ">
                    قیمت خرید فعلی هر واحد:
                    {float(selected_product['purchase_price']):,.2f}
                    افغانی
                </p>
                """,
                unsafe_allow_html=True
            )

            with st.form("purchase_form"):
                st.markdown(
                    '<p style="color:#17345e;opacity:1;'
                    'font-weight:700;font-size:16px;">'
                    'تعداد خریداری‌شده</p>',
                    unsafe_allow_html=True
                )

                purchase_quantity = st.number_input(
                    "تعداد خریداری‌شده",
                    min_value=1,
                    value=1,
                    step=1,
                    key=f"purchase_quantity_{selected_purchase_product_id}",
                    label_visibility="collapsed"
                )

                st.markdown(
                    '<p style="color:#17345e;opacity:1;'
                    'font-weight:700;font-size:16px;">'
                    'قیمت خرید هر واحد (افغانی)</p>',
                    unsafe_allow_html=True
                )

                purchase_unit_price = st.number_input(
                    "قیمت خرید هر واحد (افغانی)",
                    min_value=0.0,
                    value=float(selected_product["purchase_price"]),
                    step=1.0,
                    format="%.2f",
                    key=f"purchase_price_{selected_purchase_product_id}",
                    label_visibility="collapsed"
                )

                purchase_submitted = st.form_submit_button(
                    "ثبت خرید",
                    use_container_width=True
                )

            if purchase_submitted:
                try:
                    quantity_to_buy = int(purchase_quantity)
                    unit_purchase_price = round(
                        float(purchase_unit_price), 2
                    )

                    if quantity_to_buy < 1:
                        raise ValueError("تعداد خرید باید حداقل یک باشد.")

                    if unit_purchase_price < 0:
                        raise ValueError("قیمت خرید نمی‌تواند منفی باشد.")

                    total_purchase_amount = round(
                        quantity_to_buy * unit_purchase_price, 2
                    )

                    with get_db_connection(DB_PATH, timeout=10) as connection:
                        connection.execute("BEGIN IMMEDIATE")

                        product_check = connection.execute(
                            "SELECT id FROM products WHERE id = ?",
                            (selected_purchase_product_id,)
                        ).fetchone()

                        if product_check is None:
                            raise ValueError(
                                "محصول انتخاب‌شده دیگر وجود ندارد."
                            )

                        previous_stock_row = connection.execute(
                            "SELECT quantity FROM products WHERE id = ?",
                            (selected_purchase_product_id,)
                        ).fetchone()

                        previous_stock = int(previous_stock_row[0])

                        purchase_cursor = connection.execute(
                            """
                            INSERT INTO purchases (
                                product_id,
                                quantity,
                                unit_price,
                                total_amount
                            )
                            VALUES (?, ?, ?, ?)
                            """,
                            (
                                selected_purchase_product_id,
                                quantity_to_buy,
                                unit_purchase_price,
                                total_purchase_amount
                            )
                        )

                        update_result = connection.execute(
                            """
                            UPDATE products
                            SET quantity = quantity + ?,
                                purchase_price = ?
                            WHERE id = ?
                            """,
                            (
                                quantity_to_buy,
                                unit_purchase_price,
                                selected_purchase_product_id
                            )
                        )

                        if update_result.rowcount != 1:
                            raise ValueError(
                                "به‌روزرسانی موجودی محصول ناموفق بود."
                            )

                        new_stock = previous_stock + quantity_to_buy

                        connection.execute(
                            """
                            INSERT INTO inventory_movements (
                                product_id,
                                movement_type,
                                quantity_change,
                                previous_quantity,
                                new_quantity,
                                movement_date,
                                description,
                                source_type,
                                source_id
                            )
                            VALUES (?, 'purchase', ?, ?, ?, CURRENT_TIMESTAMP,
                                    ?, 'purchase', ?)
                            """,
                            (
                                selected_purchase_product_id,
                                quantity_to_buy,
                                previous_stock,
                                new_stock,
                                "ثبت خرید جدید",
                                purchase_cursor.lastrowid
                            )
                        )

                        connection.commit()

                    load_products.clear()

                    st.session_state["purchase_success_message"] = (
                        f"خرید {quantity_to_buy} واحد ثبت شد. "
                        f"مبلغ کل: {total_purchase_amount:,.2f} افغانی. "
                        "موجودی و آخرین قیمت خرید به‌روزرسانی شدند."
                    )

                    st.rerun()

                except ValueError as error:
                    st.error(str(error))

                except sqlite3.Error as error:
                    st.error("ثبت خرید در دیتابیس ناموفق بود.")
                    st.code(str(error))


        st.markdown(
            """
            <h3 style="
                color: #17345e;
                opacity: 1;
                font-size: 22px;
                font-weight: 700;
                margin-top: 32px;
                margin-bottom: 16px;
            ">📋 تاریخچه خریدها</h3>
            """,
            unsafe_allow_html=True
        )

        try:
            with get_db_connection(DB_PATH) as connection:
                product_rows = connection.execute(
                    """
                    SELECT DISTINCT p.name
                    FROM purchases pu
                    LEFT JOIN products p ON p.id = pu.product_id
                    WHERE p.name IS NOT NULL
                    ORDER BY p.name
                    """
                ).fetchall()

            product_options = ["همه محصولات"] + [
                row[0] for row in product_rows
            ]

            st.markdown(
                """
                <div style="
                    color: #17345e;
                    font-size: 16px;
                    font-weight: 600;
                    margin-top: 12px;
                    margin-bottom: 6px;
                ">🔍 فیلتر تاریخچه بر اساس محصول</div>
                """,
                unsafe_allow_html=True
            )

            selected_product = st.selectbox(
                "انتخاب محصول",
                options=product_options,
                label_visibility="collapsed",
                key="purchase_history_filter"
            )

            with get_db_connection(DB_PATH) as connection:
                purchase_history = pd.read_sql_query(
                    """
                    SELECT
                        pu.id AS purchase_id,
                        p.name AS product_name,
                        pu.quantity AS quantity,
                        pu.unit_price AS unit_price,
                        pu.total_amount AS total_amount,
                        pu.purchase_date AS purchase_date
                    FROM purchases pu
                    LEFT JOIN products p ON p.id = pu.product_id
                    WHERE (? = 'همه محصولات' OR p.name = ?)
                    ORDER BY pu.purchase_date DESC, pu.id DESC
                    """,
                    connection,
                    params=(selected_product, selected_product)
                )

            if purchase_history.empty:
                st.info("هنوز خریدی ثبت نشده است.")
            else:
                purchase_history = purchase_history.rename(
                    columns={
                        "purchase_id": "شماره خرید",
                        "product_name": "نام محصول",
                        "quantity": "تعداد",
                        "unit_price": "قیمت هر واحد (افغانی)",
                        "total_amount": "مبلغ کل (افغانی)",
                        "purchase_date": "تاریخ خرید"
                    }
                )

                purchase_history["قیمت هر واحد (افغانی)"] = (
                    purchase_history["قیمت هر واحد (افغانی)"].map(
                        lambda value: f"{value:,.2f}"
                    )
                )

                purchase_history["مبلغ کل (افغانی)"] = (
                    purchase_history["مبلغ کل (افغانی)"].map(
                        lambda value: f"{value:,.2f}"
                    )
                )

                st.dataframe(
                    purchase_history,
                    hide_index=True,
                    use_container_width=True
                )

                total_purchase_amount = purchase_history[
                    "مبلغ کل (افغانی)"
                ].map(
                    lambda value: float(value.replace(",", ""))
                ).sum()

                total_purchase_units = int(
                    purchase_history["تعداد"].sum()
                )

                st.markdown(
                    f"""
                    <div style="
                        color: #17345e;
                        opacity: 1;
                        font-size: 16px;
                        font-weight: 600;
                        margin-top: 12px;
                    ">
                        تعداد کل واحدهای خریداری‌شده:
                        {total_purchase_units:,}
                        <br>
                        مجموع مبالغ خرید:
                        {total_purchase_amount:,.2f}
                        افغانی
                    </div>
                    """,
                    unsafe_allow_html=True
                )

        except (sqlite3.Error, KeyError, ValueError) as error:
            st.error("نمایش تاریخچه خریدها با مشکل مواجه شد.")
            st.code(str(error))

    elif page == "🧾 فاکتور فروش":
        st.markdown(
            '<div class="section-title">🧾 فاکتور فروش</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            """
            <div style="
                color: #263238;
                background-color: #F1F5F9;
                padding: 14px;
                border-radius: 10px;
                text-align: right;
                font-size: 16px;
                font-weight: 600;
                margin-bottom: 12px;
            ">
                فروش ثبت‌شده را انتخاب کنید تا فاکتور آن نمایش داده شود.
            </div>
            """,
            unsafe_allow_html=True
        )

        try:
            with get_db_connection(DB_PATH) as connection:
                invoice_data = pd.read_sql_query(
                    """
                    SELECT
                        s.id AS sale_id,
                        s.sale_date AS sale_date,
                        p.name AS product_name,
                        s.quantity AS quantity,
                        s.unit_price AS unit_price,
                        s.total_amount AS total_amount,
                        s.purchase_price_at_sale AS purchase_price
                    FROM sales AS s
                    LEFT JOIN products AS p
                        ON p.id = s.product_id
                    ORDER BY s.id DESC
                    """,
                    connection
                )

            if invoice_data.empty:
                st.info("هنوز فروشی برای صدور فاکتور ثبت نشده است.")
            else:
                invoice_data["invoice_label"] = invoice_data.apply(
                    lambda row: (
                        f"فاکتور شماره {int(row['sale_id'])} — "
                        f"{row['product_name']} — "
                        f"{float(row['total_amount']):,.2f} افغانی"
                    ),
                    axis=1
                )

                selected_sale_id = st.selectbox(
                    "انتخاب فاکتور",
                    options=invoice_data["sale_id"].astype(int).tolist(),
                    format_func=lambda sale_id: invoice_data.loc[
                        invoice_data["sale_id"] == sale_id,
                        "invoice_label"
                    ].iloc[0],
                    key="invoice_sale_selector"
                )

                selected_invoice = invoice_data.loc[
                    invoice_data["sale_id"] == selected_sale_id
                ].iloc[0]

                invoice_profit = round(
                    float(selected_invoice["total_amount"])
                    - (
                        float(selected_invoice["purchase_price"] or 0)
                        * int(selected_invoice["quantity"])
                    ),
                    2
                )

                st.divider()

                st.markdown(
                    f"""
                    <div style="
                        color: #FFFFFF;
                        background-color: #176B5B;
                        padding: 16px;
                        border-radius: 10px;
                        text-align: right;
                        font-size: 22px;
                        font-weight: bold;
                        margin-bottom: 16px;
                    ">
                        🧾 فاکتور شماره {int(selected_invoice['sale_id'])}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                invoice_details = [
                    ("تاریخ فروش", selected_invoice["sale_date"]),
                    ("نام محصول", selected_invoice["product_name"]),
                    ("تعداد", int(selected_invoice["quantity"])),
                    (
                        "قیمت هر واحد",
                        f"{float(selected_invoice['unit_price']):,.2f} افغانی"
                    ),
                ]

                for label, value in invoice_details:
                    st.markdown(
                        f"""
                        <div style="
                            color: #263238;
                            background-color: #F1F5F9;
                            padding: 12px;
                            border-radius: 8px;
                            margin-bottom: 8px;
                            text-align: right;
                            font-size: 16px;
                            line-height: 1.8;
                            overflow-wrap: anywhere;
                        ">
                            <span style="font-weight: bold;">
                                {label}:
                            </span>
                            <span>{value}</span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                st.divider()

                col1, col2 = st.columns(2)

                with col1:
                    st.markdown(
                        f"""
                        <div style="
                            color: #17365D;
                            background-color: #E8F1FF;
                            padding: 18px 10px;
                            border-radius: 10px;
                            text-align: center;
                            min-height: 110px;
                        ">
                            <div style="
                                font-size: 15px;
                                font-weight: bold;
                                margin-bottom: 12px;
                            ">
                                مبلغ کل فاکتور
                            </div>
                            <div style="
                                font-size: 20px;
                                font-weight: bold;
                                overflow-wrap: anywhere;
                            ">
                                {float(selected_invoice['total_amount']):,.2f}
                                افغانی
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                with col2:
                    st.markdown(
                        f"""
                        <div style="
                            color: #14532D;
                            background-color: #E8F8ED;
                            padding: 18px 10px;
                            border-radius: 10px;
                            text-align: center;
                            min-height: 110px;
                        ">
                            <div style="
                                font-size: 15px;
                                font-weight: bold;
                                margin-bottom: 12px;
                            ">
                                سود ناخالص
                            </div>
                            <div style="
                                font-size: 20px;
                                font-weight: bold;
                                overflow-wrap: anywhere;
                            ">
                                {invoice_profit:,.2f} افغانی
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )


            # ایجاد PDF برای فاکتور انتخاب‌شده
            pdf_font_path = (
                _os.path.join(_HERE, "fonts", "NotoNaskhArabic-Regular.ttf")
            )

            pdf_buffer = BytesIO()
            font_name = "PersianInvoiceFont"

            if font_name not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(
                    TTFont(font_name, pdf_font_path)
                )

            def invoice_persian_text(text):
                return get_display(
                    arabic_reshaper.reshape(str(text))
                )

            pdf_doc = SimpleDocTemplate(
                pdf_buffer,
                pagesize=A4,
                rightMargin=45,
                leftMargin=45,
                topMargin=45,
                bottomMargin=45,
            )

            pdf_title_style = ParagraphStyle(
                "InvoicePDFTitle",
                fontName=font_name,
                fontSize=20,
                leading=32,
                alignment=TA_RIGHT,
                spaceAfter=18,
            )

            pdf_body_style = ParagraphStyle(
                "InvoicePDFBody",
                fontName=font_name,
                fontSize=13,
                leading=27,
                alignment=TA_RIGHT,
                spaceAfter=8,
            )

            invoice_number = int(selected_invoice["sale_id"])
            invoice_date = str(selected_invoice["sale_date"])
            product_name = str(selected_invoice["product_name"])
            quantity = int(selected_invoice["quantity"])
            unit_price = float(selected_invoice["unit_price"])
            total_amount = float(selected_invoice["total_amount"])

            invoice_story = [
                Paragraph(
                    invoice_persian_text(
                        f"فاکتور فروش شماره {invoice_number}"
                    ),
                    pdf_title_style,
                ),
                Paragraph(
                    invoice_persian_text(f"تاریخ فروش: {invoice_date}"),
                    pdf_body_style,
                ),
                Paragraph(
                    invoice_persian_text(f"نام محصول: {product_name}"),
                    pdf_body_style,
                ),
                Paragraph(
                    invoice_persian_text(f"تعداد: {quantity}"),
                    pdf_body_style,
                ),
                Paragraph(
                    invoice_persian_text(
                        f"قیمت هر واحد: {unit_price:,.2f} افغانی"
                    ),
                    pdf_body_style,
                ),
                Spacer(1, 12),
                Paragraph(
                    invoice_persian_text(
                        f"مبلغ کل فاکتور: {total_amount:,.2f} افغانی"
                    ),
                    pdf_body_style,
                ),
                Paragraph(
                    invoice_persian_text(
                        f"سود ناخالص: {invoice_profit:,.2f} افغانی"
                    ),
                    pdf_body_style,
                ),
            ]

            pdf_doc.build(invoice_story)
            pdf_bytes = pdf_buffer.getvalue()
            pdf_buffer.close()

            st.download_button(
                label="📄 دانلود فاکتور PDF",
                data=pdf_bytes,
                file_name=f"invoice_{invoice_number}.pdf",
                mime="application/pdf",
                key=f"download_invoice_pdf_{invoice_number}",
                use_container_width=True,
            )

        except (sqlite3.Error, KeyError, ValueError, TypeError) as error:
            st.error("نمایش فاکتور با مشکل مواجه شد.")
            st.code(str(error))


    elif page == "📊 گزارش‌های فروش":
        st.markdown(
            '<div class="section-title">گزارش‌های فروش و سود</div>',
            unsafe_allow_html=True
        )

        st.markdown("""
        <p style="
            font-family: Vazirmatn, sans-serif;
            color: #425777;
            font-size: 15px;
            margin-top: 0;
            margin-bottom: 18px;
            opacity: 1;
        ">
            در این صفحه می‌توانید درآمد، تعداد فروش و سود
            ناخالص محصولات را برای بازه‌های زمانی مختلف بررسی کنید.
        </p>
        """, unsafe_allow_html=True)

        st.markdown(
            '<div style="color:#17345e !important; '
            'font-family:Vazirmatn,sans-serif; '
            'font-size:16px; font-weight:700; margin-bottom:8px;">'
            'نوع گزارش را انتخاب کنید'
            '</div>',
            unsafe_allow_html=True
        )

        report_mode = st.selectbox(
            "نوع گزارش",
            [
                "همه فروش‌ها",
                "گزارش روزانه",
                "گزارش ماهانه"
            ],
            label_visibility="collapsed",
            key="sales_report_mode"
        )

        join_condition = "p.id = s.product_id"
        query_params = []

        month_names = [
            "January", "February", "March", "April",
            "May", "June", "July", "August",
            "September", "October", "November", "December"
        ]
        today = pd.Timestamp.today().date()
        year_options = list(range(1900, today.year + 6))

        if report_mode == "گزارش روزانه":
            year_col, month_col, day_col = st.columns(3)

            year_col.markdown(
                '<div style="color:#17345e !important; font-family:Vazirmatn,sans-serif; '
                'font-size:16px; font-weight:700; margin-bottom:8px;">سال میلادی</div>',
                unsafe_allow_html=True
            )
            selected_year = year_col.selectbox(
                "سال",
                options=year_options,
                label_visibility="collapsed",
                index=year_options.index(today.year),
                key="daily_report_year"
            )

            month_col.markdown(
                '<div style="color:#17345e !important; font-family:Vazirmatn,sans-serif; '
                'font-size:16px; font-weight:700; margin-bottom:8px;">ماه میلادی</div>',
                unsafe_allow_html=True
            )
            selected_month_number = month_col.selectbox(
                "ماه",
                options=list(range(1, 13)),
                label_visibility="collapsed",
                index=today.month - 1,
                format_func=lambda month: month_names[month - 1],
                key="daily_report_month"
            )

            days_in_month = pd.Period(
                f"{selected_year}-{selected_month_number:02d}"
            ).days_in_month

            day_col.markdown(
                '<div style="color:#17345e !important; font-family:Vazirmatn,sans-serif; '
                'font-size:16px; font-weight:700; margin-bottom:8px;">روز ماه</div>',
                unsafe_allow_html=True
            )
            selected_day = day_col.selectbox(
                "روز",
                options=list(range(1, days_in_month + 1)),
                label_visibility="collapsed",
                index=min(today.day, days_in_month) - 1,
                key=f"daily_report_day_{selected_year}_{selected_month_number}"
            )

            selected_date = pd.Timestamp(
                year=selected_year,
                month=selected_month_number,
                day=selected_day
            ).date()

            join_condition += " AND date(s.sale_date) = ?"
            query_params.append(
                selected_date.strftime("%Y-%m-%d")
            )

        elif report_mode == "گزارش ماهانه":
            year_col, month_col = st.columns(2)

            year_col.markdown(
                '<div style="color:#17345e !important; font-family:Vazirmatn,sans-serif; '
                'font-size:16px; font-weight:700; margin-bottom:8px;">سال میلادی</div>',
                unsafe_allow_html=True
            )
            selected_month_year = year_col.selectbox(
                "سال",
                options=year_options,
                label_visibility="collapsed",
                index=year_options.index(today.year),
                key="monthly_report_year"
            )

            month_col.markdown(
                '<div style="color:#17345e !important; font-family:Vazirmatn,sans-serif; '
                'font-size:16px; font-weight:700; margin-bottom:8px;">ماه میلادی</div>',
                unsafe_allow_html=True
            )
            selected_month_number = month_col.selectbox(
                "ماه",
                options=list(range(1, 13)),
                label_visibility="collapsed",
                index=today.month - 1,
                format_func=lambda month: month_names[month - 1],
                key="monthly_report_month"
            )

            selected_month = (
                f"{selected_month_year}-{selected_month_number:02d}"
            )

            join_condition += " AND strftime('%Y-%m', s.sale_date) = ?"
            query_params.append(selected_month)

            st.caption(
                f"گزارش مربوط به ماه {month_names[selected_month_number - 1]} "
                f"سال {selected_month_year} است."
            )

        report_query = f"""
            SELECT
                p.id AS product_id,
                p.name AS product_name,
                COALESCE(SUM(s.quantity), 0) AS units_sold,
                COALESCE(SUM(s.total_amount), 0) AS total_revenue,
                COALESCE(
                    SUM(
                        s.total_amount -
                        (
                            s.purchase_price_at_sale *
                            s.quantity
                        )
                    ),
                    0
                ) AS gross_profit
            FROM products AS p
            LEFT JOIN sales AS s
                ON {join_condition}
            GROUP BY p.id, p.name
            ORDER BY total_revenue DESC
        """

        with get_db_connection(DB_PATH) as connection:
            report = pd.read_sql_query(
                report_query,
                connection,
                params=query_params
            )

        if report.empty:
            st.info("هنوز محصولی برای گزارش وجود ندارد.")
        else:
            total_revenue = report["total_revenue"].sum()
            total_profit = report["gross_profit"].sum()
            total_units = report["units_sold"].sum()

            col1, col2, col3 = st.columns(3)

            col1.markdown(
                '<div style="color:#17345e;font-size:15px;font-weight:600;'
                'margin-bottom:6px;">مجموع درآمد فروش</div>',
                unsafe_allow_html=True
            )
            col1.metric(
                "مجموع درآمد فروش",
                f"{total_revenue:,.2f} افغانی",
                label_visibility="collapsed"
            )

            col2.markdown(
                '<div style="color:#17345e;font-size:15px;font-weight:600;'
                'margin-bottom:6px;">مجموع سود ناخالص</div>',
                unsafe_allow_html=True
            )
            col2.metric(
                "مجموع سود ناخالص",
                f"{total_profit:,.2f} افغانی",
                label_visibility="collapsed"
            )

            col3.markdown(
                '<div style="color:#17345e;font-size:15px;font-weight:600;'
                'margin-bottom:6px;">تعداد واحدهای فروخته‌شده</div>',
                unsafe_allow_html=True
            )
            col3.metric(
                "تعداد واحدهای فروخته‌شده",
                f"{int(total_units):,}",
                label_visibility="collapsed"
            )

            st.divider()

            display_report = report.rename(columns={
                "product_id": "شناسه محصول",
                "product_name": "نام محصول",
                "units_sold": "تعداد فروخته‌شده",
                "total_revenue": "درآمد فروش (افغانی)",
                "gross_profit": "سود ناخالص (افغانی)"
            })

            st.markdown(
                '<h3 style="color:#17345e;font-family:Vazirmatn,sans-serif;'
                'font-size:22px;font-weight:700;">گزارش کامل محصولات</h3>',
                unsafe_allow_html=True
            )

            st.dataframe(
                display_report.round(2),
                use_container_width=True,
                hide_index=True
            )

            st.markdown(
                '<h3 style="color:#17345e;font-family:Vazirmatn,sans-serif;'
                'font-size:22px;font-weight:700;">رتبه‌بندی محصولات بر اساس درآمد</h3>',
                unsafe_allow_html=True
            )

            revenue_ranking = (
                report[
                    ["product_name", "total_revenue"]
                ]
                .sort_values(
                    "total_revenue",
                    ascending=False
                )
                .rename(columns={
                    "product_name": "نام محصول",
                    "total_revenue": "درآمد فروش (افغانی)"
                })
            )

            st.dataframe(
                revenue_ranking.round(2),
                use_container_width=True,
                hide_index=True
            )

            st.markdown(
                '<h3 style="color:#17345e;font-family:Vazirmatn,sans-serif;'
                'font-size:22px;font-weight:700;">رتبه‌بندی محصولات بر اساس سود</h3>',
                unsafe_allow_html=True
            )

            profit_ranking = (
                report[
                    ["product_name", "gross_profit"]
                ]
                .sort_values(
                    "gross_profit",
                    ascending=False
                )
                .rename(columns={
                    "product_name": "نام محصول",
                    "gross_profit": "سود ناخالص (افغانی)"
                })
            )

            st.dataframe(
                profit_ranking.round(2),
                use_container_width=True,
                hide_index=True
            )



        st.markdown(
            '<div class="section-title" style="font-size:21px;'
            'font-weight:700; margin-top:30px;">'
            'دریافت گزارش کامل اکسل</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="section-title" style="font-size:14px;'
            'font-weight:400; margin-bottom:12px;">'
            'دانلود اطلاعات کامل محصولات، خریدها، فروش‌ها، سود و موجودی.'
            '</div>',
            unsafe_allow_html=True
        )

        excel_file = create_excel_export()

        st.download_button(
            label="📥 دانلود گزارش کامل اکسل",
            data=excel_file,
            file_name=(
                "SmartInventory_Report_"
                + __import__("datetime").datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )
                + ".xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_complete_excel",
            use_container_width=True
        )

    elif page == "🛒 ثبت فروش":
        st.markdown(
            '<div style="font-family: Vazirmatn, sans-serif; '
            'color: #17345e; font-size: 22px; font-weight: 700; '
            'margin: 25px 0 15px 0;">'
            '🛒 ثبت فروش'
            '</div>',
            unsafe_allow_html=True
        )
        st.markdown("""
        <p style="
            font-family: Vazirmatn, sans-serif;
            color: #425777;
            font-size: 15px;
            margin-top: 0;
            margin-bottom: 18px;
            opacity: 1;
        ">
            ثبت فروش جدید، محاسبه مبلغ و به‌روزرسانی موجودی محصول
        </p>
        """, unsafe_allow_html=True)

        if "sale_success_message" in st.session_state:
            st.success(st.session_state.pop("sale_success_message"))

        available_products = products[
            products["quantity"] > 0
        ].copy()

        if available_products.empty:
            st.warning(
                "هیچ محصولی با موجودی قابل فروش وجود ندارد. "
                "ابتدا موجودی محصولات را بررسی کنید."
            )
        else:
            product_labels = {
                (
                    f"{row['name']} "
                    f"(شناسه {int(row['id'])}) — "
                    f"موجودی: {int(row['quantity'])}"
                ): int(row["id"])
                for _, row in available_products.iterrows()
            }

            st.markdown("""
            <p style="font-family: Vazirmatn, sans-serif;
                      color: #17345e; font-size: 15px;
                      font-weight: 600; margin: 0 0 6px 0;">
                انتخاب محصول
            </p>
            """, unsafe_allow_html=True)

            selected_label = st.selectbox(
                "انتخاب محصول",
                options=list(product_labels.keys()),
                key="sale_product_selector",
                label_visibility="collapsed"
            )

            selected_product_id = product_labels[selected_label]

            selected_product = products[
                products["id"] == selected_product_id
            ].iloc[0]

            current_stock = int(selected_product["quantity"])

            st.info(
                f"موجودی فعلی: {current_stock} واحد | "
                f"قیمت خرید: "
                f"{float(selected_product['purchase_price']):,.2f} افغانی"
            )

            with st.form("register_sale_form"):
                st.markdown("""
                <p style="font-family: Vazirmatn, sans-serif;
                          color: #17345e; font-size: 15px;
                          font-weight: 600; margin: 0 0 6px 0;">
                    تعداد فروش
                </p>
                """, unsafe_allow_html=True)

                quantity_to_sell = st.number_input(
                    "تعداد فروش",
                    min_value=1,
                    max_value=current_stock,
                    value=1,
                    step=1,
                    label_visibility="collapsed"
                )

                st.markdown("""
                <p style="font-family: Vazirmatn, sans-serif;
                          color: #17345e; font-size: 15px;
                          font-weight: 600; margin: 0 0 6px 0;">
                    قیمت فروش هر واحد (افغانی)
                </p>
                """, unsafe_allow_html=True)

                unit_price = st.number_input(
                    "قیمت فروش هر واحد (افغانی)",
                    min_value=0.0,
                    value=float(selected_product["selling_price"]),
                    step=1.0,
                    format="%.2f",
                    label_visibility="collapsed"
                )

                sale_submitted = st.form_submit_button(
                    "ثبت فروش",
                    use_container_width=True,
                    disabled=bool(
                        st.session_state.get("pending_loss_sale")
                    )
                )

            # بررسی فروش زیان‌ده پیش از ثبت
            if sale_submitted:
                purchase_price_for_warning = float(
                    selected_product["purchase_price"]
                )

                if unit_price < purchase_price_for_warning:
                    st.session_state["pending_loss_sale"] = {
                        "product_id": int(selected_product_id),
                        "product_name": str(selected_product["name"]),
                        "quantity": int(quantity_to_sell),
                        "unit_price": float(unit_price),
                        "purchase_price": purchase_price_for_warning,
                    }
                    st.rerun()

            # نمایش هشدار و دریافت تأیید کاربر
            pending_loss_sale = st.session_state.get(
                "pending_loss_sale"
            )

            if pending_loss_sale:
                loss_per_unit = round(
                    pending_loss_sale["purchase_price"]
                    - pending_loss_sale["unit_price"],
                    2
                )

                total_loss = round(
                    loss_per_unit * pending_loss_sale["quantity"],
                    2
                )

                import html

                safe_product_name = html.escape(
                    str(pending_loss_sale["product_name"])
                )

                st.markdown(
                    f"""
                    <div style="
                        background-color: #FFF3CD;
                        color: #212529;
                        border: 2px solid #B8860B;
                        border-radius: 12px;
                        padding: 18px;
                        margin: 12px 0;
                        direction: rtl;
                        text-align: right;
                    ">
                        <div style="
                            color: #7A2E00;
                            font-size: 22px;
                            font-weight: 800;
                            margin-bottom: 10px;
                        ">
                            ⚠️ هشدار فروش زیان‌ده
                        </div>
                        <div style="
                            color: #212529;
                            font-size: 17px;
                            font-weight: 600;
                            line-height: 2;
                        ">
                            محصول «{safe_product_name}» با قیمتی کمتر
                            از قیمت خرید فروخته می‌شود.
                            <br>
                            زیان هر واحد:
                            <strong>{loss_per_unit:,.2f} افغانی</strong>
                            <br>
                            زیان کل:
                            <strong>{total_loss:,.2f} افغانی</strong>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                confirm_col, cancel_col = st.columns(2)

                with confirm_col:
                    confirm_loss_sale = st.button(
                        "تأیید و ثبت فروش زیان‌ده",
                        key="confirm_loss_sale",
                        use_container_width=True
                    )

                with cancel_col:
                    cancel_loss_sale = st.button(
                        "انصراف از فروش",
                        key="cancel_loss_sale",
                        use_container_width=True
                    )

                if cancel_loss_sale:
                    st.session_state.pop(
                        "pending_loss_sale",
                        None
                    )
                    st.rerun()

                if confirm_loss_sale:
                    selected_product_id = int(
                        pending_loss_sale["product_id"]
                    )
                    quantity_to_sell = int(
                        pending_loss_sale["quantity"]
                    )
                    unit_price = float(
                        pending_loss_sale["unit_price"]
                    )

                    st.session_state.pop(
                        "pending_loss_sale",
                        None
                    )
                    sale_submitted = True

            if sale_submitted:
                try:
                    quantity_to_sell = int(quantity_to_sell)
                    unit_price = float(unit_price)

                    if quantity_to_sell < 1:
                        raise ValueError(
                            "تعداد فروش باید حداقل یک واحد باشد."
                        )

                    if unit_price < 0:
                        raise ValueError(
                            "قیمت فروش نمی‌تواند منفی باشد."
                        )

                    with get_db_connection(
                        DB_PATH,
                        timeout=10
                    ) as connection:
                        connection.execute("BEGIN IMMEDIATE")

                        current_product = connection.execute(
                            """
                            SELECT id, purchase_price, quantity, name
                            FROM products
                            WHERE id = ?
                            """,
                            (selected_product_id,)
                        ).fetchone()

                        if current_product is None:
                            raise ValueError(
                                "محصول انتخاب‌شده دیگر وجود ندارد."
                            )

                        current_purchase_price = float(
                            current_product[1]
                        )
                        actual_stock = int(current_product[2])
                        current_product_name = current_product[3]

                        if quantity_to_sell > actual_stock:
                            raise ValueError(
                                "موجودی کافی نیست. "
                                f"موجودی فعلی: {actual_stock} واحد."
                            )

                        total_amount = round(
                            unit_price * quantity_to_sell,
                            2
                        )

                        sale_cursor = connection.execute(
                            """
                            INSERT INTO sales (
                                product_id,
                                quantity,
                                unit_price,
                                total_amount,
                                purchase_price_at_sale
                            )
                            VALUES (?, ?, ?, ?, ?)
                            """,
                            (
                                selected_product_id,
                                quantity_to_sell,
                                unit_price,
                                total_amount,
                                current_purchase_price
                            )
                        )

                        update_result = connection.execute(
                            """
                            UPDATE products
                            SET quantity = quantity - ?
                            WHERE id = ?
                              AND quantity >= ?
                            """,
                            (
                                quantity_to_sell,
                                selected_product_id,
                                quantity_to_sell
                            )
                        )

                        if update_result.rowcount != 1:
                            raise ValueError(
                                "موجودی به‌روزرسانی نشد؛ "
                                "فروش ثبت نشد."
                            )

                        new_stock = actual_stock - quantity_to_sell

                        connection.execute(
                            """
                            INSERT INTO inventory_movements (
                                product_id,
                                movement_type,
                                quantity_change,
                                previous_quantity,
                                new_quantity,
                                movement_date,
                                description,
                                source_type,
                                source_id
                            )
                            VALUES (?, 'sale', ?, ?, ?, CURRENT_TIMESTAMP,
                                    ?, 'sale', ?)
                            """,
                            (
                                selected_product_id,
                                -quantity_to_sell,
                                actual_stock,
                                new_stock,
                                "ثبت فروش جدید",
                                sale_cursor.lastrowid
                            )
                        )

                        connection.commit()

                    load_products.clear()
                    load_sales_summary.clear()

                    estimated_profit = round(
                        (
                            unit_price - current_purchase_price
                        ) * quantity_to_sell,
                        2
                    )

                    st.session_state["sale_success_message"] = (
                        f"فروش {quantity_to_sell} واحد از "
                        f"«{current_product_name}» ثبت شد. "
                        f"مبلغ کل: {total_amount:,.2f} افغانی. "
                        f"سود ناخالص این فروش: "
                        f"{estimated_profit:,.2f} افغانی."
                    )

                    st.rerun()

                except ValueError as error:
                    st.error(str(error))

                except sqlite3.Error as error:
                    st.error("ثبت فروش در دیتابیس ناموفق بود.")
                    st.code(str(error))



    elif page == "👥 مشتریان":
        st.markdown('<h2 style="color:#17345e;">👥 مدیریت مشتریان</h2>', unsafe_allow_html=True)
        conn = get_db_connection(DB_PATH)
        try:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS customers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    phone TEXT DEFAULT '',
                    email TEXT DEFAULT '',
                    address TEXT DEFAULT '',
                    notes TEXT DEFAULT '',
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()

            st.markdown('<h3 style="color:#17345e;">ثبت مشتری جدید</h3>', unsafe_allow_html=True)
            with st.form("customer_add_form", clear_on_submit=True):
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">نام مشتری *</div>', unsafe_allow_html=True)
                customer_name = st.text_input("نام مشتری *", label_visibility="collapsed")
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">شماره تماس</div>', unsafe_allow_html=True)
                customer_phone = st.text_input("شماره تماس", label_visibility="collapsed")
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">ایمیل</div>', unsafe_allow_html=True)
                customer_email = st.text_input("ایمیل", label_visibility="collapsed")
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">آدرس</div>', unsafe_allow_html=True)
                customer_address = st.text_input("آدرس", label_visibility="collapsed")
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">یادداشت</div>', unsafe_allow_html=True)
                customer_notes = st.text_area("یادداشت", label_visibility="collapsed")
                customer_submit = st.form_submit_button("ثبت مشتری")

            if customer_submit:
                if not customer_name.strip():
                    st.error("نام مشتری را وارد کنید.")
                else:
                    conn.execute("""
                        INSERT INTO customers
                        (name, phone, email, address, notes)
                        VALUES (?, ?, ?, ?, ?)
                    """, (
                        customer_name.strip(), customer_phone.strip(),
                        customer_email.strip(), customer_address.strip(),
                        customer_notes.strip()
                    ))
                    conn.commit()
                    st.success("مشتری با موفقیت ثبت شد.")
                    st.rerun()

            rows = conn.execute("""
                SELECT id, name, phone, email, address, notes, created_at
                FROM customers ORDER BY id DESC
            """).fetchall()

            st.markdown('<h3 style="color:#17345e;">فهرست مشتریان</h3>', unsafe_allow_html=True)
            if rows:
                st.dataframe([{
                    "شناسه": r[0], "نام": r[1], "شماره تماس": r[2],
                    "ایمیل": r[3], "آدرس": r[4], "یادداشت": r[5],
                    "تاریخ ثبت": r[6]
                } for r in rows], use_container_width=True)

                options = {f"{r[0]} — {r[1]}": r for r in rows}
                selected_label = st.selectbox(
                    "مشتری موردنظر برای ویرایش یا حذف",
                    list(options.keys()), key="customer_manage_select"
                )
                selected = options[selected_label]

                with st.expander("ویرایش اطلاعات مشتری"):
                    with st.form("customer_edit_form"):
                        st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">نام مشتری</div>', unsafe_allow_html=True)
                        edit_name = st.text_input("نام", value=selected[1], label_visibility="collapsed")
                        st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">شماره تماس</div>', unsafe_allow_html=True)
                        edit_phone = st.text_input("شماره تماس", value=selected[2] or "", label_visibility="collapsed")
                        st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">ایمیل</div>', unsafe_allow_html=True)
                        edit_email = st.text_input("ایمیل", value=selected[3] or "", label_visibility="collapsed")
                        st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">آدرس</div>', unsafe_allow_html=True)
                        edit_address = st.text_input("آدرس", value=selected[4] or "", label_visibility="collapsed")
                        st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">یادداشت</div>', unsafe_allow_html=True)
                        edit_notes = st.text_area("یادداشت", value=selected[5] or "", label_visibility="collapsed")
                        edit_submit = st.form_submit_button("ذخیره تغییرات")

                    if edit_submit:
                        if not edit_name.strip():
                            st.error("نام مشتری الزامی است.")
                        else:
                            conn.execute("""
                                UPDATE customers
                                SET name=?, phone=?, email=?, address=?, notes=?
                                WHERE id=?
                            """, (
                                edit_name.strip(), edit_phone.strip(),
                                edit_email.strip(), edit_address.strip(),
                                edit_notes.strip(), selected[0]
                            ))
                            conn.commit()
                            st.success("اطلاعات مشتری به‌روزرسانی شد.")
                            st.rerun()

                if st.button("حذف مشتری انتخاب‌شده", key="customer_delete_btn"):
                    conn.execute("DELETE FROM customers WHERE id=?", (selected[0],))
                    conn.commit()
                    st.success("مشتری حذف شد.")
                    st.rerun()
            else:
                st.info("هنوز مشتری‌ای ثبت نشده است.")
        finally:
            conn.close()

    elif page == "🏢 تأمین‌کنندگان":
        st.markdown('<h2 style="color:#17345e;">🏢 مدیریت تأمین‌کنندگان</h2>', unsafe_allow_html=True)
        conn = get_db_connection(DB_PATH)
        try:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS suppliers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    phone TEXT DEFAULT '',
                    email TEXT DEFAULT '',
                    address TEXT DEFAULT '',
                    notes TEXT DEFAULT '',
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()

            st.markdown('<h3 style="color:#17345e;">ثبت تأمین‌کننده جدید</h3>', unsafe_allow_html=True)
            with st.form("supplier_add_form", clear_on_submit=True):
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">نام تأمین‌کننده *</div>', unsafe_allow_html=True)
                supplier_name = st.text_input("نام تأمین‌کننده *", label_visibility="collapsed")
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">شماره تماس</div>', unsafe_allow_html=True)
                supplier_phone = st.text_input("شماره تماس", label_visibility="collapsed")
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">ایمیل</div>', unsafe_allow_html=True)
                supplier_email = st.text_input("ایمیل", label_visibility="collapsed")
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">آدرس</div>', unsafe_allow_html=True)
                supplier_address = st.text_input("آدرس", label_visibility="collapsed")
                st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">یادداشت</div>', unsafe_allow_html=True)
                supplier_notes = st.text_area("یادداشت", label_visibility="collapsed")
                supplier_submit = st.form_submit_button("ثبت تأمین‌کننده")

            if supplier_submit:
                if not supplier_name.strip():
                    st.error("نام تأمین‌کننده را وارد کنید.")
                else:
                    conn.execute("""
                        INSERT INTO suppliers
                        (name, phone, email, address, notes)
                        VALUES (?, ?, ?, ?, ?)
                    """, (
                        supplier_name.strip(), supplier_phone.strip(),
                        supplier_email.strip(), supplier_address.strip(),
                        supplier_notes.strip()
                    ))
                    conn.commit()
                    st.success("تأمین‌کننده با موفقیت ثبت شد.")
                    st.rerun()

            rows = conn.execute("""
                SELECT id, name, phone, email, address, notes, created_at
                FROM suppliers ORDER BY id DESC
            """).fetchall()

            st.markdown('<h3 style="color:#17345e;">فهرست تأمین‌کنندگان</h3>', unsafe_allow_html=True)
            if rows:
                st.dataframe([{
                    "شناسه": r[0], "نام": r[1], "شماره تماس": r[2],
                    "ایمیل": r[3], "آدرس": r[4], "یادداشت": r[5],
                    "تاریخ ثبت": r[6]
                } for r in rows], use_container_width=True)

                options = {f"{r[0]} — {r[1]}": r for r in rows}
                selected_label = st.selectbox(
                    "تأمین‌کننده موردنظر برای ویرایش یا حذف",
                    list(options.keys()), key="supplier_manage_select"
                )
                selected = options[selected_label]

                with st.expander("ویرایش اطلاعات تأمین‌کننده"):
                    with st.form("supplier_edit_form"):
                        st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">نام تأمین‌کننده</div>', unsafe_allow_html=True)
                        edit_name = st.text_input("نام", value=selected[1], label_visibility="collapsed")
                        st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">شماره تماس</div>', unsafe_allow_html=True)
                        edit_phone = st.text_input("شماره تماس", value=selected[2] or "", label_visibility="collapsed")
                        st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">ایمیل</div>', unsafe_allow_html=True)
                        edit_email = st.text_input("ایمیل", value=selected[3] or "", label_visibility="collapsed")
                        st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">آدرس</div>', unsafe_allow_html=True)
                        edit_address = st.text_input("آدرس", value=selected[4] or "", label_visibility="collapsed")
                        st.markdown('<div style="font-family: Vazirmatn, sans-serif; color: #17345e; font-size: 15px; font-weight: 600; margin-bottom: 6px;">یادداشت</div>', unsafe_allow_html=True)
                        edit_notes = st.text_area("یادداشت", value=selected[5] or "", label_visibility="collapsed")
                        edit_submit = st.form_submit_button("ذخیره تغییرات")

                    if edit_submit:
                        if not edit_name.strip():
                            st.error("نام تأمین‌کننده الزامی است.")
                        else:
                            conn.execute("""
                                UPDATE suppliers
                                SET name=?, phone=?, email=?, address=?, notes=?
                                WHERE id=?
                            """, (
                                edit_name.strip(), edit_phone.strip(),
                                edit_email.strip(), edit_address.strip(),
                                edit_notes.strip(), selected[0]
                            ))
                            conn.commit()
                            st.success("اطلاعات تأمین‌کننده به‌روزرسانی شد.")
                            st.rerun()

                if st.button("حذف تأمین‌کننده انتخاب‌شده", key="supplier_delete_btn"):
                    conn.execute("DELETE FROM suppliers WHERE id=?", (selected[0],))
                    conn.commit()
                    st.success("تأمین‌کننده حذف شد.")
                    st.rerun()
            else:
                st.info("هنوز تأمین‌کننده‌ای ثبت نشده است.")
        finally:
            conn.close()
except Exception as error:
    st.error("هنگام بارگذاری اطلاعات خطایی رخ داد.")
    st.code(str(error))
