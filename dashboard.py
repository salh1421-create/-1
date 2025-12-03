import pandas as pd
import dash
from dash import dcc, html, Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import os

# استخدام ثيم "Flatly" لتصميم عصري ونظيف
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
app.title = "لوحة متابعة الغياب"
server = app.server

# مسار ملف البيانات
DATA_FILE = 'الغياب اليومي.xlsx'

# دالة لقراءة البيانات وتنظيفها
def load_data():
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame() # إرجاع جدول فارغ إذا لم يوجد الملف
    
    try:
        # قراءة البيانات (تحديد الشيت المناسب)
        df = pd.read_excel(DATA_FILE, sheet_name='البيانات_الخام')
        
        # تنظيف أسماء الأعمدة (إزالة المسافات الزائدة التي قد تسبب أخطاء)
        df.columns = df.columns.str.strip()
        
        # التأكد من تحويل التاريخ لصيغة مفهومة
        if 'التاريخ' in df.columns:
            df['التاريخ'] = pd.to_datetime(df['التاريخ'], errors='coerce').dt.date
            
        return df
    except Exception as e:
        print(f"Error loading data: {e}")
        return pd.DataFrame()

# تخطيط الصفحة (Layout)
app.layout = dbc.Container([
    # 1. العنوان الرئيسي
    dbc.Row([
        dbc.Col(html.H1("لوحة متابعة الغياب اليومية", className="text-center text-primary mb-4"), width=12)
    ], className="mt-4"),

    # 2. فلاتر البحث (الصف والفصل)
    dbc.Row([
        dbc.Col([
            html.Label("اختر الصف:", className="fw-bold"),
            dcc.Dropdown(id="grade-dropdown", clearable=False, className="mb-2")
        ], width=12, md=6),
        
        dbc.Col([
            html.Label("اختر الفصل:", className="fw-bold"),
            dcc.Dropdown(id="class-dropdown", clearable=False, className="mb-2")
        ], width=12, md=6),
    ], className="mb-4"),

    # 3. بطاقات المؤشرات (KPIs)
    dbc.Row(id="kpi-cards", className="mb-4"),

    # 4. الرسوم البيانية
    dbc.Row([
        dbc.Col(dcc.Graph(id="absence-by-date"), width=12, lg=6),
        dbc.Col(dcc.Graph(id="absence-by-student"), width=12, lg=6),
    ]),

    # مكوّن التحديث التلقائي (كل 60 ثانية)
    dcc.Interval(
        id="interval-component",
        interval=60*1000, 
        n_intervals=0
    )
], fluid=True, dir="rtl") # dir="rtl" مهم جداً لدعم العربية

# --- Callbacks ---

# تحديث القوائم المنسدلة بناءً على البيانات
@app.callback(
    [Output("grade-dropdown", "options"),
     Output("grade-dropdown", "value"),
     Output("class-dropdown", "options"),
     Output("class-dropdown", "value")],
    [Input("interval-component", "n_intervals"),
     Input("grade-dropdown", "value")]
)
def update_dropdowns(n, selected_grade):
    df = load_data()
    
    if df.empty:
        return [], None, [], None

    # خيارات الصفوف
    grades = sorted(df["الصف"].unique().tolist())
    grade_options = [{"label": g, "value": g} for g in grades]
    
    # تحديد القيمة الافتراضية للصف إذا لم يتم اختياره
    current_grade = selected_grade if selected_grade in grades else grades[0]

    # فلترة الفصول بناءً على الصف المختار
    classes_in_grade = df[df["الصف"] == current_grade]["الفصل"].unique()
    classes = sorted(classes_in_grade.tolist()) # ترتيب الفصول
    class_options = [{"label": str(c), "value": c} for c in classes]
    
    # قيمة الفصل الافتراضية
    current_class = classes[0] if classes else None
    
    return grade_options, current_grade, class_options, current_class

# تحديث المحتوى (البطاقات والرسوم)
@app.callback(
    [Output("kpi-cards", "children"),
     Output("absence-by-date", "figure"),
     Output("absence-by-student", "figure")],
    [Input("grade-dropdown", "value"),
     Input("class-dropdown", "value"),
     Input("interval-component", "n_intervals")]
)
def update_dashboard(selected_grade, selected_class, n):
    df = load_data()
    
    if df.empty or not selected_grade or not selected_class:
        # حالة عدم وجود بيانات
        return [], go.Figure(), go.Figure()

    # فلترة البيانات
    filtered_df = df[
        (df["الصف"] == selected_grade) & 
        (df["الفصل"] == selected_class)
    ]

    # --- 1. حساب المؤشرات (KPIs) ---
    total_absence = len(filtered_df)
    
    # أعلى طالب غياباً
    if not filtered_df.empty:
        student_counts = filtered_df["الاسم"].value_counts().reset_index()
        student_counts.columns = ["الاسم", "عدد الغياب"]
        top_student_name = student_counts.iloc[0]["الاسم"]
        top_student_count = student_counts.iloc[0]["عدد الغياب"]
        top_student_text = f"{top_student_name} ({top_student_count})"
    else:
        top_student_text = "لا يوجد"

    # تصميم البطاقات باستخدام Bootstrap Cards
    cards = [
        dbc.Col(dbc.Card([
            dbc.CardHeader("إجمالي الغياب", className="text-center text-white bg-danger"),
            dbc.CardBody(html.H2(str(total_absence), className="text-center text-danger"))
        ], className="shadow-sm"), width=12, md=6),
        
        dbc.Col(dbc.Card([
            dbc.CardHeader("الأكثر غياباً", className="text-center text-white bg-primary"),
            dbc.CardBody(html.H4(top_student_text, className="text-center text-primary"))
        ], className="shadow-sm"), width=12, md=6),
    ]

    # --- 2. الرسم البياني: الغياب حسب التاريخ ---
    if not filtered_df.empty:
        absence_by_date = filtered_df.groupby("التاريخ").size().reset_index(name="عدد الغياب")
        fig_date = px.line(
            absence_by_date, x="التاريخ", y="عدد الغياب", markers=True,
            title=f"اتجاه الغياب اليومي - {selected_grade} / {selected_class}"
        )
        fig_date.update_layout(template="simple_white", xaxis_title="التاريخ", yaxis_title="العدد")
    else:
        fig_date = px.line(title="لا توجد بيانات")

    # --- 3. الرسم البياني: الغياب حسب الطالب ---
    if not filtered_df.empty:
        absence_by_student = filtered_df["الاسم"].value_counts().reset_index()
        absence_by_student.columns = ["الطالب", "عدد الغياب"]
        # عرض أعلى 10 طلاب فقط لتجنب الازدحام
        absence_by_student = absence_by_student.head(10)
        
        fig_student = px.bar(
            absence_by_student, x="الطالب", y="عدد الغياب",
            title="أكثر الطلاب غياباً (أعلى 10)",
            text="عدد الغياب",
            color="عدد الغياب",
            color_continuous_scale="Reds"
        )
        fig_student.update_layout(template="simple_white", xaxis_title="اسم الطالب", yaxis_title="العدد")
    else:
        fig_student = px.bar(title="لا توجد بيانات")

    return cards, fig_date, fig_student

if __name__ == "__main__":
    app.run_server(debug=True)