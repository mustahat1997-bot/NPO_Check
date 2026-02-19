from flask import Flask, render_template, request, send_file, session
from Q_Actions_Script import get_actions, apply_rule_local
from Q_Action_Database import save_to_sqlite
import pandas as pd

app = Flask(__name__)
app.secret_key = "supersecret"  # مهم لتخزين الجلسة

@app.route("/", methods=["GET", "POST"])
def index():
    message = None
    excel_file = None
    point_result = None

    if request.method == "POST":
        action = request.form.get("action")

        # ✨ رفع الملف من المستخدم وتحديث قاعدة البيانات
        if action == "upload_excel":
            uploaded_file = request.files.get("excel_file")
            if uploaded_file:
                try:
                    df = pd.read_excel(uploaded_file, sheet_name="All Repeaters & Affiliates")
                    df.columns = df.columns.str.strip()
                    df = df.rename(columns={
                        "RepeaterName / Affiliates Name": "name",
                        "Site code": "site_code",
                        "Q Action": "q_action",
                        "Repeater Action": "repeater_action"
                    })
                    required_columns = ["name", "site_code", "q_action", "repeater_action"]
                    df = df[required_columns]
                    df["name"] = df["name"].astype(str).str.strip()
                    df["site_code"] = df["site_code"].astype(str).str.strip()
                    df["q_action"] = df["q_action"].fillna("No Action")
                    df["repeater_action"] = df["repeater_action"].fillna("No Action")

                    save_to_sqlite(df)
                    message = "✅ Database updated successfully from uploaded file"
                except Exception as e:
                    message = f"❌ Error processing file: {e}"

        # الحصول على النتائج
        elif action == "get_actions":
            points = request.form.get("points")
            if points:
                point_result = get_actions(points)
                session['step1_results'] = point_result

                special_dc = any(
                    item.get("province", "").strip().lower() in ["najaf", "basrah"]
                    for item in point_result if not item.get("not_found")
                )
                session['show_special_note'] = special_dc

        # تطبيق القاعدة وحفظ الملف النهائي
        elif action == "apply_rule":
            step1_results = session.get('step1_results')
            rule = request.form.get("rule")

            if step1_results and rule:
                output_file = apply_rule_local(step1_results, rule)
                return send_file(output_file, as_attachment=True)

    return render_template(
        "WEB_Q_Action.html",
        message=message,
        excel_file=excel_file,
        point_result=point_result,
        show_special_note=session.get("show_special_note", False)
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
