from flask import Flask, render_template, request, send_file, session
from Q_Actions_Script import get_actions, apply_rule_local
from Q_Action_Database import excel_to_dataframe, save_to_sqlite

app = Flask(__name__)
app.secret_key = "supersecret"  # مهم لتخزين الجلسة

@app.route("/set_special_rule/<int:index>", methods=["POST"])
def set_special_rule(index):
    import json
    step1_results = session.get("step1_results", [])
    data = request.get_json()
    rule = data.get("rule")

    if 0 <= index < len(step1_results):
        step1_results[index]["special_rule"] = rule
        session["step1_results"] = step1_results

    return "", 204

@app.route("/mark_no_action/<int:index>", methods=["POST"])
def mark_no_action(index):
    step1_results = session.get("step1_results", [])

    if 0 <= index < len(step1_results):
        step1_results[index]["no_action_needed"] = True
        session["step1_results"] = step1_results

    return "", 204

@app.route("/", methods=["GET", "POST"])

def index():
    message = None
    excel_file = None
    point_result = None

    if request.method == "POST":
        action = request.form.get("action")

        # ✅ 1️⃣ رفع الاكسل وتحديث الداتابيس
        if action == "update_db":
            uploaded_file = request.files.get("excel_file")

            if uploaded_file and uploaded_file.filename.endswith(".xlsx"):
                try:
                    df = excel_to_dataframe(uploaded_file)
                    save_to_sqlite(df)
                    message = "✅ Database updated successfully"
                    excel_file = uploaded_file.filename
                except Exception as e:
                    message = f"❌ Error processing file: {str(e)}"
            else:
                message = "❌ Please upload a valid Excel file (.xlsx)"

        # ✅ 2️⃣ جلب الأكشنات
        elif action == "get_actions":
            points = request.form.get("points")
            if points:
                point_result = get_actions(points)
                session['step1_results'] = point_result
                session['total_points'] = len(point_result)

                special_dc = any(
                    item.get("province", "").strip().lower() in ["najaf", "basrah"]
                    for item in point_result if not item.get("not_found")
                )

                session['show_special_note'] = special_dc

        # ✅ 3️⃣ تطبيق القاعدة وتحميل الاكسل
        elif action == "apply_rule":
            step1_results = session.get('step1_results')
            rule = request.form.get("rule")

            if step1_results and rule:
                filtered_results = [
                    r for r in step1_results
                    if not r.get("not_found") and not r.get("no_action_needed")
                ]

                if filtered_results:
                    output_file = apply_rule_local(filtered_results, rule)
                    return send_file(output_file, as_attachment=True)
                else:
                    message = "⚠️ No valid points to apply rule"
                    
    return render_template(
        "WEB_Q_Action.html",
        message=message,
        excel_file=excel_file,
        point_result=point_result,
        show_special_note=session.get("show_special_note", False),
        total_points=session.get("total_points", 0)
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)