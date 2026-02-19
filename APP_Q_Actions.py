from flask import Flask, render_template, request, send_file, session
from Q_Actions_Script import get_actions, apply_rule_local
from Q_Action_Database import excel_to_dataframe, save_to_sqlite

app = Flask(__name__)
app.secret_key = "supersecret"  # مهم لتخزين الجلسة

@app.route("/", methods=["GET", "POST"])
def index():
    message = None
    excel_file = None
    point_result = None

    if request.method == "POST":
        action = request.form.get("action")

        if action == "update_db":
            df, excel_file = excel_to_dataframe()
            save_to_sqlite(df)
            message = "✅ Database updated successfully"

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
