from flask import Flask, request, jsonify, render_template_string
import os
import time
from codes import send_email, copy_ex, edit_ex

app = Flask(__name__)
app.config['JSON_AS_ASCII'] = False

@app.route("/")
def home():
    return '''
    <!DOCTYPE html>
    <html lang="zh-Hant">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta name="theme-color" content="#74ebd5">
        <meta name="apple-mobile-web-app-capable" content="yes">
        <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
        <title>教室與教師設定</title>
        <style>
            body {
                font-family: "Microsoft JhengHei", sans-serif;
                background: linear-gradient(to right, #74ebd5, #acb6e5);
                display: flex;
                justify-content: center;
                align-items: center;
                min-height: 100vh;
                margin: 0;
                padding: 20px;
            }
            .container {
                background-color: white;
                padding: 30px;
                border-radius: 15px;
                box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
                width: 100%;
                max-width: 500px;
            }
            h1 {
                text-align: center;
                color: #333;
            }
            select, input[type="text"] {
                width: 100%;
                padding: 10px;
                font-size: 16px;
                margin-bottom: 15px;
                border-radius: 8px;
                border: 1px solid #ccc;
            }
            button {
                background-color: #00c300;
                color: white;
                border: none;
                padding: 10px 15px;
                font-size: 16px;
                border-radius: 8px;
                cursor: pointer;
                margin-right: 10px;
            }
            button:hover {
                background-color: #00a000;
            }
        </style>
        <script>
            function addTeacherField() {
                const container = document.getElementById('teacherFields');
                const input = document.createElement('input');
                input.type = 'text';
                input.name = 'teachers';
                input.placeholder = '教師姓名';
                container.appendChild(input);
            }
        </script>
    </head>
    <body>
        <div class="container">
            <h1>設定教室與教師</h1>
            <form action="/edit" method="POST">
                <label for="classroom">選擇教室：</label>
                <select name="classroom" required>
                    <option value="萬芳">萬芳</option>
                    <option value="龍門">龍門</option>
                    <option value="多一間廁所">多一間廁所</option>
                </select>

                <div id="teacherFields">
                    <input type="text" name="teachers" placeholder="教師姓名" required>
                </div>
                <button type="button" onclick="addTeacherField()">增加教師</button>
                <button type="submit">送出</button>
            </form>
        </div>
    </body>
    </html>
    '''

@app.route("/edit", methods=["POST"])
def edit():
    classroom = request.form.get("classroom")
    teachers = [t for t in request.form.getlist("teachers") if t.strip()]

    return render_template_string('''
    <!DOCTYPE html>
    <html lang="zh-Hant">
    <head>
        <meta charset="UTF-8">
        <title>處理中...</title>
        <style>
            body {
                font-family: "Microsoft JhengHei";
                text-align: center;
                background-color: #f0f0f0;
                padding: 50px;
            }
            .spinner {
                border: 10px solid #f3f3f3;
                border-top: 10px solid #3498db;
                border-radius: 50%;
                width: 80px;
                height: 80px;
                animation: spin 2s linear infinite;
                margin: 20px auto;
            }
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
            #success {
                display: none;
                font-size: 20px;
                color: green;
                margin-top: 20px;
            }
        </style>
    </head>
    <body>
        <h2>你選擇的教室是：{{ classroom }}</h2>
        <p>教師名單：{{ teachers|join(", ") }}</p>
        <div class="spinner"></div>
        <div id="success">✅ Email 已成功傳送！</div>

        <script>
            fetch("/send_email", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    classroom: "{{ classroom }}",
                    teachers: {{ teachers | tojson }}
                })
            }).then(response => response.json())
              .then(data => {
                document.querySelector(".spinner").style.display = "none";
                document.getElementById("success").style.display = "block";
              });
        </script>
    </body>
    </html>
    ''', classroom=classroom, teachers=teachers)


@app.route("/send_email", methods=["POST"])
def send_email_async():
    data = request.get_json()
    classroom = data.get("classroom")
    teachers = data.get("teachers", [])
    t = time.time()
    tt = time.localtime(t)

    copy_ex(classroom)
    send_email(
        f"{classroom}教室 {time.strftime('%Y%m%d', tt)} 加退保excel",
        classroom,
        os.listdir(classroom),
        "\n".join(edit_ex(classroom, teachers)),
        "justin520cheng@gmail.com"
    )
    return jsonify({"status": "success"})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)