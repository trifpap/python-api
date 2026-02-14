from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route("/")
def home():
    return "API is running!"

@app.route("/process", methods=["GET", "POST"])
def process_text():
    if request.method == "POST":
        data = request.json
        text = data.get("text", "")
    else:
        text = request.args.get("text", "")
    
    return jsonify({"result": text.upper()})