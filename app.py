from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app, resources={r"/process": {"origins": "*"}})

@app.route("/")
def home():
    return "API is running!"

@app.route("/process", methods=["POST", "OPTIONS"])
def process_text():
    if request.method == "OPTIONS":
        return '', 200

    data = request.json
    text = data.get("text", "")
    return jsonify({"result": text.upper()})
