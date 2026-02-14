from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "API is running!"

@app.route("/process", methods=["POST"])
def process_text():
    data = request.json
    text = data.get("text", "")
    return jsonify({"result": text.upper()})