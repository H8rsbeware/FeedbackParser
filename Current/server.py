from flask import Flask, render_template, url_for, request
import json
import flask
from flask_cors import CORS

app = Flask(__name__)
CORS(app)


@app.route("/", methods=['POST', 'GET'])
def index():
    if request.method == 'POST':
        spreadsheetDetails = request.form['spreadsheet']
        
    else:
        return render_template('index.html')

@app.route("/postmethod", methods =["POST"])
def postmethod():
    data = request.get_json()
    print(data)
    return (data)


if __name__ == "__main__":
    app.run("localhost", 8486)

