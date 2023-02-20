from flask import Flask, render_template, url_for, request
import json
import flask
from flask_cors import CORS

import backend.creator as cr 


app = Flask(__name__)
CORS(app)

RESULT = "/home/user/Documents/FeedbackOutput"


@app.route("/", methods=['POST', 'GET'])
def index():
    if request.method == 'POST':
        spreadsheetDetails = request.form['spreadsheet']
        
    else:
        return render_template('index.html')


@app.route("/postmethod", methods =["POST"])
def postmethod():
    data = request.get_json()
    print(data["Keys"], data["Catagories"])
    sheet = cr.spreadsheet(data["Keys"], data["Catagories"], "Created")
    sheet.create(f"{RESULT}")
    return render_template('created.html')
    
    


if __name__ == "__main__":
    app.run("localhost", 8486)

