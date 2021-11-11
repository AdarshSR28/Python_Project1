import re
from flask import Flask, render_template, request
import os
from flask.helpers import url_for
import shutil
from werkzeug.wrappers import response
import sys
import rwisegenerator as rwise

app = Flask(__name__)
ALLOWED_EXTS = {"csv"}

def check_file(file):
    return '.' in file and file.rsplit('.',1)[1].lower() in ALLOWED_EXTS


@app.route("/", methods = ["GET", "POST"])
def index():

    error = None
    success = None
    if request.method == 'POST':

        if 'submit_files' in request.form:
            if 'master_roll' not in request.files or 'responses' not in request.files:
                error = "Please upload both files!"
                return render_template('index.html', error=error)

            master_file = request.files["master_roll"]
            master_filename = master_file.filename

            response_file = request.files["responses"]
            response_filename = response_file.filename

            if response_filename == '' or master_filename == '':
                error = "Filename is invalid! \n Make sure the name of the file is not blank and retry."
                return render_template('index.html',error=error)
            elif master_filename == '':
                error = "Master Roll number filename is invalid"
                return render_template('index.html',error=error)
            elif check_file(response_filename) == False or check_file(master_filename) == False:
                error = "Please upload a csv type file!!!"
                return render_template('index.html', error = error)
            
            upload_path = "./uploads"
            if os.path.exists(upload_path):
                shutil.rmtree(upload_path)
                os.mkdir(upload_path)
            else:
                os.mkdir(upload_path)
            
            master_file.save(os.path.join(upload_path,"master_roll.csv"))
            response_file.save(os.path.join(upload_path,"responses.csv"))
            return render_template("index.html", error=error, success = "Upload success" )
        
        if "gen_mark" in request.form:
            error_rwise = None
            f_path = ".\\uploads"
            if not os.path.exists(f_path):
                error_rwise = "Files uploaded cannot be found!"
                return render_template("index.html",error_rwise=error_rwise)
            else:
                # check whether marks are input by the user or not
                positive = 0
                negative = 0
                if 'positive' in request.form and 'negative' in request.form:
                    
                    try:
                        positive = int(request.form['positive'])
                        negative = int(request.form['negative'])
                    except Exception:
                        error_rwise = "Scoring value error! Please enter valid marking scheme!"
                        return render_template("index.html", error_rwise = error_rwise)

                try:
                    rwise.generate_roll_no_wise_marksheet(positive,negative)
                except Exception:
                    error_rwise = "Error in manipulating files! Please try again!"
                    return render_template("index.html", error_rwise = error_rwise)
                return render_template("index.html",success_rwise="1")


    return render_template('index.html')

app.run(debug=True)