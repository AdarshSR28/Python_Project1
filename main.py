import re
from flask import Flask, render_template, request
import os
from flask.helpers import url_for
import shutil
from werkzeug.wrappers import response
import sys
import rwisegenerator as rwise
from flask_mail import *
import scrapemail
import concise

app = Flask(__name__)
app.config['MAIL_SERVER']='smtp.gmail.com'  
app.config['MAIL_PORT']=465  
app.config['MAIL_USERNAME'] = 'ee04ee05pythonproj@gmail.com'  
app.config['MAIL_PASSWORD'] = "ee04ee05python"  
app.config['MAIL_USE_TLS'] = False  
app.config['MAIL_USE_SSL'] = True 
mail = Mail(app) 
ALLOWED_EXTS = {"csv"}

def check_file(file):
    return '.' in file and file.rsplit('.',1)[1].lower() in ALLOWED_EXTS

def gen_email():

    try:
        mail_mapping = scrapemail.roll_mail_mapping()
        output_path = ".\\outputs"
        for key in mail_mapping.keys():
            roll_no = key
            recipient_list = mail_mapping[key]
            resource_location = os.path.join(output_path,roll_no + ".xlsx")
            msg = Message(subject = "Quiz Marks for student " + roll_no,
            body = "Please file attached marks of quiz.", 
            sender = "ee04ee05pythonproj@gmail.com", 
            recipients = recipient_list)


            with app.open_resource(resource_location) as fp:
                msg.attach(roll_no + ".xlsx","text/xlsx",fp.read())
                mail.send(msg)
        
        return True
    except Exception:
        return False

def refresh_server():
    output_path = ".\\outputs"
    uploads_path = ".\\uploads"
    try:
        if os.path.exists(output_path):
            shutil.rmtree(output_path)
        if os.path.exists(uploads_path):
            shutil.rmtree(uploads_path)
        
        return True

    except Exception:
        return False
    
    

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
                    rwise.generate_blankfile(positive,negative)
                except Exception:
                    error_rwise = "Error in manipulating files! Please try again!"
                    return render_template("index.html", error_rwise = error_rwise)
                return render_template("index.html",success_rwise="1")
        
        if 'gen_email' in request.form:
            if gen_email():
                return render_template('index.html',success_mail = '1')
            return render_template('index.html', error_mail = "Error in sending mails!")
        
        if 'refresh' in request.form:
            if refresh_server():
                return render_template('index.html',success_refresh = '1')
            return render_template('index.html',error_refresh = "There was some error in refreshing, please restart the server!")

        if 'gen_concise' in request.form:
            error_cwise = None
            f_path = ".\\uploads"
            if not os.path.exists(f_path):
                error_cwise = "Files uploaded cannot be found!"
                return render_template("index.html",error_cwise=error_cwise)
            else:
                # check whether marks are input by the user or not
                positive = 0
                negative = 0
                if 'positive' in request.form and 'negative' in request.form:
                    
                    try:
                        positive = int(request.form['positive'])
                        negative = int(request.form['negative'])
                    except Exception:
                        error_cwise = "Scoring value error! Please enter valid marking scheme!"
                        return render_template("index.html", error_cwise = error_cwise)

                try:
                    concise.concise_marksheet(positive,negative)
                    concise.concise_blanks(positive,negative)
                except Exception:
                    error_cwise = "Error in manipulating files! Please try again!"
                    return render_template("index.html", error_cwise = error_cwise)
                return render_template("index.html",success_cwise="1")


    return render_template('index.html')

app.run(debug=True)