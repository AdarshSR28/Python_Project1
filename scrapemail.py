import pandas as pd
import os

def roll_mail_mapping():
    
    folder = ".\\uploads"
    input_path = os.path.join(folder,"responses.csv")
    dict = {}
    df = pd.read_csv(input_path)
    for index,row in df.iterrows():
        roll_no = row["Roll Number"]
        email = row["Email address"]
        webmail = row["IITP webmail"]
        if roll_no == "ANSWER":
            continue
        dict[roll_no] = [email,webmail]
    return dict

# roll_mail_mapping()
