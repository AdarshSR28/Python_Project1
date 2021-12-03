import os
import csv
import shutil
import pandas as pd
os.system ("cls")

def csv_modifier(path,list,pos_list):
    df_temp=pd.read_csv(path)
    df = df_temp.drop("Score",axis=1,inplace=False)
    df.insert(5,"Score_After_Negative",list)
    df.insert(2,"Score",pos_list)
    output_path = ".\\outputs"

    if os.path.exists(output_path):

        if os.path.exists(os.path.join(output_path,"concise_marksheet.csv")):
            os.remove(os.path.join(output_path,"concise_marksheet.csv"))

        df.to_csv(os.path.join(output_path,"concise_marksheet.csv"),index="")
    else:
        os.mkdir(output_path)
        df.to_csv( os.path.join(output_path,"concise_marksheet.csv") , index = "")

def csv_modifier_blanks(concise_df):
    output_path = ".\\outputs"
    # print(concise_df)
    if os.path.exists(output_path):

        if os.path.exists(os.path.join(output_path,"concise_marksheet.csv")):
            os.remove(os.path.join(output_path,"concise_marksheet.csv"))

        concise_df.to_csv(os.path.join(output_path,"concise_marksheet.csv"),index="")
    else:
        os.mkdir(output_path)
        concise_df.to_csv( os.path.join(output_path,"concise_marksheet.csv") , index = "")

def answer_extractor(file,length):
    answer_record={}
    for lines in file:
        if (lines[6]!='Roll Number'):
            list=[]
            i=7
            while(i<length):
                list.append(lines[i])
                i+=1
            answer_record[lines[6]]=list

    return answer_record   
def result_generator(file):
    final_record={}
    answer=file["ANSWER"]
    for line in file:
        result={"right":0,"wrong":0,"not answered":0}
        i=0
        for ans in file[line]:    
            if ans==answer[i]:
                result["right"]+=1
                i=i+1
            elif ans=='' :
                result["not answered"]+=1
                i=i+1
            else :
                result["wrong"]+=1
                i=i+1
                
            final_record[line]=result
    return final_record
def information_extractor(file):
    dict=[]
    for line in file:
        temp_dict2={}
        for key,value in line.items():
            if (key==''):
                continue
            else:
                temp_dict2[key]=value
        
        dict.append(temp_dict2)
    return dict

def concise_blanks(right_points = 5,wrong_points = 1):
    input_path = "./outputs"
    concise_path="concise_marksheet.csv"
    master_roll_path = "master_roll.csv"
    master=pd.read_csv(os.path.join("./uploads",master_roll_path))
    concise_marksheet=pd.read_csv(os.path.join("./outputs",concise_path))
    if not  os.path.exists(".\outputs"):
        os.mkdir(".\outputs")
    blank_list_total = []
    master_df=pd.read_csv(os.path.join("./uploads",master_roll_path),index_col="roll")
    for roll in master["roll"]:
    
        if roll.upper() not in (item.upper() for item in concise_marksheet["Roll Number"]):
            individual_blank = []
            for cols in concise_marksheet:
                if cols=="Roll Number":
                    individual_blank.append(roll)
                elif cols == "Score":
                    individual_blank.append("ABSENT")
                elif cols == "Name":
                    individual_blank.append(master_df.loc[roll]["name"])
                elif cols == "Score_After_Negative":
                    individual_blank.append("ABSENT")
                else:
                    individual_blank.append(str("-"))

            blank_list_total.append(individual_blank)

    df2 = pd.DataFrame(blank_list_total,columns=concise_marksheet.columns)
    # print(df2)
    final_df_concise = concise_marksheet.append(df2, ignore_index=True)
    csv_modifier_blanks(final_df_concise)
    return


def concise_marksheet(right_points=5,wrong_points=1):
    input_path=".\\uploads"
    answer_record={}
    final_record={}
    dict=[]
    list=[]
    path=os.path.join(input_path,"responses.csv")
    print(path)
    file=open(path,"r")
    file=csv.reader(file)
    reader=open(path,'r')
    line_of_reader=csv.DictReader(reader)
    df = pd.read_csv(path)
    dict = information_extractor(line_of_reader)
    answer_record = answer_extractor(file,len(df.columns))
    final_record = result_generator(answer_record)
    list_pos = []    
    for file in dict:
        right=int(final_record[file["Roll Number"]]["right"])
        wrong=int(final_record[file["Roll Number"]]["wrong"])
        not_attempted=int(final_record[file["Roll Number"]]["not answered"])
        marks=str(right*right_points+wrong*wrong_points*-1)+" / "+(str(right_points*(right+not_attempted+wrong))) + "\t"
        pos_marks = str(right*right_points)+" / "+(str(right_points*(right+not_attempted+wrong))) + "\t"
        list.append(marks)
        list_pos.append(pos_marks)
    csv_modifier(path,list,list_pos)
# consise_marksheet()