import os
import csv , openpyxl
import pandas as pd
from openpyxl import Workbook
import shutil
from openpyxl import Workbook
from openpyxl.styles import Font,Alignment,Border,Side
from openpyxl.drawing.image import Image
os.system('cls')
def IsEmpty(dict):
    for element in dict:
        if element:
            return True
        return False

def answer_extractor(file):
    answer_record={}
    for lines in file:
        if (lines[6]!='Roll Number'):
            list=[]
            i=7
            while(i<35):
                list.append(lines[i])
                i+=1
            answer_record[lines[6]]=list

    return answer_record


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


def generate_blankfile(right_points=5, wrong_points=1):
    input_path=".\\uploads"
    master_roll_path=os.path.join(input_path,"master_roll.csv")
    responses_path=os.path.join(input_path,"responses.csv")
    master=pd.read_csv(master_roll_path)
    response=pd.read_csv(responses_path)
    master_df=pd.read_csv(master_roll_path,index_col="roll")
    if not  os.path.exists(".\outputs"):
        os.mkdir(".\outputs")
    for roll in master["roll"]:
        if roll.upper() not in (item.upper() for item in response["Roll Number"]):
            workbook=Workbook()
            sheet=workbook.active
            img=Image("project_header.PNG")
            alignment_heading=Alignment(horizontal='right',vertical='bottom')
            alignment_content=Alignment(horizontal='left',vertical='bottom')
            alignment_ans=Alignment(horizontal='center',vertical='bottom')
            font_heading=Font(name='Century',size=14,bold=False)
            font_content=Font(name='Century',size=14,bold=True)
            right_color=Font(color="00FF00",name='Century',size=14,bold=False)
            wrong_color=Font(color="ff0000",name='Century',size=14,bold=False)
            give_color=Font(color="0000FF",name='Century',size=14,bold=False)
            border_style=Side(border_style="medium",color="000000")
            border=Border(top=border_style,bottom=border_style,left=border_style,right=border_style)
            sheet.add_image(img,"A1")
            sheet["A6"]="Name :"
            sheet["A6"].font=font_heading
            sheet["A6"].alignment=alignment_heading
            sheet["B6"]=master_df.loc[roll]["name"]
            sheet["B6"].font=font_content
            sheet["B6"].alignment=alignment_content
            sheet["D6"]="Exam :"
            sheet["D6"].font=font_heading
            sheet["D6"].alignment=alignment_heading
            sheet["E6"]="quiz"
            sheet["E6"].font=font_content
            sheet["E6"].alignment=alignment_content
            sheet["A7"]="Roll Number"
            sheet["A7"].font=font_heading
            sheet["A7"].alignment=alignment_heading
            sheet["B7"]=roll
            sheet["B7"].font=font_content
            sheet["B7"].alignment=alignment_content
            sheet.row_dimensions[5].height=20
            right=" "
            wrong=" "
            not_attempted=" "
            data={
                "Right":[right,right_points," "],
                "Wrong":[wrong,wrong_points, " "],
                "Not Attempt":[not_attempted,0,''],
                "Max":[right+wrong+not_attempted ,'',
                        "Absent"]    
            }
            rownumber=["B","C","D","E"]
            column_number=["10","11","12"]
            sheet["A9"]=""
            sheet["A10"]= "No."
            sheet["A11"]="Marking"
            sheet["A12"]="Total"
            sheet["B9"]="Right"
            sheet["C9"]="Wrong"
            sheet["D9"]="Not Attempt"
            sheet["E9"]="Max"
            
            space=["A9","A10","A11","A12","B9","C9","D9","E9"]
            for p in space:
                sheet[p].font=font_content
                sheet[p].alignment=alignment_ans
                sheet[p].border=border
            for text in data:
                i=0
                if text=="Right":
                    row=rownumber[0]
                    for column in column_number:
                        sheet[row + column]=data[text][i]
                        sheet[row + column].font=right_color
                        sheet[row + column].alignment=alignment_ans
                        sheet[row + column].border=border
                        i+=1
                elif text=="Wrong":
                    row=rownumber[1]
                    for column in column_number:
                        sheet[row + column]=data[text][i]
                        sheet[row + column].font=wrong_color
                        sheet[row + column].alignment=alignment_ans
                        sheet[row + column].border=border
                        i+=1
                elif text=="Not Attempt":
                    row=rownumber[2]
                    for column in column_number:
                        sheet[row + column]=data[text][i]
                        sheet[row + column].alignment=alignment_ans
                        sheet[row + column].font=font_content
                        sheet[row + column].border=border
                        i+=1
                elif text=="Max":     
                    row=rownumber[3]
                    for column in column_number:
                        sheet[row + column]=data[text][i]
                        sheet[row + column].font=font_content
                        sheet[row + column].alignment=alignment_ans
                        sheet[row + column].border=border
                        i+=1
                    sheet["E12"].font=give_color
            sheet.column_dimensions['A'].width=20
            sheet.column_dimensions['B'].width=20
            sheet.column_dimensions['C'].width=15
            sheet.column_dimensions['D'].width=20
            sheet.column_dimensions['E'].width=20
            workbook.save(os.path.join(".\outputs",roll+".xlsx"))



def generate_roll_no_wise_marksheet(right_points=5, wrong_points=1):
    input_path=".\\uploads"
    count=0
    dict=[]
    answer_record={}
    final_record={}
    
    path=os.path.join(input_path,"responses.csv")
    reader=open(path,'r')
    line_of_reader=csv.DictReader(reader)
    file=open(path,"r")
    file=csv.reader(file)
    answer_record = answer_extractor(file)
    dict = information_extractor(line_of_reader)
    final_record = result_generator(answer_record)

    if not os.path.exists(".\outputs"):
        os.mkdir(".\outputs")

    for file in dict:
        workbook=Workbook()
        sheet=workbook.active
        img=Image("project_header.PNG")
        alignment_heading=Alignment(horizontal='right',vertical='bottom')
        alignment_content=Alignment(horizontal='left',vertical='bottom')
        alignment_ans=Alignment(horizontal='center',vertical='bottom')
        font_heading=Font(name='Century',size=14,bold=False)
        font_content=Font(name='Century',size=14,bold=True)
        right_color=Font(color="00FF00",name='Century',size=14,bold=False)
        wrong_color=Font(color="ff0000",name='Century',size=14,bold=False)
        give_color=Font(color="0000FF",name='Century',size=14,bold=False)
        border_style=Side(border_style="medium",color="000000")
        border=Border(top=border_style,bottom=border_style,left=border_style,right=border_style)
        sheet.add_image(img,"A1")
        sheet["A6"]="Name :"
        sheet["A6"].font=font_heading
        sheet["A6"].alignment=alignment_heading
        sheet["B6"]=file["Name"]
        sheet["B6"].font=font_content
        sheet["B6"].alignment=alignment_content
        sheet["D6"]="Exam :"
        sheet["D6"].font=font_heading
        sheet["D6"].alignment=alignment_heading
        sheet["E6"]="Quiz"
        sheet["E6"].font=font_content
        sheet["E6"].alignment=alignment_content
        sheet["A7"]="Roll Number"
        sheet["A7"].font=font_heading
        sheet["A7"].alignment=alignment_heading
        sheet["B7"]=file["Roll Number"]
        sheet["B7"].font=font_content
        sheet["B7"].alignment=alignment_content
        sheet.row_dimensions[5].height=20
        right=final_record[file["Roll Number"]]["right"]
        wrong=final_record[file["Roll Number"]]["wrong"]
        not_attempted=final_record[file["Roll Number"]]["not answered"]
        data={
            "Right":[right,right_points,right*right_points],
            "Wrong":[wrong,wrong_points,wrong*wrong_points],
            "Not Attempt":[not_attempted,0,''],
            "Max":[right+wrong+not_attempted ,'',
                    str(right*right_points+wrong*wrong_points*-1)+"/"+(str(right_points*(right+not_attempted+wrong)))]    
        }
        rownumber=["B","C","D","E"]
        column_number=["10","11","12"]
        sheet["A9"]=""
        sheet["A10"]= "No."
        sheet["A11"]="Marking"
        sheet["A12"]="Total"
        sheet["B9"]="Right"
        sheet["C9"]="Wrong"
        sheet["D9"]="Not Attempt"
        sheet["E9"]="Max"
        sheet["A15"]="Student Ans"
        sheet["B15"]="Correct Ans"
        sheet["D15"]="Student Ans"
        sheet["E15"]="Correct Ans"
        space=["A9","A10","A11","A12","B9","C9","D9","E9","A15","B15","D15","E15"]
        for p in space:
            sheet[p].font=font_content
            sheet[p].alignment=alignment_ans
            sheet[p].border=border
        for text in data:
            i=0
            if text=="Right":
                row=rownumber[0]
                for column in column_number:
                    sheet[row + column]=data[text][i]
                    sheet[row + column].font=right_color
                    sheet[row + column].alignment=alignment_ans
                    sheet[row + column].border=border
                    i+=1
            elif text=="Wrong":
                row=rownumber[1]
                for column in column_number:
                    sheet[row + column]=data[text][i]
                    sheet[row + column].font=wrong_color
                    sheet[row + column].alignment=alignment_ans
                    sheet[row + column].border=border
                    i+=1
            elif text=="Not Attempt":
                row=rownumber[2]
                for column in column_number:
                    sheet[row + column]=data[text][i]
                    sheet[row + column].alignment=alignment_ans
                    sheet[row + column].font=font_content
                    sheet[row + column].border=border
                    i+=1
            elif text=="Max":     
                row=rownumber[3]
                for column in column_number:
                    sheet[row + column]=data[text][i]
                    sheet[row + column].font=font_content
                    sheet[row + column].alignment=alignment_ans
                    sheet[row + column].border=border
                    i+=1
                sheet["E12"].font=give_color
        correct_answer=answer_record["ANSWER"]
        student_answer=answer_record[file["Roll Number"]]
        student_answer_space=[]
        i1=16
        j1=16
        for p in correct_answer:
        
            if (i1<41):
                sheet["B"+str(i1)]=p
                sheet["B"+str(i1)].font=give_color
                sheet["B"+str(i1)].alignment=alignment_ans
                sheet["B"+str(i1)].border=border
                i1+=1
            else:
                sheet["E"+str(j1)]=p
                sheet["E"+str(j1)].font=give_color
                sheet["E"+str(j1)].alignment=alignment_ans
                sheet["E"+str(j1)].border=border
                j1+=1
        i=16
        j=16
        k=0
        sheet.column_dimensions['A'].width=20
        sheet.column_dimensions['B'].width=20
        sheet.column_dimensions['C'].width=15
        sheet.column_dimensions['D'].width=20
        sheet.column_dimensions['E'].width=20
        for p in student_answer:
            
            if i<41 and p==correct_answer[k]:
                sheet["A"+str(i)]=p
                sheet["A"+str(i)].font=right_color
                sheet["A"+str(i)].alignment=alignment_ans
                sheet["A"+str(i)].border=border
                
                i+=1
            elif i<41 and p!=correct_answer[k]:
                sheet["A"+str(i)]=p
                sheet["A"+str(i)].font=wrong_color
                sheet["A"+str(i)].alignment=alignment_ans
                sheet["A"+str(i)].border=border
                i+=1
            elif i>=41 and p==correct_answer[k]:
                sheet["D"+str(j)]=p
                sheet["D"+str(j)].font=right_color
                sheet["D"+str(j)].alignment=alignment_ans
                sheet["D"+str(j)].border=border
                j+=1
            elif i>=41 and p!=correct_answer[k]:
                sheet["D"+str(j)]=p
                sheet["D"+str(j)].font=wrong_color
                sheet["D"+str(j)].alignment=alignment_ans
                sheet["D"+str(j)].border=border
                j+=1
                
            k+=1
        if os.path.exists(os.path.join(".\outputs",file["Roll Number"]+".xlsx")):
            os.remove(os.path.join(".\outputs",file["Roll Number"]+".xlsx"))
        
        workbook.save(os.path.join(".\outputs",file["Roll Number"]+".xlsx"))

# generate_roll_no_wise_marksheet(4,0)