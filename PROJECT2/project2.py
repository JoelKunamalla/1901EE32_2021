# # from reportlab.pdfgen import canvas

# # pd=canvas.Canvas("jk.pdf")
# # pd.setTitle("joel")
# # pd.save()

# from typing import Any
# from reportlab.pdfgen import canvas
# from reportlab.pdfbase.ttfonts import TTFont
# from reportlab.pdfbase import pdfmetrics
# c = canvas.Canvas("hello11.pdf")
# c.setTitle("joeLLL")
# c.drawString(10,700,"Welcome to Reportlab!1")

# c.drawString(10,700,"Welcome to Reportlab!2")
# c.drawString(10,600,"Welcome to Reportlab!3")
# c.drawString(10,500,"Welcome to Reportlab!4")
# c.drawString(10,70,"Welcome to Reportlab!")
# c.drawBoundary(Any)
# c.save()
import shutil
from PIL.Image import Image
from fpdf import FPDF


# pdf.cell(-50)
import shutil


import csv
import os
if(os.path.exists("./templates/output")) : 
    shutil.rmtree("./templates/output")
else :
    pass
os.mkdir("./templates/output")
if(os.path.exists("generated")) : 
    shutil.rmtree("generated")
else :
    pass
os.mkdir("generated")
with open(f"./templates/names-roll.csv" , "r") as f:
        reader = csv.reader(f  , delimiter=',')
        l=[]
        name=[]
        for r in reader:
            if r[0]== "Roll":
                continue
            l.append(r[0])
            name.append(r[1])
with open(f"./templates/subjects_master.csv" , "r") as f:
        reader = csv.reader(f , delimiter=',')
        d={}
        for r in reader:
            if r[0]== 'subno':
                continue
            d[r[0]]=[r[1] , r[2] , r[3] ]

l11=l.copy()
l2=[]
BRAND=l11[-1]

dictim={x:[] for x in  l11}

with open(f"./templates/grades.csv" , 'r') as f:
        i=2
        rea = csv.reader(f)
        lis=[]
        for r in rea:

            if r[0] == 'Roll':
                continue
            if(os.path.exists(f"./templates/output/{r[0]+r[1]}.csv")):
                    with open(f"./templates/output/{r[0]+r[1]}.csv" , "a" , newline='') as f:
                        fie = [ 'Subject No' , 'Subject Name' , 'L-T-P' , 'Credit' , 'Grade']
                        writer = csv.DictWriter(f , fieldnames=fie)
                        writer.writerow({ 'Subject No':r[2]  , 'Subject Name':d[r[2]][0] , 'L-T-P':d[r[2]][1] , 'Credit':r[3]   , 'Grade':r[4]})
            else:
                lis.append(r[1])
                dictim[r[0]].append(r[1])
                with open(f"./templates/output/{r[0]+r[1]}.csv" , "w" , newline='') as f:
                        fie = [ 'Subject No' , 'Subject Name' , 'L-T-P' , 'Credit' , 'Grade']
                        writer = csv.DictWriter(f , fieldnames=fie)
                        writer.writerow({ 'Subject No':r[2]  , 'Subject Name':d[r[2]][0] , 'L-T-P':d[r[2]][1] ,'Credit':r[3]  , 'Grade':r[4]})









list111=lis.copy()
dict={}
l=True
log=l11.copy()
while(len(log)>0 ):
    l11[0]= FPDF("L" , "mm" ,"A3")
    l11[0].add_page() 
    l11[0].image('./templates/Picture.png' , 9,5,30,25)
    l11[0].image('./templates/Picture.png',375,5,30,25 , title="INTERIM TRANSCRIPT")
    l11[0].image("./templates/Picture3.png",48,7,300,20 , title='transcript')
    l11[0].set_font('times' ,'u',10)
    l11[0].cell(-5.5)
    l11[0].cell(70,45,"INTERIM TRANSCRIPT" )
    l11[0].cell(-66)
    l11[0].set_font('times' ,'',27)
    l11[0].cell(370,38,"Transcript ", align="C" )
    l11[0].cell(-7.5)
    l11[0].set_font('times' ,'u',10)
    l11[0].cell(80,45,"INTERIM TRANSCRIPT")
    l11[0].set_y(-281)
    l11[0].set_x(110)
    l11[0].rect(5.0, 5.0,405.0,285.0)
    l11[0].rect(5.0, 5.0, 405.0,30.0)
    l11[0].rect(5.0, 5.0, 40.0,30.0)
    l11[0].rect(370.0, 5.0, 40.0,30.0)
    l11[0].rect(105.0, 37, 200.0,13.0)
    l11[0].set_font('times' ,'B',12)
    l11[0].cell(160,48,"Roll No:")
    l11[0].rect(130.0, 39, 25.0,4.0)
    l11[0].cell(-110)
    l11[0].set_font('times' ,'B',12)
    l11[0].cell(90,48,"Name:")
    l11[0].rect(180.0, 39, 44.0,4.0)
    l11[0].cell(-20)
    l11[0].cell(90,48,"Year of Admission:")
    l11[0].rect(270.0, 39 , 20.0 ,4.0 )
    l11[0].cell(-210)
    l11[0].cell(90,60,"Programme: ")
    l11[0].set_font('times' ,'',12)
    l11[0].cell(-60)
    l11[0].cell(0,61,"Bachelor of Technology")
    l11[0].cell(-220)
    l11[0].set_font('times' ,'B',12)
    l11[0].cell(0,60,"Course:")
    l11[0].cell(-200)
    l11[0].set_font('times' ,'',12)
    l11[0].cell(0,62,"hh")
    l11[0].line(5, 110, 410,110)
    l11[0].line(5, 170, 410,170)
    l11[0].line(5, 231, 410,231)
    l11[0].cell(-280)
    l11[0].cell(40,50,f"{log[0]}")
    l11[0].cell(18)
    l11[0].cell(100,50,f"{name[0]}")
    l11[0].cell(-12)
    l11[0].cell(10,50,"2019")
    

    zuxxy=0
    luxxy=0
    total_credits=0

    joel=True
    while(True):
        Grades_convert={'AA':10 , 'AB':9 , 'BB':8 ,'BC':7 , 'CC':6 , 'CD':5 , 'DD':4 , 'F':0 , 'I':0  , 'I*':0 , "F*":0 , 'DD*':0}
        xll=3
        while(xll>0):



            if len(list111) >0 and len(log)>0:
                if( not os.path.exists(f"./templates/output/{log[0]+(list111[0])}.csv")):
                    joel=False
                    break
            else:
                break
            l11[0].set_x(30)
            l11[0].set_y(60)
            lo=[]
            l11[0].set_font("times","B",10)
            if(len(dictim[log[0]])>0 and xll==3):
                l11[0].cell(4)
                
                l11[0].cell(9.8,-5,f"semester {dictim[log[0]][0]}")
                l11[0].rect(10,98,100,7)
                dictim[log[0]].pop(0)
            elif ( len(dictim[log[0]])>0 and xll==2):
                l11[0].cell(130)
                l11[0].cell(140.9,-5,f"semester {dictim[log[0]][0]}")
                l11[0].rect(141,98,100,7)
                dictim[log[0]].pop(0)
            elif( len(dictim[log[0]]) >0 and xll==1):
                l11[0].cell(270)
                l11[0].cell(340,-5,f"semester {dictim[log[0]][0]}")
                l11[0].rect(280,98,100,7)
                dictim[log[0]].pop(0)
            
            
            header=['Subject No' , 'Subject Name' , 'L-T-P' , 'Credit' , 'Grade']
            with open(f"./templates/output/{log[0] + list111[0]}.csv" , "r") as f:
                r= csv.reader(f)
                current_sem_credits=0
                luxxy1=0
                for x in r:
                    lo.append(x)
                    total_credits+=int(x[3])
                    current_sem_credits+=int(x[3])
                    luxxy1+=(int(x[3])*Grades_convert[x[4]])
                luxxy1/=current_sem_credits
                zz=round(luxxy1,2)
                luxxy+=(luxxy1*current_sem_credits)

                zuxxy=luxxy/total_credits
                az=round(zuxxy,2)
                l11[0].cell(-12)  
                if xll==3:
                    l11[0].cell(4,82,f"Credits taken:{current_sem_credits}     Credits cleared: {current_sem_credits}    SPI: {zz}    CPI: {az}")
                elif xll==2:
                    l11[0].cell(-126)
                    l11[0].cell(60,82,f"Credits taken:{current_sem_credits}     Credits cleared: {current_sem_credits}    SPI: {zz}    CPI: {az}")
                elif xll==1:
                    l11[0].cell(-326)
                    l11[0].cell(200,82,f"Credits taken:{current_sem_credits}     Credits cleared: {current_sem_credits}    SPI: {zz}    CPI: {az}")








            col_width = 'even'
            if(xll==3):
                l11[0].set_x(9.8)
            elif(xll==2):
                l11[0].set_x(140.9)
            elif(xll==1):
                l11[0].set_x(280)
            col_widths=[20.0,55.0,15.0,15.0,15.0]
            x_start=l11[0].get_x()
            line_height=4
            y1=l11[0].get_y()
            x_left=l11[0].get_x()
            data=lo
            align_header='L'
            
            for i in range(len(header)):
                        datum = header[i]
                        l11[0].set_font("times",'B',7)
                        l11[0].multi_cell(col_widths[i], line_height, datum, border=1, align="C", ln=3, max_line_height=l11[0].font_size)
                        x_right = l11[0].get_x()
            l11[0].ln(line_height)
            y2 = l11[0].get_y()
            l11[0].line(x_left,y1,x_right,y1)
            l11[0].line(x_left,y2,x_right,y2)
            for i in range(len(data)):
            
                if(xll==3):
                    l11[0].set_x(9.8)
                elif(xll==2):
                    l11[0].set_x(140.9)
                elif(xll==1):
                    l11[0].set_x(280)
                row = data[i]
                for i in range(0,len(row)):
                    datum = row[i]
                    l11[0].set_font("times",'',7)
                    l11[0].multi_cell(col_widths[i], line_height, datum, border=1, align="C", ln=3, max_line_height=l11[0].font_size)
                l11[0].ln(line_height)
                y3 = l11[0].get_y()
                l11[0].line(x_left,y3,x_right,y3)
            os.remove(f"./templates/output/{log[0]+list111[0]}.csv")
            list111.pop(0)
            xll-=1
        

        xll=3
        while(xll>0):
            if len(list111) >0 and len(log)>0:
                if( not os.path.exists(f"./templates/output/{log[0]+(list111[0])}.csv")):
                    joel=False
                    break
            else:
                break
            lo=[]
            l11[0].set_x(30)
            l11[0].set_y(120)
            header=['Subject No' , 'Subject Name' , 'L-T-P' , 'Credit' , 'Grade']
            with open(f"./templates/output/{log[0] + list111[0]}.csv" , "r") as f:
                r= csv.reader(f)
                current_sem_credits=0
                luxxy1=0
                for x in r:
                    lo.append(x)
                    total_credits+=int(x[3])
                    current_sem_credits+=int(x[3])
                    luxxy1+=(int(x[3])*Grades_convert[x[4]])
                luxxy1/=current_sem_credits
                zz=round(luxxy1,2)
                luxxy+=(luxxy1*current_sem_credits)

                zuxxy=luxxy/total_credits
                az=round(zuxxy,2)
                l11[0].set_font("times","B",10)
                if xll==3:
                    
                    l11[0].cell(4,90,f"Credits taken:{current_sem_credits}     Credits cleared: {current_sem_credits}    SPI: {zz}    CPI: {az}")
                elif xll==2:
                    l11[0].cell(132)
                    l11[0].cell(20,90,f"Credits taken:{current_sem_credits}     Credits cleared: {current_sem_credits}    SPI: {zz}    CPI: {az}")
                elif xll==1:
                    l11[0].cell(272)
                    l11[0].cell(20,90,f"Credits taken:{current_sem_credits}     Credits cleared: {current_sem_credits}    SPI: {zz}    CPI: {az}")
            col_width = 'even'
            l11[0].set_font("times","B",10)
            if(len(dictim[log[0]])>0 and xll==3):
                l11[0].cell(4)
                l11[0].rect(10,161,100,7)
                l11[0].cell(9.8,-5,f"semester {dictim[log[0]][0]}")
                dictim[log[0]].pop(0)
            elif ( len(dictim[log[0]])>0 and xll==2):
                l11[0].cell(130)
                l11[0].rect(141,161,100,7)
                l11[0].cell(140.9,-5,f"semester {dictim[log[0]][0]}")
                dictim[log[0]].pop(0)
            elif( len(dictim[log[0]]) >0 and xll==1):
                l11[0].cell(270)
                l11[0].rect(280,161,100,7)
                l11[0].cell(340,-5,f"semester {dictim[log[0]][0]}")
                dictim[log[0]].pop(0)



            if(xll==3):
                l11[0].set_x(9.8)
            elif(xll==2):
                l11[0].set_x(140.9)
            elif(xll==1):
                l11[0].set_x(280)
            col_widths=[20.0,55.0,15.0,15.0,15.0]
            x_start=l11[0].get_x()
            line_height=4
            y1=l11[0].get_y()
            x_left=l11[0].get_x()
            data=lo
            align_header='C'
            for i in range(len(header)):
                        datum = header[i]
                        l11[0].set_font("times",'B',7)
                        l11[0].multi_cell(col_widths[i], line_height, datum, border=1, align="C", ln=3, max_line_height=l11[0].font_size)
                        x_right = l11[0].get_x()
            l11[0].ln(line_height)
            y2 = l11[0].get_y()
            l11[0].line(x_left,y1,x_right,y1)
            l11[0].line(x_left,y2,x_right,y2)
            for i in range(len(data)):
            
                if(xll==3):
                    l11[0].set_x(9.8)
                elif(xll==2):
                    l11[0].set_x(140.9)
                elif(xll==1):
                    l11[0].set_x(280)
                row = data[i]
                for i in range(0,len(row)):
                    datum = row[i]
                    l11[0].set_font("times",'',7)
                    l11[0].multi_cell(col_widths[i], line_height, datum, border=1, align="C", ln=3, max_line_height=l11[0].font_size)
                l11[0].ln(line_height)
                y3 = l11[0].get_y()
                l11[0].line(x_left,y3,x_right,y3)
            os.remove(f"./templates/output/{log[0]+list111[0]}.csv")
            list111.pop(0)
            xll-=1




        x12=2
        while(x12>0):
            if len(list111) >0 and len(log)>0:
                if( not os.path.exists(f"./templates/output/{log[0]+(list111[0])}.csv")):
                    joel=False
                    break
                else:
                    pass
            elif len(list111)==0 or len(log)==0:
                break
            else:
                pass
            lo=[]
            l11[0].set_x(30)
            l11[0].set_y(180)
            header=['Subject No' , 'Subject Name' , 'L-T-P' , 'Credit' , 'Grade']
            with open(f"./templates/output/{log[0] + list111[0]}.csv" , "r") as f:

                r= csv.reader(f)
                current_sem_credits=0
                luxxy1=0
                for x in r:
                    lo.append(x)
                    total_credits+=int(x[3])
                    current_sem_credits+=int(x[3])
                    luxxy1+=(int(x[3])*Grades_convert[x[4].strip()])
                luxxy1/=current_sem_credits
                zz=round(luxxy1,2)
                luxxy+=(luxxy1*current_sem_credits)

                zuxxy=luxxy/total_credits
                az=round(zuxxy,2)
                l11[0].set_font("times","B",10)
                if x12==2:
                    
                    l11[0].cell(4,90,f"Credits taken:{current_sem_credits}     Credits cleared: {current_sem_credits}    SPI: {zz}    CPI: {az}")
                elif x12==1:
                    l11[0].cell(132)
                    l11[0].cell(20,90,f"Credits taken:{current_sem_credits}     Credits cleared: {current_sem_credits}    SPI: {zz}    CPI: {az}")
            col_width = 'even'
            l11[0].set_font("times","B",10)
            if(len(dictim[log[0]])>0 and x12==2):
                    l11[0].cell(4)
                    l11[0].rect(10,221.5,100,7)
                    l11[0].cell(9.8,-5,f"semester {dictim[log[0]][0]}")
                    dictim[log[0]].pop(0)
            elif ( len(dictim[log[0]])>0 and x12==1):
                    l11[0].cell(130)
                    l11[0].rect(141,221.5,100,7)
                    l11[0].cell(-153)
                    l11[0].cell(100.9,-5,f"semester {dictim[log[0]][0]}")
                    dictim[log[0]].pop(0)
            if(x12==2):
                l11[0].set_x(9.8)
            elif(x12==1):
                l11[0].set_x(140.9)
            col_widths=[20.0,55.0,15.0,15.0,15.0]
            x_start=l11[0].get_x()
            line_height=4
            y1=l11[0].get_y()
            x_left=l11[0].get_x()
            data=lo
            align_header='L'
            for i in range(len(header)):
                        datum = header[i]
                        l11[0].set_font("times",'B',7)
                        l11[0].multi_cell(col_widths[i], line_height, datum, border=1, align="C", ln=3, max_line_height=l11[0].font_size)
                        x_right = l11[0].get_x()
            l11[0].ln(line_height)
            y2 = l11[0].get_y()
            l11[0].line(x_left,y1,x_right,y1)
            l11[0].line(x_left,y2,x_right,y2)
            for i in range(len(data)):
            
                if(x12==2):
                    l11[0].set_x(9.8)
                elif(x12==1):
                    l11[0].set_x(140.9)
                row = data[i]
                for i in range(0,len(row)):
                    datum = row[i]
                    l11[0].set_font("times",'',7)
                    l11[0].multi_cell(col_widths[i], line_height, datum, border=1, align="C", ln=3, max_line_height=l11[0].font_size)
                l11[0].ln(line_height)
                y3 = l11[0].get_y()
                l11[0].line(x_left,y3,x_right,y3)
            os.remove(f"./templates/output/{log[0]+list111[0]}.csv")
            list111.pop(0)
            x12-=1


        if log[0]==BRAND:
            hh=log[0]
            l11[0].set_y(-30)
            l11[0].cell(300)
            l11[0].set_font('times','',15)
            l11[0].cell(200,10,"Assistant Registar" , ln=0)
            l11[0].output(f"./generated/{hh}.pdf")



        if(len(list111)==0):
            l= False
            break

        if( not os.path.exists(f"./templates/output/{log[0]+list111[0]}.csv")):
            hh=log[0]
            l11[0].set_y(-30)
            l11[0].cell(300)
            l11[0].set_font('times','',15)
            l11[0].cell(200,10,"Assistant Registar" , ln=0)
            l11[0].cell(-500)
            l11[0].cell(20,10,"Date of issue:")
            l11[0].cell(150)
            l11[0].cell(20,10,'stamp')
            l11[0].image("./storage/Picture3.png",48,200,20,50 , title='transcript')
            l11[0].cell(20,10,"Date of Issue:")
            l11[0].output(f"./generated/{hh}.pdf")
            print(hh)
            print(log)
            print(list111)
            log.pop(0)
            break
    name.pop(0)
    if l == False:
        break












l11[0]= FPDF("L" , "mm" ,"A4")
l11[0].add_page() 
l11[0].image('./templates/Picture.png' , 9,5,30,25)
l11[0].image('./templates/Picture.png',257,5,30,25 , title="INTERIM TRANSCRIPT")
l11[0].image("./templates/Picture3.png",48,7,200,20 , title='transcript')
l11[0].set_font('times' ,'u',10)
l11[0].cell(-5.5)
l11[0].cell(70,45,"INTERIM TRANSCRIPT" )
l11[0].cell(-66)
l11[0].set_font('times' ,'',27)
l11[0].cell(270,38,"Transcript ", align="C" )
l11[0].cell(-27)
l11[0].set_font('times' ,'u',10)
l11[0].cell(90,45,"INTERIM TRANSCRIPT")
l11[0].set_y(-192)
l11[0].set_x(50)
l11[0].rect(5.0, 5.0,287.0,200.0)
l11[0].rect(5.0, 5.0, 287.0,30.0)
l11[0].rect(5.0, 5.0, 40.0,30.0)
l11[0].rect(252.0, 5.0, 40.0,30.0)
l11[0].rect(50.0, 37, 200.0,13.0)
l11[0].set_font('times' ,'B',12)
l11[0].cell(90,48,"Roll No:")
l11[0].rect(70.0, 39, 25.0,4.0)
l11[0].cell(-30)
l11[0].set_font('times' ,'B',12)
l11[0].cell(90,48,"Name:")
l11[0].rect(130.0, 39, 40.0,4.0)
l11[0].cell(-20)
l11[0].cell(90,48,"Year of Admission:")
l11[0].rect(220.0, 39, 20.0,4.0)
l11[0].cell(-220)
l11[0].cell(90,58,"Programme: ")
l11[0].set_font('times' ,'',12)
l11[0].cell(-60)
l11[0].cell(0,58,"Bachelor of Technology")
l11[0].cell(-150)
l11[0].set_font('times' ,'B',12)
l11[0].cell(0,58,"Course:")
l11[0].cell(-20)
l11[0].set_font('times' ,'',12)
l11[0].cell(0,74,"")
l11[0].line(5, 110, 292,110)
l11[0].line(5, 170, 292,170)