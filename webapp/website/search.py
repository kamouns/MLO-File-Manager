import openpyxl
import pandas as pd
from flask import Blueprint , render_template,request
import glob

def Filelistfun() :
    fileslist=glob.glob("website\ExcelFiles/*.xlsx")
    fileslistname=[]
    for f in fileslist: 
        fileslistname.append(str(f).replace('website\ExcelFiles\\',''))
    return fileslistname

search=Blueprint('search', __name__)
@search.route('/search' , methods=['GET', 'POST'])
def searching():
    return render_template("search.html")


@search.route('/result',methods=['GET', 'POST'])
def result():

    if request.method == 'POST':   
        fn =(request.form['button'])
        fileNameNoXlsx=fn.replace('.xlsx','')
        f ='website/ExcelFiles/'+ fn
        ps =openpyxl.load_workbook(str(f))
        sh= ps['Feuil1']
    return render_template("result.html" , sh=sh , fnx=fileNameNoXlsx)

@search.route('/all')
def all():
    fileslistname= Filelistfun()
    return render_template("all.html" ,fln=fileslistname)

@search.route('/searchResult',methods=['GET', 'POST'])
def searchresult(): 
    fileslistname= Filelistfun()
    if request.method=='POST':
        input = request.form['mlo']
    return render_template( "searchres.html", mloName=input, fln=fileslistname)