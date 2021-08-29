
from flask import Flask,render_template,request,send_file,send_from_directory
import pandas as pd
import time
import os
import win32gui, win32con
import pytesseract
import base64
import pywinauto
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import keyboard as kb
import pytesseract
import base64
from io import BytesIO
from PIL import Image
import pyautogui as pg
from functools import wraps
import datetime
app=Flask(__name__)
ts = datetime.datetime.now()
import uuid
v=str(uuid.uuid1())
from pymongo import MongoClient
import pymongo
connection_url = "mongodb+srv://Nikhil:mlg@mca.6gfn0.mongodb.net/myFirstDatabase?retryWrites=true&w=majority"
client = pymongo.MongoClient(connection_url)
try:
  connection_url = "mongodb+srv://Nikhil:mlg@mca.6gfn0.mongodb.net/myFirstDatabase?retryWrites=true&w=majority"
  client = pymongo.MongoClient(connection_url)
  print("Connected successfully!!!")
except:
	print("vishnuu")
db = client.get_database('WebScraping')


app.config["UPLOAD_FOLDER1"]="static/excel"
@app.route("/",methods=['GET','POST'])

def upload():

    if request.method == 'POST':
        upload_file = request.files['upload_excel']
        collection = db.PDF_data
        iii=pd.read_excel(upload_file)
        link=iii['DIN/PAN(1)'].to_list()
        alldetails=[]
        if upload_file.filename != '':

            file_path = os.path.join(app.config["UPLOAD_FOLDER1"], upload_file.filename)
            program_path=r"Acrobat Reader DC\Reader\AcroRd32.exe"
            file_path="Form_INC-7.pdf"
            Application().start(r'{} "{}"'.format(program_path,file_path))
            appl=Application().start(r'{} "{}"'.format(program_path,file_path))
            time.sleep(10)
            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            for i in range(0,len(link),1):
                start=time.time()
                pg.dragTo(1527,432, button='left')
                pg.click(button='right')
                pg.click(1474,460,1)
                pg.moveTo(1062,270,1)
                pg.click(1062,270,3)
                pg.typewrite(str('0')+str(link[i]))
                pg.moveTo(1062,320,1)
                pg.click(1062,320,3)
                time.sleep(1)
                pg.moveTo(1297,578,1)
                pg.click(1297,578,1)



                left  = 376
                top   = 404
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                b= pytesseract.image_to_string(area)
                bb=b[:b.find("\n")]

                left  = 375
                top   = 359
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                a= pytesseract.image_to_string(area)
                aa=a[:a.find("\n")]

                left  = 374
                top   = 451
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                c= pytesseract.image_to_string(area)
                cc=c[:c.find("\n")]

                left  = 225
                top   = 579
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                zz= pytesseract.image_to_string(area)
                z=zz[:zz.find("\n")]

                left  = 964
                top   = 621
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                d= pytesseract.image_to_string(area)
                dd=d[:d.find("\n")]

                left  = 681
                top   = 923
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                e= pytesseract.image_to_string(area)
                ee=e[:e.find("\n")]

                pg.dragTo(1530,506, button='left')
                pg.click(button='left')

                left  = 485
                top   = 477
                width = 600
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                f= pytesseract.image_to_string(area)
                ff=f[:f.find("\n")]

                left  = 486
                top   = 519
                width = 600
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                g= pytesseract.image_to_string(area)
                gg=g[:g.find("\n")]

                left  = 486
                top   = 569
                width = 600
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                h= pytesseract.image_to_string(area)
                hh=h[:h.find("\n")]

                left  = 484
                top   = 608
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                i= pytesseract.image_to_string(area)
                ii=i[:i.find("\n")]

                left  = 483
                top   = 696
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                j= pytesseract.image_to_string(area)
                jj=j[:j.find("\n")]

                left  = 1107
                top   = 605
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                k= pytesseract.image_to_string(area)
                kk=k[:k.find("\n")]

                left  = 485
                top   = 831
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                l= pytesseract.image_to_string(area)
                ll=l[:l.find("\n")]

                left  = 375
                top   = 359
                width = 300
                height = 37
                box = (int(left), int(top), int(left+width), int(top+height))
                img =pg.screenshot()
                area = img.crop(box)
                pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"
                m= pytesseract.image_to_string(area)
                mm=m[:m.find("\n")]

                end = time.time()
                temp={'first name':aa,'middle name':bb,'surname':cc,'father name':z,'date of birth':dd,'pan':ee,'line1':ff,'line2':gg,'city':hh,'state':ii,'country':jj,'pincode':kk,'email':ll,'mobile':mm}
                tempp={'database':v,'timestamp':ts,'upload_name':upload_file.filename,'first name':aa,'middle name':bb,'surname':cc,'father name':z,'date of birth':dd,'pan':ee,'line1':ff,'line2':gg,'city':hh,'state':ii,'country':jj,'pincode':kk,'email':ll,'mobile':mm}
                rec_id1 = collection.insert_one(tempp)
                alldetails.append(temp)
                vi=pd.DataFrame(alldetails)
                df=vi.drop_duplicates(subset='pan',keep='first')
                df.to_excel("data.xlsx")
            temp={'first name':'-','middle name':'-','surname':'-','father name':'-','date of birth':'-','pan':'-','line1':'-','line2':'-','city':'-','state':'-','country':'-','pincode':'-','email':'-','mobile':'-'}
            return render_template("download.html")
    return render_template("UploadExcel.html",variable=v)
@app.route('/download')
def download():
	path = "data.xlsx"
	return send_file(path, as_attachment=True,cache_timeout=0)
if __name__=='__main__':
    app.run(host="localhost", port=50002, debug=True)
