import requests
import urllib
import telepot
from telegram.ext import Updater
from telegram.ext import CommandHandler, CallbackQueryHandler
from pprint import pprint
from flask import Flask,redirect, url_for,request,render_template
from threading import Thread
from datetime import datetime,timedelta
from time import sleep
import telegram
import os 
from telepot.loop import MessageLoop
import comtypes.client
import sys
import pythoncom


app = Flask(__name__)

domain='docxtopdf'
TOKEN = '1067124222:AAEy8JP89d65o55wPwqFPMqEAmF_bLMRZjI'
URL = "https://api.telegram.org/bot{}/".format(TOKEN)
bot = telegram.Bot(TOKEN)
print('Started')
@app.route('/')
def main():
    return f'''<html><head><title>Telegram Echobot</title></head><body>Telegram Echobot</body></html>'''
def alink(s,k=None):
    if not k:k=s
    return '<a href="{}">{}</a>'.format(s,k)
def slink(s,k=None,v=1):
    ss=s
    for i,j in (('__','\n'),('--',' '),('-','.'),('_','@'),):
            s=s.replace(j,i)
    if k:ss=k
    s=f'https://t.me/smartmanojbot?start={s}'
    return '<a href="{}">{}</a>'.format(s,ss) if v else s
def tclink(s,d=None):
    if d:return f"+91{s[-10:]} | {slink(f'w.{s}',d)} |"
    return slink(f'w.{s}',f"+91{s[-10:]}")
def get_updates(offset=None):
    url = URL + "getUpdates?timeout=100"
    if offset:
        url += "&offset={}".format(offset)
    return  requests.get(url).json()

def get_last_update_id(updates):
    update_ids = []
    for update in updates["result"]:
        update_ids.append(int(update["update_id"]))
    return max(update_ids)



def main():
    last_update_id = None
    while True:
        try:
            updates = get_updates(last_update_id)           
            z=updates.get("result")
            if z and len(z) > 0:
                last_update_id = get_last_update_id(updates) + 1
                echo_all(updates)
            sleep(0.5)
        except Exception as e:  
            print(e)       


def echo_all(updates):
    for update in updates['result']:
        try:
        	print(update)
        	updates = bot.get_updates()
        	chat_id = bot.get_updates()[-1].message.chat_id
        	file_name=bot.get_updates()[-1].message.document.file_name
        	file_id=bot.get_updates()[-1].message.document.file_id
       		newfile=bot.get_file(file_id)
       		newfile.download(file_name)
       		w2p(name=file_name)
       		#bot.send_document(chat_id=chat_id,document=open("./"+file_name,"rb"))
       		bot.send_message(chat_id=chat_id,text="Thank you for your patience 😊.")
       	except Exception as e:
       		print(e)
def w2p(name):
	wdFormatPDF = 17
	in_file ="./"+name
	out_file ="./"+name+".pdf"
	pythoncom.CoInitialize()
	word = comtypes.client.CreateObject('Word.Application')
	doc = word.Documents.open(in_file)
	doc.SaveAs(out_file, FileFormat=wdFormatPDF)
	doc.Close()
	word.Quit()
	bot.send_document(chat_id=chat_id,document=open("./"+name+".pdf","rb"))


def msg(text,fname='',chat=0):
    if text:
        print(chat,text)
    text=text.lower()
    text=text.replace('-','')
    text=text.strip('/')
    if text.startswith('start'):
        v=text.split()
        if len(v)==1:
            text='Welcome {}'.format(fname)
            send_message(text,chat)
            text='/help'
        else:
            msgf(v[1],chat,name=name);return
    elif text.startswith('help'):
        h='''
/help
'''
        text=(h)
    elif text.startswith('hi'):text='Hi buddy'
            #customize here
    if text: print(chat,text)
    send_message(text,chat)

def snt(f,a,b=None):
  try:
    Thread(None,f,None,a,b).start()
  except Exception as e:        
    return str(e)


def restart():
  requests.head(f'http://{domain}.herokuapp.com/gtcheck',timeout=50)
  while True:
    try:
      v=(datetime.utcnow()+timedelta(hours=5,minutes=30))
      if(1 or 5*60<v.hour*60+v.minute<21*60+30):
        requests.head(f'http://{domain}.herokuapp.com/up/pys',timeout=50)
      sleep(25*60)
    except Exception as e:
      exception(e)
      sleep(2*60)
      continue

snt(main,())
snt(restart,())



if __name__ == '__main__':
    app.run()