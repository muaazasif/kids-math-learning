from flask import Flask,render_template,request,send_file
from openpyxl import Workbook,load_workbook
import os
app=Flask(__name__)
F='data/results.xlsx'
os.makedirs('data',exist_ok=True)
@app.get('/')
def home(): return render_template('index.html')
@app.post('/save')
def save():
 d=request.json;ans=int(d['answer']);correct=5;marks=100 if ans==correct else 0
 wb=load_workbook(F) if os.path.exists(F) else Workbook();ws=wb.active
 if ws.max_row==1 and ws['A1'].value is None: ws.append(['Question','Answer','Correct','AI Marks'])
 ws.append(['2+3',ans,ans==correct,marks]);wb.save(F);return {'marks':marks}
@app.get('/download')
def dl(): return send_file(F,as_attachment=True)
app.run()
