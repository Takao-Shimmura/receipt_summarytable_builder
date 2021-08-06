#!python3.9.1
from flask import Flask, render_template, request, session, \
    redirect, jsonify, current_app, g,send_file,make_response,



import openpyxl # 外部ライブラリ　pip install openpyxl
import sqlite3
import json
from datetime import datetime

import pathlib

from sqlalchemy import create_engine, Column, Integer, String, \
    Text, DateTime, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from sqlalchemy.orm.exc import NoResultFound

from myutil import User,Calculate,ErrorMsg,Search_condition,InsurerData,\
    get_dic_schCond2calAttr,get_search_condition,get_by_list,\
        get_cellno_2list,define_soukatsu1Desti,sort_insureName_4Sokatsu1,\
        soukatsu1Desti_List_set,get_soukatsu1Desti_count,get_soukatsu1Desti_insur_dic,\
        kensuu_insert,kingaku_insert



app = Flask(__name__)
app.secret_key = b'random string...'



# access top page.
@app.route('/',methods=['GET'])
def index2():
        

        return render_template('index2.html')
#post message
@app.route('/preget',methods=['GET'])
def get_filename():
    alt_data=[]
    list4=[]
    list4.append('★読み込みができなかったシート')
    list4.append([1,'「療養を受けた者の氏名」の記入漏れ'])
    list4.append([2,'「療養を受けた者の氏名」(フリガナ))の記入漏れ'])
    list4.append([3,'「保険者番号」の記入漏れ'])
    list4.append([12,'施術管理者の「登録記号番号」の記入漏れもしくは記入ミス'])
    list4.append([13,'申請書冒頭の申請「年」か、施術期間の「年」のいずれかの記入漏れもしくは記入ミス'])
    list4.append([14,'申請書冒頭の申請「月」か、施術期間の「月」のいずれかの記入漏れもしくは記入ミス'])
    list4.append([15,'「保険者番号」から保険者が特定できません　保険者番号の記入ミスもしくは、ホームページ管理者による「保険者番号の登録漏れ」です'])
    alt_data.append(list4)
    return jsonify(alt_data)

@app.route('/get',methods=['GET'])
def get_msg():
    
    return send_file('soukatsuTemp.xlsx',
                    attachment_filename='soukatsuTemp.xlsx',
                    as_attachment=True,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    
""" @app.route('/get',methods=['GET'])
def get_msg():
    response = make_response()

    # ★ポイント2
    response.data = open("soukatsuTemp.xlsx", "rb").read()

    # ★ポイント3
    downloadFileName = 'soukatsuTemp.xlsx'    
    response.headers['Content-Disposition'] = 'attachment; filename=' + downloadFileName

    # ★ポイント4
    response.mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response """

if __name__=='__main__':
    app.debug = True

    #app.run(host='0.0.0.0')
    app.run(host='localhost')