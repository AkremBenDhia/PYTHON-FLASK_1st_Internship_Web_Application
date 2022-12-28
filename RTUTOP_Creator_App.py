from flask import Flask, request, render_template, send_file
import os.path
import xlwt
import xlsxwriter
import numpy as np
import io


import pandas as pd




app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')


@app.route('/data', methods=['GET', 'POST'])
def data():
    if request.method == 'POST':


        excel_filepath1 = request.form['t1']
        excel_filepath2 = request.form['t2']

        columns = ['Name', 'Code', 'type', 'Devtyp', 'Device', 'Point', 'desc', 'Phyadr', '--', '//', '..', '$', '///',
                   '////', '/////']

        dfliste = pd.read_excel(excel_filepath1, header=None)
        list = []
        list2 = []
        for k in range(len(dfliste)):
            list.append(dfliste.loc[k, 7])
        s = 0
        for l in range(len(dfliste)):
            s = 0
            for m in range(len(list)):
                if list[m] == dfliste.loc[l, 7]:
                    s = s + 1
            if s > 1:
                x = ("doublon de", dfliste.loc[l, 7], "ligne", l+1)
                list2.append(x)

        dfrtutop = pd.read_excel(excel_filepath2)

        dfrtutop = pd.ExcelFile(excel_filepath2)

        with pd.ExcelWriter('output.xlsx') as writer:

            for i in range(len(dfliste)):

                cle = dfliste.loc[i, 3] + dfliste.loc[i, 4] + dfliste.loc[i, 5]

                for sheet in dfrtutop.sheet_names:

                    df = pd.read_excel(excel_filepath2, sheet_name=sheet, header=None)

                    for j in range(len(df)):

                        cle2 = str(df.loc[j, 6]) + str(df.loc[j, 7]) + str(df.loc[j, 8])

                        if cle == cle2:

                            column = dfliste.loc[i, 7]
                            df.loc[j, 3] = column
                            df.to_excel(writer, sheet_name=sheet)
                        else:

                            df.to_excel(writer, sheet_name=sheet)

        return render_template('index.html', data=list2)


@app.route('/download')
def download_file():
    p = "output.xlsx"
    return send_file(p, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
