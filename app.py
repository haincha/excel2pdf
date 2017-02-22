import sys
import os
from flask import Flask, request, render_template, jsonify, make_response, send_file, flash, url_for, redirect, Markup, Response
import pyexcel
import HTML
import pdfkit
import zipfile
import datetime

app = Flask(__name__)
app.secret_key = 'some_secret'

@app.route("/checker", methods=['GET', 'POST'])
def checker():
    today = datetime.date.today().strftime("%m-%d-%Y")
    if request.method == 'POST':
        numbers = request.form.getlist('accounts')
        current_date = request.form.getlist('date')
        accountlist = numbers[0].splitlines()
        accountlist = [i.strip() for i in accountlist]
        missing_account = []
        missing_count = 0
        for i in range(0,len(accountlist)):
            if os.path.exists('/mnt/consentorders/' + str(current_date[0]) + '/' + str(accountlist[i]) + '.pdf') == False:
                flash(Markup(str(accountlist[i]).strip()))
                missing_count += 1
        flash(Markup("There was " + str(missing_count) + " missing account(s)"))
        return render_template('checker.html')
    return render_template('checker.html', today=today)

@app.route("/delete", methods=['GET', 'POST'])
def delete():
    today = datetime.date.today().strftime("%m-%d-%Y")
    if request.method == 'POST':
        numbers = request.form.getlist('accounts')
        current_date = request.form.getlist('date')
        accountlist = numbers[0].splitlines()
        accountlist = [i.strip() for i in accountlist]
        delete_account = []
        delete_count = 0
        for i in range(0,len(accountlist)):
            if os.path.exists('/mnt/consentorders/' + str(current_date[0]) + '/' + str(accountlist[i]) + '.pdf') == True:
                os.remove('/mnt/consentorders/' + str(current_date[0]) + '/' + str(accountlist[i]) + '.pdf')
                flash(Markup(str(accountlist[i]).strip()))
                delete_count += 1
        flash(Markup("There was " + str(delete_count) + " account(s) deleted."))
        return render_template('delete.html')
    return render_template('delete.html', today=today)

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST' and 'excel' in request.files:
        try:
            filename = request.files['excel'].filename
            extension = filename.split(".")
            extension = extension[len(extension)-1]
            content = request.files['excel'].read()
            numbers = request.form.getlist('accounts')
            starttab = request.form.getlist('starttab')
            accountlist = numbers[0].splitlines()
            accountlist = [i.strip() for i in accountlist]
            styling = 'display: inline; page-break-before: auto; padding-bottom: 50%; font-family: Calibri; font-size: 8.76;'
            wb = pyexcel.get_book(file_type=extension, file_content=content)
            sheets = wb.to_dict()
            all_sheets = []
            found_accounts = 0
            account_column = 0
            header_column = 0
            header_found = False
            all_accounts = []
            for name in sheets.keys():
                all_sheets.append(name)
            if isinstance(starttab, int) == False:
                starttab = 0
            if starttab > len(all_sheets):
                starttab = 0
            if starttab > 1:
                for z in range(0,starttab-1):
                    all_sheets.remove(all_sheets[z])
            for k in all_sheets:
                for i in range(0,len(wb[k].column[1])):
                    if len(accountlist) == 0:
                        return render_template('upload.html')
                    for l in range(0,len(wb[k].row[0])):
                        if ("acct" in str(wb[k][i,l]).lower() or "account" in str(wb[k][i,l]).lower() or "userid" in str(wb[k][i,l]).lower() or "loanid" in str(wb[k][i,l]).lower() or "agyid" in str(wb[k][i,l]).lower() or "chd" in str(wb[k][i,l]).lower() or "t_num" in str(wb[k][i,l]).lower() or "loan" in str(wb[k][i,l]).lower()) and str(wb[k][i+1,l]).strip().isdigit() == True and header_found == False:
                            header_column = i
                            header_found = True
                        if ("acct" in str(wb[k][header_column,l]).lower() or "account" in str(wb[k][header_column,l]).lower() or "userid" in str(wb[k][header_column,l]).lower() or "loanid" in str(wb[k][header_column,l]).lower() or "agyid" in str(wb[k][header_column,l]).lower() or "chd" in str(wb[k][header_column,l]).lower() or "t_num" in str(wb[k][header_column,l]).lower() or "loan" in str(wb[k][header_column,l]).lower()) and (str(wb[k][i,l]).strip() in accountlist):
                    		    account_column = l
                for m in range(1,len(wb[k].column[1])):
                    all_accounts.append([str(wb[k][m,account_column]).strip(),m])
                for m in range(0,len(accountlist)):
                    if found_accounts == len(accountlist):
                        break
                    for n in range(0,len(all_accounts)):
                        if accountlist[m] in all_accounts[n][0]:
                            htmlcode = HTML.table()
                            wbname = str(accountlist[m]).strip()
                            if not os.path.exists('/mnt/consentorders/' + str(datetime.date.today().strftime("%m-%d-%Y")) + '/'):
                                os.makedirs('/mnt/consentorders/' + str(datetime.date.today().strftime("%m-%d-%Y")) + '/')
                            if os.path.exists('/mnt/consentorders/' + str(datetime.date.today().strftime("%m-%d-%Y")) + '/' + str(wbname) + '.pdf') == False:
                                for o in range(0,len(wb[k].row[0])):
                                    header = str(wb[k][header_column,o])
                                    row = str(wb[k][all_accounts[n][1],o])
                                    if ("ssn" in str(header).lower() or "tax" in str(header).lower() or "social" in str(header).lower() or str(header).lower() == 'tin' or "soc_sec_num" in str(header).lower() or "ss #" in str(header).lower()) and (len(str(row)) != "0" or len(str(row)) != "1"):
                                        if len(str(row)) == 8:
                                            htmlcode += HTML.table([[str(header)],[str('XXX-XX-X' + str(row)[5:])]],border=0,style=(styling))
                                        elif len(str(row)) == 11:
                                            htmlcode += HTML.table([[str(header)],[str('XXX-XX-X' + str(row)[-3:])]],border=0,style=(styling))
                                        elif len(str(row).strip()) <= 1:
                                            htmlcode += HTML.table([[str(header)],[str(row).strip()]],border=0,style=(styling))
                                        else:
                                            htmlcode += HTML.table([[str(header)],[str('XXX-XX-X' + str(row)[6:])]],border=0,style=(styling))
                                    elif len(str(row)) == 9 and is_int(str(row)) == True:
                                        htmlcode += HTML.table([[str(header)],[str(str(row)[0:5] + "-" + str(row)[5:])]],border=0,style=(styling))
                                    elif len(str(row)) == 10:
                                        if isinstance(row, int) == True:
                                            if "ph" in str(header).lower() and int(row) > 10000000:
                                                htmlcode += HTML.table([[str(header)],[str("(" + str(row)[0:3] + ") " + str(row)[3:6] + "-" + str(row)[6:])]],border=0,style=(styling))
                                        elif isinstance(row, datetime.date) == True:
                                            htmlcode += HTML.table([[str(header)],[str(row.strftime("%m-%d-%Y")).strip()]],border=0,style=(styling))
                                        else:
                                            htmlcode += HTML.table([[str(header)],[str(row).strip()]],border=0,style=(styling))
                                    elif "email" in str(header).lower():
                                        if len(str(row)) >= 1:
                                            htmlcode += HTML.table([[str(header)],[str('XXXXX')]],border=0,style=(styling))
                                        else:
                                            htmlcode += HTML.table([[str(header)],[str(row).strip()]],border=0,style=(styling))
                                    elif "sale_price" in str(header).lower() or "proceeds" in str(header).lower():
                                        if len(str(row)) >= 1:
                                            htmlcode += HTML.table([[str(header)],[str('XXXXX')]],border=0,style=(styling))
                                        else:
                                            htmlcode += HTML.table([[str(header)],[str(row).strip()]],border=0,style=(styling))
                                    elif is_float(str(row)) == True and is_int(str(row)) == False:
                                        floatnumber = row.split(".")
                                        htmlcode += HTML.table([[str(header)],[str(floatnumber[0]) + "." + str(floatnumber[1][0:2]).strip()]],border=0,style=(styling))
                                    else:
                                        htmlcode += HTML.table([[str(header)],[str(row).strip()]],border=0,style=(styling))
                                pdfkit.from_string(htmlcode, '/mnt/consentorders/' + str(datetime.date.today().strftime("%m-%d-%Y")) + '/' + str(wbname).strip() + '.pdf', options={'orientation': 'Landscape', 'quiet': ''})
                                found_accounts += 1
            flash(Markup(str(found_accounts) + " file(s) have been converted into PDF."))
            return render_template("upload.html")
        except:
            return render_template('upload.html')
    return render_template("upload.html")

def is_float(input):
    try:
        num = float(input)
    except ValueError:
        return False
    return True

def is_int(input):
    try:
        num = int(input)
    except ValueError:
        return False
    return True

if __name__ == "__main__":
    # start web server
    app.run(
        #debug=True
        threaded=True
    )
