import sys
import os
import random
import time
from flask import Flask, request, render_template, jsonify, make_response, send_file, flash, url_for, redirect, Markup, Response
import pyexcel
import HTML
import pdfkit
import zipfile
import datetime
from celery import Celery
import uuid

app = Flask(__name__)
app.secret_key = 'some_secret'
app.config['CELERY_BROKER_URL'] = 'redis://localhost:6379/0'
app.config['CELERY_RESULT_BACKEND'] = 'redis://localhost:6379/0'
app.config['UPLOAD_FOLDER'] = 'uploads/'

celery = Celery(app.name, broker=app.config['CELERY_BROKER_URL'])
celery.conf.update(app.config)


@celery.task(bind=True)
def long_task(self):
	"""Background task that runs a long function with progress reports."""
	array = os.listdir(app.config['UPLOAD_FOLDER'])
	self.update_state(state='PROGRESS',meta={'current': 0, 'total': 1, 'status': 'working...'})
	for i in array:
		if '.xlsx' in i:
			filename = i
	wb = pyexcel.get_sheet(file_name=str(os.path.join(app.config['UPLOAD_FOLDER'], filename)))
	for i in range(0,10):
		number = i * i
		self.update_state(state='PROGRESS',meta={'current': i, 'total': 10, 'status': number})
		time.sleep(1)
	return {'current': 100, 'total': 100, 'status': 'Task completed!','result': 200}

@app.route('/longtask', methods=['POST'])
def longtask():
	task = long_task.apply_async()
	return jsonify({}), 202, {'Location': url_for('taskstatus', task_id=task.id)}

@app.route('/status/<task_id>')
def taskstatus(task_id):
	task = long_task.AsyncResult(task_id)
	if task.state == 'PENDING':
		response = {
			'state': task.state,
			'current': 0,
			'total': 1,
			'status': 'Pending...'
		}
	elif task.state != 'FAILURE':
		response = {
			'state': task.state,
			'current': task.info.get('current', 0),
			'total': task.info.get('total', 1),
			'status': task.info.get('status', '')
		}
		if 'result' in task.info:
			response['result'] = task.info['result']
	else:
		# something went wrong in the background job
		response = {
			'state': task.state,
			'current': 1,
			'total': 1,
			'status': str(task.info),  # this is the exception raised
		}
	return jsonify(response)

@app.route('/', methods=['GET', 'POST'])
def upload():
	if request.method == 'POST' and 'excel' in request.files:
		file = request.files['excel']
		filename = request.files['excel'].filename
		file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
	return render_template("upload.html")

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

if __name__ == "__main__":
	# start web server
	app.run(
		#debug=True
		threaded=True,
		host='0.0.0.0',
		port=80
	)
