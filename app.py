#!/usr/bin/env python
import os,sys,json,xlrd,requests,logging
from flask import Flask
from flask import request
from flask import make_response

def excel2json(workbook):
	book = xlrd.open_workbook(workbook)
	sheets = book.sheet_names()
	source={}
	sessiondetails = {"signedin":False,"phonenumber":""}
	for sheet in sheets:
		source[sheet] = []
		page = book.sheet_by_name(sheet)
		for row in range(1,page.nrows):
			data = {}
			for col in range(0,page.ncols):
				data[str(page.cell(0,col).value)] = str(page.cell(row,col).value)
			source[sheet].append(data)
	return source,sessiondetails

def Remove(duplicate): 
    final_list = [] 
    for num in duplicate: 
        if num not in final_list: 
            final_list.append(num) 
    return final_list

workbook = "Ahold.xlsx"
book,session = excel2json(workbook)

app = Flask(__name__)
app.logger.addHandler(logging.StreamHandler(sys.stdout))
app.logger.setLevel(logging.ERROR)

@app.route('/', methods=['POST', 'GET'])
def webhook():
	if request.method == 'POST':
		req = json.loads((request.data).decode("utf-8"))
		request_data = {"known":{}, "unknown":"", "fulfillmentText":"", "result" : ""}
		for key in req['queryResult']['parameters'].keys():
			request_data["known"].update({key : req['queryResult']['parameters'][key]})
		request_data["unknown"] = str(req['queryResult']['intent']['displayName'])
		request_data["fulfillmentText"] = str(req['queryResult']['fulfillmentText'])
		print (request_data)
		
		if request_data["unknown"] == "welcome":
			if session["signedin"]:
				request_data["result"] = "how may i assist you?"
			else:
				request_data["result"] = request_data["fulfillmentText"]
		
		if request_data["unknown"] == "phonenumber-yes" or request_data["unknown"] ==  "phonenumber":
			session["phonenumber"],session["signedin"],request_data["result"] = request_data['known']['phone-number'],True,request_data["fulfillmentText"].replace("*result",str([request_data['known']['phone-number'][i:i+1] for i in range(0,len(request_data['known']['phone-number']),1)]).replace(" ","").replace("'","").replace("[","").replace("]","").replace(","," "))
		
		if request_data["unknown"] == "phonenumber-no":
			session["phonenumber"],session["signedin"],request_data["result"] = "",True,request_data["fulfillmentText"]
			
		if request_data["unknown"] == "Thankyou" or request_data["unknown"] == "nothing":
			session["phonenumber"],session["signedin"],request_data["result"] = "",False, request_data["fulfillmentText"]
		"""
		if request_data["unknown"] == "order-yes":
			request_data["result"] = request_data["fulfillmentText"].replace("*phone-number",str(session["phonenumber"][i:i+1] for i in range(0,len(session["phonenumber"]),1)]).replace(" ","").replace("'","").replace("[","").replace("]","").replace(","," "))
		"""	
		if request_data["unknown"] == "product":
			if session["signedin"]:
				availables,outofstocks = [],[]
				for sheet in book.keys():
					for row in book[sheet]:
						headers = row.keys()
						if "quantity" in headers:
							for key in request_data['known']:
								for header in headers:
									if len(str(request_data['known'][key]))>0:
										if request_data['known'][key] in row[header]:
											if int(float(row["quantity"]))>0:
												availables.append(row[request_data["unknown"]])
											else:
												outofstocks.append(row[request_data["unknown"]])
				availables = Remove(availables)
				outofstocks = Remove(outofstocks)
				if len(availables)>0:
					request_data["result"] = str(request_data["fulfillmentText"]).replace('*result','available').replace('*availables ',str(availables).replace("[","").replace("]","").replace("'","").replace('"','').replace('.',' ')+" ")
				else:
					request_data["result"] = str(request_data["fulfillmentText"]).replace('*result','not available').replace("you can find *availables in stock"," sorry for inconvenience!")
				if len(outofstocks)>0:
					request_data["result"] = str(request_data["result"]).replace("*outofstocks",str(outofstocks).replace("[","").replace("]","").replace("'","").replace('"','').replace('.',' ')+" ")
				else:
					request_data["result"] = str(request_data["result"]).replace(" *outofstocks currently unavailable","")
			else:
				request_data["result"] = "Hi! Please share your Phone number for personalized Assistance!"
				
		if request_data["unknown"] == "order":
			if session["signedin"]:
				detail = {"match":False,"availability":"","offer":""}
				for sheet in book.keys():
					for row in book[sheet]:
						headers = row.keys()
						if "quantity" in headers:
							if request_data["known"]["product"] == row["product"]:
								detail["match"] = True
								if int(float(row["quantity"]))>0:
									if int(float(row["quantity"]))>=int(request_data["known"]["quantity"]):
										detail["availability"] = "available"
									else:
										detail["availability"] = "available till "+str(int(float(row["quantity"])))+" quantities"
								else:
									detail["availability"] = "not available currently"
						if "offer" in headers:
							if request_data["known"]["product"] == row["product"]:
								detail["offer"] = row["offer"]+" due to "+row["description"]
				#print(detail)
				if detail["match"]:
					if "not available currently" in request_data["result"]:
						request_data["result"] = request_data["fulfillmentText"].replace(request_data["known"]["product"]+" is *result!","sorry!").replace("& *offer will be applied on "+request_data["known"]["product"]+"! would you like to place the order for sure?","")
					else:
						request_data["result"] = request_data["fulfillmentText"].replace("*availability",detail["availability"])
					if len(detail["offer"])>0:
						request_data["result"] = request_data["result"].replace("*offer",detail["offer"])
					else:
						request_data["result"] = request_data["result"].replace("*offer","no offer")
				else:
					request_data["result"] = "No such product found! please specify the exact product name"
			else:
				request_data["result"] = "Hi! Please share your Phone number for personalized Assistance!"
		"""
		else:
			request_data["result"] = request_data["fulfillmentText"]
		"""
		print (request_data["result"])
		return json.dumps({"fulfillmentText":request_data["result"]})
	else:
		return "<h1>Home</h1>"

if __name__ == '__main__':
	port = int(os.getenv('PORT', 5000))
	print("Starting app on port %d" % port)
	app.run(debug=True, port=port, host='0.0.0.0')
