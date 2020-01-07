
from flask import Flask, render_template, request, jsonify
import pyodbc
import pandas as pd

cnxn = pyodbc.connect(
    Trusted_Connection='Yes',
    Driver='{ODBC Driver 17 for SQL Server}',
    Server='TOCIONDWSQL01P',
    Database='ION_ADW'
)

df_AE = pd.read_sql_query('select DISTINCT Full_Name from ION_ADW.dbo.Dim_Account_Exec', cnxn) 
df_Advertiser = pd.read_sql_query('select DISTINCT Advertiser_Name from ION_ADW.dbo.Dim_Advertiser', cnxn) 
df_Agency = pd.read_sql_query('select DISTINCT Agency_Name from ION_ADW.dbo.Dim_Agency', cnxn) 


items = {'AE': df_AE['Full_Name'].tolist(),
         'Advertiser': df_Advertiser['Advertiser_Name'].tolist(),
         'Agency': df_Agency['Agency_Name'].tolist()}


app = Flask(__name__)

@app.route("/")
def index():
    return render_template('index.html', items= items)


## get the filled form from web
get_back = {}
@app.route("/info", methods=['GET', 'POST'])
def get_info():
      get_back['AE'] = request.form['ae_name']
      get_back['Advertiser'] = request.form['ad_name']
      get_back['Agency'] = request.form['ag_name']
      get_back['Start_Date'] = request.form['start_dt']
      get_back['End_Date'] = request.form['end_dt']
      get_back['Dayparts'] = request.form.getlist('check')
      get_back['Days'] = request.form.getlist('days')
      get_back['Start_Time'] = request.form.getlist('start_time')
      get_back['End_Time'] = request.form.getlist('end_time')
      get_back['Bidding_Amount'] = request.form['bid_amount']
      
      
     # return get_back
      return jsonify(get_back)



if __name__ == "__main__":
    app.run(debug= True, use_reloader= False)


## convert to DataFrame    
web_output = pd.DataFrame([get_back], columns= get_back.keys())

## save as xlsx file    
web_output.to_excel(r'M:\Work_Rishi\Pricing_Tool\web_output.xlsx', index= False)
    
    
    
    


