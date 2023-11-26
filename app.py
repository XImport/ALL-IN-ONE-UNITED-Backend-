# import pandas as pd
# import simplejson as json

# df = pd.read_excel('10-23.xlsx',sheet_name="API")

# json_data = json.loads(df.to_json(orient='records'))

# print(json_data)


from api import app


if __name__ == "__main__":
    app.run(debug=True)
