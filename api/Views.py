from api import app
from flask import request, jsonify
from pathlib import Path
import pandas as pd
import json
from datetime import datetime
import os
import locale
from api.Functions import OrderExcelFilesByNames

FrenchMonthNames = {
    "Janvier": "01",
    "Février": "02",
    "Mars": "03",
    "Avril": "04",
    "Mai": "05",
    "Juin": "06",
    "Juillet": "07",
    "Août": "08",
    "Septembre": "09",
    "Octobre": "10",
    "Novembre": "11",
    "Décembre": "12",
}


@app.route("/", methods=["POST", "GET"])
def Testing():
    if request.method == "POST":
        return "Post request that you sent"
    return "GET request that you sent"


@app.route("/api/db/data", methods=["POST", "GET"])
def CommercialData():
    if request.method == "POST":
        data = request.get_json()
        YearQuery = data.get("Year")
        MonthQuery = data.get("Month")

        ColumnsValues = {
            "clientNameState": True,
            "categorieState": True,
            "qntEnTState": False,
            "caBrutState": True,
            "caNetState": False,
            "caNetFactState": False,
            "caNetHFactState": False,
            "CreanceGlobalState": True,
            "creanceFacturerState": True,
            "creanceHfactuerState": True,
            "ModalitePaimentState": False,
            "UniteState": False,
            "cautionState": True,
            "plafondState": True,
            "soldeRestState": False,
            "voyRestState": False,
            "transportState": True,
        }
        ConvertedMonthNumber = FrenchMonthNames.get(MonthQuery, "")
        print(YearQuery, MonthQuery)
        if not ConvertedMonthNumber:
            return jsonify({"error": "Opps!! Le Nom du Mois Non validé"})

        YearConvertedShortCut = YearQuery[-2:]

        SearchQuery = f"{ConvertedMonthNumber}-{YearConvertedShortCut}"
        current_base_path = Path.cwd()
        file_path = current_base_path / "api" / "database" / (SearchQuery +
                                                              ".xlsx")

        if file_path.exists():
            df = pd.read_excel(file_path, sheet_name="API")
            DATA = df.to_json(orient="records", lines=False)
            ConvertedToJson = json.loads(DATA)
            for Values in ConvertedToJson:
                Values.update(ColumnsValues)
            return jsonify(ConvertedToJson)
        else:
            return jsonify({
                "error":
                "Opps !! Les données que vous avez entrées sont introuvables."
            })


@app.route("/api/db/espece", methods=["POST", "GET"])
def EspeceData():
    if request.method == "POST":
        data = request.get_json()
        YearQuery = data.get("Year")
        MonthQuery = data.get("Month")

        ConvertedMonthNumber = FrenchMonthNames.get(MonthQuery, "")
        if not ConvertedMonthNumber:
            return jsonify({"error": "Opps!! Le Nom du Mois Non validé"})

        YearConvertedShortCut = YearQuery[-2:]
        SearchQuery = f"{ConvertedMonthNumber}-{YearConvertedShortCut}"
        current_base_path = Path.cwd()
        file_path = current_base_path / "api" / "database" / (SearchQuery +
                                                              ".xlsx")

        if file_path.exists():
            df = pd.read_excel(file_path, sheet_name="ESP")
            DATA = df.to_json(orient="records", lines=False)
            ConvertedToJson = json.loads(DATA)
            NewDateFormatter = []
            Cash = []
            Numberofclients = []
            for Date in ConvertedToJson:
                # formatted_date = datetime.utcfromtimestamp(
                #     Date["Date"] / 1000
                # ).strftime("%d-%m")
                # NewDateFormatter.append(formatted_date)
                NewDateFormatter.append(Date["Date"])
            for cash in ConvertedToJson:
                Cash.append(cash["CA CAISSE"])

            for client in ConvertedToJson:
                Numberofclients.append(client["Nombre de Clients"])

            return jsonify(date=NewDateFormatter,
                           CashCaisse=Cash,
                           Numberofclients=Numberofclients)
        else:

            return jsonify({
                "error":
                "Opps !! Les données que vous avez entrées sont introuvables."
            })


@app.route("/api/db/transport", methods=["POST", "GET"])
def TransportData():
    if request.method == "POST":
        data = request.get_json()
        YearQuery = data.get("Year")
        MonthQuery = data.get("Month")

        ConvertedMonthNumber = FrenchMonthNames.get(MonthQuery, "")

        if not ConvertedMonthNumber:
            return jsonify({"error": "Opps!! Le Nom du Mois Non validé"})

        YearConvertedShortCut = YearQuery[-2:]
        SearchQuery = f"{ConvertedMonthNumber}-{YearConvertedShortCut}"
        current_base_path = Path.cwd()
        file_path = current_base_path / "api" / "database" / (SearchQuery +
                                                              ".xlsx")

        if file_path.exists():
            df = pd.read_excel(file_path, sheet_name="API")
            Transport_df = df[df["transport"] > 1]
            Transport_df = Transport_df.reset_index(drop=True)

            TransprtName = {
                "TransportNames": Transport_df["clientName"].values.tolist()
            }
            TransportCash = {
                "TransportCash": Transport_df["transport"].values.tolist()
            }
            return jsonify(TransprtName, TransportCash)
        else:

            return jsonify({
                "error":
                "Opps !! Les données que vous avez entrées sont introuvables."
            })

@app.route("/api/db/Plafond", methods=["POST", "GET"])
def PlafondData():
    if request.method == "POST":
        data = request.get_json()
        YearQuery = data.get("Year")
        MonthQuery = data.get("Month")

        ConvertedMonthNumber = FrenchMonthNames.get(MonthQuery, "")

        if not ConvertedMonthNumber:
            return jsonify({"error": "Opps!! Le Nom du Mois Non validé"})

        YearConvertedShortCut = YearQuery[-2:]
        SearchQuery = f"{ConvertedMonthNumber}-{YearConvertedShortCut}"
        current_base_path = Path.cwd()
        file_path = current_base_path / "api" / "database" / (SearchQuery +
                                                              ".xlsx")

        if file_path.exists():
            df = pd.read_excel(file_path, sheet_name="Plafond")
            ColumnsValues = {
                "clientNameState": True,
                "categorieState": True,
                "caBrutState": True,
                "plafondState": True,
                "soldeRestState": False,
                "voyRestState": True,
            }
            DATA = df.to_json(orient="records", lines=False)
            ConvertedToJson = json.loads(DATA)
            for Values in ConvertedToJson:
                Values.update(ColumnsValues)

            return ConvertedToJson

        else:
            return jsonify({
                "error":
                "Opps !! Les données que vous avez entrées sont introuvables."
            })

@app.route("/api/db/recap", methods=["POST", "GET"])
def recap_data():
    if request.method == "GET":
        current_base_path = Path.cwd()
        file_path = current_base_path / "api" / "database"
        sorted_files = OrderExcelFilesByNames(file_path)

        if not sorted_files:
            return jsonify(
                {"error": "No Excel files found in the specified directory"})

        latest_file = sorted_files[-1]
        file_path = current_base_path / "api" / "database" / latest_file

        df_rccomm = pd.read_excel(file_path, sheet_name="RCCOMM")
        df_comm = pd.read_excel(file_path, sheet_name="COMM")
        df_CA = pd.read_excel(file_path, sheet_name="API")
        df_DATE = pd.read_excel(file_path, sheet_name="BASE")
        DATETIME = df_DATE.iloc[
            7:, df_DATE.columns.get_loc("Unnamed: 0")].to_list()[-1]
        date_string = str(DATETIME).split(" ")[0]
        # print(str(DATETIME).split(" ")[0])
        locale.setlocale(locale.LC_TIME, "fr_FR.utf8")
        date_object = datetime.strptime(date_string, "%Y-%m-%d")
        formatted_date = date_object.strftime("%A %d %B %Y")

        selected_columns = ["clientName", "qntEnT", "caBrut"]
        print(formatted_date)
        output = df_CA[selected_columns].to_dict(orient="records")
        OutputArray = []
        for CA_scanner in output:
            if CA_scanner["caBrut"] > 0:
                OutputArray.append(CA_scanner)
            else:
                pass

        # Extracting values for Card1
        columns_values = {
            "CABRUTE": 0,
            "TRANSPORT": 0,
            "VOLUMELIVRE": 0,
            "CAESPECE": 0,
        }

        if not df_rccomm.empty:
            first_row = df_rccomm.iloc[0]
            columns_values["CABRUTE"] = first_row.get("CABRUTE", 0)
            columns_values["TRANSPORT"] = first_row.get("CA TRANSPORT", 0)
            columns_values["VOLUMELIVRE"] = first_row.get("Volume Vendus", 0)
            columns_values["CAESPECE"] = first_row.get("TOTAL ESPECE", 0)

        # Extracting values for Card2
        columns_values2 = {
            "CREANCEGLOBAL": 0,
            "COMMANDESLIVRE": 0,
            "PMV": 0,
        }

        if not df_rccomm.empty:
            first_row = df_rccomm.iloc[0]
            columns_values2["CREANCEGLOBAL"] = first_row.get(
                "CREANCE A CREDIT", 0)+first_row.get("CREANCE NOCIVE", 0)
            columns_values2["COMMANDESLIVRE"] = first_row.get(
                "COMMANDES RENDU LIVRE", 0)

        # Avoid division by zero
        if columns_values["VOLUMELIVRE"] != 0:
            columns_values2["PMV"] = (first_row.get("CA NET", 0) /
                                      columns_values["VOLUMELIVRE"])

        # Extracting data for the graph
        date = pd.to_datetime(df_comm["DATE"]).dt.strftime("%d-%m").tolist()
        # date = df_comm["DATE"].tolist()
        n_voyages_comm = df_comm["Nombre des Voyages Commandé"].tolist()
        n_voyages_liv = df_comm["Nombre des Voyages Livré"].tolist()

        result = {
            "Card1": columns_values,
            "Card2": columns_values2,
            "datee": date,
            "NVoyagesComm": n_voyages_comm,
            "NVoyagesLiv": n_voyages_liv,
            "items": OutputArray,
            "Date": formatted_date
        }

        return jsonify(result)
