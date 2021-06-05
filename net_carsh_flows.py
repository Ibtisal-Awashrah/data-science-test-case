import pandas as pd
import numpy as np
from IPython.display import display
import numpy_financial as fnp
from flask import Flask, request


# Flask and HTML for Handle GUI
app = Flask(__name__)


@app.route("/upload", methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        df = read_file(f)
        list_of_funds = generate_fund_names(df)

        list_of_count_sum = calculate_contribution(df)
        list_of_dist_sum = calculate_distribution(df)

        list_of_last_valuation = generate_last_valuation(df, list_of_funds)
        list_of_tvpi = calculate_tvpi(
            list_of_funds, list_of_dist_sum, list_of_last_valuation, list_of_count_sum)
        list_of_irr = calculate_irr(df, list_of_funds)

        first_sheet = generate_first_sheet(list_of_funds, list_of_count_sum, list_of_dist_sum,
                                           list_of_last_valuation, list_of_tvpi, list_of_irr)

        second_sheet = generate_second_sheet(df)

        write_file(first_sheet, second_sheet)

        data_xls = pd.read_excel("output.xlsx")

        return data_xls.to_html()
    return '''
    <!doctype html>
    <title>Upload an excel file</title>
    <h1>Excel file upload xlsx</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Upload>
    </form>
    '''


@app.route("/second")
def read_second_sheet():
    data_xls = pd.read_excel("output.xlsx", "Second Task")

    return data_xls.to_html()


def read_file(file_name):
    df = pd.read_excel(file_name, 'Funds Net CFs')
    return df


def write_file(first_sheet, second_sheet):
    writer = pd.ExcelWriter('output.xlsx', engine="xlsxwriter",
                            options={'strings_to_urls': False})
    first_sheet.to_excel(writer, "First Task")
    second_sheet.to_excel(writer, "Second Task")
    writer.save()


def generate_fund_names(df):
    list_of_funds = df["Fund Name"].drop_duplicates(
    ).sort_values().values.tolist()

    return list_of_funds


def calculate_contribution(df):
    cont_list = df[df["Cashflow Type"] == "Contribution"]
    list_of_count_sum = cont_list.groupby(cont_list["Fund Name"])[
        "Amount"].sum().values
    return list_of_count_sum


def calculate_distribution(df):
    dist_list = df[df["Cashflow Type"] == "Distribution"]
    list_of_dist_sum = dist_list.groupby(dist_list["Fund Name"])[
        "Amount"].sum().values
    return list_of_dist_sum


def generate_last_valuation(df, list_of_funds):
    val_list = df[df["Cashflow Type"] == "Valuation"]
    list_of_last_valuation = []
    for fund in list_of_funds:
        # Take Last Valuation for each fund
        list_of_last_valuation.append(
            val_list[val_list["Fund Name"] == fund].iloc[-1]["Amount"])
    return list_of_last_valuation


def calculate_irr(df, list_of_funds):
    list_of_irr = []
    for fund in list_of_funds:
        # Calculate irr
        fun_list = df[df["Fund Name"] == fund]
        list_of_irr.append(fnp.irr(fun_list["Amount"]))
    return list_of_irr


def calculate_tvpi(list_of_funds, list_of_dist_sum, list_of_last_valuation, list_of_count_sum):
    list_of_tvpi = []
    counter = 0
    for fund in list_of_funds:
        # Calculate tvpi
        list_of_tvpi.append(abs(
            (list_of_dist_sum[counter]+list_of_last_valuation[counter])/list_of_count_sum[counter]))
        counter += 1

    return list_of_tvpi


def generate_first_sheet(list_of_funds, list_of_count_sum, list_of_dist_sum, list_of_last_valuation, list_of_tvpi, list_of_irr):
    # Creat output
    dict = {"Fund Name": list_of_funds,
            "Total Contribution": list_of_count_sum,
            "Total Distribution": list_of_dist_sum,
            "Valuation": list_of_last_valuation,
            "TVPI": list_of_tvpi,
            "IRR": list_of_irr}
    dfe = pd.DataFrame(dict)
    display(dfe)
    return dfe

# Task 2


def generate_second_sheet(df):
    df_grouped_by_Date = df.groupby([df["Date"].dt.year, df["Date"].dt.month])
    amount_list = df_grouped_by_Date.sum()["Amount"].values

    date_dict = pd.DataFrame(df.groupby(
        [df["Date"].dt.year, df["Date"].dt.month]))[0].values

    task2_dict = {"Date": date_dict,
                  "Amount": amount_list
                  }
    task2_output = pd.DataFrame(task2_dict)
    display(task2_output)
    return task2_output


if __name__ == "__main__":
    app.run()
