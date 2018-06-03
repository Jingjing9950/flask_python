from flask import Flask,render_template,request,send_file,Response,send_file
import uuid
from werkzeug.utils import secure_filename
import os
import pandas as pd
import xlsxwriter

app = Flask(__name__)

@app.route("/upload")
def upload_file():
    return render_template("upload.html")

@app.route("/return-file/")
def return_file():
    return send_file("Tax_provision.xlsx")

@app.route("/uploader", methods=['GET','POST'])
def uploader_file():
    if request.method == 'POST':
        f = request.files['file']
        process(f)
        #filename = secure_filename(f.filename)
        # return Response("Hello jingjing",mimetype="text/xls",
        #                headers={"Content-Disposition":
        #                             "attachment;filename=test.xlsx"})
        return render_template('downloat.html')

def process(file):
    data_frame = pd.read_excel(file)
    data_frame = data_frame.fillna(value="missing")
    #file_name = str(uuid.uuid1()) +  ".xlsx"
    write_excel = pd.ExcelWriter("Tax_provision.xlsx")

    add_back_items = {}

    def auto_sum_expense(accountdata_filter, column_name, account_name, key_name):
        accountdata_filter = data_frame[data_frame[column_name] == account_name]
        add_back_items[key_name] = accountdata_filter["Amount in local currency"].sum()

    def auto_sum_expense_contain(accountdata_filter, column_name1, account_name, filtered_data, column_name2,
                                 string_contain, key_name, ):
        accountdata_filter = data_frame[data_frame[column_name1] == account_name]
        filtered_data = accountdata_filter[accountdata_filter[column_name2].str.contains("|".join(string_contain))]
        add_back_items[key_name] = filtered_data["Amount in local currency"].sum()

    def breakdown_list_1(accountdata_filter, column_name, account_name, key_name):
        accountdata_filter = data_frame[data_frame[column_name] == account_name]
        return accountdata_filter

    def breakdown_list_2(accountdata_filter, column_name1, account_name, filtered_data, column_name2, string_contain,
                         key_name, ):
        accountdata_filter = data_frame[data_frame[column_name1] == account_name]
        filtered_data = accountdata_filter[accountdata_filter[column_name2].str.contains("|".join(string_contain))]
        return filtered_data

    def write_into_excel_sheet(summary_breakdown_list, work_sheet):
        summary_breakdown_list.to_excel(write_excel, work_sheet)

    # Vehicel expense_add back
    auto_sum_expense("vehicle_exp", "Account Text", "Vehicle exp", "vehical_expenses")

    # Vehicel road tax_add back
    auto_sum_expense_contain("tax_dues", "Account Text", "Taxes and dues", "road_tax", "Text", ["Road"],
                             "Vehicle_road_tax")
    write_into_excel_sheet(breakdown_list_2("tax_dues", "Account Text", "Taxes and dues", "road_tax", "Text", ["Road"],
                                            "Vehicle_road_tax"), "Vehicle_road_tax")

    # Vehicel insurance_add back
    auto_sum_expense("vehicle_insu", "Account Text", "Vehicle insurance", "Vehicle_insurance")

    # Rental_expense_vehicle
    auto_sum_expense("rental_vehicle", "Account Text", "Rental exp vehicle", "Rental_expenses_vehicle")
    write_into_excel_sheet(
        breakdown_list_1("rental_vehicle", "Account Text", "Rental exp vehicle", "Rental_expenses_vehicle"),
        "rental_expe_vehicle")

    # Property tax & employment pass fee
    auto_sum_expense_contain("tax_dues", "Account Text", "Taxes and dues", "property_tax", "Text", ["Property"],
                             "Property_tax")
    auto_sum_expense_contain("tax_dues", "Account Text", "Taxes and dues", "employ_pass", "Text",
                             ["pass", "visa", "permit"], "Employee_pass_fee")
    write_into_excel_sheet(
        breakdown_list_2("tax_dues", "Account Text", "Taxes and dues", "property_tax", "Text", ["Property"],
                         "Property_tax"), "property_tax")
    write_into_excel_sheet(breakdown_list_2("tax_dues", "Account Text", "Taxes and dues", "employ_pass", "Text",
                                            ["pass", "visa", "permit"], "Employee_pass_fee"), "employee_pass_fee")

    # professional_fee
    auto_sum_expense("profession_fee", "Account Text", "Professional fees", "Professional_services")
    write_into_excel_sheet(
        breakdown_list_1("profession_fee", "Account Text", "Professional fees", "Professional_services"),
        "professional_service")

    # Welfare expensed _add back
    welfare_expense = data_frame[data_frame["Account Text"] == "Welfare exp"]
    search_medical = ["Medical", "medical"]
    medical_expen = welfare_expense[welfare_expense["Assignment"].str.contains("|".join(search_medical))]

    search_medical_fee = ["Medical", "medical", "Dental", "dental", "vac", "Vac"]
    medical_fee = medical_expen[medical_expen["Text"].str.contains("|".join(search_medical_fee))]
    add_back_items["Medical_fee"] = medical_fee["Amount in local currency"].sum()
    write_into_excel_sheet(medical_fee, "medical_fee")

    search_medical_insu = ["Hospi", "hospi"]
    medical_insurance = medical_expen[medical_expen["Text"].str.contains("|".join(search_medical_insu))]
    add_back_items["Medical_insurance"] = medical_insurance["Amount in local currency"].sum()
    write_into_excel_sheet(medical_insurance, "medical_insurance")

    # Misc_income
    auto_sum_expense("misllan_income", "Account Text", "Misc income", "Misllanence_income")
    write_into_excel_sheet(breakdown_list_1("misllan_income", "Account Text", "Misc income", "Misllanence_income"),
                           "misc_income")

    # convert dictionary to Series in pandas
    df_result = pd.Series(add_back_items, name="Amount")

    header_name = "THK KM SYSTEM PTE LTD"
    header = "TAX SCHEDULES - YA 2019"

    # header_name.to_excel(write_excel, "provision_summary")
    df_result.to_excel(write_excel, sheet_name="provision_summary", startrow=6)

    # write_into_excel_sheet(df_result, "provision_summary")
    write_excel.save()
    write_excel.close()

if __name__=="__main__":
    app.run()
