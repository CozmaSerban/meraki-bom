# from typing_extensions import Required
from flask import *
import xlsxwriter
import math

app = Flask(__name__)
app.secret_key = "asd"

products = {
    "mr": {
        "indoor": [
            {
                "pn": "MR45-HW",
                "lic": "LIC-ENT-3YR",
                "gpl_product": 1200,
                "gpl_license": 123
            },
        ],
        "outdoor": [
            {
                "pn": "MR76-HW",
                "lic": "LIC-ENT-3YR",
                "gpl_product": 1200,
                "gpl_license": 123,
                "antenna": "MA-ANT-20",
                "number_antenna": 2,
                "gpl_antenna": 123
            },
        ]
    },
    "ms": {
        "24": [
            {
                "pn": "MS120-24P-HW",
                "lic": "LIC-MS120-24P-3YR",
                "gpl_product": 1200,
                "gpl_license": 123
            },
        ],
        "48": [
            {
                "pn": "MS120-48LP-HW",
                "lic": "LIC-MS120-48LP-3YR",
                "gpl_product": 1200,
                "gpl_license": 123
            },
            {
                "pn": "MS120-48FP-HW",
                "lic": "LIC-MS120-48FP-3YR",
                "gpl_product": 1200,
                "gpl_license": 123
            },
        ]
    },
    "mx": {
        "virtual": {
            "pn": "LIC-VMX-S-ENT-3Y",
            "gpl": 123
        },
        "4g": {
            "pn": "MX75-HW",
            "lic": "LIC-MX75-SEC-3Y",
            "gpl_product": 200,
            "gpl_license": 200
        },
        "redudant":
            [
                {
                    "pn": "MX64-HW",
                    "lic": "LIC-MX64-SEC-3YR",
                    "gpl_product": 1200,
                    "gpl_license": 123,
                    "tput": 250
                },
                {
                    "pn": "MX68-HW",
                    "lic": "LIC-MX68-SEC-3YR",
                    "gpl_product": 1200,
                    "gpl_license": 123,
                    "tput": 450
                },
                {
                    "pn": "MX75-HW",
                    "lic": "LIC-MX75-SEC-3Y",
                    "gpl_product": 1200,
                    "gpl_license": 123,
                    "tput": 1000
                }
        ]
    },
    "mv": {
        "outdoor": {
            "pn": "MV72-HW",
            "lic": "LIC-MV-3YR",
            "gpl_product": 1200,
            "wall": "MA-MNT-MV-10",
            "gpl_license": 1200,
            "gpl_wall": 1200,
        },
        "indoor": {
            "pn": "MV32-HW",
            "lic": "LIC-MV-3YR",
            "gpl_product": 1200,
            "wall": "MA-MNT-MV-30",
            "gpl_license": 1200,
            "gpl_wall": 1200,
        }
    },
    "mt": {
        "temp": {
            "pn": "MT10-HW",
            "lic": "LIC-MT-3Y",
            "gpl_product": 1200,
            "gpl_license": 1200,
        },
        "hum": {
            "pn": "MT12-HW",
            "lic": "LIC-MT-3Y",
            "gpl_product": 1200,
            "gpl_license": 1200,
        },
    }
}


@app.route("/", methods=['POST', 'GET'])
def index():
    if request.method == 'POST':
        user = request.form['username']
        password = request.form['password']
        if user == "admin" and password == "serban":
            user = request.form['username']
            session["name"] = request.form.get("username")
            print("DA")
            return "DA"
    else:
        return render_template("login.html")


@app.route("/dashboard", methods=['POST', 'GET'])
def dashboard():
    return render_template('dashboard.html')


@app.route("/generate_bom", methods=['POST'])
def generate_bom():
    data = request.form.to_dict()
    generate_bom(data)
    return {"status": 201}


@app.route("/bom", methods=['POST', 'GET'])
def bom():
    return send_file("hello.xlsx")


def generate_bom(data):
    workbook = xlsxwriter.Workbook('hello.xlsx')

    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Informatii Generale')
    worksheet.write('A2', 'PN')
    worksheet.write('B2', 'Quantity')
    worksheet.write('C2', 'GPL')
    worksheet.write('D2', 'Discount')

    if data["tput_br"] != "" and data["tput_hq"] != "":
        if int(data["tput_br"]) <= 250 and int(data["tput_hq"]) <= 250:
            worksheet.write('A4', products["mx"]["redudant"][0]["pn"])
            worksheet.write('B4', 1)
            worksheet.write('C4', products["mx"]["redudant"][0]["gpl_product"])
            worksheet.write('A5', products["mx"]["redudant"][0]["lic"])
            worksheet.write('B5', 1)
            worksheet.write('C5', products["mx"]["redudant"][0]["gpl_license"])
        elif int(data["tput_br"]) <= 450 and int(data["tput_hq"]) <= 450:
            worksheet.write('A4', products["mx"]["redudant"][1]["pn"])
            worksheet.write('B4', 1)
            worksheet.write('C4', products["mx"]["redudant"][1]["gpl_product"])
            worksheet.write('A5', products["mx"]["redudant"][1]["lic"])
            worksheet.write('B5', 1)
            worksheet.write('C5', products["mx"]["redudant"][1]["gpl_license"])
        elif int(data["tput_br"]) <= 1000 and int(data["tput_hq"]) <= 1000:
            worksheet.write('A4', products["mx"]["redudant"][2]["pn"])
            worksheet.write('B4', 1)
            worksheet.write('C4', products["mx"]["redudant"][2]["gpl_product"])
            worksheet.write('A5', products["mx"]["redudant"][2]["lic"])
            worksheet.write('B5', 1)
            worksheet.write('C5', products["mx"]["redudant"][2]["gpl_license"])
        else:
            worksheet.write('A4', products["mx"]["redudant"][1]["pn"])
            worksheet.write('B4', 1)
            worksheet.write('C4', products["mx"]["redudant"][1]["gpl_product"])
            worksheet.write('A5', products["mx"]["redudant"][1]["lic"])
            worksheet.write('B5', 1)
            worksheet.write('C5', products["mx"]["redudant"][1]["gpl_license"])
    

    if data["4g_locations"] != 0 and data["4g_locations"] != '':
        worksheet.write('A6', 'MX cu LTE')
        worksheet.write('A7', products["mx"]["4g"]["pn"])
        worksheet.write('B7', data["4g_locations"])
        worksheet.write('C7', products["mx"]["4g"]["gpl_product"])
        worksheet.write('A8', products["mx"]["4g"]["lic"])
        worksheet.write('B8', data["4g_locations"])
        worksheet.write('C8', products["mx"]["4g"]["gpl_license"])

    worksheet.write('A10', 'Informatii specifice HQ')
    if data["hq_ap"] != 0 and data["hq_ap"] != "":
        worksheet.write('A11', 'Access Points')
        worksheet.write('A11', products["mr"]["indoor"][0]["pn"])
        worksheet.write('B11', math.ceil(float(data["hq_ap"])/100))
        worksheet.write('C11', products["mr"]["indoor"][0]["gpl_product"])
        worksheet.write('A12', products["mr"]["indoor"][0]["lic"])
        worksheet.write('B12', math.ceil(float(data["hq_ap"])/100))
        worksheet.write('C12', products["mr"]["indoor"][0]["gpl_license"])

    worksheet.write('A13', 'Switches')
    worksheet.write('A14', 'MS210-48FP-HW')
    worksheet.write('A15', 'LIC-MS-ENT-3YR')

    if data["hq_azure"] == "Da":
        worksheet.write('A16', 'MX virtual pentru cloud')
        worksheet.write('A17', products["mx"]["virtual"]["pn"])
        worksheet.write('C17', products["mx"]["virtual"]["gpl"])
        worksheet.write('B17', "1")

    if data["hq_cam_int"] != "":
        if data["hq_cam_int"] != 0 :
            worksheet.write('A18', 'Camera supraveghere interior')
            worksheet.write('A19', products["mv"]["indoor"]["pn"])
            worksheet.write('B19', data["hq_cam_int"])
            worksheet.write('C19', products["mv"]["indoor"]["gpl_product"])
            worksheet.write('A20', products["mv"]["indoor"]["lic"])
            worksheet.write('B20', data["hq_cam_int"])
            worksheet.write('C20', products["mv"]["indoor"]["gpl_license"])
    if data["hq_cam_ext"] != "":
        if data["hq_cam_ext"] != 0:
            worksheet.write('A21', 'Camera supraveghere exterior')
            worksheet.write('A22', products["mv"]["outdoor"]["pn"])
            worksheet.write('B22', data["hq_cam_ext"])
            worksheet.write('C22', products["mv"]["outdoor"]["gpl_product"])
            worksheet.write('A23', products["mv"]["outdoor"]["lic"])
            worksheet.write('B23', data["hq_cam_ext"])
            worksheet.write('C23', products["mv"]["outdoor"]["gpl_license"])
            worksheet.write('A24', products["mv"]["outdoor"]["wall"])
            worksheet.write('B24', data["hq_cam_ext"])
            worksheet.write('C24', products["mv"]["outdoor"]["gpl_wall"])
    
    if data["hq_senz_temp"] != "":
        if data["hq_senz_temp"] !=0:
            worksheet.write('A25', 'Senzori Temp')
            worksheet.write('A26', products["mt"]["temp"]["pn"])
            worksheet.write('B26', data["hq_senz_temp"])
            worksheet.write('C26', products["mt"]["temp"]["gpl_product"])
            worksheet.write('A27', products["mt"]["temp"]["lic"])
            worksheet.write('B27', data["hq_senz_temp"])
            worksheet.write('C27', products["mt"]["temp"]["gpl_license"])
    if data["hq_senz_temp"] != "":
        if data["hq_senz_hum"] !=0:
            worksheet.write('A28', 'Senzori Umiditate')
            worksheet.write('A29', products["mt"]["hum"]["pn"])
            worksheet.write('B29', data["hq_senz_hum"])
            worksheet.write('C29', products["mt"]["hum"]["gpl_product"])
            worksheet.write('A30', products["mt"]["hum"]["lic"])
            worksheet.write('B30', data["hq_senz_hum"])
            worksheet.write('C30', products["mt"]["hum"]["gpl_license"])

    worksheet.write('A31', 'Informatii specifice Branch')
    if data["br_ap"] != "":
        if data["br_ap"] != 0:
            worksheet.write('A32', 'Access Points')
            worksheet.write('A33', products["mr"]["indoor"][0]["pn"])
            worksheet.write('B33', math.ceil(float(data["br_ap"])/100))
            worksheet.write('C33', products["mr"]["indoor"][0]["gpl_product"])
            worksheet.write('A34', products["mr"]["indoor"][0]["lic"])
            worksheet.write('B34', math.ceil(float(data["br_ap"])/100))
            worksheet.write('C34', products["mr"]["indoor"][0]["gpl_license"])

    worksheet.write('A35', 'Switches')
    worksheet.write('A36', 'MS210-48FP-HW')
    worksheet.write('A37', 'LIC-MS-ENT-3YR')

    if data["hq_azure"] == "Da":
        worksheet.write('A38', 'MX virtual pentru cloud')
        worksheet.write('A39', products["mx"]["virtual"]["pn"])
        worksheet.write('C39', products["mx"]["virtual"]["gpl"])
        worksheet.write('B39', "1")

    if data["br_cam_int"] != "":
        if data["br_cam_int"] != 0:
            worksheet.write('A40', 'Camera supraveghere interior')
            worksheet.write('A41', products["mv"]["indoor"]["pn"])
            worksheet.write('B41', data["br_cam_int"])
            worksheet.write('C41', products["mv"]["indoor"]["gpl_product"])
            worksheet.write('A42', products["mv"]["indoor"]["lic"])
            worksheet.write('B42', data["br_cam_int"])
            worksheet.write('C42', products["mv"]["indoor"]["gpl_license"])
    if data["br_cam_ext"] != "":
        if data["br_cam_ext"] != 0:
            worksheet.write('A43', 'Camera supraveghere exterior')
            worksheet.write('A44', products["mv"]["outdoor"]["pn"])
            worksheet.write('B44', data["br_cam_ext"])
            worksheet.write('C44', products["mv"]["outdoor"]["gpl_product"])
            worksheet.write('A45', products["mv"]["outdoor"]["lic"])
            worksheet.write('B45', data["br_cam_ext"])
            worksheet.write('C45', products["mv"]["outdoor"]["gpl_license"])
            worksheet.write('A46', products["mv"]["outdoor"]["wall"])
            worksheet.write('B46', data["br_cam_ext"])
            worksheet.write('C46', products["mv"]["outdoor"]["gpl_wall"])
    if data["br_senz_temp"] != '':
        if data["br_senz_temp"] !=0:
            worksheet.write('A47', 'Senzori Temp')
            worksheet.write('A48', products["mt"]["temp"]["pn"])
            worksheet.write('B48', data["br_senz_temp"])
            worksheet.write('C48', products["mt"]["temp"]["gpl_product"])
            worksheet.write('A49', products["mt"]["temp"]["lic"])
            worksheet.write('B49', data["br_senz_temp"])
            worksheet.write('C49', products["mt"]["temp"]["gpl_license"])
    if data["br_senz_hum"] != '':
        if data["br_senz_hum"] !=0:
            worksheet.write('A50', 'Senzori Umiditate')
            worksheet.write('A51', products["mt"]["hum"]["pn"])
            worksheet.write('B51', data["br_senz_hum"])
            worksheet.write('C51', products["mt"]["hum"]["gpl_product"])
            worksheet.write('A52', products["mt"]["hum"]["lic"])
            worksheet.write('B52', data["br_senz_hum"])
            worksheet.write('C52', products["mt"]["hum"]["gpl_license"])


    worksheet.write('A53', 'Informatii specifice depozit')
    if data["dep_ap_int"] != "":
        if data["dep_ap_int"] != 0:
            worksheet.write('A54', 'Access Points Interior')
            worksheet.write('A55', products["mr"]["indoor"][0]["pn"])
            worksheet.write('B55', math.ceil(float(data["dep_ap_int"])/100))
            worksheet.write('C55', products["mr"]["indoor"][0]["gpl_product"])
            worksheet.write('A56', products["mr"]["indoor"][0]["lic"])
            worksheet.write('B56', math.ceil(float(data["dep_ap_int"])/100))
            worksheet.write('C56', products["mr"]["indoor"][0]["gpl_license"])

    if data["dep_ap_ext"] != "":
        if data["dep_ap_ext"] != 0:
            worksheet.write('A57', 'Access Points Exterior')
            worksheet.write('A58', products["mr"]["outdoor"][0]["pn"])
            worksheet.write('B58', math.ceil(float(data["dep_ap_ext"])/100))
            worksheet.write('C58', products["mr"]["outdoor"][0]["gpl_product"])
            worksheet.write('A59', products["mr"]["outdoor"][0]["lic"])
            worksheet.write('B59', math.ceil(float(data["dep_ap_ext"])/100))
            worksheet.write('C59', products["mr"]["outdoor"][0]["gpl_license"])
            worksheet.write('A60', products["mr"]["outdoor"][0]["antenna"])
            worksheet.write('B60', products["mr"]["outdoor"][0]["number_antenna"])
            worksheet.write('C60', products["mr"]["outdoor"][0]["gpl_antenna"])

    worksheet.write('A61', 'Switches')
    worksheet.write('A62', 'MS210-48FP-HW')
    worksheet.write('A63', 'LIC-MS-ENT-3YR')

    if data["dep_cam_int"] != 0 and data["dep_cam_int"] != "":
        worksheet.write('A64', 'Camera supraveghere interior')
        worksheet.write('A65', products["mv"]["indoor"]["pn"])
        worksheet.write('B65', data["dep_cam_int"])
        worksheet.write('C65', products["mv"]["indoor"]["gpl_product"])
        worksheet.write('A66', products["mv"]["indoor"]["lic"])
        worksheet.write('B66', data["dep_cam_int"])
        worksheet.write('C66', products["mv"]["indoor"]["gpl_license"])

    if data["dep_cam_ext"] != 0 and data["dep_cam_ext"] != "":
        worksheet.write('A67', 'Camera supraveghere exterior')
        worksheet.write('A68', products["mv"]["outdoor"]["pn"])
        worksheet.write('B68', data["dep_cam_ext"])
        worksheet.write('C68', products["mv"]["outdoor"]["gpl_product"])
        worksheet.write('A69', products["mv"]["outdoor"]["lic"])
        worksheet.write('B69', data["dep_cam_ext"])
        worksheet.write('C69', products["mv"]["outdoor"]["gpl_license"])
        worksheet.write('A70', products["mv"]["outdoor"]["wall"])
        worksheet.write('B70', data["dep_cam_ext"])
        worksheet.write('C70', products["mv"]["outdoor"]["gpl_wall"])

    if data["dep_senz_temp"] !=0 and data["dep_senz_temp"] != '':
        worksheet.write('A71', 'Senzori Temp')
        worksheet.write('A72', products["mt"]["temp"]["pn"])
        worksheet.write('B72', data["dep_senz_temp"])
        worksheet.write('C72', products["mt"]["temp"]["gpl_product"])
        worksheet.write('A73', products["mt"]["temp"]["lic"])
        worksheet.write('B73', data["dep_senz_temp"])
        worksheet.write('C73', products["mt"]["temp"]["gpl_license"])
    if data["dep_senz_hum"] !=0 and data["dep_senz_hum"] != '':
        worksheet.write('A74', 'Senzori Umiditate')
        worksheet.write('A75', products["mt"]["hum"]["pn"])
        worksheet.write('B75', data["dep_senz_hum"])
        worksheet.write('C75', products["mt"]["hum"]["gpl_product"])
        worksheet.write('A76', products["mt"]["hum"]["lic"])
        worksheet.write('B76', data["dep_senz_hum"])
        worksheet.write('C76', products["mt"]["hum"]["gpl_license"])

    worksheet.write('A80', '20')
    worksheet.write('A81', '30')

    worksheet.write('A90', '=SUM(A80 + A81)')

    workbook.close()
    return None


if __name__ == "__main__":
    app.run()
