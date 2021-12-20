# from typing_extensions import Required
from flask import *
import xlsxwriter
import math
import requests, xmltodict

app = Flask(__name__)
app.secret_key = "asd"

products = {
    "mr": {
        "indoor": [
            {
                "pn": "MR45-HW",
                "lic": "LIC-ENT-3YR",
                "gpl_product": 1200,
                "gpl_license": 321
            },
        ],
        "outdoor": [
            {
                "pn": "MR76-HW",
                "lic": "LIC-ENT-3YR",
                "gpl_product": 1887.72,
                "gpl_license": 321,
                "antenna": "MA-ANT-20",
                "number_antenna": 2,
                "gpl_antenna": 213.44
            },
        ]
    },
    "ms": {
        "24": [
            {
                "pn": "MS120-24P-HW",
                "lic": "LIC-MS120-24P-3YR",
                "gpl_product": 2831.58,
                "gpl_license": 304.95
            },
        ],
        "48": [
            {
                "pn": "MS120-48LP-HW",
                "lic": "LIC-MS120-48LP-3YR",
                "gpl_product": 4419.98,
                "gpl_license": 476.15
            },
            {
                "pn": "MS120-48FP-HW",
                "lic": "LIC-MS120-48FP-3YR",
                "gpl_product": 5218.04,
                "gpl_license": 567.1
            },
        ]
    },
    "mx": {
        "virtual": {
            "pn": "LIC-VMX-S-ENT-3Y",
            "gpl": 1000
        },
        "4g": {
            "pn": "MX75-HW",
            "lic": "LIC-MX75-SEC-3Y",
            "gpl_product": 2139.77,
            "gpl_license": 4000
        },
        "redudant":
            [
                {
                    "pn": "MX64-HW",
                    "lic": "LIC-MX64-SEC-3YR",
                    "gpl_product": 638.18,
                    "gpl_license": 1200,
                    "tput": 250
                },
                {
                    "pn": "MX68-HW",
                    "lic": "LIC-MX68-SEC-3YR",
                    "gpl_product": 1067.21,
                    "gpl_license": 1500,
                    "tput": 450
                },
                {
                    "pn": "MX75-HW",
                    "lic": "LIC-MX75-SEC-3Y",
                    "gpl_product": 2139.77,
                    "gpl_license": 4000,
                    "tput": 1000
                }
        ]
    },
    "mv": {
        "outdoor": {
            "pn": "MV72-HW",
            "lic": "LIC-MV-3YR",
            "gpl_product": 1703.1,
            "wall": "MA-MNT-MV-10",
            "gpl_license": 600,
            "gpl_wall": 267.1,
        },
        "indoor": {
            "pn": "MV32-HW",
            "lic": "LIC-MV-3YR",
            "gpl_product": 1502.6,
            "wall": "MA-MNT-MV-30",
            "gpl_license": 600,
            "gpl_wall": 267.1,
        }
    },
    "mt": {
        "temp": {
            "pn": "MT10-HW",
            "lic": "LIC-MT-3Y",
            "gpl_product": 249.6,
            "gpl_license": 300,
        },
        "hum": {
            "pn": "MT12-HW",
            "lic": "LIC-MT-3Y",
            "gpl_product": 399.96,
            "gpl_license": 300,
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
    worksheet.write('E2', 'Price')

    if data["tput_br"] != "" and data["tput_hq"] != "":
        tput = int(data["br_locations"]) + int(data["hq_locations"])
        if int(data["tput_br"]) <= 250 and int(data["tput_hq"]) <= 250:
            worksheet.write('A4', products["mx"]["redudant"][0]["pn"])
            worksheet.write('B4', 1*(tput - int(data["4g_locations"])))
            worksheet.write('C4', products["mx"]["redudant"][0]["gpl_product"])
            worksheet.write('D4', 65)
            worksheet.write('E4', '=B4*C4*D4/100')

            worksheet.write('A5', products["mx"]["redudant"][0]["lic"])
            worksheet.write('B5', 1*(tput - int(data["4g_locations"])))
            worksheet.write('C5', products["mx"]["redudant"][0]["gpl_license"])
            worksheet.write('D5', 65)
            worksheet.write('E5', '=B5*C5*D5/100')
            
        elif int(data["tput_br"]) <= 450 and int(data["tput_hq"]) <= 450:
            worksheet.write('A4', products["mx"]["redudant"][1]["pn"])
            worksheet.write('B4', 1*(tput - int(data["4g_locations"])))
            worksheet.write('C4', products["mx"]["redudant"][1]["gpl_product"])
            worksheet.write('D4', 65)
            worksheet.write('E4', '=B4*C4*D4/100')
            worksheet.write('A5', products["mx"]["redudant"][1]["lic"])
            worksheet.write('B5', 1*(tput - int(data["4g_locations"])))
            worksheet.write('C5', products["mx"]["redudant"][1]["gpl_license"])
            worksheet.write('D5', 65)
            worksheet.write('E5', '=B5*C5*D5/100')
            
        elif int(data["tput_br"]) <= 1000 and int(data["tput_hq"]) <= 1000:
            worksheet.write('A4', products["mx"]["redudant"][2]["pn"])
            worksheet.write('B4', 1*(tput - int(data["4g_locations"])))
            worksheet.write('C4', products["mx"]["redudant"][2]["gpl_product"])
            worksheet.write('D4', 65)
            worksheet.write('E4', '=B4*C4*D4/100')
        
            worksheet.write('A5', products["mx"]["redudant"][2]["lic"])
            worksheet.write('B5', 1*(tput - int(data["4g_locations"])))
            worksheet.write('C5', products["mx"]["redudant"][2]["gpl_license"])
            worksheet.write('D5', 65)
            worksheet.write('E5', '=B5*C5*D5/100')
            
        else:
            worksheet.write('A4', products["mx"]["redudant"][1]["pn"])
            worksheet.write('B4', 1*(tput - int(data["4g_locations"])))
            worksheet.write('C4', products["mx"]["redudant"][1]["gpl_product"])
            worksheet.write('D4', 65)
            worksheet.write('E4', '=B4*C4*D4/100')
            
            worksheet.write('A5', products["mx"]["redudant"][1]["lic"])
            worksheet.write('B5', 1*(tput - int(data["4g_locations"])))
            worksheet.write('C5', products["mx"]["redudant"][1]["gpl_license"])
            worksheet.write('D5', 65)
            worksheet.write('E5', '=B5*C5*D5/100')
            
    

    if data["4g_locations"] != 0 and data["4g_locations"] != '':
        worksheet.write('A6', 'MX cu LTE')
        worksheet.write('A7', products["mx"]["4g"]["pn"])
        worksheet.write('B7', int(data["4g_locations"]))
        worksheet.write('C7', products["mx"]["4g"]["gpl_product"])
        worksheet.write('D7', 65)
        worksheet.write('E7', '=B7*C7*D7/100')
            
        worksheet.write('A8', products["mx"]["4g"]["lic"])
        worksheet.write('B8', int(data["4g_locations"]))
        worksheet.write('C8', products["mx"]["4g"]["gpl_license"])
        worksheet.write('D8', 65)
        worksheet.write('E8', '=B8*C8*D8/100')
            

    worksheet.write('A10', 'Informatii specifice HQ')
    if data["hq_ap"] != 0 and data["hq_ap"] != "":
        worksheet.write('A11', 'Access Points')
        worksheet.write('A11', products["mr"]["indoor"][0]["pn"])
        worksheet.write('B11', math.ceil(float(data["hq_ap"])/100)*int(data["hq_locations"]))
        worksheet.write('C11', products["mr"]["indoor"][0]["gpl_product"])
        worksheet.write('D11', 65)
        worksheet.write('E11', '=B11*C11*D11/100')
            
        worksheet.write('A12', products["mr"]["indoor"][0]["lic"])
        worksheet.write('B12', math.ceil(float(data["hq_ap"])/100)*int(data["hq_locations"]))
        worksheet.write('C12', products["mr"]["indoor"][0]["gpl_license"])
        worksheet.write('D12', 65)
        worksheet.write('E12', '=B12*C12*D12/100')
            

    
    if data['hq_ports'] != '' and data['hq_ports'] !=0:
        worksheet.write('A13', 'Switches')
        ports = int(int(data["hq_ports"]) * 1.2)
        if ports / 48 == 0 and ports / 24 == 0:
             #de 24
            worksheet.write('A14', products['ms']['24']['pn'])
            worksheet.write('B14', 1*int(data["hq_locations"]))
            worksheet.write('C14', products['ms']['24']["gpl_product"])
            worksheet.write('D14', 65)
            worksheet.write('E14', '=B14*C14*D14/100')
            
            worksheet.write('A15', products['ms']['24']['pn'])
            worksheet.write('B15', 1*int(data["hq_locations"]))
            worksheet.write('C15', products['ms']['24']["gpl_product"])
            worksheet.write('D15', 65)
            worksheet.write('E15', '=B15*C15*D15/100')
            
        if ports / 48 == 0 and ports / 24 != 0:
            
             #de 48 1
            worksheet.write('A14', products['ms']['48'][0]['pn'])
            worksheet.write('B14', 1*int(data["hq_locations"]))
            worksheet.write('C14', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D14', 65)
            worksheet.write('E14', '=B14*C14*D14/100')
            
            worksheet.write('A15', products['ms']['48'][0]['lic'])
            worksheet.write('B15', 1*int(data["hq_locations"]))
            worksheet.write('C15', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D15', 65)
            worksheet.write('E15', '=B15*C15*D15/100')
            
        if ports / 48 != 0 and ports % 48!=0:
            worksheet.write('A14', products['ms']['48'][0]['pn'])
            worksheet.write('B14', int((ports / 48))*int(data["hq_locations"]))
            worksheet.write('C14', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D14', 65)
            worksheet.write('E14', '=B14*C14*D14/100')
            
            worksheet.write('A15', products['ms']['48'][0]['lic'])
            worksheet.write('B15', int((ports / 48))*int(data["hq_locations"]))
            worksheet.write('C15', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D15', 65)
            worksheet.write('E15', '=B15*C15*D15/100')
            
        if ports / 48 != 0 and ports % 48 ==0:
            worksheet.write('A14', products['ms']['48'][0]['pn'])
            worksheet.write('B14', int((ports / 48))*int(data["hq_locations"]))
            worksheet.write('C14', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D14', 65)
            worksheet.write('E14', '=B14*C14*D14/100')
            
            worksheet.write('A15', products['ms']['48'][0]['lic'])
            worksheet.write('B15', int((ports / 48))*int(data["hq_locations"]))
            worksheet.write('C15', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D15', 65)
            worksheet.write('E15', '=B15*C15*D15/100')
            
             #numar + 1


 
    if data["hq_azure"] == "Da":
        worksheet.write('A16', 'MX virtual pentru cloud')
        worksheet.write('A17', products["mx"]["virtual"]["pn"])
        worksheet.write('C17', products["mx"]["virtual"]["gpl"])
        worksheet.write('D17', 65)
        worksheet.write('E17', '=B17*C17*D17/100')
            
        worksheet.write('B17', 1*int(data["hq_locations"]))

    if data["hq_cam_int"] != "":
        if int(data["hq_cam_int"]) != 0 :
            worksheet.write('A18', 'Camera supraveghere interior')
            worksheet.write('A19', products["mv"]["indoor"]["pn"])
            worksheet.write('B19', int(data["hq_cam_int"])*int(data["hq_locations"]))
            worksheet.write('C19', products["mv"]["indoor"]["gpl_product"])
            worksheet.write('D19', 65)
            worksheet.write('E19', '=B19*C19*D19/100')
            
            worksheet.write('A20', products["mv"]["indoor"]["lic"])
            worksheet.write('B20', int(data["hq_cam_int"])*int(data["hq_locations"]))
            worksheet.write('C20', products["mv"]["indoor"]["gpl_license"])
            worksheet.write('D20', 65)
            worksheet.write('E20', '=B20*C20*D20/100')
            
    if data["hq_cam_ext"] != "":
        if data["hq_cam_ext"] != 0:
            worksheet.write('A21', 'Camera supraveghere exterior')
            worksheet.write('A22', products["mv"]["outdoor"]["pn"])
            worksheet.write('B22', int(data["hq_cam_ext"])*int(data["hq_locations"]))
            worksheet.write('C22', products["mv"]["outdoor"]["gpl_product"])
            worksheet.write('D22', 65)
            worksheet.write('E22', '=B22*C22*D22/100')
            
            worksheet.write('A23', products["mv"]["outdoor"]["lic"])
            worksheet.write('B23', int(data["hq_cam_ext"])*int(data["hq_locations"]))
            worksheet.write('C23', products["mv"]["outdoor"]["gpl_license"])
            worksheet.write('D23', 65)
            worksheet.write('E23', '=B23*C23*D23/100')
            
            worksheet.write('A24', products["mv"]["outdoor"]["wall"])
            worksheet.write('B24', int(data["hq_cam_ext"])*int(data["hq_locations"]))
            worksheet.write('C24', products["mv"]["outdoor"]["gpl_wall"])
            worksheet.write('D24', 65)
            worksheet.write('E24', '=B24*C24*D24/100')
            
    
    if data["hq_senz_temp"] != "":
        if data["hq_senz_temp"] !=0:
            worksheet.write('A25', 'Senzori Temp')
            worksheet.write('A26', products["mt"]["temp"]["pn"])
            worksheet.write('B26', int(data["hq_senz_temp"])*int(data["hq_locations"]))
            worksheet.write('C26', products["mt"]["temp"]["gpl_product"])
            worksheet.write('D26', 65)
            worksheet.write('E26', '=B26*C26*D26/100')
            
            worksheet.write('A27', products["mt"]["temp"]["lic"])
            worksheet.write('B27', int(data["hq_senz_temp"])*int(data["hq_locations"]))
            worksheet.write('C27', products["mt"]["temp"]["gpl_license"])
            worksheet.write('D27', 65)
            worksheet.write('E27', '=B27*C27*D27/100')
            
    if data["hq_senz_temp"] != "":
        if data["hq_senz_hum"] !=0:
            worksheet.write('A28', 'Senzori Umiditate')
            worksheet.write('A29', products["mt"]["hum"]["pn"])
            worksheet.write('B29', int(data["hq_senz_hum"])*int(data["hq_locations"]))
            worksheet.write('C29', products["mt"]["hum"]["gpl_product"])
            worksheet.write('D29', 65)
            worksheet.write('E29', '=B29*C29*D29/100')
            
            worksheet.write('A30', products["mt"]["hum"]["lic"])
            worksheet.write('B30', int(data["hq_senz_hum"])*int(data["hq_locations"]))
            worksheet.write('C30', products["mt"]["hum"]["gpl_license"])
            worksheet.write('D30', 65)
            worksheet.write('E30', '=B30*C30*D30/100')
            

    worksheet.write('A31', 'Informatii specifice Branch')
    if data["br_ap"] != "":
        if data["br_ap"] != 0:
            worksheet.write('A32', 'Access Points')
            worksheet.write('A33', products["mr"]["indoor"][0]["pn"])
            worksheet.write('B33', math.ceil(float(int(data["br_ap"]))/100)*int(data["br_locations"]))
            worksheet.write('C33', products["mr"]["indoor"][0]["gpl_product"])
            worksheet.write('D33', 65)
            worksheet.write('E33', '=B33*C33*D33/100')
            
            worksheet.write('A34', products["mr"]["indoor"][0]["lic"])
            worksheet.write('B34', math.ceil(float(data["br_ap"])/100)*int(data["br_locations"]))
            worksheet.write('C34', products["mr"]["indoor"][0]["gpl_license"])
            worksheet.write('D34', 65)
            worksheet.write('E34', '=B34*C34*D34/100')
            

    if data['br_ports'] != '' and data['br_ports'] !=0:
        worksheet.write('A35', 'Switches')
        ports = int(int(data["br_ports"]) * 1.2)
        if ports / 48 == 0 and ports / 24 == 0:
           
            worksheet.write('A36', products['ms']['24']['pn'])
            worksheet.write('B36', 1*int(data["br_locations"]))
            worksheet.write('C36', products['ms']['24']["gpl_product"])
            worksheet.write('D36', 65)
            worksheet.write('E36', '=B36*C36*D36/100')
            
            worksheet.write('A37', products['ms']['24']['pn'])
            worksheet.write('B37', 1*int(data["br_locations"]))
            worksheet.write('C37', products['ms']['24']["gpl_product"])
            worksheet.write('D37', 65)
            worksheet.write('E37', '=B37*C37*D37/100')
            
        if ports / 48 == 0 and ports / 24 != 0:
            
           
            worksheet.write('A36', products['ms']['48'][0]['pn'])
            worksheet.write('B36', 1*int(data["br_locations"]))
            worksheet.write('C36', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D36', 65)
            worksheet.write('E36', '=B36*C36*D36/100')
            
            worksheet.write('A37', products['ms']['48'][0]['lic'])
            worksheet.write('B37', 1*int(data["br_locations"]))
            worksheet.write('C37', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D37', 65)
            worksheet.write('E37', '=B37*C37*D37/100')
            
        if ports / 48 != 0 and ports % 48!=0:
            worksheet.write('A36', products['ms']['48'][0]['pn'])
            worksheet.write('B36', int((ports / 48))*int(data["br_locations"]))
            worksheet.write('C36', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D36', 65)
            worksheet.write('E36', '=B36*C36*D36/100')
            
            worksheet.write('A37', products['ms']['48'][0]['lic'])
            worksheet.write('B37', int((ports / 48))*int(data["br_locations"]))
            worksheet.write('C37', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D37', 65)
            worksheet.write('E37', '=B37*C37*D37/100')
        
        if ports / 48 != 0 and ports % 48==0:
            worksheet.write('A36', products['ms']['48'][0]['pn'])
            worksheet.write('B36', int((ports / 48))*int(data["br_locations"]))
            worksheet.write('C36', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D36', 65)
            worksheet.write('E36', '=B36*C36*D36/100')
            
            worksheet.write('A37', products['ms']['48'][0]['lic'])
            worksheet.write('B37', int((ports / 48))*int(data["br_locations"]))
            worksheet.write('C37', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D37', 65)
            worksheet.write('E37', '=B37*C37*D37/100')
            
             #numar + 1

    if data["hq_azure"] == "Da":
        worksheet.write('A38', 'MX virtual pentru cloud')
        worksheet.write('A39', products["mx"]["virtual"]["pn"])
        worksheet.write('C39', products["mx"]["virtual"]["gpl"])
        worksheet.write('B39', 1*int(data["br_locations"]))
        worksheet.write('D39', 65)
        worksheet.write('E39', '=B39*C39*D39/100')
            

    if data["br_cam_int"] != "":
        if int(data["br_cam_int"]) != 0:
            worksheet.write('A40', 'Camera supraveghere interior')
            worksheet.write('A41', products["mv"]["indoor"]["pn"])
            worksheet.write('B41', int(data["br_cam_int"])*int(data["br_locations"]))
            worksheet.write('C41', products["mv"]["indoor"]["gpl_product"])
            worksheet.write('D41', 65)
            worksheet.write('E41', '=B41*C41*D41/100')
            
            worksheet.write('A42', products["mv"]["indoor"]["lic"])
            worksheet.write('B42', int(data["br_cam_int"])*int(data["br_locations"]))
            worksheet.write('C42', products["mv"]["indoor"]["gpl_license"])
            worksheet.write('D42', 65)
            worksheet.write('E42', '=B42*C42*D42/100')
            
    if data["br_cam_ext"] != "":
        if int(data["br_cam_ext"]) != 0:
            worksheet.write('A43', 'Camera supraveghere exterior')
            worksheet.write('A44', products["mv"]["outdoor"]["pn"])
            worksheet.write('B44', int(data["br_cam_ext"])*int(data["br_locations"]))
            worksheet.write('C44', products["mv"]["outdoor"]["gpl_product"])
            worksheet.write('D44', 65)
            worksheet.write('E44', '=B44*C44*D44/100')
            
            worksheet.write('A45', products["mv"]["outdoor"]["lic"])
            worksheet.write('B45', int(data["br_cam_ext"])*int(data["br_locations"]))
            worksheet.write('C45', products["mv"]["outdoor"]["gpl_license"])
            worksheet.write('D45', 65)
            worksheet.write('E45', '=B45*C45*D45/100')
            
            worksheet.write('A46', products["mv"]["outdoor"]["wall"])
            worksheet.write('B46', int(data["br_cam_ext"])*int(data["br_locations"]))
            worksheet.write('C46', products["mv"]["outdoor"]["gpl_wall"])
            worksheet.write('D46', 65)
            worksheet.write('E46', '=B46*C46*D46/100')
            
    if data["br_senz_temp"] != '':
        if int(data["br_senz_temp"]) !=0:
            worksheet.write('A47', 'Senzori Temp')
            worksheet.write('A48', products["mt"]["temp"]["pn"])
            worksheet.write('B48', int(data["br_senz_temp"])*int(data["br_locations"]))
            worksheet.write('C48', products["mt"]["temp"]["gpl_product"])
            worksheet.write('D48', 65)
            worksheet.write('E48', '=B48*C48*D48/100')
            
            worksheet.write('A49', products["mt"]["temp"]["lic"])
            worksheet.write('B49', int(data["br_senz_temp"])*int(data["br_locations"]))
            worksheet.write('C49', products["mt"]["temp"]["gpl_license"])
            worksheet.write('D49', 65)
            worksheet.write('E49', '=B49*C49*D49/100')
            
    if data["br_senz_hum"] != '':
        if int(data["br_senz_hum"]) !=0:
            worksheet.write('A50', 'Senzori Umiditate')
            worksheet.write('A51', products["mt"]["hum"]["pn"])
            worksheet.write('B51', int(data["br_senz_hum"])*int(data["br_locations"]))
            worksheet.write('C51', products["mt"]["hum"]["gpl_product"])
            worksheet.write('D51', 65)
            worksheet.write('E51', '=B51*C51*D51/100')
            
            worksheet.write('A52', products["mt"]["hum"]["lic"])
            worksheet.write('B52', int(data["br_senz_hum"])*int(data["br_locations"]))
            worksheet.write('C52', products["mt"]["hum"]["gpl_license"])
            worksheet.write('D52', 65)
            worksheet.write('E52', '=B52*C52*D52/100')
            


    worksheet.write('A53', 'Informatii specifice depozit')
    if data["dep_ap_int"] != "":
        if data["dep_ap_int"] != 0:
            worksheet.write('A54', 'Access Points Interior')
            worksheet.write('A55', products["mr"]["indoor"][0]["pn"])
            worksheet.write('B55', math.ceil(float(data["dep_ap_int"])/100))
            worksheet.write('C55', products["mr"]["indoor"][0]["gpl_product"])
            worksheet.write('D55', 65)
            worksheet.write('E55', '=B55*C55*D55/100')
            
            worksheet.write('A56', products["mr"]["indoor"][0]["lic"])
            worksheet.write('B56', math.ceil(float(data["dep_ap_int"])/100))
            worksheet.write('C56', products["mr"]["indoor"][0]["gpl_license"])
            worksheet.write('D56', 65)
            worksheet.write('E56', '=B56*C56*D56/100')
            

    if data["dep_ap_ext"] != "":
        if data["dep_ap_ext"] != 0:
            worksheet.write('A57', 'Access Points Exterior')
            worksheet.write('A58', products["mr"]["outdoor"][0]["pn"])
            worksheet.write('B58', math.ceil(float(data["dep_ap_ext"])/100))
            worksheet.write('C58', products["mr"]["outdoor"][0]["gpl_product"])
            worksheet.write('D58', 65)
            worksheet.write('E58', '=B58*C58*D58/100')
            
            worksheet.write('A59', products["mr"]["outdoor"][0]["lic"])
            worksheet.write('B59', math.ceil(float(data["dep_ap_ext"])/100))
            worksheet.write('C59', products["mr"]["outdoor"][0]["gpl_license"])
            worksheet.write('D59', 65)
            worksheet.write('E59', '=B59*C59*D59/100')
            
            worksheet.write('A60', products["mr"]["outdoor"][0]["antenna"])
            worksheet.write('B60', products["mr"]["outdoor"][0]["number_antenna"])
            worksheet.write('C60', products["mr"]["outdoor"][0]["gpl_antenna"])
            worksheet.write('D60', 65)
            worksheet.write('E60', '=B60*C60*D60/100')
            

    if data['dep_ports'] != '' and data['dep_ports'] !=0:
        worksheet.write('A61', 'Switches')
        ports = int(int(data["dep_ports"]) * 1.2)
        if ports / 48 == 0 and ports / 24 == 0:
             #de 24
            worksheet.write('A62', products['ms']['24']['pn'])
            worksheet.write('B62', 1)
            worksheet.write('C62', products['ms']['24']["gpl_product"])
            worksheet.write('D62', 65)
            worksheet.write('E62', '=B62*C62*D62/100')
            
            worksheet.write('A63', products['ms']['24']['pn'])
            worksheet.write('B63', 1)
            worksheet.write('C63', products['ms']['24']["gpl_product"])
            worksheet.write('D63', 65)
            worksheet.write('E63', '=B63*C63*D63/100')
            
        if ports / 48 == 0 and ports / 24 != 0:
            
             #de 48 1
            worksheet.write('A62', products['ms']['48'][0]['pn'])
            worksheet.write('B62', 1)
            worksheet.write('C62', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D62', 65)
            worksheet.write('E62', '=B62*C62*D62/100')
            
            worksheet.write('A63', products['ms']['48'][0]['lic'])
            worksheet.write('B63', 1)
            worksheet.write('C63', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D63', 65)
            worksheet.write('E63', '=B63*C63*D63/100')
            
        if ports / 48 != 0 and ports % 48!=0:
            worksheet.write('A62', products['ms']['48'][0]['pn'])
            worksheet.write('B62', int((ports / 48)))
            worksheet.write('C62', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D62', 65)
            worksheet.write('E62', '=B62*C62*D62/100')
            
            worksheet.write('A63', products['ms']['48'][0]['lic'])
            worksheet.write('B63', int((ports / 48)))
            worksheet.write('C63', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D63', 65)
            worksheet.write('E63', '=B63*C63*D63/100')
        
        if ports / 48 != 0 and ports % 48==0:
            worksheet.write('A62', products['ms']['48'][0]['pn'])
            worksheet.write('B62', int((ports / 48)))
            worksheet.write('C62', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D62', 65)
            worksheet.write('E62', '=B62*C62*D62/100')
            
            worksheet.write('A63', products['ms']['48'][0]['lic'])
            worksheet.write('B63', int((ports / 48)))
            worksheet.write('C63', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D63', 65)
            worksheet.write('E63', '=B63*C63*D63/100')
            
             #numar + 1

    if data["dep_cam_int"] != 0 and data["dep_cam_int"] != "":
        worksheet.write('A64', 'Camera supraveghere interior')
        worksheet.write('A65', products["mv"]["indoor"]["pn"])
        worksheet.write('B65', int(data["dep_cam_int"]))
        worksheet.write('C65', products["mv"]["indoor"]["gpl_product"])
        worksheet.write('D65', 65)
        worksheet.write('E65', '=B65*C65*D65/100')
            
        worksheet.write('A66', products["mv"]["indoor"]["lic"])
        worksheet.write('B66', int(data["dep_cam_int"]))
        worksheet.write('C66', products["mv"]["indoor"]["gpl_license"])
        worksheet.write('D66', 65)
        worksheet.write('E66', '=B66*C66*D66/100')
            

    if data["dep_cam_ext"] != 0 and data["dep_cam_ext"] != "":
        worksheet.write('A67', 'Camera supraveghere exterior')
        worksheet.write('A68', products["mv"]["outdoor"]["pn"])
        worksheet.write('B68', int(data["dep_cam_ext"]))
        worksheet.write('C68', products["mv"]["outdoor"]["gpl_product"])
        worksheet.write('D68', 65)
        worksheet.write('E68', '=B68*C68*D68/100')
            
        worksheet.write('A69', products["mv"]["outdoor"]["lic"])
        worksheet.write('B69', int(data["dep_cam_ext"]))
        worksheet.write('C69', products["mv"]["outdoor"]["gpl_license"])
        worksheet.write('D69', 65)
        worksheet.write('E69', '=B69*C69*D69/100')
            
        worksheet.write('A70', products["mv"]["outdoor"]["wall"])
        worksheet.write('B70', int(data["dep_cam_ext"]))
        worksheet.write('C70', products["mv"]["outdoor"]["gpl_wall"])
        worksheet.write('D70', 65)
        worksheet.write('E70', '=B70*C70*D70/100')
            

    if data["dep_senz_temp"] !=0 and data["dep_senz_temp"] != '':
        worksheet.write('A71', 'Senzori Temp')
        worksheet.write('A72', products["mt"]["temp"]["pn"])
        worksheet.write('B72', int(data["dep_senz_temp"]))
        worksheet.write('C72', products["mt"]["temp"]["gpl_product"])
        worksheet.write('D72', 65)
        worksheet.write('E72', '=B72*C72*D72/100')
            
        worksheet.write('A73', products["mt"]["temp"]["lic"])
        worksheet.write('B73', int(data["dep_senz_temp"]))
        worksheet.write('C73', products["mt"]["temp"]["gpl_license"])
        worksheet.write('D73', 65)
        worksheet.write('E73', '=B73*C73*D73/100')
            
    if data["dep_senz_hum"] !=0 and data["dep_senz_hum"] != '':
        worksheet.write('A74', 'Senzori Umiditate')
        worksheet.write('A75', products["mt"]["hum"]["pn"])
        worksheet.write('B75', int(data["dep_senz_hum"]))
        
        worksheet.write('C75', products["mt"]["hum"]["gpl_product"])
        worksheet.write('D75', 65)
        worksheet.write('E75', '=B75*C75*D75/100')
            
        worksheet.write('A76', products["mt"]["hum"]["lic"])
        worksheet.write('B76', int(data["dep_senz_hum"]))
        worksheet.write('C76', products["mt"]["hum"]["gpl_license"])
        worksheet.write('D76', 65)
        worksheet.write('E76', '=B76*C76*D76/100')
            

    worksheet.write('E78', 'TOTAL USD')
    worksheet.write('E79', '=SUM(E4:E76)')

    worksheet.write('A90', '=SUM(A80 + A81)')
    url = "https://www.bnr.ro/nbrfxrates.xml"

    payload={}
    headers = {
    'Authorization': 'Token {{Token_value}}',
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Cookie': 'TS01cc05a4=0187e48c1b36104d94a3b1033b1b6766457e31ebff81d11c164980a17e4f193ce787d84692c35b0f5a02c4d2a73d86a66c6738a6fc'
    }

    response = requests.request("GET", url, headers=headers, data=payload)

    print(response.text)
    doc = xmltodict.parse(response.text)
    rates =  json.loads(json.dumps(doc))
    euro = rates["DataSet"]["Body"]["Cube"]["Rate"][10]["#text"]
    print(euro)
    usd = rates["DataSet"]["Body"]["Cube"]["Rate"][28]["#text"]
    print(usd)
    print(float(float(euro)/float(usd)))
    worksheet.write('E80', 'TOTAL EURO')
    worksheet.write('E81', '=SUM(E4:E76)*'+str(float(euro)/float(usd)))
    workbook.close()
    return None


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=80)
