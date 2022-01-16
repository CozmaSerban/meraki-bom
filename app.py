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
                "pn": "AP MR45-HW",
                "lic": "Licenta LIC-ENT-3YR",
                "gpl_product": 1200,
                "gpl_license": 321
            },
        ],
        "outdoor": [
            {
                "pn": "AP MR76-HW",
                "lic": "Licenta LIC-ENT-3YR",
                "gpl_product": 1887.72,
                "gpl_license": 321,
                "antenna": "Antena MA-ANT-20",
                "number_antenna": 2,
                "gpl_antenna": 213.44
            },
        ]
    },
    "ms": {
        "24": [
            {
                "pn": "Switch MS120-24P-HW",
                "lic": "Licenta LIC-MS120-24P-3YR",
                "gpl_product": 2831.58,
                "gpl_license": 304.95
            },
        ],
        "48": [
            {
                "pn": "Switch MS120-48LP-HW",
                "lic": "Licenta LIC-MS120-48LP-3YR",
                "gpl_product": 4419.98,
                "gpl_license": 476.15
            },
            {
                "pn": "Switch MS120-48FP-HW",
                "lic": "Licenta LIC-MS120-48FP-3YR",
                "gpl_product": 5218.04,
                "gpl_license": 567.1
            },
        ]
    },
    "mx": {
        "virtual": {
            "pn": "Router/Firewall Virtual LIC-VMX-S-ENT-3Y",
            "gpl": 1000
        },
        "4g": {
            "pn": "Router/Firewall 4G MX75-HW",
            "lic": "LIC-MX75-SEC-3Y",
            "gpl_product": 2139.77,
            "gpl_license": 4000
        },
        "redudant":
            [
                {
                    "pn": "Router/Firewall MX64-HW",
                    "lic": "Licenta LIC-MX64-SEC-3YR",
                    "gpl_product": 638.18,
                    "gpl_license": 1200,
                    "tput": 250
                },
                {
                    "pn": "Router/Firewall MX68-HW",
                    "lic": "Licenta LIC-MX68-SEC-3YR",
                    "gpl_product": 1067.21,
                    "gpl_license": 1500,
                    "tput": 450
                },
                {
                    "pn": "Router/Firewall MX75-HW",
                    "lic": "Licenta LIC-MX75-SEC-3Y",
                    "gpl_product": 2139.77,
                    "gpl_license": 4000,
                    "tput": 1000
                }
        ]
    },
    "mv": {
        "outdoor": {
            "pn": "Camera MV72-HW",
            "lic": "Licenta LIC-MV-3YR",
            "gpl_product": 1703.1,
            "wall": "Suport MA-MNT-MV-10",
            "gpl_license": 600,
            "gpl_wall": 267.1,
        },
        "indoor": {
            "pn": "Camera MV32-HW",
            "lic": "Licenta LIC-MV-3YR",
            "gpl_product": 1502.6,
            "wall": "Suport MA-MNT-MV-30",
            "gpl_license": 600,
            "gpl_wall": 267.1,
        }
    },
    "mt": {
        "temp": {
            "pn": "Senzor Temperatura MT10-HW",
            "lic": "Licenta LIC-MT-3Y",
            "gpl_product": 249.6,
            "gpl_license": 300,
        },
        "hum": {
            "pn": "Senzor Umiditate MT12-HW",
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
    worksheet.write('A2', 'PN')
    worksheet.write('B2', 'Quantity')
    worksheet.write('C2', 'GPL')
    worksheet.write('D2', 'Discount')
    worksheet.write('E2', 'Price')
    ##INFORMATII HQ##################################################################################
    worksheet.write('A3', 'Informatii specifice HQ x'+str(data["hq_locations"]))
    
    ## MX HQ
    if data["tput_hq"] != "":
        tput = int(data["hq_locations"])
        if int(data["tput_hq"]) <= 250:
            worksheet.write('A4', products["mx"]["redudant"][0]["pn"])
            worksheet.write('B4', 1*(tput ))
            worksheet.write('C4', products["mx"]["redudant"][0]["gpl_product"])
            worksheet.write('D4', 65)
            worksheet.write('E4', '=B4*C4*D4/100')

            worksheet.write('A5', products["mx"]["redudant"][0]["lic"])
            worksheet.write('B5', 1*(tput ))
            worksheet.write('C5', products["mx"]["redudant"][0]["gpl_license"])
            worksheet.write('D5', 65)
            worksheet.write('E5', '=B5*C5*D5/100')
            
        elif int(data["tput_hq"]) <= 450:
            worksheet.write('A4', products["mx"]["redudant"][1]["pn"])
            worksheet.write('B4', 1*(tput ))
            worksheet.write('C4', products["mx"]["redudant"][1]["gpl_product"])
            worksheet.write('D4', 65)
            worksheet.write('E4', '=B4*C4*D4/100')
            worksheet.write('A5', products["mx"]["redudant"][1]["lic"])
            worksheet.write('B5', 1*(tput ))
            worksheet.write('C5', products["mx"]["redudant"][1]["gpl_license"])
            worksheet.write('D5', 65)
            worksheet.write('E5', '=B5*C5*D5/100')
            
        elif int(data["tput_hq"]) <= 1000:
            worksheet.write('A4', products["mx"]["redudant"][2]["pn"])
            worksheet.write('B4', 1*(tput ))
            worksheet.write('C4', products["mx"]["redudant"][2]["gpl_product"])
            worksheet.write('D4', 65)
            worksheet.write('E4', '=B4*C4*D4/100')
        
            worksheet.write('A5', products["mx"]["redudant"][2]["lic"])
            worksheet.write('B5', 1*(tput ))
            worksheet.write('C5', products["mx"]["redudant"][2]["gpl_license"])
            worksheet.write('D5', 65)
            worksheet.write('E5', '=B5*C5*D5/100')
            
        else:
            worksheet.write('A4', products["mx"]["redudant"][1]["pn"])
            worksheet.write('B4', 1*(tput ))
            worksheet.write('C4', products["mx"]["redudant"][1]["gpl_product"])
            worksheet.write('D4', 65)
            worksheet.write('E4', '=B4*C4*D4/100')
            
            worksheet.write('A5', products["mx"]["redudant"][1]["lic"])
            worksheet.write('B5', 1*(tput ))
            worksheet.write('C5', products["mx"]["redudant"][1]["gpl_license"])
            worksheet.write('D5', 65)
            worksheet.write('E5', '=B5*C5*D5/100')
    ##MS HQ####################################

    if data['hq_ports'] != '' and data['hq_ports'] !=0:
        worksheet.write('A6', 'Switches')
        ports = int(int(data["hq_ports"]) * 1.2)
        if ports / 48 == 0 and ports / 24 == 0:
             #de 24
            worksheet.write('A7', products['ms']['24']['pn'])
            worksheet.write('B7', 1*int(data["hq_locations"]))
            worksheet.write('C7', products['ms']['24']["gpl_product"])
            worksheet.write('D7', 65)
            worksheet.write('E7', '=B7*C7*D7/100')
            
            worksheet.write('A8', products['ms']['24']['pn'])
            worksheet.write('B8', 1*int(data["hq_locations"]))
            worksheet.write('C8', products['ms']['24']["gpl_product"])
            worksheet.write('D8', 65)
            worksheet.write('E8', '=B8*C8*D8/100')
            
        if ports / 48 == 0 and ports / 24 != 0:
            
             #de 48 1
            worksheet.write('A7', products['ms']['48'][0]['pn'])
            worksheet.write('B7', 1*int(data["hq_locations"]))
            worksheet.write('C7', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D7', 65)
            worksheet.write('E7', '=B7*C7*D7/100')
            
            worksheet.write('A8', products['ms']['48'][0]['lic'])
            worksheet.write('B8', 1*int(data["hq_locations"]))
            worksheet.write('C8', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D8', 65)
            worksheet.write('E8', '=B8*C8*D8/100')
            
        if ports / 48 != 0 and ports % 48!=0:
            worksheet.write('A7', products['ms']['48'][0]['pn'])
            worksheet.write('B7', int((ports / 48))*int(data["hq_locations"]))
            worksheet.write('C7', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D7', 65)
            worksheet.write('E7', '=B7*C7*D7/100')
            
            worksheet.write('A8', products['ms']['48'][0]['lic'])
            worksheet.write('B8', int((ports / 48))*int(data["hq_locations"]))
            worksheet.write('C8', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D8', 65)
            worksheet.write('E8', '=B8*C8*D8/100')
            
        if ports / 48 != 0 and ports % 48 ==0:
            worksheet.write('A7', products['ms']['48'][0]['pn'])
            worksheet.write('B7', int((ports / 48))*int(data["hq_locations"]))
            worksheet.write('C7', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D7', 65)
            worksheet.write('E7', '=B7*C7*D7/100')
            
            worksheet.write('A8', products['ms']['48'][0]['lic'])
            worksheet.write('B8', int((ports / 48))*int(data["hq_locations"]))
            worksheet.write('C8', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D8', 65)
            worksheet.write('E8', '=B8*C8*D8/100')
            
             #numar + 1
    ##MR HQ ################################################
    if data["hq_ap"] != 0 and data["hq_ap"] != "":
            worksheet.write('A9', 'Access Points')
            worksheet.write('A9', products["mr"]["indoor"][0]["pn"])
            worksheet.write('B9', math.ceil(float(data["hq_ap"])/100)*int(data["hq_locations"]))
            worksheet.write('C9', products["mr"]["indoor"][0]["gpl_product"])
            worksheet.write('D9', 65)
            worksheet.write('E9', '=B9*C9*D9/100')
                
            worksheet.write('A10', products["mr"]["indoor"][0]["lic"])
            worksheet.write('B10', math.ceil(float(data["hq_ap"])/100)*int(data["hq_locations"]))
            worksheet.write('C10', products["mr"]["indoor"][0]["gpl_license"])
            worksheet.write('D10', 65)
            worksheet.write('E10', '=B10*C10*D10/100')
    
    ##AZURE HQ ################################################
    if data["hq_azure"] == "Da":
            worksheet.write('A11', 'MX virtual pentru cloud')
            worksheet.write('A12', products["mx"]["virtual"]["pn"])
            worksheet.write('C12', products["mx"]["virtual"]["gpl"])
            worksheet.write('D12', 65)
            worksheet.write('E12', '=B12*C12*D12/100')
                
            worksheet.write('B12', 1*int(data["hq_locations"]))
    ##MV HQ ###################################################
    if data["hq_cam_int"] != "":
            if int(data["hq_cam_int"]) != 0 :
                worksheet.write('A13', 'Camera supraveghere interior')
                worksheet.write('A14', products["mv"]["indoor"]["pn"])
                worksheet.write('B14', int(data["hq_cam_int"])*int(data["hq_locations"]))
                worksheet.write('C14', products["mv"]["indoor"]["gpl_product"])
                worksheet.write('D14', 65)
                worksheet.write('E14', '=B14*C14*D14/100')
                
                worksheet.write('A15', products["mv"]["indoor"]["lic"])
                worksheet.write('B15', int(data["hq_cam_int"])*int(data["hq_locations"]))
                worksheet.write('C15', products["mv"]["indoor"]["gpl_license"])
                worksheet.write('D15', 65)
                worksheet.write('E15', '=B15*C15*D15/100')
                
    if data["hq_cam_ext"] != "":
            if data["hq_cam_ext"] != 0:
                worksheet.write('A16', 'Camera supraveghere exterior')
                worksheet.write('A17', products["mv"]["outdoor"]["pn"])
                worksheet.write('B17', int(data["hq_cam_ext"])*int(data["hq_locations"]))
                worksheet.write('C17', products["mv"]["outdoor"]["gpl_product"])
                worksheet.write('D17', 65)
                worksheet.write('E17', '=B17*C17*D17/100')
                
                worksheet.write('A18', products["mv"]["outdoor"]["lic"])
                worksheet.write('B18', int(data["hq_cam_ext"])*int(data["hq_locations"]))
                worksheet.write('C18', products["mv"]["outdoor"]["gpl_license"])
                worksheet.write('D18', 65)
                worksheet.write('E18', '=B18*C18*D18/100')
                
                worksheet.write('A19', products["mv"]["outdoor"]["wall"])
                worksheet.write('B19', int(data["hq_cam_ext"])*int(data["hq_locations"]))
                worksheet.write('C19', products["mv"]["outdoor"]["gpl_wall"])
                worksheet.write('D19', 65)
                worksheet.write('E19', '=B19*C19*D19/100')
    ##MT HQ #####################################################

    
    if data["hq_senz_temp"] != "":
            if data["hq_senz_temp"] !=0:
                worksheet.write('A20', 'Senzori Temp')
                worksheet.write('A21', products["mt"]["temp"]["pn"])
                worksheet.write('B21', int(data["hq_senz_temp"])*int(data["hq_locations"]))
                worksheet.write('C21', products["mt"]["temp"]["gpl_product"])
                worksheet.write('D21', 65)
                worksheet.write('E21', '=B21*C21*D21/100')
                
                worksheet.write('A22', products["mt"]["temp"]["lic"])
                worksheet.write('B22', int(data["hq_senz_temp"])*int(data["hq_locations"]))
                worksheet.write('C22', products["mt"]["temp"]["gpl_license"])
                worksheet.write('D22', 65)
                worksheet.write('E22', '=B22*C22*D22/100')
                
    if data["hq_senz_hum"] != "":
            if data["hq_senz_hum"] !=0:
                worksheet.write('A23', 'Senzori Umiditate')
                worksheet.write('A24', products["mt"]["hum"]["pn"])
                worksheet.write('B24', int(data["hq_senz_hum"])*int(data["hq_locations"]))
                worksheet.write('C24', products["mt"]["hum"]["gpl_product"])
                worksheet.write('D24', 65)
                worksheet.write('E24', '=B24*C24*D24/100')
                
                worksheet.write('A25', products["mt"]["hum"]["lic"])
                worksheet.write('B25', int(data["hq_senz_hum"])*int(data["hq_locations"]))
                worksheet.write('C25', products["mt"]["hum"]["gpl_license"])
                worksheet.write('D25', 65)
                worksheet.write('E25', '=B25*C25*D25/100')

    #######################################
    #######################################
    ## INFORMATII BRANCH #######################################
    worksheet.write('A26', 'Informatii specifice Branch x'+str(data["br_locations"]))

    ## MX BRANCH#####################################################
    if data["tput_br"] != "" :
        tput = int(data["br_locations"])
        if int(data["tput_br"]) <= 250:
            worksheet.write('A27', products["mx"]["redudant"][0]["pn"])
            worksheet.write('B27', 1*(tput ))
            worksheet.write('C27', products["mx"]["redudant"][0]["gpl_product"])
            worksheet.write('D27', 65)
            worksheet.write('E27', '=B27*C27*D27/100')

            worksheet.write('A28', products["mx"]["redudant"][0]["lic"])
            worksheet.write('B28', 1*(tput ))
            worksheet.write('C28', products["mx"]["redudant"][0]["gpl_license"])
            worksheet.write('D28', 65)
            worksheet.write('E28', '=B28*C28*D28/100')
            
        elif int(data["tput_br"]) <= 450:
            worksheet.write('A27', products["mx"]["redudant"][1]["pn"])
            worksheet.write('B27', 1*(tput ))
            worksheet.write('C27', products["mx"]["redudant"][1]["gpl_product"])
            worksheet.write('D27', 65)
            worksheet.write('E27', '=B27*C27*D27/100')
            worksheet.write('A28', products["mx"]["redudant"][1]["lic"])
            worksheet.write('B28', 1*(tput ))
            worksheet.write('C28', products["mx"]["redudant"][1]["gpl_license"])
            worksheet.write('D28', 65)
            worksheet.write('E28', '=B28*C28*D28/100')
            
        elif int(data["tput_br"]) <= 1000:
            worksheet.write('A27', products["mx"]["redudant"][2]["pn"])
            worksheet.write('B27', 1*(tput ))
            worksheet.write('C27', products["mx"]["redudant"][2]["gpl_product"])
            worksheet.write('D27', 65)
            worksheet.write('E27', '=B27*C27*D27/100')
        
            worksheet.write('A28', products["mx"]["redudant"][2]["lic"])
            worksheet.write('B28', 1*(tput ))
            worksheet.write('C28', products["mx"]["redudant"][2]["gpl_license"])
            worksheet.write('D28', 65)
            worksheet.write('E28', '=B28*C28*D28/100')
            
        else:
            worksheet.write('A27', products["mx"]["redudant"][1]["pn"])
            worksheet.write('B27', 1*(tput ))
            worksheet.write('C27', products["mx"]["redudant"][1]["gpl_product"])
            worksheet.write('D27', 65)
            worksheet.write('E27', '=B27*C27*D27/100')
            
            worksheet.write('A28', products["mx"]["redudant"][1]["lic"])
            worksheet.write('B28', 1*(tput ))
            worksheet.write('C28', products["mx"]["redudant"][1]["gpl_license"])
            worksheet.write('D28', 65)
            worksheet.write('E28', '=B28*C28*D28/100')
    ## MS BRANCH #############################################################
    if data['br_ports'] != '' and data['br_ports'] !=0:
        worksheet.write('A29', 'Switches')
        ports = int(int(data["br_ports"]) * 1.2)
        if ports / 48 == 0 and ports / 24 == 0:
           
            worksheet.write('A30', products['ms']['24']['pn'])
            worksheet.write('B30', 1*int(data["br_locations"]))
            worksheet.write('C30', products['ms']['24']["gpl_product"])
            worksheet.write('D30', 65)
            worksheet.write('E30', '=B30*C30*D30/100')
            
            worksheet.write('A31', products['ms']['24']['pn'])
            worksheet.write('B31', 1*int(data["br_locations"]))
            worksheet.write('C31', products['ms']['24']["gpl_product"])
            worksheet.write('D31', 65)
            worksheet.write('E31', '=B31*C31*D31/100')
            
        if ports / 48 == 0 and ports / 24 != 0:
            
           
            worksheet.write('A30', products['ms']['48'][0]['pn'])
            worksheet.write('B30', 1*int(data["br_locations"]))
            worksheet.write('C30', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D30', 65)
            worksheet.write('E30', '=B30*C30*D30/100')
            
            worksheet.write('A31', products['ms']['48'][0]['lic'])
            worksheet.write('B31', 1*int(data["br_locations"]))
            worksheet.write('C31', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D31', 65)
            worksheet.write('E31', '=B31*C31*D31/100')
            
        if ports / 48 != 0 and ports % 48!=0:
            worksheet.write('A30', products['ms']['48'][0]['pn'])
            worksheet.write('B30', int((ports / 48))*int(data["br_locations"]))
            worksheet.write('C30', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D30', 65)
            worksheet.write('E30', '=B30*C30*D30/100')
            
            worksheet.write('A31', products['ms']['48'][0]['lic'])
            worksheet.write('B31', int((ports / 48))*int(data["br_locations"]))
            worksheet.write('C31', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D31', 65)
            worksheet.write('E31', '=B31*C31*D31/100')
        
        if ports / 48 != 0 and ports % 48==0:
            worksheet.write('A30', products['ms']['48'][0]['pn'])
            worksheet.write('B30', int((ports / 48))*int(data["br_locations"]))
            worksheet.write('C30', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D30', 65)
            worksheet.write('E30', '=B30*C30*D30/100')
            
            worksheet.write('A31', products['ms']['48'][0]['lic'])
            worksheet.write('B31', int((ports / 48))*int(data["br_locations"]))
            worksheet.write('C31', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D31', 65)
            worksheet.write('E31', '=B31*C31*D31/100')
            
             #numar + 1
    ## MR BRANCH ###########################################################
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
    
    ## MV BRANCH #########################################################
    if data["br_cam_int"] != "":
        if int(data["br_cam_int"]) != 0:
            worksheet.write('A35', 'Camera supraveghere interior')
            worksheet.write('A36', products["mv"]["indoor"]["pn"])
            worksheet.write('B36', int(data["br_cam_int"])*int(data["br_locations"]))
            worksheet.write('C36', products["mv"]["indoor"]["gpl_product"])
            worksheet.write('D36', 65)
            worksheet.write('E36', '=B36*C36*D36/100')
            
            worksheet.write('A37', products["mv"]["indoor"]["lic"])
            worksheet.write('B37', int(data["br_cam_int"])*int(data["br_locations"]))
            worksheet.write('C37', products["mv"]["indoor"]["gpl_license"])
            worksheet.write('D37', 65)
            worksheet.write('E37', '=B37*C37*D37/100')
            
    if data["br_cam_ext"] != "":
        if int(data["br_cam_ext"]) != 0:
            worksheet.write('A38', 'Camera supraveghere exterior')
            worksheet.write('A39', products["mv"]["outdoor"]["pn"])
            worksheet.write('B39', int(data["br_cam_ext"])*int(data["br_locations"]))
            worksheet.write('C39', products["mv"]["outdoor"]["gpl_product"])
            worksheet.write('D39', 65)
            worksheet.write('E39', '=B39*C39*D39/100')
            
            worksheet.write('A40', products["mv"]["outdoor"]["lic"])
            worksheet.write('B40', int(data["br_cam_ext"])*int(data["br_locations"]))
            worksheet.write('C40', products["mv"]["outdoor"]["gpl_license"])
            worksheet.write('D40', 65)
            worksheet.write('E40', '=B40*C40*D40/100')
            
            worksheet.write('A41', products["mv"]["outdoor"]["wall"])
            worksheet.write('B41', int(data["br_cam_ext"])*int(data["br_locations"]))
            worksheet.write('C41', products["mv"]["outdoor"]["gpl_wall"])
            worksheet.write('D41', 65)
            worksheet.write('E41', '=B41*C41*D41/100')
    
    ##MT BRANCH #####################################################################
    if data["br_senz_temp"] != '':
        if int(data["br_senz_temp"]) !=0:
            worksheet.write('A42', 'Senzori Temp')
            worksheet.write('A43', products["mt"]["temp"]["pn"])
            worksheet.write('B43', int(data["br_senz_temp"])*int(data["br_locations"]))
            worksheet.write('C43', products["mt"]["temp"]["gpl_product"])
            worksheet.write('D43', 65)
            worksheet.write('E43', '=B43*C43*D43/100')
            
            worksheet.write('A44', products["mt"]["temp"]["lic"])
            worksheet.write('B44', int(data["br_senz_temp"])*int(data["br_locations"]))
            worksheet.write('C44', products["mt"]["temp"]["gpl_license"])
            worksheet.write('D44', 65)
            worksheet.write('E44', '=B44*C44*D44/100')
            
    if data["br_senz_hum"] != '':
        if int(data["br_senz_hum"]) !=0:
            worksheet.write('A45', 'Senzori Umiditate')
            worksheet.write('A46', products["mt"]["hum"]["pn"])
            worksheet.write('B46', int(data["br_senz_hum"])*int(data["br_locations"]))
            worksheet.write('C46', products["mt"]["hum"]["gpl_product"])
            worksheet.write('D46', 65)
            worksheet.write('E46', '=B46*C46*D46/100')
            
            worksheet.write('A47', products["mt"]["hum"]["lic"])
            worksheet.write('B47', int(data["br_senz_hum"])*int(data["br_locations"]))
            worksheet.write('C47', products["mt"]["hum"]["gpl_license"])
            worksheet.write('D47', 65)
            worksheet.write('E47', '=B47*C47*D47/100')
    
    #######################################
    #######################################
    ## INFORMATII DEPOZIT #######################################
    worksheet.write('A48', 'Informatii specifice Depozit x'+str(data["dep_locations"]))
    ## MX Depozit#####################################################
    if data["tput_br"] != "" :
        tput = int(data["dep_locations"])
        if int(data["tput_br"]) <= 250:
            worksheet.write('A49', products["mx"]["redudant"][0]["pn"])
            worksheet.write('B49', 1*(tput ))
            worksheet.write('C49', products["mx"]["redudant"][0]["gpl_product"])
            worksheet.write('D49', 65)
            worksheet.write('E49', '=B49*C49*D49/100')

            worksheet.write('A50', products["mx"]["redudant"][0]["lic"])
            worksheet.write('B50', 1*(tput ))
            worksheet.write('C50', products["mx"]["redudant"][0]["gpl_license"])
            worksheet.write('D50', 65)
            worksheet.write('E50', '=B50*C50*D50/100')
            
        elif int(data["tput_hq"]) <= 450:
            worksheet.write('A49', products["mx"]["redudant"][1]["pn"])
            worksheet.write('B49', 1*(tput ))
            worksheet.write('C49', products["mx"]["redudant"][1]["gpl_product"])
            worksheet.write('D49', 65)
            worksheet.write('E49', '=B49*C49*D49/100')

            worksheet.write('A50', products["mx"]["redudant"][1]["lic"])
            worksheet.write('B50', 1*(tput ))
            worksheet.write('C50', products["mx"]["redudant"][1]["gpl_license"])
            worksheet.write('D50', 65)
            worksheet.write('E50', '=B50*C50*D50/100')
            
        elif int(data["tput_hq"]) <= 1000:
            worksheet.write('A49', products["mx"]["redudant"][2]["pn"])
            worksheet.write('B49', 1*(tput ))
            worksheet.write('C49', products["mx"]["redudant"][2]["gpl_product"])
            worksheet.write('D49', 65)
            worksheet.write('E49', '=B49*C49*D49/100')

            worksheet.write('A50', products["mx"]["redudant"][2]["lic"])
            worksheet.write('B50', 1*(tput ))
            worksheet.write('C50', products["mx"]["redudant"][2]["gpl_license"])
            worksheet.write('D50', 65)
            worksheet.write('E50', '=B50*C50*D50/100')
            
        else:
            worksheet.write('A49', products["mx"]["redudant"][1]["pn"])
            worksheet.write('B49', 1*(tput ))
            worksheet.write('C49', products["mx"]["redudant"][1]["gpl_product"])
            worksheet.write('D49', 65)
            worksheet.write('E49', '=B49*C49*D49/100')

            worksheet.write('A50', products["mx"]["redudant"][1]["lic"])
            worksheet.write('B50', 1*(tput ))
            worksheet.write('C50', products["mx"]["redudant"][1]["gpl_license"])
            worksheet.write('D50', 65)
            worksheet.write('E50', '=B50*C50*D50/100')
    
    ## MS Depozit#####################################################

    if data['dep_ports'] != '' and data['dep_ports'] !=0:
        ports = int(int(data["dep_ports"]) * 1.2 * int(data['dep_locations']))
        if ports / 48 == 0 and ports / 24 == 0:
             #de 24
            worksheet.write('A51', products['ms']['24']['pn'])
            worksheet.write('B51', 1)
            worksheet.write('C51', products['ms']['24']["gpl_product"])
            worksheet.write('D51', 65)
            worksheet.write('E51', '=B51*C51*D51/100')
            
            worksheet.write('A52', products['ms']['24']['pn'])
            worksheet.write('B52', 1)
            worksheet.write('C52', products['ms']['24']["gpl_product"])
            worksheet.write('D52', 65)
            worksheet.write('E52', '=B52*C52*D52/100')
            
        if ports / 48 == 0 and ports / 24 != 0:
            
             #de 48 1
            worksheet.write('A51', products['ms']['48'][0]['pn'])
            worksheet.write('B51', 1)
            worksheet.write('C51', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D51', 65)
            worksheet.write('E51', '=B51*C51*D51/100')
            
            worksheet.write('A52', products['ms']['48'][0]['lic'])
            worksheet.write('B52', 1)
            worksheet.write('C52', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D52', 65)
            worksheet.write('E52', '=B52*C52*D52/100')
            
        if ports / 48 != 0 and ports % 48!=0:
            worksheet.write('A51', products['ms']['48'][0]['pn'])
            worksheet.write('B51', int((ports / 48)))
            worksheet.write('C51', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D51', 65)
            worksheet.write('E51', '=B51*C51*D51/100')
            
            worksheet.write('A52', products['ms']['48'][0]['lic'])
            worksheet.write('B52', int((ports / 48)))
            worksheet.write('C52', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D52', 65)
            worksheet.write('E52', '=B52*C52*D52/100')
        
        if ports / 48 != 0 and ports % 48==0:
            worksheet.write('A51', products['ms']['48'][0]['pn'])
            worksheet.write('B51', int((ports / 48)))
            worksheet.write('C51', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D51', 65)
            worksheet.write('E51', '=B51*C51*D51/100')
            
            worksheet.write('A52', products['ms']['48'][0]['lic'])
            worksheet.write('B52', int((ports / 48)))
            worksheet.write('C52', products['ms']['48'][0]["gpl_product"])
            worksheet.write('D52', 65)
            worksheet.write('E52', '=B52*C52*D52/100')
    
    ## MR Depozit#####################################################

    if data["dep_ap_int"] != "":
        if data["dep_ap_int"] != 0:
            worksheet.write('A53', 'Access Points Interior')
            worksheet.write('A54', products["mr"]["indoor"][0]["pn"])
            worksheet.write('B54', math.ceil(float(data["dep_ap_int"])/100)* int(data['dep_locations']))
            worksheet.write('C54', products["mr"]["indoor"][0]["gpl_product"])
            worksheet.write('D54', 65)
            worksheet.write('E54', '=B54*C54*D54/100')
            
            worksheet.write('A55', products["mr"]["indoor"][0]["lic"])
            worksheet.write('B55', math.ceil(float(data["dep_ap_int"])/100)* int(data['dep_locations']))
            worksheet.write('C55', products["mr"]["indoor"][0]["gpl_license"])
            worksheet.write('D55', 65)
            worksheet.write('E55', '=B55*C55*D55/100')
            

    if data["dep_ap_ext"] != "":
        if data["dep_ap_ext"] != 0:
            worksheet.write('A56', 'Access Points Exterior')
            worksheet.write('A57', products["mr"]["outdoor"][0]["pn"])
            worksheet.write('B57', math.ceil(float(data["dep_ap_ext"])/100)* int(data['dep_locations']))
            worksheet.write('C57', products["mr"]["outdoor"][0]["gpl_product"])
            worksheet.write('D57', 65)
            worksheet.write('E57', '=B57*C57*D57/100')
            
            worksheet.write('A58', products["mr"]["outdoor"][0]["lic"])
            worksheet.write('B58', math.ceil(float(data["dep_ap_ext"])/100)* int(data['dep_locations']))
            worksheet.write('C58', products["mr"]["outdoor"][0]["gpl_license"])
            worksheet.write('D58', 65)
            worksheet.write('E58', '=B58*C58*D58/100')
            
            worksheet.write('A59', products["mr"]["outdoor"][0]["antenna"])
            worksheet.write('B59', products["mr"]["outdoor"][0]["number_antenna"]* int(data['dep_locations']))
            worksheet.write('C59', products["mr"]["outdoor"][0]["gpl_antenna"])
            worksheet.write('D59', 65)
            worksheet.write('E59', '=B59*C59*D59/100')
    
    ## MV Depozit#####################################################
    if data["dep_cam_int"] != 0 and data["dep_cam_int"] != "":
        worksheet.write('A60', 'Camera supraveghere interior')
        worksheet.write('A61', products["mv"]["indoor"]["pn"])
        worksheet.write('B61', int(data["dep_cam_int"])* int(data['dep_locations']))
        worksheet.write('C61', products["mv"]["indoor"]["gpl_product"])
        worksheet.write('D61', 65)
        worksheet.write('E61', '=B61*C61*D61/100')
            
        worksheet.write('A62', products["mv"]["indoor"]["lic"])
        worksheet.write('B62', int(data["dep_cam_int"])* int(data['dep_locations']))
        worksheet.write('C62', products["mv"]["indoor"]["gpl_license"])
        worksheet.write('D62', 65)
        worksheet.write('E62', '=B62*C62*D62/100')
            

    if data["dep_cam_ext"] != 0 and data["dep_cam_ext"] != "":
        worksheet.write('A63', 'Camera supraveghere exterior')
        worksheet.write('A64', products["mv"]["outdoor"]["pn"])
        worksheet.write('B64', int(data["dep_cam_ext"])* int(data['dep_locations']))
        worksheet.write('C64', products["mv"]["outdoor"]["gpl_product"])
        worksheet.write('D64', 65)
        worksheet.write('E64', '=B64*C64*D64/100')
            
        worksheet.write('A65', products["mv"]["outdoor"]["lic"])
        worksheet.write('B65', int(data["dep_cam_ext"]* int(data['dep_locations'])))
        worksheet.write('C65', products["mv"]["outdoor"]["gpl_license"])
        worksheet.write('D65', 65)
        worksheet.write('E65', '=B65*C65*D65/100')
            
        worksheet.write('A66', products["mv"]["outdoor"]["wall"])
        worksheet.write('B66', int(data["dep_cam_ext"])* int(data['dep_locations']))
        worksheet.write('C66', products["mv"]["outdoor"]["gpl_wall"])
        worksheet.write('D66', 65)
        worksheet.write('E66', '=B66*C66*D66/100')

    
    ## MT Depozit#####################################################
    
    if data["dep_senz_temp"] !=0 and data["dep_senz_temp"] != '':
        worksheet.write('A67', 'Senzori Temp')
        worksheet.write('A68', products["mt"]["temp"]["pn"])
        worksheet.write('B68', int(data["dep_senz_temp"])* int(data['dep_locations']))
        worksheet.write('C68', products["mt"]["temp"]["gpl_product"])
        worksheet.write('D68', 65)
        worksheet.write('E68', '=B68*C68*D68/100')
            
        worksheet.write('A69', products["mt"]["temp"]["lic"])
        worksheet.write('B69', int(data["dep_senz_temp"])* int(data['dep_locations']))
        worksheet.write('C69', products["mt"]["temp"]["gpl_license"])
        worksheet.write('D69', 65)
        worksheet.write('E69', '=B69*C69*D69/100')
            
    if data["dep_senz_hum"] !=0 and data["dep_senz_hum"] != '':
        worksheet.write('A70', 'Senzori Umiditate')
        worksheet.write('A71', products["mt"]["hum"]["pn"])
        worksheet.write('B71', int(data["dep_senz_hum"])* int(data['dep_locations']))
        
        worksheet.write('C71', products["mt"]["hum"]["gpl_product"])
        worksheet.write('D71', 65)
        worksheet.write('E71', '=B71*C71*D71/100')
            
        worksheet.write('A72', products["mt"]["hum"]["lic"])
        worksheet.write('B72', int(data["dep_senz_hum"])* int(data['dep_locations']))
        worksheet.write('C72', products["mt"]["hum"]["gpl_license"])
        worksheet.write('D72', 65)
        worksheet.write('E72', '=B72*C72*D72/100')
    

    #######################################
    #######################################
    ## INFORMATII 4G Locatii #######################################
    worksheet.write('A73', 'Informatii specifice locatii 4G x'+str(data["4g_locations"]))
    
    if data["4g_locations"] != 0 and data["4g_locations"] != '':
        worksheet.write('A74', 'MX cu LTE')
        worksheet.write('A75', products["mx"]["4g"]["pn"])
        worksheet.write('B75', int(data["4g_locations"]))
        worksheet.write('C75', products["mx"]["4g"]["gpl_product"])
        worksheet.write('D75', 65)
        worksheet.write('E75', '=B75*C75*D75/100')
            
        worksheet.write('A76', products["mx"]["4g"]["lic"])
        worksheet.write('B76', int(data["4g_locations"]))
        worksheet.write('C76', products["mx"]["4g"]["gpl_license"])
        worksheet.write('D76', 65)
        worksheet.write('E76', '=B76*C76*D76/100')
    

    worksheet.write('E79', 'TOTAL USD')
    worksheet.write('E80', '=SUM(E4:E76)')

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
    worksheet.write('F81', 'Curs '+str(float(float(euro)/float(usd))))
    worksheet.write('E81', 'TOTAL EURO')
    worksheet.write('E82', '=SUM(E4:E76)*'+str(float(euro)/float(usd)))
    workbook.close()
    return None


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=80)
