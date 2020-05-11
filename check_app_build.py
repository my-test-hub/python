import xlrd
import datetime
import requests
import xml.etree.ElementTree as ET
import os
from urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)
base_url = "https://xld.com/deployit"
http_headers = {"Authorization": "Basic xxxxxxxkeyxxxxx=", "Content-Type": "application/xml"}
# source file location in full path
input_excel = r'\\fullfilepath\filename.xlsx'
#target folder for output file
output_folder = r'C:\folderpath'

'''reads excel file to get app & build names'''

def read_excel(path):
    wb = xlrd.open_workbook(path)   # open file
    sheet = wb.sheet_by_index(0)    #open first sheet
    app_build = []
    for x in range(1,sheet.nrows):
        app_name = sheet.cell_value(x,0)
        build_name = sheet.cell_value(x,1)
        temp_dcit = {'app': app_name, 'build': build_name}
        app_build.append(temp_dcit)
    return app_build

'''gets list of applications from xl-deploy'''

def get_applications():
    end_point = base_url + "/repository/query?Type=core.Directory&parent=Application&resultsPerPage=-1"
    application_response = requests.get(end_point, headers=http_headers, verify=False)
    if application_response.status_code != 200:
        return "Failure"
    else:
        apps_ci_list = ET.fromstring(application_response.text)
        app_output_list = []
        for child in apps_ci_list:
            temp_app_dict = child.attrib
            app = str(temp_app_dict['ref'])
            app_output_list.append(app.split("/")[1])
        return app_output_list

'''check if build and deployables are available'''

def get_build_details(app_name, build_name):
    end_point = f'{base_url}/repository/ci/Applications/{app_name}/{app_name}/{build_name}'
    build_response = requests.get(end_point, headers=http_headers, verify=False)
    if build_response.status_code != 200:
        print(build_response.text)
        return "Failure"
    else:
        pkg_ci = ET.fromstring(build_response.text)
        deployables = pkg_ci.find("deployables")
        if len(deployables) > 0:
            for ci in deployables:
                end_point = base_url + "/repository/ci" + ci.attrib["ref"]
                pkg_response = requests.get(end_point, headers=http_headers, verify=False)
                if pkg_response.status_code != 200:
                    return False
                pkg_dtls = ET.fromstring(pkg_response.text)
                if pkg_dtls.find("fileUri") is not None:
                    return True
        else:
            print('No deployables', build_name)
            return False

'''return list of build deployed env'''

def get_deploy_report(app_name, build_name):
    today_date = datetime.datetime.now()
    three_months_ago = datetime.datetime.utcnow() - datetime.timedelta(days=90)
    start_date = three_months_ago.strftime("%Y-%m-%d")
    end_date_time = today_date.strftime("%Y-%m-%d") + "T23%3A59%3A59.000%2B000"
    request_body = '[{"id":"' + app_name + '/' + app_name + '","type":"udm.Application"}]'
    end_point = f'{base_url}/internal/reports/tasks?begin={start_date}T00%3A00%3A00.000%2B000' \
                f'end={end_date_time}&filterType=application&states=DONE&order=startDate%3Adesc'
    http_headers = {"Authorization": "Basic xxxxxxxkeyxxxxx=", "Content-Type": "application/json"}
    temp_response = requests.post(end_point, headers=http_headers, verify=False, data=request_body)
    if temp_response.status_code != 200:
        print("Error: Error while fetching report for app: {}".format(app_name))
        return "Failure"
    #proceed to process for a successful call
    report_entries = ET.fromstring(temp_response.text)
    if len(report_entries) == 0:
        print("Info: No deployments reported for app: {}".format(app_name))
        return "Failure"
    env_list = ""
    temp_env = {}
    for values in report_entries:
        for lines in values:
            line_dict = {}
            for line in lines:
                if line.find("key").text == "environment" or line.find("key").text == "package":
                    line_dict.update({line.find('key').text: line.find('value').text})
            temp_list = [line_dict]
            for x in temp_list:
                temp_package = str(x['package'])
                build_tmp = temp_package.split("/")[1]
                if build_tmp == build_name:
                    temp_env.update({x['environment']: x['environment']})
    report_entries = ""
    return ' -- '.join(temp_env)

'''calls all above fn and writes ouyput to csv file'''

def check_app_build(file):
    app_build = read_excel(file)
    xld_apps = get_applications()
    date_time = datetime.datetime.now()
    file_name = "xld_report_" + str(date_time.strftime("%Y%m%d")) + "-" + str(date_time.strftime("%H%M%S")) + ".csv"
    file_path = os.path.join(output_folder, file_name)
    f = open(file_path, "a")
    f.write("App Name,App Name Exists,Build Name,Build Exists,Env Deployed,Deployed in Accp \n")
    for x in app_build:
        if x['app'] in xld_apps:
            check_build = get_build_details(x['app'], x['build'])
            if check_build is True:
                env = get_deploy_report(x['app'], x['build'])
                if env.lower().__contains__('accp'):
                    deployed_in_accp = "Yes"
                else:
                    deployed_in_accp = "No"
                f.write(
                    x['app'] + "," + 'Yes' + "," + x['build'] + "," + 'Yes' + "," + env + "," + deployed_in_accp + "\n")
            else:
                f.write(x['app'] + "," + 'Yes' + "," + x['build'] + "," + 'No' + "," + "N/A" + "," + "N/A" + "\n")
        else:
            f.write(x['app'] + "," + 'No' + "," + "N/A" + "," + 'No' + "," + "N/A" + "," + "N/A" + "\n")
    f.close()

#check_app_build(input_excel)
