#   The script is designed to update/reinstall tpp/user portal to the newer version
#
#   Steps to execute:
#   Choose branch between PROD_URL, DEV_URL_JAGUAR, DEV_URL_FEATURE, DEV_URL_DEV and pass it 
#   to starter() as argument. The script will automatically uninstall and then install newer version.
#   If reboot has happend you need to run this script again. Or you can make automatic execution 
#   of this script at windows startup by create a windows schedule. Portal is set up on the same
#   server as tpp.
#   
#   ! YOU NEED TO HAVE PROPER ANSWER TPP XML(SCHEMA) FILE !
#   ! SET IN AS XML_SCHEMA CONST !
#
#   libraries needed : 
#       beautifulsoup4==4.6.3
#       pywin32==224
#       pywinauto==0.6.5
#       requests==2.21.0
#       urllib3==1.24.1
#       WMI==1.4.9
#

import os
from datetime import datetime as dt

import wmi
import ctypes
from bs4 import BeautifulSoup
import requests
import win32serviceutil
from pywinauto import Application


TPP_LOG = r'C:\tpp_inst_log.txt'
TMP_DIR = r'C:\tmp'
XML_SCHEMA = r'C:\TPP_ANSWER_FILE.xml'

TPP_CONFIG_PATH = r'C:\Program Files\Venafi\Platform'
PORTAL_CONFIG_PATH = r'C:\Program Files\Venafi\User Portal'

TPP_BUILD_NAME = 'VenafiTPPInstallx64'
PORTAL_BUILD_NAME = 'UserPortalInstallx64'

PORTAL_NAME = "Venafi User Portal"
TPP_NAME = "Venafi Trust Protection Platform"

TPP_CONFIG_UTIL = 'TppConfiguration.exe'
PORTAL_CONFIG_UTIL = 'PortalSetup.exe'

# Branches
PROD_URL = "https://files.prod.ca.eng.venafi.com/builds-prod/TPP_19.1_Build_Prod_W2012/"
DEV_URL_JAGUAR = 'https://files.prod.ca.eng.venafi.com/builds-dev/TPP_19.1_Build_Jaguar_W2012/'

DEV_URL_FEATURE = 'https://files.prod.ca.eng.venafi.com/builds-dev/TPP_19.1_Build_Feature_W2012/'
DEV_URL_DEV = 'https://files.prod.ca.eng.venafi.com/builds-dev/TPP_19.1_Build_Dev_W2012/'


def get_latest_build_folder(build_url, f):
    f.write("Getiing latest build...\n")
    hreflist = []
    r = requests.get(build_url)
    soup = BeautifulSoup(r.text, 'html.parser')
    index = soup.findAll("td", {"class": "indexcolname"})
    for link in index:
        a = link.find('a')
        hreflist.append(a.get('href'))
    f.write("Build url: {}\n".format(build_url))
    f.write("Latest build folder: {}\n".format(hreflist[-1]))
    return hreflist[-1]


def download_latest_build(name, build_url, f):
    f.write("Downloading latest build...\n")
    lbuild = get_latest_build_folder(build_url, f)
    r = requests.get("{}{}/Product/{}.msi".format(
        build_url, lbuild, name), stream=True
    )
    try:
        os.mkdir(TMP_DIR)
        print("Directory " , TMP_DIR ,  " Created ") 
    except OSError:
        print("Directory " , TMP_DIR ,  " already exists")
	with open(r'{}\{}.msi'.format(TMP_DIR, name), 'wb') as fd:
        for chunk in r.iter_content(2000):
            fd.write(chunk)


def delete_tmp_files():
    try:
        for x in os.listdir(TMP_DIR):
            os.remove(os.path.join(TMP_DIR, x))
    except OSError:
        pass


def install_build(name, f):
    f.write("Installing build...\n")
    os.system('msiexec /i %s /qn' % r'{}\{}.msi'.format(TMP_DIR, name))
    delete_tmp_files()


def configure_tpp(f):
    f.write("Configuring tpp...\n")
    os.chdir(TPP_CONFIG_PATH)
    status_code = os.system(
        r'{} -add -install:{}'.format(TPP_CONFIG_UTIL, XML_SCHEMA))
    if status_code != 0 and status_code != 5:
        raise Exception(
            "Tpp configuration failed with error level: {}".format(status_code)
        )


def configure_portal(f):
    f.write("Configuring portal...\n")
    os.chdir(PORTAL_CONFIG_PATH)
    app = Application(backend='uia').start(PORTAL_CONFIG_UTIL)
    app.dlg.child_window(title="Aperture is already installed on this server.", auto_id="rdoLocal").select()
    # app.dlg.child_window(auto_id="rdoLocal").select()
    app.dlg.Install.click()


def start_tpp_services(f):
    f.write("starting tpp services...\n")
    log_service = "VenafiLogServer"
    tpp_service = "VED"
    win32serviceutil.RestartService(tpp_service)
    win32serviceutil.RestartService(log_service)


def restart_iis(f):
    f.write("restarting iis...\n")
    os.system('iisreset.exe')


def already_installed(name):
    w = wmi.WMI()
    for app in w.Win32_Product():
        if app.Name.startswith(name):
            return app.Name


def uninstall(app, f):
    w = wmi.WMI()
    for product in w.Win32_Product(Name=app):
        f.write("Uninstalling " + app + "...\n")
        product.Uninstall()
        f.write("The app uninstalled\n")


def need_uninstall(name, f):
    n = already_installed(name)
    if n:
        uninstall(n, f)


def exec_tpp_update(name, url, f):
    if not already_installed(name):
        download_latest_build(TPP_BUILD_NAME, url, f)
        install_build(TPP_BUILD_NAME, f)
        configure_tpp(f)
        start_tpp_services(f)
        f.write('TPP successfully installed!\n')
        # ctypes.windll.user32.MessageBoxW(0, "TPP Installed!", "TPP", 0)
    else:
        f.write("First uninstall previous tpp version!\n")


def exec_portal_update(name, url, f):
    if not already_installed(name):
        download_latest_build(PORTAL_BUILD_NAME, url, f)
        install_build(PORTAL_BUILD_NAME, f)
        configure_portal(f)
        f.write('Portal successfully installed!\n')
        # ctypes.windll.user32.MessageBoxW(0, "Portal Installed!", "Portal", 0)
    else:
        f.write("First uninstall previous portal version!\n")


def starter(branch_name):
    time = dt.now()
    with open(TPP_LOG, 'w') as f:
        f.write('{}\n'.format(time.strftime('%d-%m-%Y-%H:%M')))
        try:
            need_uninstall(PORTAL_NAME, f)
            need_uninstall(TPP_NAME, f)
            exec_tpp_update(TPP_NAME, branch_name, f)
            exec_portal_update(PORTAL_NAME, branch_name, f)
            restart_iis(f)
        except Exception as e:
            f.write('During commands execution the exception occur: {}\n'.format(e))

if __name__ == "__main__":
    # Function argument is a repo's branch name
    starter(DEV_URL_JAGUAR)
