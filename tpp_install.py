import os
from datetime import datetime as dt

import wmi
import ctypes
from bs4 import BeautifulSoup
import requests
import win32serviceutil
from pywinauto import Application

tpp_log = r'C:\tpp_inst_log.txt'
tmp_dir = r'C:\tmp'
xml_schema = r'C:\TPP_ANSWER_FILE.xml'

tpp_config_path = r'C:\Program Files\Venafi\Platform'
portal_config_path = r'C:\Program Files\Venafi\User Portal'

portal_name = "Venafi User Portal"
tpp_name = "Venafi Trust Protection Platform"

tpp_config_util = 'TppConfiguration.exe'
portal_config_util = 'PortalSetup.exe'

prod_url = "https://files.prod.ca.eng.venafi.com/builds-prod/TPP_19.1_Build_Prod_W2012/"
dev_url_jaguar = 'https://files.prod.ca.eng.venafi.com/builds-dev/TPP_19.1_Build_Jaguar_W2012/'

dev_url_feature = 'https://files.prod.ca.eng.venafi.com/builds-dev/TPP_19.1_Build_Feature_W2012/'
dev_url_dev = 'https://files.prod.ca.eng.venafi.com/builds-dev/TPP_19.1_Build_Dev_W2012/'


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
    with open(r'{}\{}.msi'.format(tmp_dir, name), 'wb') as fd:
        for chunk in r.iter_content(2000):
            fd.write(chunk)


def delete_tmp_files():
    try:
        for x in os.listdir(tmp_dir):
            os.remove(os.path.join(tmp_dir, x))
    except OSError:
        pass


def install_build(name, f):
    f.write("Installing build...\n")
    os.system('msiexec /i %s /qn' % r'{}\{}.msi'.format(tmp_dir, name))
    delete_tmp_files()


def configure_tpp(f):
    f.write("Configuring tpp...\n")
    os.chdir(tpp_config_path)
    status_code = os.system(
        r'{} -add -install:{}'.format(tpp_config_util, xml_schema))
    if status_code != 0 and status_code != 5:
        raise Exception(
            "Tpp configuration failed with error level: {}".format(status_code)
        )


def configure_portal(f):
    f.write("Configuring portal...\n")
    os.chdir(portal_config_path)
    app = Application(backend='uia').start(portal_config_util)
    app.dlg.child_window(auto_id="rdoLocal").select()
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
            return True


def exec_update(name, url):
    time = dt.now()
    with open(tpp_log, 'w') as f:
        print('{}\n'.format(time.strftime('%d-%m-%Y-%H:%M')))
        f.write('{}\n'.format(time.strftime('%d-%m-%Y-%H:%M')))
        if not already_installed(name):
            download_latest_build('VenafiTPPInstallx64', url, f)
            install_build('VenafiTPPInstallx64', f)
            download_latest_build('UserPortalInstallx64', url, f)
            install_build('UserPortalInstallx64', f)
            configure_tpp(f)
            start_tpp_services(f)
            configure_portal(f)
            restart_iis(f)
            f.write('TPP successfully installed!')
            ctypes.windll.user32.MessageBoxW(0, "TPP Installed!", "TPP", 0)
        else:
            f.write("First uninstall previous version!\n")


if __name__ == "__main__":
    exec_update(tpp_name, dev_url_jaguar)
