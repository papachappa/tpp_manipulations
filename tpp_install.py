import os

from bs4 import BeautifulSoup
import requests
import win32serviceutil
from pywinauto import Application


tmp_dir = r'C:\tmp'
xml_schema = r'C:\TPP_ANSWER_FILE.xml'

tpp_config_path = r'C:\Program Files\Venafi\Platform'
portal_config_path = r'C:\Program Files\Venafi\User Portal'

tpp_config_util = 'TppConfiguration.exe'
portal_config_util = 'PortalSetup.exe'

prod_url = "https://files.prod.ca.eng.venafi.com/builds-prod/TPP_19.1_Build_Prod_W2012/"
dev_url_jaguar = 'https://files.prod.ca.eng.venafi.com/builds-dev/TPP_19.1_Build_Jaguar_W2012/'

dev_url_feature = 'https://files.prod.ca.eng.venafi.com/builds-dev/TPP_19.1_Build_Feature_W2012/'
dev_url_dev = 'https://files.prod.ca.eng.venafi.com/builds-dev/TPP_19.1_Build_Dev_W2012/'


def get_latest_build_folder(build_url):
    print("Getiing latest build...")
    hreflist = []
    r = requests.get(build_url)
    soup = BeautifulSoup(r.text, 'html.parser')
    index = soup.findAll("td", {"class": "indexcolname"})
    for link in index:
        a = link.find('a')
        hreflist.append(a.get('href'))
    print("Build url: {}".format(build_url))
    print("Latest build folder: {}".format(hreflist[-1]))
    return hreflist[-1]


def download_latest_build(name, build_url):
    print("Downloading latest build...")
    lbuild = get_latest_build_folder(build_url)
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


def install_build(name):
    print("Installing build...")
    os.system('msiexec /i %s /qn' % r'{}\{}.msi'.format(tmp_dir, name))
    delete_tmp_files()


def configure_tpp():
    print("Configuring tpp...")
    os.chdir(tpp_config_path)
    status_code = os.system(
        r'{} -add -install:{}'.format(tpp_config_util, xml_schema))
    if status_code != 0 and status_code != 5:
        raise Exception(
            "Tpp configuration failed with error level: {}".format(status_code)
        )


def configure_portal():
    print("Configuring portal...")
    os.chdir(portal_config_path)
    app = Application(backend='uia').start(portal_config_util)
    app.dlg.child_window(auto_id="rdoLocal").select()
    app.dlg.Install.click()


def start_tpp_services():
    print("starting tpp services...")
    log_service = "VenafiLogServer"
    tpp_service = "VED"
    win32serviceutil.RestartService(tpp_service)
    win32serviceutil.RestartService(log_service)


def restart_iis():
    print("restarting iis...")
    os.system('iisreset.exe')


if __name__ == "__main__":
    download_latest_build('VenafiTPPInstallx64', dev_url_jaguar)
    install_build('VenafiTPPInstallx64')
    download_latest_build('UserPortalInstallx64', dev_url_jaguar)
    install_build('UserPortalInstallx64')
    configure_tpp()
    start_tpp_services()
    configure_portal()
    restart_iis()
