from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By
from pathlib import Path
from datetime import datetime, timedelta
import csv, json

# get base root
BASE_DIR=Path(__file__).resolve().parent
MAX_WAIT_TIME=60
URL='https://login.xero.com/identity/user/login'

# load login info
with open(BASE_DIR / 'data/credential.json','r') as rf:
    cred = json.load(rf)
    USERNAME=cred['username']
    PASSWORD=cred['password']

with open(BASE_DIR / 'data/entities.json','r') as rf:
    entity_dict = json.load(rf)

# fetch xero website contents
all_entities=['placeholder',*[x for x in entity_dict.keys()]]

# initialize webdriver
def create_driver(root_folder):
    driver_path=root_folder / 'drivers/chromedriver.exe'
    service = Service(executable_path=driver_path)
    options = webdriver.ChromeOptions()
    driver=webdriver.Chrome(service=service, options=options)
    return driver

# define get element with wait func
def wait_til_get_elem(driver,elem_id,max_wait,mode='id'):
    wait = WebDriverWait(driver, max_wait)
    if mode == 'id':
        element = wait.until(expected_conditions.presence_of_all_elements_located((By.ID, elem_id)))
    elif mode == 'class':
        element = wait.until(expected_conditions.presence_of_all_elements_located((By.CLASS_NAME, elem_id)))

    return element

def login(driver,username,password):
    # type in username and password then login
    
    email=wait_til_get_elem(driver,'xl-form-email',MAX_WAIT_TIME)[0]
    pw=wait_til_get_elem(driver,'xl-form-password',MAX_WAIT_TIME)[0]
    login=wait_til_get_elem(driver,'xl-form-submit',MAX_WAIT_TIME)[0]

    email.clear()
    email.send_keys(username)
    pw.send_keys(password)
    login.click()



if __name__=='__main__':
    # initialize container
    container=[['farm','date','item','value','link']]

    # initialize driver
    driver=create_driver(BASE_DIR)

    # get today's date
    reference_day=datetime.today()-timedelta(days=1)
    today=reference_day.strftime('%Y-%m-%d')
    today_date=reference_day.strftime('(%b %d)')
    print(f"program running...today is {today}, keeping only records on {today_date}")

    for i,entity in enumerate(all_entities):
        print(container[-1])
        if i==0:
            # open website
            driver.get(URL)
            login(driver,USERNAME,PASSWORD)
        else:
            DASHBOARD_URL=f'https://go.xero.com/app/{entity}/dashboard'
            driver.get(DASHBOARD_URL)

        link=driver.current_url
        farm=wait_til_get_elem(driver,'xui-pageheading--title',MAX_WAIT_TIME,'class')[0].text.replace('\n','')
        statement_bal=wait_til_get_elem(driver,'bankWidget-balanceTable__summary--label--N4zpS',MAX_WAIT_TIME,'class')
        # statement balance are all elements with class name: bankWidget-balanceTable__summary--label--N4zpS
        for i,found in enumerate(statement_bal):
            item=found.text.strip()
            if item.lower().startswith('statement balance') and item.endswith(today_date):
                val=float(statement_bal[i+1].text.replace(',',''))
                # save to container
                container.append([farm,today,item,val,link])
    print(container[-1])

    print('All data retrieved, now saving to csv...')

    # write to csv
    fpath=BASE_DIR / f'statement_balance.csv'
    with open(fpath,mode='w',newline='',encoding='utf-8') as rf:
        writer=csv.writer(rf)
        writer.writerows(container)

    print(f'Data saved to {fpath}')

# pyinstaller --icon="C:\Users\RancoXu\OneDrive - Argyle Capital Partners Pty Ltd\Desktop\Ranco\Python\xerobot\image\robot.png" --noconfirm "C:\Users\RancoXu\OneDrive - Argyle Capital Partners Pty Ltd\Desktop\Ranco\Python\xerobot\xerobot.py" --paths "C:\Users\RancoXu\OneDrive - Argyle Capital Partners Pty Ltd\Desktop\Ranco\Python\xerobot\venv_xerobot\Lib\site-packages"  --add-data 'drivers\chromedriver.exe;.\drivers' --add-data 'data\credential.json;.\data' --add-data 'data\save_path.json;.\data' --add-data 'data\entities.json;.\data' --add-data 'requirement.txt;.' --add-data 'xero_bank_reconciler.py;.' --add-data 'image\argyle.jpg;.\image' --add-data 'image\folder.png;.\image' --add-data 'image\robot.png;.\image'
