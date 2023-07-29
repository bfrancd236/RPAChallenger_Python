from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd



def scriptJs(driver, data):
    js_code = """
        // Clique no botão para adicionar um novo registro
        document.querySelector("[class='waves-effect col s12 m12 l12 btn-large uiColorButton']").click();
        
        var rows = arguments[0];
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            document.querySelector("[ng-reflect-name='labelFirstName']").value = row['First Name'];
            document.querySelector("[ng-reflect-name='labelLastName']").value = row['Last Name '];
            document.querySelector("[ng-reflect-name='labelCompanyName']").value = row['Company Name'];
            document.querySelector("[ng-reflect-name='labelRole']").value = row['Role in Company'];
            document.querySelector("[ng-reflect-name='labelAddress']").value = row['Address'];
            document.querySelector("[ng-reflect-name='labelEmail']").value = row['Email'];
            document.querySelector("[ng-reflect-name='labelPhone']").value = row['Phone Number'];
            document.querySelector("[value='Submit']").click();
        }
    """

    driver.execute_script(js_code, data)

def main():
    #webDriver init
    driver = webdriver.Chrome()

    try:
        # Abra a página do RPA Challenge
        driver.get('https://rpachallenge.com/')

        #read excel file
        data = pd.read_excel('challenge.xlsx')
        print(data)
        
        #preencher dados
        scriptJs(driver, data.to_dict('records'))

        # Aguarde alguns segundos para verificar o time
        time.sleep(5)
    finally:
        driver.quit()


main()