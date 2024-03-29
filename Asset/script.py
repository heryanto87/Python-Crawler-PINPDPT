from math import atanh
from operator import truediv
import os
from unicodedata import name
import pyodbc
import pandas as pd
from openpyxl import load_workbook
from os import system
from time import sleep
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

import schedule
import configparser
import pathlib
import keyboard

# Headless Crawler Settings
chrome_options = Options()
# chrome_options.headless = True
# chrome_options.add_argument("--headless")
chrome_options.add_argument('log-level=3')
chrome_options.add_argument('window-size=1920,1080')
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--no-sandbox");
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument(
    r"D:\PROJECT-PIN\Asset\Download"
)
prefs = {'download.default_directory':
         r'D:\PROJECT-PIN\Asset\Download'}
chrome_options.add_experimental_option('prefs', prefs)
install = ChromeDriverManager().install()
print(install)
driver = webdriver.Chrome(service=Service(install), options=chrome_options)
driver.set_window_size(1920, 1080)
driver.get('https://pin.kemdikbud.go.id/pin/index.php/login')

# driver.get('http://103.56.190.37/pin/')

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=edm-pdpt.binus.db;'
                      'Database=PDPT_POSTING;'
                      'UID=app_pin;'
                      'PWD=B!4PpP!nNU$;'
                      'Trusted_Connection=no;')


'''
####################################################################################################
# Modul Function untuk mempermudah pemanggilan fungsi yang dipakai berulang
####################################################################################################
'''


def ClickXPATH(xPATH, wait = 60):
    try:
        button = WebDriverWait(driver, wait).until(
            lambda driver: driver.find_element("xpath", xPATH))
        driver.execute_script("arguments[0].scrollIntoView(true);", button)
        button.click()
        return True
    except Exception as err:
        return False


def SendXPATH(xPATH, text, wait = 60):
    try:
        element = WebDriverWait(driver, wait).until(
            lambda driver: driver.find_element("xpath", xPATH))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        element.clear()
        element.send_keys(text)
        return True
    except Exception as err:
        return False


def SelectCSS(cssSelector, wait = 60):
    element = WebDriverWait(driver, wait).until(
        lambda driver: driver.find_elements("css selector", cssSelector))
    return element


def GetXPATHElement(xPATH, wait = 60):
    element = WebDriverWait(driver, wait).until(
        lambda driver: driver.find_element("xpath", xPATH)
    )

    return element

def GetXPATHElements(xPATH):    
    elements = driver.find_elements("xpath", xPATH)
    return elements

def GetCSSElements(cssSelector):
    elements = driver.find_elements("css selector", cssSelector)
    return elements

def Homepage():
    # driver.get('http://103.56.190.37/pin/index.php/home/batal')
    driver.get('https://pin.kemdikbud.go.id/pin/index.php/home/batal')

def Login():
    user = "031038"
    password = "Binus2010"
    usernameXPATH = "//input[@placeholder='Masukan Username Anda']"
    passwordXPATH = "//input[@placeholder='Masukan Password Anda']"
    buttonXPATH = "//button[@class='btn-login btn-primary-login block-login full-width-login m-b']"

    try:
        GetXPATHElement(usernameXPATH)
        SendXPATH(usernameXPATH, user)
        SendXPATH(passwordXPATH, password)

        if ClickXPATH(buttonXPATH) is True:
            print("[INFO]: Login Success")
        else:
            raise Exception("[INFO]: Login Failed")
    except TimeoutException:
        print("[INFO] Already logged in")
        pass 


'''
# Print iterations progress
'''


def ProgressBar(iteration, total, prefix='', suffix='', decimals=1, length=100, fill='█', printEnd="\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end=printEnd)
    # Print New Line on Complete
    if iteration == total:
        print()


'''
####################################################################################################
# End of Modul
####################################################################################################
'''


'''
# Mereset reservasi
'''


def Reset(skipEnterKey = False):
    # Menuju list data Arsip Reservasi
    driver.get("https://pin.kemdikbud.go.id/pin/index.php/history")

    # Mengambil data jumlah reservasi di list table
    print("[INFO]: Count Reserved Program")
    view_XPATH = "//select[@name='DataTables_Table_0_length']/option[text()='100']"
    ClickXPATH(view_XPATH)
    search_XPATH = "//input[contains(@placeholder,'Ketik Kata Kunci')]"
    SendXPATH(search_XPATH, "BELUM")
    button_length = 0

    # Memulai progress Reset pada reservasi
    try:
        button_css = ".btn.btn-danger"
        buttons = SelectCSS(button_css, 5)
        button_length = len(buttons)
        print("[INFO]: Reset is on progress")

        ProgressBar(0, button_length, prefix='Progress:', suffix='Complete', length=50)
        i = 0

        while True:
            sleep(1)
            if i <= button_length:
                ProgressBar(i + 1, button_length, prefix='Progress:', suffix='Complete', length=50)
                i += 1

            view_XPATH = "//select[@name='DataTables_Table_0_length']/option[text()='100']"
            ClickXPATH(view_XPATH)
            search_XPATH = "//input[contains(@placeholder,'Ketik Kata Kunci')]"
            SendXPATH(search_XPATH, "BELUM")
            buttons_left = SelectCSS(button_css)

            if len(buttons_left) < 1:
                break

            reset_XPATH = "//button[contains(@class,'btn btn-danger')]"
            while ClickXPATH(reset_XPATH, 5) is True:
                continue

    except TimeoutException:
        ProgressBar(1, 1, prefix='Progress:', suffix='Complete', length=50)
        print("No Data...")
        if not skipEnterKey:
            input("Press Enter to Continue...")


'''
# Update table dan database dari hasil crawling data
'''


def Update(grad_year = None, wisuda = None):
    # Pindah ke page reservasi
    # driver.get("https://pin.kemdikbud.go.id/pin/index.php/prodi")
    reservation_xpath = "/html/body/div[1]/div/div[3]/div/div[2]/div[1]/form/button"
    ClickXPATH(reservation_xpath)

    # Input tahun wisuda
    if grad_year is None:
        grad_year = input("[INPUT] Graduation Year : ")

    # Deklarasi variable yang diperlukan
    FINAL_NOT_EG_TABLE = "NOT_EG_PIN"
    FINAL_HNOT_EG_TABLE = "H_NOT_EG_PIN"
    FINAL_PIN_TABLE = "Tbl_PIN_Mahasiswa_Lulusan"
    FINAL_NINA_TABLE = "Tbl_PIN_Nomor_Ijazah_Lulusan"
    FINAL_HPIN_TABLE = "HTbl_PIN_Mahasiswa_Lulusan"
    FINAL_HNINA_TABLE = "HTbl_PIN_Nomor_Ijazah_Lulusan"
    SP_BACKUP = "SPBackUpTablePIN"

    if wisuda is None:
        wisuda = input("[INPUT] Wisuda : ")

    # Deklarasi cursor ke database dan tablenya
    cursor = conn.cursor()     
    print("[INFO] Clearing tables")
    cursor.execute("TRUNCATE TABLE " + FINAL_NOT_EG_TABLE)
    cursor.execute("TRUNCATE TABLE " + FINAL_PIN_TABLE)
    cursor.execute("TRUNCATE TABLE " + FINAL_NINA_TABLE)

    # Deklarasi XPath untuk operasi yang akan dilakukan dengan function yang sesuai
    print("[INFO] Preparing Table")
    viewXPath = "//select[@name='DataTables_Table_0_length']/option[text()='100']"
    chooseTableCSS = ".btn.btn-xs.btn-block.btn-primary"
    ClickXPATH(viewXPath)
    sleep(2)
    button = SelectCSS(chooseTableCSS)
    buttonLength = len(button)

    # Deklarasi Variable tambahan untuk adaptasi crawler dengan web PIN
    i = 0
    skipIdx = 0
    jumlahNotEg = 0
    jumlahMhs = 0
    jumlahNina = 0

    # Operasi berdasarkan flow bisnis yang dikerjakan secara manual
    while i != buttonLength:
        # titleXPath = '//*[@id="page-wrapper"]/div[3]/div[1]/div[2]/div/div/div/div[1]/h2/strong'
        # title = GetXPATHElement(titleXPath).text
        viewXPath = "//select[@name='DataTables_Table_0_length']/option[text()='100']"
        if ClickXPATH(viewXPath) == False:
            driver.get("https://pin.kemdikbud.go.id/pin/index.php/prodi")
            ClickXPATH(backXPath)
            continue

        sleep(2)

        kodeProdiXPath = "/html/body/div[1]/div/div[3]/div[1]/div[2]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr[" + str(
            i + 1) + "]/td[2]"
        namaProdiXPath = "/html/body/div[1]/div/div[3]/div[1]/div[2]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr[" + str(
            i + 1) + "]/td[3]"
        backXPath = '//*[@id="page-wrapper"]/div[3]/div[1]/div[1]/div/a[1]'
        kodeProdi = GetXPATHElement(kodeProdiXPath).text
        namaProdi = GetXPATHElement(namaProdiXPath).text
        print("")
        print("[INFO] Prodi = ", end=" ")
        print(i + 1, end=". ")
        print(kodeProdi, end=" - ")
        print(namaProdi)

    
        # 1
        sleep(2)
        try:
            operationButton = SelectCSS(chooseTableCSS) 
            operationButton[skipIdx].click()
        except TimeoutException:
            print(f"[WARNING] Could not find (step 1): {chooseTableCSS}")
            ClickXPATH(backXPath)            
            continue

        # 2
        sleep(1.5)
        try:
            graduationXPATH = "//input[@type='number']"
            SendXPATH(graduationXPATH, grad_year)
            if graduationXPATH:
                print(f"[INFO] graduationXPATH OK")
            # pilihXPath = "//form[@class='form-inline text-center']//button[@class='btn btn-primary'][contains(text(),'Submit')]"
            pilihXPath = '//*[@id="page-wrapper"]/div[3]/div/div[2]/div/div/div/div[2]/form/button'
            ClickXPATH(pilihXPath)
            if pilihXPath:
                print(f"[INFO] pilihXPath OK")
        except TimeoutException:
            print(f"[WARNING] Could not find (step 2): {pilihXPath}")
            ClickXPATH(backXPath)
            continue

        # 3
        # NOT EG
        sleep(1.5)
        print("[INFO] Start Processing Not Eligible Data")
        notEligibleData = ""
        notEligibleTableNextXPATH = "//li[@id='DataTables_Table_0_next']//a[contains(text(),'Selanjutnya')]"
        notEligibleDisableNextXPath = "//li[@class='paginate_button next disabled' and @id='DataTables_Table_0_next']"
        tBodyXPATH = "//table[@id='DataTables_Table_0']//tbody"

        try:
            notEligibleTableNextButton = GetXPATHElement(notEligibleTableNextXPATH)
        except TimeoutException as err:
            print(f"[WARNING] Could not find (step 3): {notEligibleTableNextXPATH}")
            ClickXPATH(backXPath)
            continue

        sleep(1)
        while True:
            tBodyElement = GetXPATHElement(tBodyXPATH)
            notEligibleData += tBodyElement.get_attribute("innerHTML")

            if len(GetXPATHElements(notEligibleDisableNextXPath)) == 0:
                notEligibleTableNextButton.click()
                notEligibleTableNextButton = GetXPATHElement(notEligibleTableNextXPATH)
            else:
                break

        soup = BeautifulSoup(notEligibleData, "html.parser")
        rows = soup.find_all('tr')

        sleep(1)
        if len(rows) > 1:
            rowLength = len(rows)
            jumlahNotEg = rowLength

            ProgressBar(0, rowLength, prefix='Progress:', suffix='Complete', length=50)
            x = 0

            for row in rows:
                ProgressBar(x + 1, rowLength, prefix='Progress:', suffix='Complete', length=50)
                x += 1
                cells = row.find_all("td")
                nama = cells[1].string
                nim = cells[2].string
                sks = cells[3].string
                ipk = cells[4].string
                alasan = cells[5].string
                tanggalTarik = datetime.today().strftime('%Y-%m-%d')

                cursor.execute("insert into " + FINAL_NOT_EG_TABLE + " (Nama, Nim, [Total SKS], IPK, Alasan, TanggalTarik) values('" + str(
                    nama.replace("'", "''")) + "','" + str(nim) + "','" + str(sks) + "','" + str(ipk) + "','" + str(alasan) + "','" + str(tanggalTarik) + "')")
                conn.commit()

        driver.refresh()

        # 4
        # MAHASISWA
        print("[INFO] Start Processing Mahasiswa Data")
        daftarCalonData = ""
        daftarCalonTableNextXPATH = "//li[@id='example_next']//a[contains(text(),'Selanjutnya')]"
        daftarCalonDisableNextXPath = "//li[@class='paginate_button next disabled' and @id='example_next']"
        tBodyXPATH = "//table[@id='example']//tbody"

        try:
            daftarCalonTableNextButton = GetXPATHElement(daftarCalonTableNextXPATH)
        except TimeoutException as err:
            print(f"[WARNING] Could not find (step 4): {daftarCalonTableNextButton}")
            ClickXPATH(backXPath)
            continue

        while True:
            TBodyElement = GetXPATHElement(tBodyXPATH)
            daftarCalonData += TBodyElement.get_attribute("innerHTML")
            if len(GetXPATHElements(daftarCalonDisableNextXPath)) == 0:
                daftarCalonTableNextButton.click()
                daftarCalonTableNextButton = GetXPATHElement(daftarCalonTableNextXPATH)
            else:
                break

        soup = BeautifulSoup(daftarCalonData, "html.parser")
        rows = soup.find_all('tr')

        if len(rows) > 1:
            rowLength = len(rows)
            jumlahMhs = rowLength

            ProgressBar(0, rowLength, prefix='Progress:', suffix='Complete', length=50)
            x = 0

            for row in rows:
                ProgressBar(x + 1, rowLength, prefix='Progress:', suffix='Complete', length=50)
                x += 1

                cells = row.find_all("td")
                nama = cells[2].string
                nim = cells[3].string
                tanggalTarik = datetime.today().strftime('%Y-%m-%d')

                cursor.execute("insert into " + FINAL_PIN_TABLE + " (NIM, Nama, KodeProdi, Wisuda, TanggalTarik) values('" + str(
                    nim) + "','" + str(nama.replace("'", "''")) + "','" + str(kodeProdi) + "','" + str(wisuda) + "','" + str(tanggalTarik) + "')")
                conn.commit()

        # 5
        # NOMOR IJAZAH
        print("[INFO]: Start Processing Nomor Ijazah Data")
        prosesIjazahXPath = "//input[@class='btn btn-primary btn-rounded text-center']"

        if ClickXPATH(prosesIjazahXPath) is True:
            nomorIjazahData = ""
            nomorIjazahTableNextXPATH = "//li[@id='DataTables_Table_0_next']//a[contains(text(),'Selanjutnya')]"
            nomorIjazahTableNextCSSSelector = ".paginate_button.next.disabled"
            tBodyXPATH = "//table[@id='DataTables_Table_0']//tbody"
            try:
                nomorIjazahTableNextButton = GetXPATHElement(nomorIjazahTableNextXPATH)
            except TimeoutException:
                print(f"[WARNING] Could not find (step 5): {nomorIjazahTableNextXPATH}")
                driver.get("https://pin.kemdikbud.go.id/pin/index.php/prodi")
                continue

            while True:
                TBodyElement = GetXPATHElement(tBodyXPATH)
                nomorIjazahData += TBodyElement.get_attribute("innerHTML")         
                if len(GetCSSElements(nomorIjazahTableNextCSSSelector)) == 0:
                    nomorIjazahTableNextButton.click()
                    nomorIjazahTableNextButton = GetXPATHElement(nomorIjazahTableNextXPATH)
                else:
                    break

            soup = BeautifulSoup(nomorIjazahData, "html.parser")
            rows = soup.find_all('tr')

            if len(rows) > 1:
                rowLength = len(rows)
                jumlahNina = rowLength

                ProgressBar(0, rowLength, prefix='Progress:', suffix='Complete', length=50)
                x = 0

                for row in rows:
                    ProgressBar(x + 1, rowLength, prefix='Progress:', suffix='Complete', length=50)
                    x += 1

                    cells = row.find_all("td")
                    nomorIjazah = cells[1].string

                    cursor.execute("insert into " + FINAL_NINA_TABLE +
                                   " (NomorIjazah, KodeProdi) values('" + str(nomorIjazah) + "','" + str(kodeProdi) + "')")
                    conn.commit()

            try:
                PengajuanNomorIjazahXPATH = "//button[contains(text(),'Akhiri Pengajuan Nomor Ijazah')]"
                ClickXPATH(PengajuanNomorIjazahXPATH)
            except TimeoutException:
                backXPath = "//*[@id='page-wrapper']/div[3]/div/div[1]/div/a[1]"
                ClickXPATH(backXPath)
                continue

            print("[Reserved]")

        else:
            skipIdx += 1
            print("[INFO] No Data For Nomor Ijazah")

        # moving Index Row
        i += 1

        # insert Record
        print(kodeProdi)
        print(namaProdi)
        tanggalTarik = datetime.today().strftime('%Y-%m-%d')
        conn.commit()
        cursor.execute("insert into crawler_record(kodeProdi, namaProdi, jumlahNotEg, jumlahMhs, jumlahNina, tanggalReservasi) values('" + str(kodeProdi) +
                       "','" + str(namaProdi) + "','" + str(jumlahNotEg) + "','" + str(jumlahMhs) + "','" + str(jumlahNina) + "','" + str(tanggalTarik) + "')")

        # commit all queries
        conn.commit()

        # back to reservation point
        backXPath = "//*[@id='page-wrapper']/div[3]/div/div[1]/div/a[1]"
        ClickXPATH(backXPath)

    cursor.execute("insert into " + FINAL_HNINA_TABLE +
                   " select *, TglBackup = GETDATE() from " + FINAL_NINA_TABLE)
    cursor.execute("insert into " + FINAL_HPIN_TABLE +
                   " select *, TglBackup = GETDATE() from " + FINAL_PIN_TABLE)
    cursor.execute("insert into " + FINAL_HNOT_EG_TABLE +
                   " select *, tglbackup = GETDATE() from " + FINAL_NOT_EG_TABLE)

    print("[INFO] Copying table")
    cursor.execute("EXEC " + SP_BACKUP)
    conn.commit()
    cursor.close()
    print("[INFO] Done")


'''
# Gunakan Validator ketika ingin menghapus data yang ada di filter SGGC
'''


def Validator():
    # Deklarasi variable yang diperlukan
    sggcView = "VIEW_SGGC_MAPPING_PELAPORAN_MASTER_TRACK_S2"
    periode = "20202"
    cursor = conn.cursor()

    # Mengambil data dari database
    cursor.execute("select [no] from Tbl_PIN_Mahasiswa_Lulusan where NIM in (select external_system_id from " +
                   sggcView + " where coalesce(periode_mata_kuliah_dilaporkan,'') in ('', '" + periode + "'))")
    dataIdx = cursor.fetchall()

    # Looping validasi dari view SGGC
    for i in dataIdx:
        cursor.execute(
            "select nama, nim from TBL_PIN_Mahasiswa_Lulusan where no = '" + str(i[0]) + "'")
        data = cursor.fetchone()
        print("[INFO] Mahasiswa Dihapus = ", end=" ")
        print(data[0])
        cursor.execute("delete from Tbl_PIN_Mahasiswa_Lulusan where no = '" + str(i[0]) + "'")
        cursor.execute("delete from Tbl_PIN_Nomor_Ijazah_Lulusan where no = '" + str(i[0]) + "'")

    print("Press enter to continue...")

    conn.commit()
    cursor.close()


'''
# Developing function Upload
'''


def Upload():
    # Membuka cursor koneksi
    cursor = conn.cursor()
    cursor.execute("truncate table tbl_nina_dipadankan")
    cursor.execute("truncate table PIN_UPLOAD_RESULT")
    cursor.execute("truncate table PIN_UPLOAD_LOG")
    conn.commit()

    # Membaca data dari excel
    df = pd.read_excel(r'Asset/upload.xlsx')
    acad_list = df["Academic Institution"].tolist()
    grad_list = df["Graduation Batch"].tolist()
    prod_list = df["Prodi Code"].tolist()
    nim_list = df["External System ID"].tolist()
    pin_list = df["National Diploma Number"].tolist()
    tanggal = datetime.today().strftime('%Y-%m-%d')

    # Panjang list
    row_length = len(acad_list)

    # Deklarasi awal untuk progress bar
    ProgressBar(0, row_length, prefix='Progress:', suffix='Complete', length=50)
    x = 0

    for i in range(row_length):
        # Menambah index progress bar
        ProgressBar(x + 1, row_length, prefix='Progress:', suffix='Complete', length=50)
        x += 1

        # Insert query ke tabel
        insert_query = "insert into tbl_nina_dipadankan ([Academic Institution], [Graduation Batch], [Prodi Code], [External System ID], [National Diploma Number], [TgldiPadankan]) values('" + str(
            acad_list[i]) + "','" + str(grad_list[i]) + "','" + str(prod_list[i]) + "','" + str(nim_list[i]) + "','" + str(pin_list[i]) + "','" + str(tanggal) + "')"
        cursor.execute(insert_query)
        conn.commit()

    print("[INFO] Student Has Been Inserted")

    # Proses menjalankan Stored Procedure
    graduation_batch = input("[INPUT] Graduation Batch : ")

    print("[INFO] Executing Stored Procedure")

    conn.autocommit = True
    cursor.execute("exec PROC_PIN_UPLOAD_TO_WEB ?", graduation_batch)
    conn.autocommit = False
    conn.commit()

    print("[INFO] Stored Procedure Has Been Executed")

    # Mengupload file satu per satu
    tggl_log = input("[INPUT] Start Tanggal [YYYY-MM-DD] : ")
    log_query = "select kodeProdi from PIN_UPLOAD_LOG where TanggalLog >= '" +  \
        tggl_log + "' and TanggalLog <= DATEADD(DAY, 1, '" + tggl_log + "') order by kodeProdi"
    cursor.execute(log_query)
    prod_list = cursor.fetchall()
    print("Total Rows : ", len(prod_list))

    upload_XPATH = "/html/body/div[1]/div/div[3]/div/div[2]/div[2]/form/button"
    ClickXPATH(upload_XPATH)

    for i in prod_list:
        cursor.execute("exec EXPORTPRODI ?", i[0])

        mhs_list = cursor.fetchall()
        nim_list = []
        nina_list = []

        for x in mhs_list:
            nim_list.append(x[0])
            nina_list.append(x[1])

        file_path = "Asset/Prodi/PIN-Template.xlsx"    
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:    
            data = {'NIM': nim_list, 'PIN': nina_list}
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name="Sheet1", index=False)

        # Upload

        search_XPATH = "//input[contains(@placeholder,'Ketik Kata Kunci')]"
        SendXPATH(search_XPATH, i[0])

        try :
            ClickXPATH("/html/body/div[1]/div/div[3]/div[1]/div[2]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr[1]/td[4]/form/input[5]")
        except :
            continue

        upload_XPATH = "/html/body/div[1]/div/div[3]/div/div[2]/div/div/div/div[2]/form/input[9]"
        submit_XPATH = "/html/body/div[1]/div/div[3]/div/div[2]/div/div/div/div[2]/form/input[10]"
        back_XPATH = "/html/body/div[1]/div/div[3]/div/div[1]/div/a"
        pasang_XPATH = "/html/body/div[1]/div/div[3]/div/div[2]/div/div/div/div[2]/form[1]/button"

        file_path = os.path.abspath(r'Asset/Prodi/PIN-Template.xlsx')

        sleep(5)
        try :
            GetXPATHElement(upload_XPATH).send_keys(file_path)
            ClickXPATH(submit_XPATH)
            ClickXPATH(pasang_XPATH)
            ClickXPATH(back_XPATH)
        except:
            continue

    conn.commit()
    cursor.close()


'''
# Mengupdate Database Arsip setelah mengupload NINA Mahasiswa
'''


def UpdateArsip():
    # Pindah ke halaman History Arsip PIN
    driver.get("https://pin.kemdikbud.go.id/pin/index.php/historypin")
    cursor = conn.cursor()

    # Input tanggal batch pin reservasi
    tanggalBatch = input("[INPUT] Tanggal Batch [YYYYMMDD] : ")
    wisuda = input("[INPUT] Wisuda Batch : ")

    try:
        # Menghitung jumlah tombol download
        viewXPath = "//select[@name='DataTables_Table_0_length']/option[text()='100']"
        ClickXPATH(viewXPath)
        sleep(1)
        searchXPATH = "//input[contains(@placeholder,'Ketik Kata Kunci')]"
        SendXPATH(searchXPATH, tanggalBatch)
        sleep(1)
        btnCssSelector = ".btn.btn-success"
        buttons = SelectCSS(btnCssSelector)
        buttonLength = len(buttons)

        ProgressBar(0, buttonLength, prefix='Progress:', suffix='Complete', length=50)

        # Proses update arsip di database
        for i in range(buttonLength):
            sleep(3)
            ProgressBar(i + 1, buttonLength, prefix='Progress:', suffix='Complete', length=50)

            viewXPath = "//select[@name='DataTables_Table_0_length']/option[text()='100']"
            if ClickXPATH(viewXPath):
                searchXPATH = "//input[contains(@placeholder,'Ketik Kata Kunci')]"
                SendXPATH(searchXPATH, tanggalBatch)

            batchCodeXPATH = "/html/body/div[1]/div/div[3]/div[1]/div/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr[" + str(
                i + 1) + "]/td[2]"
            batchCode = GetXPATHElement(batchCodeXPATH).text
            dstring = batchCode[0:13]
            sleep(1)
            operationButton = SelectCSS(btnCssSelector)
            operationButton[i].click()
            sleep(5)

            while True:
                try:
                    df = pd.read_excel(r'Asset/Download/Daftar Nomor Ijazah-' +
                                       str(dstring) + str(batchCode) + '.xlsx', header=5)
                    no_list = df["NO"].tolist()
                    nim_list = df["NIM"].tolist()
                    nama_list = df["NAMA"].tolist()
                    nina_list = df["NOMOR IJAZAH"].tolist()
                    tanggal = datetime.today().strftime('%Y-%m-%d')
                    break
                except Exception:
                    continue

            row_length = len(no_list)
            for i in range(row_length):
                query = "insert into tblMahasiswaBerhasildiPadankan values('" + str(no_list[i]) + "','" + str(nim_list[i]) + "','" + str(
                    nama_list[i]) + "','" + str(nina_list[i]) + "','" + str(tanggal) + "','Crawler','Berhasil di padankan','" + wisuda + "','" + str(batchCode) + "')"
                cursor.execute(query)
                conn.commit()

    except TimeoutException as e:
        print("[INFO] No Data...")
        print(e)

    cursor.close()

'''
# Scheduler Modules
'''
def ReadConfig():
    Config = configparser.ConfigParser()
    Config.read(str(pathlib.Path(__file__).parent.resolve()) + '/config.ini')

    return Config

def ConfigSectionMap(section):
    Config = ReadConfig()
    dict1 = {}
    options = Config.options(section)
    for option in options:
        try:
            dict1[option] = Config.get(section, option)
            # if dict1[option] == -1:
            #     DebugPrint("skip: %s" % option)
        except:
            print("exception on %s!" % option)
            dict1[option] = None
    return dict1
    
def at_day(job, day_name):
    day_name = day_name.lower()
    if day_name == "monday":
        return job.monday
    if day_name == "tuesday":
        return job.tuesday
    if day_name == "wednesday":
        return job.wednesday
    if day_name == "thursday":
        return job.thursday
    if day_name == "friday":
        return job.friday
    if day_name == "saturday":
        return job.saturday
    if day_name == "sunday":
        return job.sunday
    raise Exception("Unknown name of day")

def Job():
    gradyear = ConfigSectionMap("Update")['gradyear']
    wisuda = ConfigSectionMap("Update")['wisuda']

    print("Checking login session")
    driver.get('https://pin.kemdikbud.go.id/pin/index.php/login')
    buttonXPATH = "//button[@class='btn-login btn-primary-login block-login full-width-login m-b']"
    loginStatus = GetXPATHElements(buttonXPATH)
    # print("Login status: ",  "OK" if len(loginStatus) > 0  else "FAILED/ALREADY LOGIN")
    if len(GetXPATHElements(buttonXPATH)) > 0:
        print("Logging in..")
        Login()   
    
    print("Running the reset job")
    Reset(skipEnterKey = True)
    print("Running the update job")
    Homepage()
    Update(gradyear, wisuda)

def RunScheduler():
    userday = ConfigSectionMap("Scheduler")['day']
    usertime = ConfigSectionMap("Scheduler")['time']

    if userday == None or usertime == None:
        print("Warning: You need to setup scheduler configuration first")
        input("Press Enter to continue...")
        Mainmenu(False)
        return

    job = at_day(schedule.every(), userday).at(usertime).do(Job)

    while True:
        system('cls')
        print("===================================")
        print("Running scheduler")
        print("===================================")
        print(f"Scheduler is set to run every {userday} at {usertime}")
        print(f"Next run: {job.next_run}")    
        print(f"Time remaining: {str(timedelta(seconds=schedule.idle_seconds()))}")
        print("\nWaiting for the next execution.. (Press `c` to cancel)")
        if keyboard.is_pressed("c"):            
            schedule.clear()
            print("Aborting job..")
            break
        
        schedule.run_pending()
        sleep(1)


def EditScheduler():
    system('cls')
    print("===================================")
    print("Edit scheduler")
    print("===================================")

    userday = input("When do you want the execute the job? [monday]: ")
    usertime = input("What time? [HH:SS]: ")
    gradyear = input("Graduation year for update [YYYY]: ")
    wisuda = input("Wisuda [1-100]: ")

    Config = ReadConfig()
    if not Config.has_section('Scheduler'):
        Config.add_section('Scheduler')
    Config.set('Scheduler', 'day', userday)
    Config.set('Scheduler', 'time', usertime)

    if not Config.has_section('Update'):
        Config.add_section('Update')
    Config.set('Update', 'gradyear', gradyear)
    Config.set('Update', 'wisuda', wisuda)

    with open(str(pathlib.Path(__file__).parent.resolve()) + '/config.ini', 'w') as configfile: 
        Config.write(configfile)

    print("Configuration set!")
    input("Press Enter to continue...")
    Mainmenu(False)


'''
####################################################################################################
# Menu untuk membuat pemakaian script lebih mudah
####################################################################################################
'''

def Mainmenu(homepage = True):
    menu = True

    while menu:
        if homepage: Homepage()
        system('cls')
        print("===================================")
        print("PIN Crawler")
        print("===================================")
        print("1. Update PIN")
        print("2. Upload PIN")
        print("3. Run Scheduler")
        print("4. Edit Scheduler")
        print("5. Exit")

        choose = True
        while choose:
            index = input("Choose[1-5] : ")

            if(index == "1"):
                UpdatePINMenu()
                choose = False
            elif(index == "2"):
                UploadPINMenu()
                choose = False
            elif(index == "3"):
                RunScheduler()
                choose = False
            elif(   index == "4"):
                EditScheduler()
                choose = False
            elif(index == "5"):
                choose = False
                menu = False
            elif(index == "6"):
                Job()
                choose = False
                input("Press Enter to Continue...")


'''
# Print menu untuk operasi Update PIN
'''

def UpdatePINMenu():
    Homepage()
    system('cls')
    print("===================================")
    print("PIN Crawler")
    print("===================================")
    print("1. Reset PIN")
    print("2. Update PIN")
    #print("3. Validate PIN")
    print("3. Back")

    choose = True
    while choose:
        index = input("Choose[1-3] : ")

        if(index == "1"):
            Reset()
            Homepage()
        elif(index == "2"):
            Update()
            Homepage()
        #elif(index == "3"):
        #    Validator()
        elif(index == "3"):
            choose = False


'''
# Print menu untuk operasi Upload PIN
'''


def UploadPINMenu():
    Homepage()
    system('cls')
    print("===================================")
    print("PIN Crawler")
    print("===================================")
    print("1. Upload PIN")
    print("2. Update Arsip")
    print("3. Back")

    choose = True
    while choose:
        index = input("Choose[1-3] : ")

        if(index == "1"):
            Upload()
        elif(index == "2"):
            UpdateArsip()
        elif(index == "3"):
            choose = False


'''
####################################################################################################
# End of Menu
####################################################################################################
'''


'''
# Function Library,
# Hapus tanda comment untuk function yang ingin digunakan
# Berikan tanda comment untuk function yang tidak ingin digunakan
'''
Login()
##########################
Mainmenu()
##########################
Homepage()

'''
# Close driver Chrome ketika selesai webcrawling
'''
driver.quit()
