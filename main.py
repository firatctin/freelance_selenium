import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
import csv
from datetime import datetime
df = pd.read_excel('son.xlsx')
MODELS = df['MANN KODU'][209:500]
excel_name = '209-500'
brands = ['AIRKO COMPRESSORS', 'SSANGYONG', 'COMPAIR', 'BUCYRUS-ERIE', 'ATMOS', 'SAME', 'FORDSON / FORD SCHLEPPER', 'SDLG', 'LINDE', 'AGCO', 'HOLDEN', 'IVECO', 'THWAITES', 'DAVID BROWN', 'HAMM', 'FANUC', 'PACCAR', 'BUNGARTZ', 'SAEFFLE', 'WESTINGHOUSE BRAKES', 'MERCEDES-BENZ (DAIMLER AG)', 'STIHL', 'SATOH', 'KENWORTH-DART', 'AS-MOTOR', 'HANOMAG-HENSCHEL', 'KING', 'SCHULZ', 'LIEBHERR', 'TROJAN', 'E.R.F.', 'VM (MOTORI VM)', 'ZF', 'PETTER', 'LEROI', 'WAUKESHA', 'HOKUETSU KOGYO', 'NELSON N.P.', 'ENGESA', 'BENFORD', 'MG MOTOR UK', 'IHI (I H I)', 'COLES', 'TOHATSU', 'BASE VAC', 'MANTSINEN', 'KALMAR LMV', 'WEBASTO - SPHEROS', 'SODICK', 'ERKUNT', 'SANDVIK', 'AGRALE-DEUTZ', 'VOLVO CARS', 'BLITZ ROTARY', 'MULTIVAC', 'MACO-MEUDON', 'WACKER', 'KOHLER', 'HEINKEL', 'ECOAIR (INGERSOLL-RAND)', 'AURORA', 'DIECI', 'VAUXHALL', 'JACOBSEN', 'BENDIX', 'PEGASO', 'AMHC', 'P.V.R. (AGILENT TECHNOLOGIES)', 'STEINBOCK', 'HONDA', 'BELAIR', 'ALLIS CHALMERS', 'ZETOR', 'SCHAEFF', 'PUTZMEISTER', 'DV SYSTEMS', 'GEV', 'ROLLS-ROYCE', 'KING LONG', 'PEINER', 'LOSENHAUSEN', 'CHELYABINSK COMPRESSOR (CHKZ)', 'SCANIA', 'AIRPLUS', 'AGIE', 'REFORM-WERKE', 'TRABANT', 'LINK-BELT', 'GOTTWALD', 'TEMSA', 'HAKO', 'PFEIFFER', 'AVELAIR COMPRESSED AIR SYSTEMS', 'HINO', 'PALFINGER', 'LS AGRICULTURAL MACHINERY', 'GUTBROD', 'BETICO', 'EVOBUS/SETRA', 'INMESOL POWER SOLUTIONS', 'SEDDON-ATKINSON', 'MUSTANG', 'VW GROUP (AUDI/SEAT/SKODA/VW)', 'SCHMITZ CARGOBULL', 'POCLAIN', 'SOLE MARINE DIESEL', 'SOGIMAIR', 'SULLIVAN', 'DODGE', 'HOLDER', 'GIANT (TOBROCO MACHINES)', 'REXROTH', 'SPERRY NEW HOLLAND', 'GMC', 'BMW', 'SMART (DAIMLER AG)', 'LOMBARDINI', 'TIMBERJACK', 'FENWICK', 'BRAY', 'GENIE', 'WEIDEMANN', 'RENAULT SAMSUNG MOTORS', 'SULLAIR', 'BUCHER SCHÖRLING', 'HATZ', 'KRUPP', 'FAUN', 'HUERLIMANN', 'SEARS', 'SAUER', 'TERBERG', 'CASE-INTERNATIONAL/CASE-IH', 'AVELING-BARFORD', 'BENTLEY (VW GROUP)', 'ANTONIO CARRARO', 'IHC / CASE IH', 'HYUNDAI', 'MOSKVITCH', 'RENK', 'BECKER-KOMPRESSOREN', 'INNIO JENBACHER', 'ZASTAVA', 'FAUN-FRISCH', 'LINDNER', 'MELROE-BOBCAT', 'LISTER', 'ALLISON', 'SCHLUETER', 'FRANKLIN EQUIP CO.', 'MAHINDRA', 'KAELBLE', 'AGGRETECH', 'BRAUD', 'FORD USA', 'JAWA', 'ANDORIA', 'SAURER', 'MEILLER', 'ROTORCOMP', 'HYDRA-MAC', 'FINI', '-', 'HUBER-WARCO', 'STEIGER', 'MINI (BMW GROUP)', 'TOYOTA', 'CHARLES MACHINE WORKS', 'BOBCAT', 'JUNGHEINRICH', 'LADA', 'OERLIKON LEYBOLD VACUUM', 'DRILTECHN. DROTT', 'DEMAG (COMPAIR)', 'FILCOM', 'DAEWOO / CHEVROLET (GM)', 'DAIHATSU', 'BAUTZ', 'MTU', 'RUGGERINI', 'ORENSTEIN + KOPPEL (O+K)', 'PULFORD', 'MERLO', 'DYNAHOE', 'CATERPILLAR', 'SANTANA MOTOR', 'JAGUAR', 'GRADALL', 'BOSS', 'MERCRUISER', 'NANNI DIESEL', 'LAMBORGHINI', 'CARRIER-TRANSICOLD', 'ONAN', 'RENNER KOMPRESSOREN', 'VERMEER', 'GNUTTI', 'RENAULT TRUCKS (OKELIA)', 'INNOCENTI', 'KRAUSS-MAFFEI', 'HYSTER', 'MANITOU', 'LIGIER', 'BLAW-KNOX', 'YALE', 'EICHER', 'CLARK MICHIGAN', 'VOEGELE', 'KAMAZ', 'ALPHA DIESEL', 'KLEEMANN', 'KIOTI', 'COLES CRANES', 'MARVEL', 'BELARUS', 'SCHNELL', 'DYNAPAC', 'MAN', 'FIAT-AGRI', 'BOMAG', 'DAEWOO CONSTRUCTION EQUIPMENT', 'HUANGHAI JINMA TRACTOR', 'DAF', 'BUCHER HYDRAULICS', 'EATON YALE + TOWNE TROJA', 'DREIHA', 'DEUTZ AG / DEUTZ-FAHR (KHD)', 'CONSOLIDATED PNEUMATIC', 'INGERSOLL-RAND', 'WARTBURG', 'NETSTAL', 'EUCLID', 'BARBER-GREENE', 'CLARK EQUIPMENT', 'CUMMINS', 'CASE CONSTRUCTION', 'MINSK MOTOR PLANT', 'YANMAR', 'IRISBUS', 'SAMBRON', 'DORMAN', 'EBERSPÄCHER SÜTRAK', 'JDM', 'BAUDOUIN', 'NINGBO XINDA', 'AMERICAN MOTORS AMC', 'DATSUN', 'MOXY', 'STEYR MOTORS', 'ROSTSELMASH', 'TEREX', 'HITACHI', 'ELGIN SWEEPER', 'GRAU BREMSE / HALDEX', 'BROOM WADE (COMPAIR)', 'LANDINI', 'RAMMAX', 'JOHNSON MARINE', 'WACKER NEUSON', 'JOY', 'AXECO', 'HOLMAN (COMPAIR)', 'EDER', 'SOLIDAIR (BOGE GROUP)', 'NEOPLAN', 'BROCKWAY TRUCKS', 'HYDROVANE (COMPAIR)', 'RENAULT AGRICULTURE', 'ZETTELMEYER', 'SCHAEFFER', 'GARDNER-DENVER', 'SDMO', 'ON-BOARD-POWER (OBP)', 'MITSUBISHI ENGINES', 'JLG', 'URALAZ', 'NEXEN', 'TÜMOSAN', 'UNION-TECH', 'HOUGH', 'LETOURNEAU WESTINGHOUSE', 'HANOMAG BAUMASCHINEN', 'SUMITOMO', 'SCARAB', 'TAYLOR MACHINE', 'DARI', 'PERKINS', 'KELLOGG AMERICAN', 'AQUA POWER', 'GUELDNER', 'ASTON-MARTIN', 'KRAMER', 'ABG (VOLVO CE)', 'MATTEI', 'MWM', 'ENGEL', 'JOHNSTON SWEEPERS', 'TAKEUCHI', 'PETERBILT', 'ISEKI', 'MICHIGAN FLUID POWER', 'CHAR-LYNN', 'ASTRA (IVECO GROUP)', 'VOITH', 'ATLAS CRANES + EXCAVATORS', 'MAZDA', 'SOLARIS BUS', 'LAVERDA (AGRI)', 'SLANZI', 'ADICOMP', 'DALGAKIRAN', 'J.C.BAMFORD', 'MANITOWOC', 'CHRYSLER USA', 'NASH ELMO', 'PETTIBONE-MULLIKEN', 'WIRTGEN GROUP', 'UNICCOMP', 'YAMZ ENGINES', 'KONVEKTA', 'GALION', 'RENAULT', 'PELLENC', 'PORSCHE', 'KTM', 'FAHR (DEUTZ-FAHR)', 'SHIBAURA', 'ARGO TRACTORS', 'PILOT AIR', 'KRONE', 'IPS', 'MAZ', 'AIFO (FIAT-IVECO)', 'MICHIGAN', 'BRILLIANCE', 'AUWAERTER', 'GAZ', 'ALSTHOM', 'NEW HOLLAND (CNH)', 'INGERSOLL-RAND/DOOSAN', 'BERKO', 'AERZENER', 'UAZ', 'DENNIS', 'OTOKAR', 'TCM', 'MOTO GUZZI', 'SANDS AGRICULTURAL MACHINERY', 'DIAMOND REO', 'AUGUST COMPRESSOR', 'ROVER', 'BERGMEISTER', 'ISUZU', 'KOBELCO', 'OMC (OUTBOARD MARINE CORP.)', 'UNIC', 'KOEHRING', 'FRISCH', 'LINAMAR', 'TATA MOTORS', 'FIAT-ALLIS', 'AHLMANN', 'BUKH', 'ROPA', 'NISSAN-MOTOR / UD TRUCKS', 'P.V.R. VACUUM PUMPS', 'MESSERSI', 'LEITNER', 'LANZ', 'GOLDONI', 'MECALAC', 'CBT', 'MACK', 'BUFFALO-SPRINGFIELD', 'ANESTA IWATA', 'RENAULT TRUCKS (RVI)', 'FARYMANN', 'RADAELLI', 'BOGE', 'KAESSBOHRER', 'LDV', 'LUPAMAT', 'HYMAC', 'CHEVROLET', 'FLOTTMANN', 'MC CORMICK', 'AMMANN', 'WABCO', 'PEL JOB', 'TOW-MOTOR', 'NISSAN', 'BEDFORD', 'BRIGGS', 'TATRA', 'KAERCHER', 'MAHLE KOMPRESSOREN', 'MASSEY-FERGUSON', 'GEHL', 'MICROCAR', 'VAN HOOL', 'GMEINDER', 'VALPADANA', 'MERCURY', 'AUSTIN', 'BELL EQUIPMENT (SA)', 'SISU DIESEL', 'FICHTEL + SACHS', 'BMC (TR)', 'MITSUBISHI FORKLIFT', 'ROBIN', 'THERMO KING', 'CLAAS', 'SCHWING', 'SUNDSTRAND', 'FLUIDAIR', 'FORD', 'LANCER BOSS', 'MASERATI', 'KUKJE MACHINERY', 'OLIVER', 'FAI (KOMATSU)', 'ZUENDAPP', 'FG WILSON', 'FIAT-KOBELCO', 'AIRMAN', 'VETUS MARINE', 'INTERNATIONAL TRUCK', 'VALMET', 'CROSS', 'AEBI', 'BAUER', 'FIAT-HITACHI', 'MERITOR (MERITOR-WABCO)', 'CHALLENGER', 'KAWASAKI', 'PIAGGIO', 'KAESSBOHRER PISTENBULLY', 'IKARUS', 'FURUKAWA', 'PAUS', 'CADILLAC', 'MENZI MUCK', 'FERRARI', 'POWER SYSTEM', 'STEYR-DAIMLER-PUCH', 'F.S.O.', 'SIRONA', 'WHITLOCK', 'FUCHS', 'DETROIT DIESEL', 'SAAB', 'AG CHEM EQUIP', 'WILLE', 'WESTERN', 'WÄRTSILÄ', 'GEELY', 'HUSQVARNA', 'ASIA MOTORS', 'TONGCHENG', 'VALTRA-VALMET', 'CLUB CAR', 'AVIA', 'FOMOCO', 'SUNBEAM', 'ALASKA DIESEL', 'FISCHER PANDA GENERATOREN', 'ATLAS (TEREX)', 'PNEUMOFORE', 'ERROR', 'MITSUBISHI', 'LONG', 'GENERAL MOTORS (GM)', 'SEAGRAVE', 'PROTON', 'KIA MOTORS', 'LEYBOLD GMBH', 'SUBARU', 'TAMROCK', 'PEUGEOT', 'NAWOOTEC', 'KNORR-BREMSE', 'JOHN DEERE', 'CHERY', 'JOSVAL COMPRESSORES', 'GREATWALL', 'RIETSCHLE', 'IDEAL', 'LOTUS', 'AIR ENERGIE BERNARD', 'DACIA (RENAULT GROUP)', 'TENNANT', 'FS CURTIS', 'VERSATILE', 'BIAO DING', 'CITROEN', 'BRITISH LEYLAND', 'CHRYSLER U.K.', 'SUZUKI', 'BORGWARD', 'FIAT-IVECO', 'MULTICAR (HAKO)', 'BMW (MOTORBIKE)', 'LENZ', 'TORO', 'HANSHIN', 'HANBELL', 'CHRIS-CRAFT', 'GROVE', 'SCHMIDT WINTERDIENST', 'BRIGGS + STRATTON', 'OWATONNA', 'WITTE', 'HALDEX', 'GLAS', 'NELSON-WINSLOW', 'NEUSON', 'OPTARE', 'KOMATSU', 'UWE VERKEN AB', 'LJUNGBY MASKIN', 'BMC (BMH)', 'DEWULF', 'VOSS', '+GF+ AGIE CHARMILLES', 'WOLFCOMP', 'LONDON TAXI INTERNATIONAL', 'AGRIA', 'TÜRK TRAKTÖR', 'HYUNDAI CONSTRUCTION', 'BUESSING', 'VDL BUS GROUP', 'RANDON', 'TALBOT', 'DITCH WITCH', 'ISARTALER', 'HEULIEZ', 'LAND ROVER', 'LEYLAND', 'ALFA ROMEO', 'MITSUBISHI FUSO', 'DOOSAN', 'HERCULES', 'BUSCH', 'ATHEY', 'RICHIER', 'BALDWIN-LIMA-HAMILTON', 'FENDT', 'SATURN (GM)', 'TRIUMPH', 'TADANO FAUN', 'ALMIG KOMPRESSOREN', 'ATLAS (WEYHAUSEN)', 'AMG', 'MILLER ELECTRIC', 'STILL', 'AUSA', 'BARREIROS', 'SIGMA', 'NAVISTAR', 'EIMCO', 'DECHENTREITER', 'BYD', 'KUBOTA', 'FREIGHTLINER', 'STENHOJ', 'VOLVO (TRUCKS/VCE/VME/PENTA)', 'OPEL', "CHINOOK INDUSTRIAL LTD.'S", 'PRINOTH', 'FPT (FIAT POWERTRAIN)', 'MAXION', 'FIAT GROUP (ALFA/FIAT/LANCIA)', 'SIEMENS VDO', 'THOMAS', 'VANAIR (SULLAIR)', 'TONG YANG MOOLSAN', 'HAULOTTE', 'FODEN', 'SENNEBOGEN', 'MOOG']

driver = webdriver.Chrome()

columns = ['MODEL'] + brands
df_new = pd.DataFrame(columns= columns)

df_new.to_excel(f'data_final_{excel_name}.xlsx',index = False)
with open(f'data_cleared_{excel_name}.csv', mode='w', newline='') as csv_file:
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(columns)

driver.get("https://catalog.mann-filter.com/EU/tur")
try:
    WebDriverWait(driver,10).until(expected_conditions.element_to_be_clickable((By.ID,'onetrust-accept-btn-handler')))
    driver.find_element(by= By.ID, value= 'onetrust-accept-btn-handler').click()
except:
    pass
timer = 500
chronometers = []
for i in MODELS:
    
    start = datetime.now().timestamp()
    data_keys = columns
    data_values = [i] + ['-' for x in brands]
    data = dict(zip(data_keys,data_values))
    

    driver.get("https://catalog.mann-filter.com/EU/tur")
    WebDriverWait(driver,10).until(expected_conditions.element_to_be_clickable((By.ID,'headerSearchForm:headerSearchForm:searchQueryInput')))
    inputarea = driver.find_element(by=By.ID, value= 'headerSearchForm:headerSearchForm:searchQueryInput')
    inputarea.send_keys(i)
    inputarea.send_keys(Keys.ENTER)
    WebDriverWait(driver,10).until(expected_conditions.element_to_be_clickable((By.ID,'productDetailTab_Compare')))
    driver.find_element(by= By.ID, value= 'productDetailTab_Compare').click()
    values_old = '-'
    present_brands = 0
    
    for index,j in enumerate(brands):
        
        
        
        
        try:
            element = driver.find_element(by= By.ID, value= f'productCompare_Manufacturer_{j}')
            driver.execute_script("arguments[0].scrollIntoView();", element)
            element.click()
            present_brands += 1
            
            
        except:
            continue
        
        
        
        try:
            WebDriverWait(driver,10).until(expected_conditions.visibility_of_element_located((By.XPATH,f'//li[@id="productCompare_Manufacturer_{j}"]/div[2]' )))
            values_parent = driver.find_element(by= By.ID, value= f'productCompare_Manufacturer_{j}')
            if values_parent != element:
                values = values_parent.text
        except:   
            try: 
            
                time.sleep(1)
                
                values = driver.find_element(by= By.XPATH, value= f'//li[@id="productCompare_Manufacturer_{j}"]/div[2]').text
            except:
                values = 'ERROR'

             

            

        
        data_values[index +1] = values
        data[j] = values
        values_old = values
        try:
            close_element = driver.find_element(by= By.ID, value= f'productCompare_Manufacturer_{j}')
            driver.execute_script("arguments[0].scrollIntoView();", close_element)
            close_element.click()
        except:
            try:
                WebDriverWait(driver,10).until(expected_conditions.element_to_be_clickable((By.ID,f'productCompare_Manufacturer_{j}')))
                close_element = driver.find_element(by= By.ID, value= f'productCompare_Manufacturer_{j}')
                driver.execute_script("arguments[0].scrollIntoView();", close_element)
                close_element.click()
            except:
                escape_element1 = driver.find_element(by = By.ID, value= 'productDetailTab_Product')
                WebDriverWait(driver,10).until(expected_conditions.element_to_be_clickable((By.ID,'productDetailTab_Product')))
                driver.execute_script("arguments[0].scrollIntoView();", escape_element1)
                escape_element1.click()
                time.sleep(0.2)
                driver.find_element(by = By.ID, value= 'productDetailTab_Compare').click()
        time.sleep(1)
    
    df_update = pd.read_excel(f'data_final_{excel_name}.xlsx') 
   
    
    df_temp = pd.DataFrame(data= data, index = [0])
    
    df_update = df_update._append(df_temp)
    df_update.to_excel(f'data_final_{excel_name}.xlsx',index = False)
    

    with open(f'data_cleared_{excel_name}.csv', mode = 'a', newline='') as csv_file:
          csv_writer = csv.writer(csv_file)
          csv_writer.writerow(data_values)
    timer -= 1
    ending_time = datetime.now().timestamp()
    if len(chronometers) == 0:
        print(f'{timer} öge kaldı. Tahmini zaman:')
    else:
        print(f'{timer} öge kaldı. Tahmini kalan zaman: {((sum(chronometers)/len(chronometers))*timer)/60} dakika')
    fark = ending_time - start
    chronometers.append(int(fark)* present_brands)
    time.sleep(0.1)


