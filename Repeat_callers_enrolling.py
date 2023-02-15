#!/usr/local/bin/python3.5

import requests
import time
import csv
import platform
import xlsxwriter
import json
import sys
import re
import os
from subprocess import call
from time import sleep
from datetime import datetime
from requests import Request, Session
from requests.structures import CaseInsensitiveDict

#----------------------------------------------
# Script :
# Repeat_callers_enrolling.py
# (c) August 2022 Jcd 
#----------------------------------------------

start_time = time.time()

# -- token pour Retail tenant

token = '00Wrmw3rZ64dWPvCqw1zockS_u-FIG3b3GfzWaqvrO'
user = dserop@tantrumcorp.com
serv = 'https://nbc.okta.com'
bncid = ''


#Json du header des requêtes vers Okta
headers = {'Authorization': 'SSWS {}'.format(token),
               'Accept': 'application/json',
               'Content-Type': 'application/json'}

#----------------------------------------------

# Fonction pour lire le fichier CSV
def read_csv(file_name):         
    data = {}
    with open(file_name, newline='', encoding='utf_8_sig') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:                       
            key = row["bngfReferenceId"]
            data[key] = row
    return data

# fonction d'écriture et formatage style printf
def print_to_stdout(*a): 
    # *a est l'array contenant les objets
    # passés comme argument à la fonction
    print(*a, file=sys.stdout)
    sys.stdout.write(str(a))
    return a
#================================================================

# fonction d'extraction des valeurs json imbriquées
def json_extract(obj, key):
    #Récupérer récursivement des valeurs à partir de JSON imbriqués
    arr = []

    def extract(obj, arr, key):
        #Recherche récursive des valeurs de la clé dans le JSON tree
        if isinstance(obj, dict):
            for k, v in obj.items():
                 if isinstance(v, (dict, list)):
                    extract(v, arr, key)
                 elif k == key:
                    arr.append(v)
        elif isinstance(obj, list):
            for item in obj:
                 extract(item, arr, key)              
        return arr

    values = extract(obj, arr, key)

    return values
#================================================================
# fonction d'extraction des valeurs json imbriquées
# version améliorée pour le choix de la valeur par indice
def json_extract_tst(obj, key, ind):
    #Récupérer récursivement des valeurs à partir de JSON imbriqués
    arr = []

    def extract(obj, arr, key):
        #Recherche récursive des valeurs de la clé dans le JSON tree
        if isinstance(obj, dict):
            for k, v in obj.items():
                 if isinstance(v, (dict, list)):
                    extract(v, arr, key)
                 elif k == key:
                    arr.append(v)
        elif isinstance(obj, list):
            for item in obj:
                 extract(item, arr, key)              
        return arr

    values = extract(obj, arr, key)
    #Extrait une sous-chaîne en spécifiant la position et le nombre de chr
    try:
        if values !=None:
            if ind == 1:
                ret_val = values[0]
            else: 
                ret_val = values[ind - 1]
        else:
            ret_val = 'null'
    except IndexError:
        ret_val =  'null'   

    #retourne le contenu de stdout vers la variable de travail
    return print_to_stdout(ret_val)    

#================================================================
def list_user_MFA(Host,oktaid):
# fonction ajout des usagers aux groupes de gestion MFA  
        
        
    url = f'{Host}/api/v1/users/{oktaid}/factors'     
    

    resp = requests.get(url, headers=headers) 

    #Si pas trouvé ou si pas accès au groupe -- > code http :  401
    #if resp == []:
       #return resp.status_code
       #return "VIDE"

    return resp.json()
#================================================================ 

# Create an new Excel file and add a worksheet.
try:
    Xls_file = 'Liste_usagers_bncid_facteurs.xlsx'
    workbook = xlsxwriter.Workbook(Xls_file)
    worksheet = workbook.add_worksheet('Primary_email')

    #============Header pour le rapport excel===============================
    # 19 champs 
    worksheet.write('A1', 'Okta_id')
    worksheet.write('B1', 'status')
    worksheet.write('C1', 'created')
    worksheet.write('D1', 'activated')
    worksheet.write('E1', 'statusChanged')
    worksheet.write('F1', 'lastLogin')   
    worksheet.write('G1', 'lastUpdated')
    worksheet.write('H1', 'passwordChanged')
    worksheet.write('I1', 'firstName')    
    worksheet.write('J1', 'lastName')    
    worksheet.write('K1', 'preferredLanguage')
    worksheet.write('L1', 'mobilePhone')    
    worksheet.write('M1', 'BNCid')
    worksheet.write('N1', 'secondEmail')
    worksheet.write('O1', 'login_id') 
    worksheet.write('P1', 'email')
    worksheet.write('Q1', 'factor_sms')
    worksheet.write('R1', 'factor_call')
    worksheet.write('S1', 'factor_email')
    

except xlsxwriter.exceptions.FileCreateError as e:
        print("Exception caught in workbook.close()")

#================================================================  

now = datetime.now()

print("========  Extraction des résultats :" , now ,"  ============" )
row = 1
col = 0

file_a = "CC_MFT_EXPORT_PERF_TMP_VoiceMFA_modif.csv"
data_a = read_csv(file_a)

#Extraction du contenu de la clé du dictionaire data_a
for key in data_a: 
 
    temp = key.strip().split(',')    
    bncid = temp[0].strip('"')

    # URL d'extraction des bncid
    url = f'{serv}/api/v1/users?search=profile.bngfReferenceId%20eq%20%22{bncid}%22'    

    try:

            events = requests.get(url, headers=headers, stream=True)   
            events.raise_for_status()
            
            i = 1
 
            for e in events.iter_lines():
                if e:
                    decoded_line = e.decode('utf-8')
                    data = json.loads(decoded_line)

                    if data==[]:
                        print("BncId : ", bncid , "  not found in OKta")                                                                        

                    for json_inner_array in data:
                                               
                        #============creation de la structure pour le rapport excel==============                           
                        
                            item1 = ' '.join(map(str,json_extract_tst(json_inner_array, 'id', 1)))
                            item2 = ' '.join(map(str,json_extract_tst(json_inner_array, 'status', 1)))
                            item3 = ' '.join(map(str,json_extract_tst(json_inner_array, 'created', 1)))
                            item4 = ' '.join(map(str,json_extract_tst(json_inner_array, 'activated', 1)))
                            item5 = ' '.join(map(str,json_extract_tst(json_inner_array, 'statusChanged', 1)))
                            item6 = ' '.join(map(str,json_extract_tst(json_inner_array, 'lastLogin', 1)))
                            item7 = ' '.join(map(str,json_extract_tst(json_inner_array, 'lastUpdated', 1)))
                            item8 = ' '.join(map(str,json_extract_tst(json_inner_array, 'passwordChanged', 1)))
                            item9 = ' '.join(map(str,json_extract_tst(json_inner_array, 'firstName', 1)))
                            item10 = ' '.join(map(str,json_extract_tst(json_inner_array, 'lastName', 1)))    
                            item11 = ' '.join(map(str,json_extract_tst(json_inner_array, 'preferredLanguage', 1)))
                            item12 = ' '.join(map(str,json_extract_tst(json_inner_array, 'mobilePhone', 1)))
                            item13 = ' '.join(map(str,json_extract_tst(json_inner_array, 'bngfReferenceId', 1)))
                            item14 = ' '.join(map(str,json_extract_tst(json_inner_array, 'secondEmail', 1)))
                            item15 = ' '.join(map(str,json_extract_tst(json_inner_array, 'login', 1)))
                            item16 = ' '.join(map(str,json_extract_tst(json_inner_array, 'email', 1)))

                            factors = list_user_MFA(serv, item1)

                            if factors!=[]:

                                sleep(0.01)
                                item17 = 'None'
                                item18 = 'None'
                                item19 = 'None'                                

                                for MFAid in factors:
                                    factor_type = ' '.join(map(str,json_extract(MFAid, 'factorType')))
                                
                                    if factor_type == 'sms':
                                        item17 = ' '.join(map(str,json_extract_tst(MFAid, 'phoneNumber', 1)))
                                        if item17==None:
                                            item17 = 'None'

                                    elif factor_type == 'call':
                                        item18 = ' '.join(map(str,json_extract_tst(MFAid, 'phoneNumber', 1)))
                                        if item18==None:
                                            item18 = 'None'

                                    elif factor_type == 'email':
                                        item19 = ' '.join(map(str,json_extract_tst(MFAid, 'email', 1)))                                    
                                        if item19==None:
                                            item19 = 'None'
                            else:
                                item17 = 'None'
                                item18 = 'None'
                                item19 = 'None'
                               
                        
                            print(i, item16 ," : " , item1 , " - ", item2, item3," - ",item4, item5, item6, item7, item8,item9,item10,item11,item12,
                                    item13,item14,item15,item16,item17,item18,item19)

                            itemlist = [item1,item2, item3,item4, item5, item6, item7, item8,item9,item10,item11,item12,item13,item14,item15,item16,item17,
                                        item18,item19]

                            print(itemlist)
                            #============body pour le rapport excel================================== 
                            #row est fixe
                            if i == 1:
                                pass
                            else: 
                                row = i
                                #col est en déplacement vers la droite
                                col = 0
                           
                            for x in itemlist: 

                                worksheet.write(row, col, itemlist[col])
                                col += 1
                            #============Fin de la structure pour le rapport excel================== 
  

                            i = i + 1

                    if not data:
                        print("")                    
                
                    sleep(0.01)
                
                
                    

            while 'next' in events.links:
                events = requests.get(events.links['next']['url'], headers=headers, stream=True)            
                events.raise_for_status()

                for e in events.iter_lines():
                    if e:
                        decoded_line = e.decode('utf-8')
                        data = json.loads(decoded_line)

                        if data==[]:
                            print("BncId : ", bncid , "  not found in OKta")                         

                        for json_inner_array in data:
                            #============creation de la structure pour le rapport excel===============                                
                            
                                item1 = ' '.join(map(str,json_extract_tst(json_inner_array, 'id', 1)))
                                item2 = ' '.join(map(str,json_extract_tst(json_inner_array, 'status', 1)))
                                item3 = ' '.join(map(str,json_extract_tst(json_inner_array, 'created', 1)))
                                item4 = ' '.join(map(str,json_extract_tst(json_inner_array, 'activated', 1)))
                                item5 = ' '.join(map(str,json_extract_tst(json_inner_array, 'statusChanged', 1)))
                                item6 = ' '.join(map(str,json_extract_tst(json_inner_array, 'lastLogin', 1)))
                                item7 = ' '.join(map(str,json_extract_tst(json_inner_array, 'lastUpdated', 1)))
                                item8 = ' '.join(map(str,json_extract_tst(json_inner_array, 'passwordChanged', 1)))
                                item9 = ' '.join(map(str,json_extract_tst(json_inner_array, 'firstName', 1)))
                                item10 = ' '.join(map(str,json_extract_tst(json_inner_array, 'lastName', 1)))    
                                item11 = ' '.join(map(str,json_extract_tst(json_inner_array, 'preferredLanguage', 1)))
                                item12 = ' '.join(map(str,json_extract_tst(json_inner_array, 'mobilePhone', 1)))
                                item13 = ' '.join(map(str,json_extract_tst(json_inner_array, 'bngfReferenceId', 1)))
                                item14 = ' '.join(map(str,json_extract_tst(json_inner_array, 'secondEmail', 1)))
                                item15 = ' '.join(map(str,json_extract_tst(json_inner_array, 'login', 1)))
                                item16 = ' '.join(map(str,json_extract_tst(json_inner_array, 'email', 1)))

                                factors = list_user_MFA(serv, item1)

                                if factors!=[]:                          

                                    sleep(0.01)
                                    item17 = 'None'
                                    item18 = 'None'
                                    item19 = 'None'

                                    for MFAid in factors:
                                        factor_type = ' '.join(map(str,json_extract(MFAid, 'factorType')))
                                    
                                        if factor_type == 'sms':
                                            item17 = ' '.join(map(str,json_extract_tst(MFAid, 'phoneNumber', 1)))
                                            if item17==None:
                                                item17 = 'None'

                                        elif factor_type == 'call':
                                            item18 = ' '.join(map(str,json_extract_tst(MFAid, 'phoneNumber', 1)))
                                            if item18==None:
                                                item18 = 'None'

                                        elif factor_type == 'email':
                                            item19 = ' '.join(map(str,json_extract_tst(MFAid, 'email', 1)))                                        
                                            if item19==None:
                                                item19 = 'None'
                                else:
                                    item17 = 'None'
                                    item18 = 'None'
                                    item19 = 'None'
                       
                        
                                print(i, item16 ," : " , item1 , " - ", item2, item3," - ",item4, item5, item6, item7, item8,item9,item10,item11,item12,
                                        item13,item14,item15,item16,item17,item18,item19)

                                itemlist = [item1,item2, item3,item4, item5, item6, item7, item8,item9,item10,item11,item12,item13,item14,item15,item16,item17,
                                            item18,item19]

                                print(itemlist)    
                                #============body pour le rapport excel================================== 
                                #row est fixe
                                if i == 1:
                                    pass
                                else:  
                                    row = i
                                    #col est en déplacement vers la droite
                                    col = 0
                                   
                                for x in itemlist:

                                    worksheet.write(row, col, itemlist[col])
                                    col += 1                  
                                #============Fin de la structure pour le rapport excel==================                                                                                       
                        
                          
                                i = i + 1

                        if not data:
                            print("")                        
                   
                        sleep(0.01)           
                        

    
    except Exception as e:
        error = "Extract failed with exception {}".format(e)
        print(error) 
        events = 'null'   

    print("=================Primary_email==================== ")
    

#Création du rapport

#================================================================

print("--- %s seconds ---" % (time.time() - start_time))
workbook.close()
