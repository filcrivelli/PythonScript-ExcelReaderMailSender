# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import openpyxl
import os

"""importo libreria per mandare mail in automatico"""
import smtplib
"""creiamo una variabile che contenga il server di gmail a cui poi ci connetteremo"""
conn= smtplib.SMTP("#mailProvider#", '#portNumber#') 
"""ci connettiamo al server"""
conn.ehlo()
"""facciamo in modo che la password che diamo nella riga dopo per connetterci al nostro
account gmail venga criptata"""
conn.starttls()
"""importo libreria per mandare mail in automatico"""
conn.login("#emailAddress#", "#pwd#")

"""andiamo a prendere la cartella dove si trova il file"""
os.chdir("#pathToWorkingDir#")

"""creiamo una variabile che contenga aperto il file excel su cui vogliamo lavorare"""
workbook = openpyxl.load_workbook("#exelFileToRead#")

"""prendiamo il foglio del file su cui vogliamo lavorare"""
sheet = workbook.get_sheet_by_name("#sheetName#")

"""creiamo una variabile che ci dirà il numero di candidati sopra un determinato punteggio 
in questo caso 39.2"""
conto = 0

"""looppiamo tutte le celle excel che contengono i punteggi dei candidati e 
in questo caso vanno dalla 2 alla 697 """
for i in range(2,697):
    valore = (sheet.cell(row=i, column=7).value)
    if valore > 39.2:
        conto +=1
        """se il conto è uguale a 1 mandiamo una mail al candidato in questo caso non ce l'ho
        sul foglio excel ma uso quella di lollo dicendogli cheè arrivato primo al test"""
    if conto == 1:
        conn.sendmail("#emailSender#", "#emailReceiver#", "Subject: Mail automatica \n\nCiao lollone ti mando mail automatica da python")
        conn.quit
        
        """se il conto è uguale a 1 mandiamo una mail al candidato in questo caso non ce l'ho
        sul foglio excel ma uso quella di lollo dicendogli cheè arrivato primo al test"""
print(conto)

    
   
    
    
  
    
        
    


