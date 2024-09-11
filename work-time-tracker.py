import requests, bs4, re, os, openpyxl, time, datetime, calendar, sys, threading, platform
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from plyer import notification
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from requests.exceptions import ConnectionError
from selenium.common.exceptions import WebDriverException
if platform.system()=='Windows':
    from pywintypes import com_error
if not getattr(__builtins__, "WindowsError", None): #nel caso in cui l'OS sia Mac, per gestione di exception
    class WindowsError(OSError): pass

def login(*args):
    try:
        loginurl =  'https://gestionale.com/logged_in'
        payload = {
            '__ac_name': user.get(),
            '__ac_password': password.get()
            }
        global s
        s = requests.session()
        global headers
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'}
        global p
        p = s.post(loginurl,data=payload,headers=headers)
        
        resp = s.get("https://gestionale.com/main/commesse/lista_commesse")
        global soup
        soup = bs4.BeautifulSoup(resp.text,'html.parser')
        loginCheck = soup.find('a', href={"https://gestionale.com/recover"})

        if loginCheck: #check se il login è andato a buon fine
            tk.messagebox.showwarning(title='Credenziali errate',message='Username o password errata, inserisci le credenziali corrette.',parent=loginWindow)
        else:
            #chiude la finestra di login e apre la finestra principale
            loginWindow.after(1200, loginWindow.withdraw)
            root.after(1200, root.deiconify)
            
            #check del giorno e dell'ora
            workTimeCheck()
            selectButton.configure(state='disabled')
    except ConnectionError: #in caso non ci fosse una connessione a internet
            tk.messagebox.showerror(title='Connessione assente',message='Connessione a internet assente. Assicurarsi di essere collegati e riprovare il login.',parent=loginWindow)

def workTimeCheck():
    currentDay=time.ctime()
    isWorkday=re.compile(r'^(Mon|Tue|Wed|Thu|Fri)')
    if re.search(isWorkday,currentDay):
        global dt,nextRow
        dt=datetime.datetime.now()
        minRoundedCheck=round(dt.minute/30)*30
        if minRoundedCheck==60:
            hourRoundedCheck=dt.hour+1
            minRoundedCheck=0
        else:
            hourRoundedCheck=dt.hour
        if hourRoundedCheck in range(9,13) or hourRoundedCheck in range(14,18):
            setup()
        elif hourRoundedCheck>=18:
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "L'ora attuale va oltre l'orario lavorativo. Registra le ore di oggi su Gestionale se non l'hai ancora fatto oppure chiudi il programma.",'tagCenter')
            mainTextBox.configure(state='disabled')
            selectDrop.configure(state='disabled')
            selectButton.configure(state='disabled')
            notesEntry.set('')
            notesText.configure(state='disabled')
            nextRow=18
            setup()
            notSetCounter()
        else:
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "Il programma si attiverà all'inizio dell'orario lavorativo (alle 9 o alle 14). Attendere...",'tagCenter')
            mainTextBox.configure(state='disabled')
            selectDrop.configure(state='disabled')
            selectButton.configure(state='disabled')
            gestionaleButton.configure(state='disabled')
            notesEntry.set('')
            notesText.configure(state='disabled')
            root.after(60000, workTimeCheck)
    else:
        mainTextBox.configure(state='normal')
        clearMainTextBox()
        mainTextBox.insert('end', 'Oggi non è un giorno feriale. Programma in chiusura...','tagCenter')
        mainTextBox.configure(state='disabled')
        selectDrop.configure(state='disabled')
        selectButton.configure(state='disabled')
        gestionaleButton.configure(state='disabled')
        notesEntry.set('')
        notesText.configure(state='disabled')
        root.after(5000,rootDestroy)

def setup():
    try:
        table = soup.find('table', {'class':'new-companylist'})
        rows = table.find_all('tr', {'valign':'top'})

        #estrae i nomi dei clienti e crea una lista
        colonnaClienti=table.select('tr td:nth-of-type(3)')
        listaClienti=[]
        for cliente in colonnaClienti:
            #listaClienti.append(cliente.text.strip())
            listaClienti.append(" ".join(cliente.text.strip().split()))

        #estrae la prima riga delle descrizioni delle commesse e crea una lista
        listaDescrizioni1 = []
        estrazDescrizioni = soup.find_all('span', {'class':'green'})
        for descrizione1 in estrazDescrizioni:
            #listaDescrizioni1.append(descrizione1.text.strip())
            listaDescrizioni1.append(" ".join(descrizione1.text.strip().split()))

        #estrae la seconda riga delle descrizioni delle commesse e crea una lista
        listaDescrizioni2 = []
        for descrizione2 in estrazDescrizioni:
            listaDescrizioni2.append(" ".join(descrizione2.next_sibling.next_sibling.next_sibling.text.strip().split()))

        #merge delle due liste e creazione di dizionario in ordine alfabetico
        global listClientSort
        listaClientiDescrizioni=[clienti+' | '+descrizioni1+' | '+descrizioni2 for clienti,descrizioni1,descrizioni2 in zip(listaClienti,listaDescrizioni1,listaDescrizioni2)]
        listClientSort=sorted(listaClientiDescrizioni)

        numberList=list(range(1,len(listaClientiDescrizioni)+1))
        numberListStr=list(map(str,numberList))
        dictNumClients=dict(zip(numberList,listClientSort))

        #crea dizionario delle task per la commessa COSTI
        global listaTaskCosti
        findCommCosti=re.compile('costi.*',flags=re.IGNORECASE)
        #commCostiURL=str.replace(re.search("(?P<url>https?://[^\s]+)", table.find(string=findCommCosti).find_parent('td').get('onclick')).group("url"),"'",'')
        listaCommCosti = table.find_all(string=findCommCosti)
        listaCommCostiURL = []
        for comm in listaCommCosti:
            listaCommCostiURL.append(str.replace(re.search("(?P<url>https?://[^\s]+)", comm.find_parent('td').get('onclick')).group("url"),"'",''))
        listaCommCostiURL=list(dict.fromkeys(listaCommCostiURL)) #elimina gli url duplicati delle commesse costi
        nCommCosti=0
        for commURL in listaCommCostiURL:
            respCosti = s.get(commURL)
            soupCosti = bs4.BeautifulSoup(respCosti.text,'html.parser')
            listaTaskCosti=[]
            for option in soupCosti.find('select', {'name':'id_task'}):
                    listaTaskCosti.append(option.text.strip())
            while("" in listaTaskCosti): #pulizia della lista di task Costi
                    listaTaskCosti.remove("")
            listaTaskCosti.remove('--Seleziona--')
            numberTasks=list(range(1,len(listaTaskCosti)+1))
            numberTasksStr=list(map(str,numberTasks))
            taskCosti=dict(zip(numberTasks,listaTaskCosti))

            clientCosti=[entry for entry in listaClientiDescrizioni if ' COSTI ' in entry or ' Costi ' in entry or ' costi ' in entry]
            #listaClientiDescrizioni.remove(clientCosti[0]) #rimuovere asterisco se si vuole rimuovere la commessa costi generica (senza task)
            for task in listaTaskCosti:
                clientCostiTask=clientCosti[nCommCosti]+' | '+task
                listClientSort.append(clientCostiTask)
            nCommCosti+=1
        listClientSort=sorted(listClientSort)
        

        #apre il file e le cartelle excel
        global wb, sheet1, sheet2
        if platform.system()=='Windows':
            if os.path.exists('AutoGestionale.xlsx'):
                wb=openpyxl.load_workbook('AutoGestionale.xlsx')
            else: #crea il file excel se non è gia presente nella cartella
                wb=openpyxl.Workbook()
                sheet1=wb.active
                sheet1['A1']='year'
                sheet1['B1']='month'
                sheet1['C1']='day'
                sheet1['D1']='client'
                sheet1['E1']='description1'
                sheet1['F1']='description2'
                sheet1['G1']='optionalTask'
                sheet1['H1']='optionalNote'
                sheet1['I1']='startHour'
                sheet1['J1']='startMin'
                sheet1['K1']='endHour'
                sheet1['L1']='endMin'
                sheet1['M1']='hoursWorked'
                sheet1['N1']='minutesWorked'
                sheet1.title='timeSheet'
                hourValue=9
                iterCount=0
                for i in sheet1['I2:I17']:
                    for cell in i:
                        if cell.row==10:
                            hourValue+=1
                        if iterCount%2==0 and iterCount!=0:
                            hourValue+=1
                        cell.value=hourValue
                        iterCount+=1
                hourValue=9
                iterCount=0
                for i in sheet1['K2:K17']:
                    for cell in i:
                        if cell.row==10:
                            hourValue+=1
                        if iterCount%2!=0:
                            hourValue+=1
                        cell.value=hourValue
                        iterCount+=1
                minValue=0
                for i in sheet1['J2:J17']:
                    for cell in i:
                        cell.value=minValue
                        minValue+=30
                        if minValue==60:
                            minValue=0
                minValue=30
                for i in sheet1['L2:L17']:
                    for cell in i:
                        cell.value=minValue
                        minValue+=30
                        if minValue==60:
                            minValue=0
                for i in sheet1['M2:M17']:
                    for cell in i:
                        cell.value=0
                for i in sheet1['N2:N17']:
                    for cell in i:
                        cell.value=30
                wb.create_sheet('clientList')
                sheet2=wb['clientList']
                sheet2['A1']='client'
                sheet2['B1']='description1'
                sheet2['C1']='description2'
                sheet2['D1']='optionalTask'
                sheet2['F1']="Se ti mancano ore da inserire su Gestionale:"
                sheet2['F2']='1- Copia i valori "client", "description1", "description2" e "optionalTask" del cliente che ti interessa'
                sheet2['F3']='2- Incolla i valori nelle apposite celle della scheda "timeSheet"'
                sheet2['F4']="3- inserisci manualmente le ore e i minuti di inizio e fine dell'attività nelle apposite celle"
                wb.save('AutoGestionale.xlsx')
        if platform.system()=='Darwin':
            global wbPath
            wbPath=os.path.join(os.getcwd(),'AutoGestionale.xlsx')
            wb=openpyxl.load_workbook(wbPath)
        sheet1=wb['timeSheet']
        sheet2=wb['clientList']
        
        #cancella valori del giorno prima in entrambi i fogli (se è la prima apertura di oggi del programma)        
        dt=datetime.datetime.now()
        if sheet1['C2'].value!=dt.day:
            for row in sheet1['A2:H17']:
                for cell in row:
                    cell.value=None
                
        #incolla nel primo sheet l'anno, il mese e il giorno
        cMonth=calendar.month_name[dt.month]
        global minRounded
        minRounded=None #setta un valore per minRounded
        global hourRounded
        hourRounded=None #setta un valore per hourRounded
        global checkBefGestionaleLoop
        checkBefGestionaleLoop=0 #setta il valore per evitare errori nella creazione della finestra delle ore che mancano fino alle 18
        sheet1['A2']=dt.year
        sheet1['B2']=cMonth
        sheet1['C2']=dt.day
        
        #incolla nel secondo sheet la lista di clienti e descrizioni e task Costi
        for row in sheet2['A2:D'+str(len(sheet2['D']))]:
            for cell in row:
                cell.value=None
        currRow=2
        key=0
        for values in listClientSort:
            listClientSplit=listClientSort[key].split(' | ')
            sheet2.cell(row=currRow,column=1).value=listClientSplit[0]
            sheet2.cell(row=currRow,column=2).value=listClientSplit[1]
            sheet2.cell(row=currRow,column=3).value=listClientSplit[2]
            if 'COSTI' in listClientSplit[1].upper() and len(listClientSplit)==4:
                sheet2.cell(row=currRow,column=4).value=listClientSplit[3]
            currRow+=1
            key+=1
        wb.save('AutoGestionale.xlsx')
        if dt.hour in range(9,14) or dt.hour in range(14,17) or (dt.hour==17 and dt.minute<45):###
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "Seleziona il cliente su cui intendi lavorare per la prossima mezz'ora.",'tagCenter')
            mainTextBox.configure(state='disabled')
            #aggiorna la lista dei clienti dei menù commesse e task costi
            selectClick.set('')
            selectDrop['values']=listClientSort
            firstNotification()
            global notSetCount
            notSetCount=0
    except PermissionError as e: #in caso il file excel fosse aperto
        if str(e)=="[Errno 13] Permission denied: 'AutoGestionale.xlsx'":
            root.after(1000,root.withdraw)
            loginWindow.after(1001, loginWindow.deiconify)
            tk.messagebox.showwarning(title='File Excel aperto',message='Il file Excel AutoGestionale risulta aperto, chiudilo prima di procedere con il login.',parent=loginWindow)

def firstNotification():
    global loopCount,nextRow,firstTime,startTime,dt,minRounded,hourRounded
    loopCount=1
    firstTime=datetime.datetime.now()#per check al momento della seconda registrazione di ore
    startTime=firstTime
    if firstTime.minute not in [30,0]:
        minRounded=round(firstTime.minute/30)*30 #crea valore per i minuti arrotondati al 30simo minuto più vicino
        if minRounded==60:
            hourRounded=firstTime.hour+1
            for cell in sheet1['I']: #seleziona la riga di excel su cui scrivere che corrisponde all'orario corrente
                if cell.value==hourRounded:
                    if sheet1.cell(row=cell.row, column=cell.column+1).value==0:
                        nextRow=cell.row
        else:
            hourRounded=firstTime.hour
            for cell in sheet1['I']: #seleziona la riga di excel su cui scrivere che corrisponde all'orario corrente
                if cell.value==firstTime.hour:
                    if sheet1.cell(row=cell.row, column=cell.column+1).value==minRounded:
                        nextRow=cell.row
        if minRounded==0: #imposta il messaggio sopra la selezione della commessa in base all'orario in cui è stato effettuato il login
            selectLabelText.set("Scegli la commessa su cui intendi lavorare la prossima mezz'ora (dalle "+str(firstTime.hour)+":00 alle "+str(firstTime.hour)+":30)")
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "In attesa della selezione della commessa...",'tagCenter')
            mainTextBox.configure(state='disabled')
            selectDrop.configure(state='normal')
            selectButton.configure(state='normal')
            notesText.configure(state='normal')
        elif minRounded==30:
            selectLabelText.set("Scegli la commessa su cui intendi lavorare la prossima mezz'ora (dalle "+str(firstTime.hour)+":30 alle "+str(firstTime.hour+1)+":00)")
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "In attesa della selezione della commessa...",'tagCenter')
            mainTextBox.configure(state='disabled')
            selectDrop.configure(state='normal')
            selectButton.configure(state='normal')
            notesText.configure(state='normal')
        elif minRounded==60:
            selectLabelText.set("Scegli la commessa su cui intendi lavorare la prossima mezz'ora (dalle "+str(firstTime.hour+1)+":00 alle "+str(firstTime.hour+1)+":30)")
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "In attesa della selezione della commessa...",'tagCenter')
            mainTextBox.configure(state='disabled')
            selectDrop.configure(state='normal')
            selectButton.configure(state='normal')
            notesText.configure(state='normal')
            minRounded=0 #per fare in modo che il confronto con il tempo corrente funzioni in timeCheck
    elif firstTime.minute in [30,0]:
        for cell in sheet1['I']:
            if cell.value==firstTime.hour:
                if sheet1.cell(row=cell.row, column=cell.column+1).value==firstTime.minute:
                    nextRow=cell.row
        if firstTime.minute==0:#imposta il messaggio sopra la selezione della commessa in base all'orario in cui è stato effettuato il login
            selectLabelText.set("Scegli la commessa su cui intendi lavorare la prossima mezz'ora (dalle "+str(startTime.hour)+":00 alle "+str(startTime.hour)+":30)")
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "In attesa della selezione della commessa...",'tagCenter')
            mainTextBox.configure(state='disabled')
            selectDrop.configure(state='normal')
            selectButton.configure(state='normal')
            notesText.configure(state='normal')
        elif firstTime.minute==30:
            selectLabelText.set("Scegli la commessa su cui intendi lavorare la prossima mezz'ora (dalle "+str(startTime.hour)+":30 alle "+str(startTime.hour+1)+":00)")
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "In attesa della selezione della commessa...",'tagCenter')
            mainTextBox.configure(state='disabled')
            selectDrop.configure(state='normal')
            selectButton.configure(state='normal')
            notesText.configure(state='normal')
    if nextRow>2: #controlla se ci sono ore non registrate tra le 9 e l'orario di inizio
        for cell_tuple in sheet1['D2:D'+str(nextRow-1)]:
            cell=cell_tuple[0]
            if sheet1.cell(row=cell.row, column=cell.column).value==None:
                sheet1.cell(row=cell.row, column=cell.column).value='NOT SET'                
        wb.save('AutoGestionale.xlsx')
        notSetCounter()
    root.after(60000,timeCheck)

def standardNotification():
    global startTime,dt
    if dt.minute in [30,0] and dt.minute!=minRounded:
        startTime=datetime.datetime.now()
        if platform.system()=='Windows':
            notification.notify(
                    title='Gestionale',
                    message="Puoi annotare la mezz'ora successiva.",
                    app_icon=resource_path("logo-gestionale-win.ico"),
                    timeout=5,
                    )
        if platform.system()=='Darwin':
            os.system("""
              osascript -e 'display notification "{}" with title "{}" sound name "beep"'
              """.format('Puoi annotare la mezzora successiva.', 'Gestionale'))
        if startTime.minute==0:#imposta il messaggio sopra la selezione della commessa in base all'orario in cui è stato effettuato il login
            selectLabelText.set("Scegli la commessa su cui intendi lavorare la prossima mezz'ora (dalle "+str(startTime.hour)+":00 alle "+str(startTime.hour)+":30)")
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "In attesa della selezione della commessa...",'tagCenter')
            mainTextBox.configure(state='disabled')
            selectDrop.configure(state='normal')
            selectButton.configure(state='normal')
            notesText.configure(state='normal')
        elif startTime.minute==30:
            selectLabelText.set("Scegli la commessa su cui intendi lavorare la prossima mezz'ora (dalle "+str(startTime.hour)+":30 alle "+str(startTime.hour+1)+":00)")
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "In attesa della selezione della commessa...",'tagCenter')
            mainTextBox.configure(state='disabled')
            selectDrop.configure(state='normal')
            selectButton.configure(state='normal')
            notesText.configure(state='normal')
        notSetCounter()

def finalNotification():
    if platform.system()=='Windows':
        notification.notify(
                    title='Gestionale',
                    message="L'ultimo slot di mezz'ora è stato raggiunto.",
                    app_icon=resource_path("logo-gestionale-win.ico"),
                    timeout=5,
                    )
    if platform.system()=='Darwin':
        os.system("""
              osascript -e 'display notification "{}" with title "{}" sound name "beep"'
              """.format('Ultimo slot di mezzora raggiunto.', 'Gestionale'))

def confirm():
    clientChoiceSplit=selectClick.get().split(' | ')
    global minRounded,nextRow,loopCount,startTime,firstTime
    try:
        if len(clientChoiceSplit)==4:
            sheet1.cell(row=nextRow,column=7).value=clientChoiceSplit[3]
        sheet1.cell(row=nextRow,column=4).value=clientChoiceSplit[0]
        sheet1.cell(row=nextRow,column=5).value=clientChoiceSplit[1]
        sheet1.cell(row=nextRow,column=6).value=clientChoiceSplit[2]
        sheet1.cell(row=nextRow,column=8).value=notesEntry.get()
        loopCount+=1
        wb.save('AutoGestionale.xlsx')
        mainTextBox.configure(state='normal')
        clearMainTextBox()
        mainTextBox.insert('end', "La mezz'ora di lavoro è stata annotata nel file Excel.",'tagCenter')
        mainTextBox.configure(state='disabled')
        selectDrop.configure(state='disabled')
        selectButton.configure(state='disabled')
        notesEntry.set('')
        notesText.configure(state='disabled')
        if startTime.hour in range(9,13) or startTime.hour in range(14,17) or (startTime.hour==17 and startTime.minute<15):
            mainTextBox.configure(state='normal')
            mainTextBox.insert('end'," In attesa della prossima mezz'ora da annotare...",'tagCenter')
            mainTextBox.configure(state='disabled')
        else:
            mainTextBox.configure(state='normal')
            mainTextBox.insert('end'," L'ultimo slot di mezz'ora è stato annotato. Compila eventuali slot mancanti e registra le ore.",'tagCenter')
            mainTextBox.configure(state='disabled')
            
    except PermissionError as e: #in caso il file excel fosse aperto
        if str(e)=="[Errno 13] Permission denied: 'AutoGestionale.xlsx'":
            loopCount-=1
            tk.messagebox.showwarning(title='File Excel aperto',message='Il file Excel AutoGestionale risulta aperto, chiudilo prima di procedere con la conferma della commessa.',parent=root)

def checkClient(*args):
    if selectClick.get() in listClientSort:
        selectButton.configure(state='normal')            
    else:
        selectButton.configure(state='disabled')

def checkNotSetClient(*args):
    global notSetComboVar,listClientSort,notSetComboVarGet,confirmNotSetButton
    global confirmNotSetButton
    notSetComboVarGet=[]
    for i in notSetComboVar:
        notSetComboVarGet.append(i.get())
    unlockNotSet=all(client in listClientSort for client in notSetComboVarGet)
    if unlockNotSet:
            confirmNotSetButton.configure(state='normal')
    else:
            confirmNotSetButton.configure(state='disabled')

def checkMissingClient(*args):
    global missingComboVar,listClientSort,missingComboVarGet,confirmMissingButton
    missingComboVarGet=[]
    for i in missingComboVar:
        missingComboVarGet.append(i.get())
    unlockMissing=all(client in listClientSort for client in missingComboVarGet)
    if unlockMissing:
            confirmMissingButton.configure(state='normal')
    else:
            confirmMissingButton.configure(state='disabled')

def timeCheck():
    global dt,nextRow,loopCount,minRounded
    dt=datetime.datetime.now()
    if dt.hour==13: #se sono le 13, metti in pausa per un'ora
        if dt.minute==0:
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "Inizio della pausa pranzo, la selezione delle commesse verrà sospesa fino alle 14. Attendere...",'tagCenter')
            mainTextBox.configure(state='disabled')
            selectDrop.configure(state='disabled')
            selectButton.configure(state='disabled')
            notesText.configure(state='disabled')
            selectLabelText.set("Scegli la commessa su cui intendi lavorare:")
            checkRegistered()
        root.after(60000,timeCheck)
    elif dt.hour in range(9,13) or dt.hour in range(14,18):#check della mezz'ora
        if dt.minute in [30,0] and dt.minute!=minRounded:
            checkRegistered()
            nextRow+=1
            standardNotification()
            if minRounded!=None:
                minRounded=None
            root.after(60000,timeCheck)
        else:
            root.after(60000,timeCheck)
    elif dt.hour==18 and dt.minute==0:
        checkRegistered()
        nextRow+=1
        notSetCounter() #aggiunto per il problema dell'ultima mezz'ora non aggiunta alle mancanti se non è stata registrata manualmente
        mainTextBox.configure(state='normal')
        mainTextBox.insert('end',"L'ultimo slot di mezz'ora è stato raggiunto. Compila eventuali slot mancanti e registra le ore.",'tagCenter')
        mainTextBox.configure(state='disabled')
        selectDrop.set('')
        selectDrop.configure(state='disabled')
        selectButton.configure(state='disabled')
        notesText.configure(state='disabled')
        selectLabelText.set("Scegli la commessa su cui intendi lavorare:")
        finalNotification()

def checkRegistered():#controlla se la mezz'ora precedente è stata inserita e in caso inserisce un placeholder
    global loopCount,nextRow,minRounded,hourRounded,startTime
    startTime=datetime.datetime.now()
    if sheet1.cell(row=nextRow,column=4).value==None:
        sheet1.cell(row=nextRow,column=4).value='NOT SET'
        wb.save('AutoGestionale.xlsx')
        loopCount+=1
    notSetCounter()

def notSetCounter():#contatore dei clienti in not set, fare finestra per reimpostarli con widget
    global listClientSort,notSetRowList,notSetCount,notSetRowList,notSetText,notSetSelectClick,notSetComboVar,confirmNotSetButton,notSetNotesVar,listNotSetVarCombo,nextRow,confirmNotSetButton
    notSetRowList=[]
    for client in (sheet1.cell(row=i,column=4) for i in range(2,nextRow)):
        if client.value=='NOT SET':
            notSetRowList.append(client.row)
    notSetCount=len(notSetRowList)
    notSetText.set(str(notSetCount)+" slot senza una commessa associata.")
    destroyNotSetWidget()#distrugge i vecchi widget della finestra degli slot mancanti
    notSetWindowRow=1
    headerFrame=tk.Frame(notSetWindow)
    clientHeader=tk.Label(headerFrame,text='Seleziona una commessa per ogni slot mancante e inserisci una nota (opzionale):')
    clientHeader.grid(row=0,column=0,padx=10,pady=5)
    headerFrame.pack(anchor="w")
    notSetComboWidgets=[]
    notSetComboVar=[]
    notSetNotesVar=[]
    notSetFrame=tk.Frame(notSetWindow)
    for notSetRow in notSetRowList:#crea una riga nella finestra per ogni slot di mezz'ora mancante
        #crea gli widget della finestra per gli slot mancanti
        notSetLabel=tk.Label(notSetFrame,text="Dalle "+str("{:02d}".format(sheet1.cell(row=notSetRow,column=9).value))+":"+str("{:02d}".format(sheet1.cell(row=notSetRow,column=10).value))+" alle "+str("{:02d}".format(sheet1.cell(row=notSetRow,column=11).value))+":"+str("{:02d}".format(sheet1.cell(row=notSetRow,column=12).value))+":")
        notSetSelectClick=tk.StringVar()
        notSetSelectClick.set('')
        notSetComboVar.append(notSetSelectClick)
        notSetSelectClick.trace('w',checkNotSetClient)
        if platform.system()=='Windows':
            notSetSelectDrop=ttk.Combobox(notSetFrame,textvariable=notSetSelectClick,width=130)
        if platform.system()=='Darwin':
            notSetSelectDrop=ttk.Combobox(notSetFrame,textvariable=notSetSelectClick,width=102)
        notSetComboWidgets.append(notSetSelectDrop)
        notSetSelectDrop['values']=listClientSort
        notSetNotes=tk.StringVar()
        notSetNotes.set('')
        notSetNotesVar.append(notSetNotes)
        if platform.system()=='Windows':
            notSetNotesBox=ttk.Entry(notSetFrame,textvariable=notSetNotes,width=40)
        if platform.system()=='Darwin':
            notSetNotesBox=ttk.Entry(notSetFrame,textvariable=notSetNotes,width=20)
        notSetNotesLabel=tk.Label(notSetFrame, text="Note:")
        #bind per la ricerca nei menu di clienti e task
        notSetSelectDrop.bind('<KeyRelease>', checkNotSetKeyClients)
        #posizionamento degli widget nella finestra per gli slot mancanti
        clientHeader.grid(row=0,column=1,padx=10,pady=5)
        notSetLabel.grid(row=notSetWindowRow,column=0,padx=10,pady=5)
        notSetSelectDrop.grid(row=notSetWindowRow,column=1,padx=10,pady=5)
        notSetNotesLabel.grid(row=notSetWindowRow,column=2,pady=5)
        notSetNotesBox.grid(row=notSetWindowRow,column=3,pady=5)
        notSetWindowRow+=1
    notSetFrame.columnconfigure(0,weight=1)
    notSetFrame.columnconfigure(1,weight=4)
    notSetFrame.columnconfigure(2,weight=1)
    notSetFrame.columnconfigure(3,weight=1)
    notSetFrame.pack(anchor="w")
    buttonFrame=tk.Frame(notSetWindow)
    confirmNotSetButton=ttk.Button(buttonFrame,text='Conferma',command=confirmNotSet,style='Mod.TButton')
    confirmNotSetButton.grid(columnspan=4,pady=5)
    confirmNotSetButton.configure(state='disabled')
    buttonFrame.pack()
    if platform.system()=='Windows':
        notSetHeight=32*(len(notSetComboVar)+1)+32
        notSetWidth=1250
    if platform.system()=='Darwin':
        notSetHeight=37*(len(notSetComboVar)+1)+37
        notSetWidth=1370
    xNotSet=(screenWidth/2)-(notSetWidth/2)
    yNotSet=(screenHeight/2)-(notSetHeight/2)
    notSetWindow.geometry('%dx%d+%d+%d' % (notSetWidth, notSetHeight, xNotSet, yNotSet))
    listNotSetVarCombo=list(zip(notSetComboVar,notSetComboWidgets))

def confirmNotSet():#evento del pulsante di conferma della finestra degli slot mancanti
    global notSetRowList,notSetText,notSetComboVarGet,notSetClientChoiceSplit,notSetNotesVar,notSetNotesVarGet,confirmNotSetButton,notSetRootButton
    try:
        confirmNotSetButton.configure(state='disabled')
        notSetNotesVarGet=[]
        for i in notSetNotesVar:
            notSetNotesVarGet.append(i.get())
        listClientsNotes=list(zip(notSetComboVarGet,notSetNotesVarGet))
        dictNotSetRowClients=dict(zip(notSetRowList,listClientsNotes))
        for row,client in dictNotSetRowClients.items():
            notSetClientChoiceSplit=client[0].split(' | ')
            if len(notSetClientChoiceSplit)==4:
                sheet1.cell(row=row,column=7).value=notSetClientChoiceSplit[3]
            sheet1.cell(row=row,column=4).value=notSetClientChoiceSplit[0]
            sheet1.cell(row=row,column=5).value=notSetClientChoiceSplit[1]
            sheet1.cell(row=row,column=6).value=notSetClientChoiceSplit[2]
            sheet1.cell(row=row,column=8).value=client[1]
        confirmNotSetButton['text']='Slot mancanti annotati.'
        notSetText.set("0 slot senza una commessa associata.")
        wb.save('AutoGestionale.xlsx')
        notSetWindow.after(1000,notSetWindow.withdraw)
    except PermissionError as e: #in caso il file excel fosse aperto
        if str(e)=="[Errno 13] Permission denied: 'AutoGestionale.xlsx'":
            notSetWindow.withdraw()
            tk.messagebox.showwarning(title='File Excel aperto',message='Il file Excel AutoGestionale risulta aperto, chiudilo prima di procedere con la conferma della commessa.',parent=root)
            notSetWindow.deiconify()
            notSetRootButton.configure(state='normal')
            notSetCount=len(notSetRowList)
            notSetText.set(str(notSetCount)+" slot senza una commessa associata.")
            confirmNotSetButton.configure(state='normal')
            confirmNotSetButton['text']='Conferma'
            
        
def destroyNotSetWidget():#distrugge gli widget della finestra degli slot mancanti
    for child in notSetWindow.winfo_children():
        child.destroy()
                                                                                                                                                    
def notSetWindowOpen():#evento del pulsante di apertura della finestra degli slot mancanti (precedenti)
    notSetWindow.deiconify()
            
def NotSetRootButtonRelease(*args):#attiva il pulsante di cui sopra e disattiva il pulsante di registrazione in gestionale se ci sono slot mancanti
    global notSetRootButton,gestionaleButton
    startsWithZero=re.compile(r'^0')
    if re.match(startsWithZero,notSetText.get()):
        notSetRootButton.configure(state='disabled')
        gestionaleButton.configure(state='normal')
    else:
        notSetRootButton.configure(state='normal')
        gestionaleButton.configure(state='disabled')

def checkBeforeGestionale():
    global listMissingVarCombo,missingComboVar,confirmMissingButton,missingNotesVar,missingRowList,checkBefGestionaleLoop
    if sheet1.cell(row=17,column=4).value==None:
        alertGestionaleWindow=tk.messagebox.askquestion(title='Slot mancanti',
                                                     message='Non sono stati annotati tutti gli slot orari fino alle 18. Vuoi associare delle commesse prima di registrare le ore in Gestionale?',)
        if alertGestionaleWindow=='yes':
            checkBefGestionaleLoop+=1
            if checkBefGestionaleLoop==1:
                missingRowList=[]
                for missingClient in (sheet1.cell(row=i,column=4) for i in range(2,18)):
                    if missingClient.value==None:
                        missingRowList.append(missingClient.row)
                missingWindowRow=1
                missingHeaderFrame=tk.Frame(lastMissingWindow)
                missingClientHeader=tk.Label(missingHeaderFrame,text='Seleziona una commessa per ogni slot mancante e inserisci una nota (opzionale):')
                missingClientHeader.grid(row=0,column=0,padx=10,pady=5)
                missingHeaderFrame.pack(anchor="w")
                missingComboWidgets=[]
                missingComboVar=[]
                missingNotesVar=[]
                missingFrame=tk.Frame(lastMissingWindow)
                for missingRow in missingRowList:#crea una riga nella finestra per ogni slot di mezz'ora mancante
                    #crea gli widget della finestra per gli slot mancanti
                    missingLabel=tk.Label(missingFrame,text="Dalle "+str("{:02d}".format(sheet1.cell(row=missingRow,column=9).value))+":"+str("{:02d}".format(sheet1.cell(row=missingRow,column=10).value))+" alle "+str("{:02d}".format(sheet1.cell(row=missingRow,column=11).value))+":"+str("{:02d}".format(sheet1.cell(row=missingRow,column=12).value))+":")
                    missingSelectClick=tk.StringVar()
                    missingSelectClick.set('')
                    missingComboVar.append(missingSelectClick)
                    missingSelectClick.trace('w',checkMissingClient)
                    if platform.system()=='Windows':
                        missingSelectDrop=ttk.Combobox(missingFrame,textvariable=missingSelectClick,width=130)
                    if platform.system()=='Darwin':
                        missingSelectDrop=ttk.Combobox(missingFrame,textvariable=missingSelectClick,width=102)
                    missingComboWidgets.append(missingSelectDrop)
                    missingSelectDrop['values']=listClientSort
                    missingNotes=tk.StringVar()
                    missingNotes.set('')
                    missingNotesVar.append(missingNotes)
                    if platform.system()=='Windows':
                        missingNotesBox=ttk.Entry(missingFrame,textvariable=missingNotes,width=40)
                    if platform.system()=='Darwin':
                        missingNotesBox=ttk.Entry(missingFrame,textvariable=missingNotes,width=20)
                    missingNotesLabel=tk.Label(missingFrame, text="Note:")
                    #bind per la ricerca nei menu di clienti e task
                    missingSelectDrop.bind('<KeyRelease>', checkMissingKeyClients)
                    #posizionamento degli widget nella finestra per gli slot mancanti
                    missingClientHeader.grid(row=0,column=1,padx=10,pady=5)
                    missingLabel.grid(row=missingWindowRow,column=0,padx=10,pady=5)
                    missingSelectDrop.grid(row=missingWindowRow,column=1,padx=10,pady=5)
                    missingNotesLabel.grid(row=missingWindowRow,column=2,pady=5)
                    missingNotesBox.grid(row=missingWindowRow,column=3,pady=5)
                    missingWindowRow+=1
                missingFrame.columnconfigure(0,weight=1)
                missingFrame.columnconfigure(1,weight=4)
                missingFrame.columnconfigure(2,weight=1)
                missingFrame.columnconfigure(3,weight=1)
                missingFrame.pack(anchor="w")
                missingButtonFrame=tk.Frame(lastMissingWindow)
                confirmMissingButton=ttk.Button(missingButtonFrame,text='Conferma',command=confirmMissing,style='Mod.TButton')
                confirmMissingButton.grid(columnspan=4,pady=5)
                confirmMissingButton.configure(state='disabled')
                missingButtonFrame.pack()
                if platform.system()=='Windows':
                    missingHeight=32*(len(missingComboVar)+1)+32
                    missingWidth=1250
                if platform.system()=='Darwin':
                    missingHeight=37*(len(missingComboVar)+1)+37
                    missingWidth=1370
                xMissing=(screenWidth/2)-(missingWidth/2)
                yMissing=(screenHeight/2)-(missingHeight/2)
                lastMissingWindow.geometry('%dx%d+%d+%d' % (missingWidth, missingHeight, xMissing, yMissing))
                listMissingVarCombo=list(zip(missingComboVar,missingComboWidgets))
                lastMissingWindow.deiconify()
            else:
                lastMissingWindow.deiconify()
        else:
            gestionaleButton.configure(state='disabled')
            threading.Thread(target=gestionaleOpen,daemon=True).start()
    else:
        gestionaleButton.configure(state='disabled')
        threading.Thread(target=gestionaleOpen,daemon=True).start()

def confirmMissing():#evento del pulsante di conferma della finestra degli slot mancanti
    global missingRowList,missingComboVarGet,missingClientChoiceSplit,missingNotesVar,missingNotesVarGet
    try:
        confirmMissingButton.configure(state='disabled')
        missingNotesVarGet=[]
        for i in missingNotesVar:
            missingNotesVarGet.append(i.get())
        listMissingClientsNotes=list(zip(missingComboVarGet,missingNotesVarGet))
        dictMissingRowClients=dict(zip(missingRowList,listMissingClientsNotes))
        for row,client in dictMissingRowClients.items():
            missingClientChoiceSplit=client[0].split(' | ')
            if ' COSTI ' in missingClientChoiceSplit[1].upper():
                sheet1.cell(row=row,column=7).value=missingClientChoiceSplit[3]
            sheet1.cell(row=row,column=4).value=missingClientChoiceSplit[0]
            sheet1.cell(row=row,column=5).value=missingClientChoiceSplit[1]
            sheet1.cell(row=row,column=6).value=missingClientChoiceSplit[2]
            sheet1.cell(row=row,column=8).value=client[1]
        confirmMissingButton['text']='Slot mancanti annotati.'
        wb.save('AutoGestionale.xlsx')
        lastMissingWindow.after(2000,lastMissingWindow.withdraw)
        gestionaleButton.configure(state='disabled')
        threading.Thread(target=gestionaleOpen,daemon=True).start()
    except PermissionError as e: #in caso il file excel fosse aperto
        if str(e)=="[Errno 13] Permission denied: 'AutoGestionale.xlsx'":
            lastMissingWindow.withdraw()
            tk.messagebox.showwarning(title='File Excel aperto',message='Il file Excel AutoGestionale risulta aperto, chiudilo prima di procedere con la conferma della commessa.',parent=root)
            lastMissingWindow.deiconify()
            confirmMissingButton.configure(state='normal')
            confirmMissingButton['text']='Conferma'

def gestionaleOpen():
    try:
        #messaggio di registrazione ore in Gestionale
        mainTextBox.configure(state='normal')
        clearMainTextBox()
        mainTextBox.insert('end', "Registrazione delle ore in Gestionale in corso...",'tagCenter')
        mainTextBox.configure(state='disabled')
        root.after(5000,root.iconify)
        global sheet1,wb
        sheet1Copy=wb.copy_worksheet(sheet1)
        checkDuplicateRow=3
        #somma l'orario della commessa attuale con la precedente se sono uguali e consecutive
        while sheet1Copy.cell(row=checkDuplicateRow,column=4).value!=None:
            if sheet1Copy.cell(row=checkDuplicateRow,column=4).value==sheet1Copy.cell(row=checkDuplicateRow-1,column=4).value and \
               sheet1Copy.cell(row=checkDuplicateRow,column=5).value==sheet1Copy.cell(row=checkDuplicateRow-1,column=5).value and \
               sheet1Copy.cell(row=checkDuplicateRow,column=6).value==sheet1Copy.cell(row=checkDuplicateRow-1,column=6).value and \
               sheet1Copy.cell(row=checkDuplicateRow,column=7).value==sheet1Copy.cell(row=checkDuplicateRow-1,column=7).value and \
               sheet1Copy.cell(row=checkDuplicateRow,column=8).value==sheet1Copy.cell(row=checkDuplicateRow-1,column=8).value and \
               (sheet1Copy.cell(row=checkDuplicateRow,column=9).value==sheet1Copy.cell(row=checkDuplicateRow-1,column=11).value and sheet1Copy.cell(row=checkDuplicateRow,column=10).value==sheet1Copy.cell(row=checkDuplicateRow-1,column=12).value):
                    sheet1Copy.cell(row=checkDuplicateRow-1,column=11).value=sheet1Copy.cell(row=checkDuplicateRow,column=11).value
                    sheet1Copy.cell(row=checkDuplicateRow-1,column=12).value=sheet1Copy.cell(row=checkDuplicateRow,column=12).value
                    sheet1Copy.cell(row=checkDuplicateRow-1,column=14).value+=30
                    if sheet1Copy.cell(row=checkDuplicateRow-1,column=14).value==60:
                        sheet1Copy.cell(row=checkDuplicateRow-1,column=13).value+=1
                        sheet1Copy.cell(row=checkDuplicateRow-1,column=14).value=0
                    sheet1Copy.delete_rows(checkDuplicateRow, 1)
            else:
                checkDuplicateRow+=1
        browser=webdriver.Chrome()
        browser.get('https://gestionale.com')
        userElem=browser.find_element(By.NAME, "__ac_name")
        userElem.send_keys(user.get())
        passwordElem=browser.find_element(By.NAME, "__ac_password")
        passwordElem.send_keys(password.get())
        passwordElem.submit()
        browser.get('https://gestionale.com/main/commesse')
        endProgram=False
        for client in sheet1Copy["D2:D17"]: #loop percorre ogni cliente registrato, se la cella è vuota si ferma
            if endProgram==True:
                    break
            for cell in client:
                if cell.value==None:
                        del wb[wb.sheetnames[2]]
                        wb.save('AutoGestionale.xlsx')
                        wb.close()
                        endProgram=True
                        break
                if cell.value=='NOT SET':
                    continue
                currClient=sheet1Copy.cell(row=cell.row,column=4).value
                currDesc1=sheet1Copy.cell(row=cell.row,column=5).value
                currDesc2=sheet1Copy.cell(row=cell.row,column=6).value
                currStartHour="{:02d}".format(int(sheet1Copy.cell(row=cell.row,column=9).value))
                currStartMin="{:02d}".format(int(sheet1Copy.cell(row=cell.row,column=10).value))
                currHoursWorked="{:02d}".format(int(sheet1Copy.cell(row=cell.row,column=13).value))
                currMinsWorked="{:02d}".format(int(sheet1Copy.cell(row=cell.row,column=14).value))
                clientElement=WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, """//*[text()[contains(normalize-space(.),"%s")]]/span[text()[contains(normalize-space(.),"%s")]]/../preceding-sibling::td[text()[contains(normalize-space(.),"%s")]]"""% (currDesc2,currDesc1,currClient))))
                clientElement.click()
                detailElement=WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='div_dettaglio_commessa']/div[2]/table/tbody")))
                detailElement.click()
                startHourElement=WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@name,'hh_inizio')]/option[contains(@value,'"+str(currStartHour)+"')]")))
                startHourElement.click()
                startMinElement=WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@name,'mm_inizio')]/option[contains(@value,'"+str(currStartMin)+"')]")))
                startMinElement.click()
                hoursWorkedElement=WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@name,'hh_')]")))
                hoursWorkedElement.send_keys(currHoursWorked)
                minsWorkedElement=WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@name,'mm_')]")))
                minsWorkedElement.send_keys(currMinsWorked)
                if sheet1Copy.cell(row=cell.row,column=7).value!=None:
                    currTask=sheet1Copy.cell(row=cell.row,column=7).value
                    taskDropElement=WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.NAME, "id_todo")))
                    taskDropElement.click()
                    taskElement=WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, """//select[contains(@name,"id_todo")]//*[contains(text(),"%s")]"""% (currTask))))
                    taskElement.click()
                if sheet1Copy.cell(row=cell.row,column=8).value!=None:
                    currNotes=sheet1Copy.cell(row=cell.row,column=8).value
                    notesElement=WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class,'tabellaInterna')]//textarea[contains(@name,'descrizione')]")))
                    notesElement.send_keys(currNotes)
                confirmElement=WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='1']//input[contains(@value,'Conferma')]")))
                confirmElement.click()
                alert_obj = browser.switch_to.alert
                alert_obj.accept()
                browser.get('https://gestionale.com/main/commesse')
        if endProgram==True:
            closingAlert=tk.Toplevel(root)
            closingAlert.title('Programma in chiusura')
            closingAlertLabel=tk.Label(closingAlert, text="Le ore della giornata sono state registrate in Gestionale. Chiusura del programma in corso...")
            closingAlertLabel.pack(padx=20, pady=20)
            if platform.system()=='Windows':
                closingHeight=65
                closingWidth=550
            if platform.system()=='Darwin':
                closingHeight=65
                closingWidth=600
            global screenWidth,screenHeight
            xClosing=(screenWidth/2)-(closingWidth/2)
            yClosing=(screenHeight/2)-(closingHeight/2)
            closingAlert.geometry('%dx%d+%d+%d' % (closingWidth, closingHeight, xClosing, yClosing))
            if platform.system()=='Windows':
                closingAlert.iconbitmap(resource_path("logo-gestionale-win.ico"))
            closingAlert.attributes("-topmost", True)
            closingAlert.resizable(False,False)
            mainTextBox.configure(state='normal')
            clearMainTextBox()
            mainTextBox.insert('end', "Le ore della giornata sono state registrate in Gestionale. Chiusura del programma in corso...",'tagCenter')
            mainTextBox.configure(state='disabled')
            s.close()
            root.after(3000,rootDestroy)
    except WebDriverException: #errore generico
        tk.messagebox.showerror(title='Errore',message='Un errore ha impedito la corretta registrazione delle ore. Assicurati di non chiudere la finestra di Chrome automatizzata e di essere collegato a internet. Se sono state già registrate alcune ore per oggi, cancellale prima di riprovare.',parent=root)
        del wb[wb.sheetnames[2]]
        wb.save('AutoGestionale.xlsx')
        gestionaleButton.configure(state='normal')
    except com_error as e:
        if e.excepinfo[5]==-2147352570:
            tk.messagebox.showerror(title='Errore',message='Un errore ha impedito la corretta registrazione delle ore. Chiudi i file Excel attualmente aperti e riprova.',parent=root)
            gestionaleButton.configure(state='normal')
    except AttributeError as e: #nel caso in cui la finestra di chrome venga chiusa
        if str(e)=="'NoneType' object has no attribute 'is_displayed'":
            tk.messagebox.showerror(title='Errore',message='Un errore ha impedito la corretta registrazione delle ore. Assicurati di non chiudere la finestra di Chrome automatizzata e di essere collegato a internet. Se sono state già registrate alcune ore per oggi, cancellale prima di riprovare.',parent=root)
            del wb[wb.sheetnames[2]]
            wb.save('AutoGestionale.xlsx')
            gestionaleButton.configure(state='normal')

def clearMainTextBox():
    mainTextBox.delete(1.0, tk.END)

def checkKeyClients(event):
    global listClientSort
    value=event.widget.get()
    if value=='':
        filterClients=listClientSort
    else:
        filterClients=[]
        for client in listClientSort:
            if value.lower() in client.lower():
                filterClients.append(client)
    updateSelectDrop(filterClients)

def updateSelectDrop(filterClients):
    selectDrop['values']=(filterClients)

def checkNotSetKeyClients(event):
    global listClientSort,notSetComboVar,listNotSetVarCombo
    for var in listNotSetVarCombo:
        value=var[0].get()
        if value=='':
            filterClients=listClientSort
        else:
            filterClients=[]
            for client in listClientSort:
                if value.lower() in client.lower():
                    filterClients.append(client)
        var[1]['values']=(filterClients)

def checkMissingKeyClients(event):
    global listClientSort,listMissingVarCombo
    for var in listMissingVarCombo:
        value=var[0].get()
        if value=='':
            filterClients=listClientSort
        else:
            filterClients=[]
            for client in listClientSort:
                if value.lower() in client.lower():
                    filterClients.append(client)
        var[1]['values']=(filterClients)

def updateMissingSelectDrop(filterClients):
    global missingSelectDrop
    missingSelectDrop['values']=(filterClients)

def rootDestroy():
    root.quit()
    root.destroy()
    sys.exit()

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

#creazione finestra principale
root=tk.Tk()
if platform.system()=='Windows':
    rootWidth=850
    rootHeight=250
    root.iconbitmap(resource_path("logo-gestionale-win.ico"))
if platform.system()=='Darwin':
    rootWidth=1000
    rootHeight=275
    image_path=tk.PhotoImage(file=os.path.join(os.getcwd(),'logo-gestionale.png'))
    root.iconphoto(True,image_path)
screenWidth=root.winfo_screenwidth()
screenHeight=root.winfo_screenheight()
xRoot=(screenWidth//2)-(rootWidth//2)
yRoot=(screenHeight//2)-(rootHeight//2)
root.geometry('%dx%d+%d+%d' % (rootWidth, rootHeight, xRoot, yRoot))
root.title('AutoGestionale')
root.resizable(False,False)
root.protocol('WM_DELETE_WINDOW', rootDestroy)
#creazione widget finestra principale
mainFrame=tk.Frame()
selectClick=tk.StringVar()
selectClick.trace('w',checkClient)
selectLabelText=tk.StringVar()
selectLabelText.set("Seleziona una commessa:")
selectLabel=ttk.Label(mainFrame,textvariable=selectLabelText,anchor='w',justify='left')
if platform.system()=='Windows':
    selectDrop=ttk.Combobox(mainFrame,textvariable=selectClick,width=130)
if platform.system()=='Darwin':
    selectDrop=ttk.Combobox(mainFrame,textvariable=selectClick,width=102)
notesLabel=ttk.Label(mainFrame,text="Inserisci una nota (opzionale):",anchor='w',justify='left')
notesEntry=tk.StringVar()
if platform.system()=='Windows':
    notesText=ttk.Entry(mainFrame,textvariable=notesEntry,width=133)
if platform.system()=='Darwin':
    notesText=ttk.Entry(mainFrame,textvariable=notesEntry,width=103,style='Mod.TEntry')
selectButton=ttk.Button(mainFrame,text='Conferma',command=confirm,style='Mod.TButton')
mainTextFrame=tk.Frame()
if platform.system()=='Windows':
    mainTextBox=tk.Text(mainTextFrame, width=133,height=1,wrap=tk.WORD,bg='#f0f0f0',font=('TkDefaultFont'))
if platform.system()=='Darwin':
    mainTextBox=tk.Text(mainTextFrame, width=103,height=1,wrap=tk.WORD,bg='#f6f5f6',font=('TkDefaultFont'))
mainTextBox.tag_configure('tagCenter', justify='center')
CommandsFrame=tk.Frame()
notSetText=tk.StringVar()
notSetText.set('')
notSetRootLabel=ttk.Label(CommandsFrame,textvariable=notSetText)
notSetRootButton=ttk.Button(CommandsFrame,text='Compila gli slot mancanti',command=notSetWindowOpen,style='Mod.TButton')
gestionaleLabel=ttk.Label(CommandsFrame,text="Quando hai terminato inserisci le ore in Gestionale:")
gestionaleButton=ttk.Button(CommandsFrame,text='Apri Gestionale',command=checkBeforeGestionale,style='Mod.TButton')
notSetText.trace('w',NotSetRootButtonRelease)
notSetText.set("0 slot senza una commessa associata.")

#applicazione widget a finestra principale
selectLabel.grid(row=0,columnspan=3,pady=(10,0))
selectDrop.grid(row=1,columnspan=3,pady=(0,10))
notesLabel.grid(row=2,columnspan=3)
notesText.grid(row=3,columnspan=3,pady=(0,10))
selectButton.grid(row=4,columnspan=3)
notSetRootLabel.grid(row=5,column=0,padx=15)
notSetRootButton.grid(row=6,column=0,padx=15)
notSetRootButton.configure(state='disabled')
gestionaleLabel.grid(row=5,column=1,padx=15)
gestionaleButton.grid(row=6,column=1,padx=15)
mainTextBox.pack(pady=20,padx=10)
mainFrame.pack()
mainTextFrame.pack()
CommandsFrame.pack()
CommandsFrame.columnconfigure(0,weight=1,uniform='second')
CommandsFrame.columnconfigure(1,weight=1,uniform='second')

#attiva il filtro dei menu drop down quando si scrive nell'apposito campo
selectDrop.bind('<KeyRelease>', checkKeyClients)

#creazione finestra secondaria login
loginWindow=tk.Toplevel()
if platform.system()=='Windows':
    loginHeight=140
    loginWidth=240
if platform.system()=='Darwin':
    loginHeight=160
    loginWidth=280
xLogin=(screenWidth/2)-(loginWidth/2)
yLogin=(screenHeight/2)-(loginHeight/2)
loginWindow.geometry('%dx%d+%d+%d' % (loginWidth, loginHeight, xLogin, yLogin))
loginWindow.title('Login')
if platform.system()=='Windows':
    loginWindow.iconbitmap(resource_path("logo-gestionale-win.ico"))
loginFrame=tk.Frame(loginWindow)
loginTextFrame=tk.Frame(loginWindow)
loginWindow.attributes("-topmost", True)
loginWindow.resizable(False,False)
loginWindow.protocol('WM_DELETE_WINDOW', rootDestroy)
#creazione widget login
loginLabel=tk.Label(loginFrame,text="Accedi a Gestionale:")
userLabel=tk.Label(loginFrame, text="User Name")
user=tk.StringVar()
userEntry=ttk.Entry(loginFrame, textvariable=user)
passwordLabel=tk.Label(loginFrame,text="Password")
password=tk.StringVar()
passwordEntry=ttk.Entry(loginFrame, textvariable=password, show='*')
loginButton=ttk.Button(loginFrame,text='Login',command=login)
#applicazione widget login a finestra secondaria
loginLabel.grid(row=0,column=0,columnspan=2,pady=10)
userLabel.grid(row=1,column=0,pady=2)
userEntry.grid(row=1,column=1,pady=2)
passwordLabel.grid(row=2,column=0,pady=2)
passwordEntry.grid(row=2,column=1,pady=2)
loginButton.grid(row=3,columnspan=2,pady=10)
loginWindow.bind('<Return>',login)
loginFrame.pack()
loginTextFrame.pack()

#creazione finestra per gli slot di mezz'ora in Not Set
notSetWindow=tk.Toplevel()
notSetWindow.title('Slot mancanti')
if platform.system()=='Windows':
    notSetWindow.iconbitmap(resource_path("logo-gestionale-win.ico"))
notSetWindow.protocol('WM_DELETE_WINDOW', notSetWindow.withdraw)
notSetWindow.resizable(False,False)

#creazione finestra di registrazione degli slot di mezz'ora che mancano fino alle 18
lastMissingWindow=tk.Toplevel()
lastMissingWindow.title('Slot mancanti')
if platform.system()=='Windows':
    lastMissingWindow.iconbitmap(resource_path("logo-gestionale-win.ico"))
lastMissingWindow.protocol('WM_DELETE_WINDOW', lastMissingWindow.withdraw)
lastMissingWindow.resizable(False,False)

confirmNotSetButton=None #dichiara l'oggetto del pulsante conferma della finestra degli slot in not set

styleButton=ttk.Style(root)
styleButton.map('Mod.TButton',
          foreground=[('disabled', 'grey')])

styleEntry=ttk.Style(root)
styleEntry.map('Mod.TEntry',
          fieldbackground=[('disabled', '#f6f5f6')])

notSetWindow.withdraw()
lastMissingWindow.withdraw()
root.withdraw()
root.mainloop()