from selenium import webdriver
from time import sleep
from selenium.webdriver.common import keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common import exceptions  
from selenium.common.exceptions import ElementNotInteractableException, NoSuchElementException  
from selenium.webdriver.chrome.options import Options 
from colorama import Fore
import warnings
import os
import xlsxwriter
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=UserWarning)


options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument("--start-maximized")
browser = webdriver.Chrome(options=options)
browser.get("https://www.nesine.com/iddaa?et=1&le=3&ocg=MS-2%2C5&gt=Popüler")

workbook = xlsxwriter.Workbook('analiz.xlsx')
sheet = workbook.add_worksheet("cal")


sheet.write("A1","Takım Başlıgı")
sheet.write("B1","(Lig) Son 6 Maçın Golleri")
sheet.write("C1","(Aralarindaki) Toplam Golleri")
sheet.write("D1","(Aralarindaki) KG Var Oranı")
sheet.write("E1","(Aralarindaki) Toplam İY Golleri")
sheet.write("F1","(Aralarindaki) İY 0.5 Üst Oranı")
sheet.write("G1","(Kariyer) Son Tüm Maçların Oranı")
sheet.write("H1","(Aralarindaki) Kazanma Sayısı")


def scroll():
    sleep(5)
    for i in range(20):
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
     
def searchlink():
    scroll()
    linklist = []
    for i in browser.find_elements_by_css_selector(".code-time-name>.name>a"):
        try:
            link = i.get_attribute("href")
            if link=="javascript:void(0);":
                pass
            else:
                linklist.append(link)
        except:
            pass


    linklistcount = 0
    excelacount = 2

    linkcounterrr = len(linklist)
    for i in range(linkcounterrr):
        browser.get(linklist[linklistcount])

        try:
            #sıralama

            


            #son 6 maç gol
            teamhome = browser.find_element_by_css_selector("#summary-point-table-content > div > div > div > div:nth-child(2) > div.table__row.home > div.team-name > a").text
            teamaway = browser.find_element_by_css_selector("#summary-point-table-content > div > div > div > div:nth-child(2) > div.table__row.away > div.team-name > a").text
            sheet.write(f"A{excelacount}",f" {teamhome} - {teamaway} ")
            sas = 2
            teamgoalhome = 0
            for i in range(6):
                
                homename =browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(2) > div > div:nth-child(3) > div > div > div > div:nth-child(1) > div.table.last-match.highlight-last-match.summary > div:nth-child({sas}) > div.table__row__middle.last-match-widget.tennis-lmw > div:nth-child(1) > span > span > a").text
                goal =browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(2) > div > div:nth-child(3) > div > div > div > div:nth-child(1) > div.table.last-match.highlight-last-match.summary > div:nth-child({sas}) > div.table__row__middle.last-match-widget.tennis-lmw > div.scoreBtn > button > div").text
                if homename==teamhome:
                    myarray = []
                    myarray.append(goal)
                    code = myarray[0][0:1]
                    teamgoalhome+=int(code)
                else:
                    myarray = []
                    myarray.append(goal)
                    code = myarray[0][4:5]
                    teamgoalhome+=int(code)
                sas+=1
            sas = 2
            teamgoalaway = 0
            for i in range(6):
                
                homename =browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(2) > div > div:nth-child(3) > div > div > div > div:nth-child(2) > div.table.last-match.highlight-last-match.summary > div:nth-child({sas}) > div.table__row__middle.last-match-widget.tennis-lmw > div:nth-child(1) > span > span > a").text
                goal =browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(2) > div > div:nth-child(3) > div > div > div > div:nth-child(2) > div.table.last-match.highlight-last-match.summary > div:nth-child({sas}) > div.table__row__middle.last-match-widget.tennis-lmw > div.scoreBtn > button > div").text
                if homename==teamaway:
                    myarray = []
                    myarray.append(goal)
                    code = myarray[0][0:1]
                    teamgoalaway+=int(code)
                else:
                    myarray = []
                    myarray.append(goal)
                    code = myarray[0][4:5]
                    teamgoalaway+=int(code)
                sas+=1
            sheet.write(f"B{excelacount}",f" {teamgoalhome} - {teamgoalaway}")
            

            #aralarındaki son maçların gol istatistigi
            try:
                browser.find_element_by_css_selector("#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.data-display-none.border > p").text
                teamgoalhome2 = "N"
                teamgoalaway2 = "N"
                deger = "N"
                kgcounter = "N"
            except:
                nthcounter = 1
                teamgoalhome2 = 0
                teamgoalaway2 = 0
                dosya = browser.find_elements_by_css_selector("#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > .table__row")
                deger = 0
                for satir in dosya:
                    deger+=1
                for x in range(deger):
                    
                    firstteam = browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > div:nth-child({nthcounter}) > div.table__row__middle > div > a").text
                    scoreteam = browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > div:nth-child({nthcounter}) > div.table__row__middle > div.scoreBtn > button > div").text
                    
                    nthcounter+=1
                    if firstteam==teamaway:
                        myarray = []
                        myarray.append(scoreteam)
                        code = myarray[0][0:1]
                        teamgoalaway2+=int(code)
                    else:
                        myarray = []
                        myarray.append(scoreteam)
                        code = myarray[0][4:5]
                        teamgoalaway2+=int(code)
                    
                    if firstteam==teamhome:
                        myarray = []
                        myarray.append(scoreteam)
                        code = myarray[0][0:1]
                        teamgoalhome2+=int(code)
                    else:
                        myarray = []
                        myarray.append(scoreteam)
                        code = myarray[0][4:5]
                        teamgoalhome2+=int(code)
            sheet.write(f"C{excelacount}",f" {teamgoalhome2} - {teamgoalaway2}")
        


            #kg kontrol 
            dosya = browser.find_elements_by_css_selector("#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > .table__row")
            deger = 0
            for satir in dosya:
                deger+=1
            nthcounter = 1
            kgcounter = 0
            for x in range(deger):
                score = browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > div:nth-child({nthcounter}) > div.table__row__middle > div.scoreBtn > button > div").text
                nthcounter+=1
                myarray = []
                myarray.append(score)
                goal1 = myarray[0][4:5]
                goal2 = myarray[0][0:1]
                if int(goal1)>1 or int(goal1)==1:
                    if int(goal2)>1 or int(goal2)==1:
                        kgcounter+=1
            sheet.write(f"D{excelacount}",f" {deger} / {kgcounter}")
        
            #aralarındaki maç win
            teamhome = browser.find_element_by_css_selector("#summary-point-table-content > div > div > div > div:nth-child(2) > div.table__row.home > div.team-name > a").text
            teamaway = browser.find_element_by_css_selector("#summary-point-table-content > div > div > div > div:nth-child(2) > div.table__row.away > div.team-name > a").text
                        
            dosya = browser.find_elements_by_css_selector("#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > .table__row")
            deger = 0
            for satir in dosya:
                deger+=1
            teamcounter = 1
            winhomecount = 0
            winawaycount = 0
            drawcount = 0
            for x in range(deger):
                try:
                    teamwin = browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > div:nth-child({teamcounter}) > div.table__row__middle > div.winner > a").text
                    if teamwin==teamhome:
                        winhomecount+=1
                    if teamwin==teamaway:
                        winawaycount+=1
                    
                except NoSuchElementException:
                    drawcount+=1
                teamcounter+=1
            sheet.write(f"H{excelacount}",f"{deger}M - {winhomecount}W {drawcount}D {winawaycount}W")

            #iy gol
            try:
                browser.find_element_by_css_selector("#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.data-display-none.border > p").text
                deger = "N"
                matchcounterveriable = "N"
                teamscorehome = "N"
                teamscoreaway = "N"
                iygoalcounter = "N"
            except:
                    
                teamhome = browser.find_element_by_css_selector("#summary-point-table-content > div > div > div > div:nth-child(2) > div.table__row.home > div.team-name > a").text
                teamaway = browser.find_element_by_css_selector("#summary-point-table-content > div > div > div > div:nth-child(2) > div.table__row.away > div.team-name > a").text
                dosya = browser.find_elements_by_css_selector("#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > .table__row")
                deger = 0
                for satir in dosya:
                    deger+=1
                scorecounter = 1
                teamscorehome = 0
                teamscoreaway = 0
                matchcounterveriable = 0
                for i in range(deger):
                    try:

                        score = browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > div:nth-child({scorecounter}) > div.iy-draw").text
                        firstteam = browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > div:nth-child({scorecounter}) > div.table__row__middle > div:nth-child(1) > a").text
                        scorecounter+=1
                        if firstteam==teamhome:
                            myarray = []
                            myarray.append(score)
                            code = myarray[0][0:1]
                            teamscorehome+=int(code)
                        else:
                            myarray = []
                            myarray.append(score)
                            code = myarray[0][2:3]
                            teamscorehome+=int(code)
                        if firstteam==teamaway:
                            myarray = []
                            myarray.append(score)
                            code = myarray[0][0:1]
                            teamscoreaway+=int(code)
                        else:
                            myarray = []
                            myarray.append(score)
                            code = myarray[0][2:3]
                            teamscoreaway+=int(code)
                        matchcounterveriable+=1
                    except:
                        pass 
                scorecounter = 1
                iygoalcounter = 0
                for x in range(deger):
                    try:

                        score = browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div:nth-child(1) > div.col-xs-12.col-md-8.padding-right > div:nth-child(1) > div > div:nth-child(2) > div > div > div > div.table_body > div:nth-child({scorecounter}) > div.iy-draw").text
                        scorecounter+=1
                        myarray = []
                        myarray.append(score)
                        goal1 = myarray[0][2:3]
                        goal2 = myarray[0][0:1]
                        goals = int(goal1)+int(goal2)
                        if goals>1 or goals==1:
                            iygoalcounter+=1
                    except:
                        pass
            sheet.write(f"F{excelacount}",f"{matchcounterveriable} / {iygoalcounter}")
            sheet.write(f"E{excelacount}",f"{deger} / {matchcounterveriable} - ({teamscorehome} - {teamscoreaway})")

            #son tüm maçların gol istatistigi
            browser.get(linklist[linklistcount]+"/son-maclari")
            dosya = browser.find_elements_by_css_selector("#root > div:nth-child(2) > div.panels.last-matches-page.clearfix.false > div.page-content > div:nth-child(4) > div:nth-child(1) > div.table.last-match.page > div.table__row.clearfix")
            degerhome = 0
            for satir in dosya:
                degerhome+=1
            classcounterhome = 2
            wincounthome = 0
            losecounthome = 0
            drawcounthome = 0
            for x in range(degerhome):
                scorclass = browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div.panels.last-matches-page.clearfix.false > div.page-content > div:nth-child(4) > div:nth-child(1) > div.table.last-match.page > div:nth-child({classcounterhome}) > div.table__row__middle > div.scoreBtn > button > div").get_attribute("class")
                if scorclass=="win":
                    wincounthome+=1
                if scorclass=="lose":
                    losecounthome+=1
                if scorclass=="draw":
                    drawcounthome+=1
                classcounterhome+=1

            dosya = browser.find_elements_by_css_selector("#root > div:nth-child(2) > div.panels.last-matches-page.clearfix.false > div.page-content > div:nth-child(4) > div:nth-child(2) > div.table.last-match.page > div.table__row.clearfix")
            degeraway = 0
            for satir in dosya:
                degeraway+=1
            classcounteraway = 2
            wincountaway = 0
            losecountaway = 0
            drawcountaway = 0
            for x in range(degeraway):
                scorclass = browser.find_element_by_css_selector(f"#root > div:nth-child(2) > div.panels.last-matches-page.clearfix.false > div.page-content > div:nth-child(4) > div:nth-child(2) > div.table.last-match.page > div:nth-child({classcounteraway}) > div.table__row__middle > div.scoreBtn > button > div").get_attribute("class")
                if scorclass=="win":
                    wincountaway+=1
                if scorclass=="lose":
                    losecountaway+=1
                if scorclass=="draw":
                    drawcountaway+=1
                classcounteraway+=1
            sheet.write(f"G{excelacount}",f"{degerhome} ({wincounthome}W {losecounthome}L {drawcounthome}D) / {degeraway} ({wincountaway}W {losecountaway}L {drawcountaway}D)")

                        
            
            
            
            
            excelacount+=1
  
        except:
            pass
        linklistcount+=1

          
searchlink()
workbook.close()