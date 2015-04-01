import clr
clr.AddReferenceToFile("WatiN.Core.dll")

from WatiN.Core import *
from System import *
from System.Threading import Thread
from math import floor

ie1 = IE.AttachTo[IE](Find.ByTitle("Gestione Giornaliera"))
ie2 = IE.AttachTo[IE](Find.ByTitle("Giustificativo Singolo"))

from datetime import datetime


fom = DateTime(int(ie1.SelectList(Find.ByName("cmbAnni")).SelectedOption.Value), int(ie1.SelectList(Find.ByName("cmbMesi")).SelectedOption.Value), 1)
eom = fom.AddMonths(1).AddDays(-1)

day1 = int(raw_input("Start=") or fom.Day)
day2 = int(raw_input("End=") or eom.Day)

print "%i - %i of %s" % (day1, day2, fom.ToString("MMM yyyy"))

for day in range(day1, day2+1):
    ie1.Link(Find.ByText(str(day))).Click()
    Thread.Sleep(1000)
    f = ie1.Frame(Find.ByName("prestazioni"))
    t = f.Tables[0].Tables[0]
    for r in t.TableRows:
        if r.TableCells.Count >= 7:
            if r.TableCells[4].Text == "Presenza Oltre Monte Ore Teorico":
                inizio = r.TableCells[5].Text.strip()
                fine = r.TableCells[6].Text.strip()
                

                date = DateTime(fom.Year, fom.Month, day)
                ie2.TextField(Find.ByName("cmbCodCausale")).Value = '001'        # straordinari
                #ie2.TextField(Find.ByName("cmbCodCausale")).Value = '008'        # banca ore
                ie2.TextField(Find.ByName("datai")).Value = date.ToString("dd/MM/yyyy")
                ie2.TextField(Find.ByName("dataf")).Value = date.ToString("dd/MM/yyyy")
                iHH, iMM = inizio.split(":")
                fHH, fMM = fine.split(":")
                
                iDT = DateTime(date.Year, date.Month, day, int(iHH), int(iMM), 0)
                fDT = DateTime(date.Year, date.Month, day, int(fHH), int(fMM), 0)
                
                ts = fDT - iDT
                # print "%i\t%s\t%s - " % (day, inizio, fine)
                
                mins = ts.TotalMinutes
                
                rmins = floor(mins/30)*30
                if rmins < 60:
                    rmins = 0
                
                print "%s - %s (%s mins rounded to %s - lost %s)" % (iDT, fDT, mins, rmins, mins - rmins)
                
                fDT = iDT.AddMinutes(rmins)
                fHH = "%02i" % fDT.Hour
                fMM = "%02i" % fDT.Minute
                
                if rmins > 0:
                    try:
                        ie2.SelectList(Find.ByName("cmbOraInizioPezzaHH")).Option(Find.ByValue(iHH)).Select()
                    except:
                        pass
                    try:
                        ie2.SelectList(Find.ByName("cmbOraInizioPezzaMM")).Option(Find.ByValue(iMM)).Select()
                    except:
                        pass
                    try:
                        ie2.SelectList(Find.ByName("cmbOraFinePezzaHH")).Option(Find.ByValue(fHH)).Select()
                    except:
                        pass
                    try:
                        ie2.SelectList(Find.ByName("cmbOraFinePezzaMM")).Option(Find.ByValue(fMM)).Select()
                    except:
                        pass
                    ie2.Button(Find.ById("buttonInserisci")).Click()
                    
                    #Console.ReadLine()
                    
                    Thread.Sleep(1000)

