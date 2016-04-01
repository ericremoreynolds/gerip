using System;
using System.Threading;
using WatiN.Core;

class Gerip
{
    [STAThread]
    public static void Main()
    {

        var ie1 = IE.AttachTo<IE>(Find.ByTitle("Gestione Giornaliera"));
        var ie2 = IE.AttachTo<IE>(Find.ByTitle("Giustificativo Singolo"));

        var fom = new DateTime(
            Int32.Parse(ie1.SelectList(Find.ByName("cmbAnni")).SelectedOption.Value),
            Int32.Parse(ie1.SelectList(Find.ByName("cmbMesi")).SelectedOption.Value),
            1
        );
        var eom = fom.AddMonths(1).AddDays(-1);

        Console.Write("Start=");
        var day1 = Int32.Parse(Console.ReadLine());
        Console.Write("End=");
        var day2 = Int32.Parse(Console.ReadLine());
        Console.Write("Banca Ore? [n]");
        var bo = Console.ReadLine().ToUpper() == "Y";

        Console.WriteLine("{0} - {1} of {2}", day1, day2, fom.ToString("MMM yyyy"));

        for(int day=day1; day<=day2; day++)
        {
            ie1.Link(Find.ByText(day.ToString())).Click();
            Thread.Sleep(1000);
            var f = ie1.Frame(Find.ByName("prestazioni"));
			var t = f.Table(Find.ById("tabellaPrestazioniSAVFR"));
            foreach(var r in t.TableRows)
            {
                if(r.TableCells.Count >= 7)
                {
                    if(r.TableCells[4].Text == "Presenza Oltre Monte Ore Teorico")
                    {
                        var inizio = r.TableCells[5].Text.Trim();
                        var fine = r.TableCells[6].Text.Trim();

                        var date = new DateTime(fom.Year, fom.Month, day);
                        if(bo)
                            ie2.TextField(Find.ByName("cmbCodCausale")).Value = "008";        // banca ore
                        else
                            ie2.TextField(Find.ByName("cmbCodCausale")).Value = "001";       // straordinari
                        ie2.TextField(Find.ByName("datai")).Value = date.ToString("dd/MM/yyyy");
                        ie2.TextField(Find.ByName("dataf")).Value = date.ToString("dd/MM/yyyy");
                        var iXX = inizio.Split(':');
                        var iHH = iXX[0];
                        var iMM = iXX[1];
                        var fXX = fine.Split(':');
                        var fHH = fXX[0];
                        var fMM = fXX[1];
                        
                        var iDT = new DateTime(date.Year, date.Month, day, Int32.Parse(iHH), Int32.Parse(iMM), 0);
                        var fDT = new DateTime(date.Year, date.Month, day, Int32.Parse(fHH), Int32.Parse(fMM), 0);
                        
                        var ts = fDT - iDT;
                        // print "%i\t%s\t%s - " % (day, inizio, fine)
                        
                        var mins = (int) ts.TotalMinutes;
                        
                        var rmins = (int) Math.Floor(mins / 30.0) * 30;
                        if(rmins < 60)
                            rmins = 0;
                        
                        Console.WriteLine("{0} - {1} ({2} mins rounded to {3} - lost {4})", iDT, fDT, mins, rmins, mins - rmins);
                        
                        fDT = iDT.AddMinutes(rmins);
                        fHH = fDT.Hour.ToString("00");
                        fMM = fDT.Minute.ToString("00");
                        
                        if(rmins > 0)
                        {
                            try
                            {
                                ie2.SelectList(Find.ByName("cmbOraInizioPezzaHH")).Option(Find.ByValue(iHH)).Select();
                            }
                            catch
                            {
                            }

                            try
                            {
                                ie2.SelectList(Find.ByName("cmbOraInizioPezzaMM")).Option(Find.ByValue(iMM)).Select();
                            }
                            catch
                            {
                            }
                            try
                            {
                                ie2.SelectList(Find.ByName("cmbOraFinePezzaHH")).Option(Find.ByValue(fHH)).Select();
                            }
                            catch
                            {
                            }
                            try
                            {
                                ie2.SelectList(Find.ByName("cmbOraFinePezzaMM")).Option(Find.ByValue(fMM)).Select();
                            }
                            catch
                            {
                            }
                            
                            ie2.Button(Find.ById("buttonInserisci")).Click();
                            
                            // Console.ReadLine()
                            
                            Thread.Sleep(1000);
                        }
                    }
                }
            }
        }
    }
}