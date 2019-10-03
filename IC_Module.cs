using System;
using System.Linq;
using System.IO;
using System.IO.IsolatedStorage;
using System.Collections.Generic;
using System.Windows.Input;

using Microsoft.LightSwitch;
using Microsoft.LightSwitch.Framework.Client;
using Microsoft.LightSwitch.Presentation;
using Microsoft.LightSwitch.Presentation.Extensions;
using System.Collections.Specialized;
using System.Runtime.InteropServices.Automation;
using System.Text.RegularExpressions;
using System.Windows.Controls;

/*
Insurance Commissions Module Main Act Screen
Victor Doshenko
2012-2019
C#: Microsoft Visual Lightswitch
All Rights Reserved
*/

namespace LightSwitchApplication
{
    public partial class Акты
    {

        partial void Акты_Created()
        {
            try
            {
                this.FindControl("QueryActGrid").ControlAvailable += Acts_ControlAvailable;
                this.FindControl("CommissionsGrid").ControlAvailable += Commissions_ControlAvailable;
            }
            catch (System.InvalidOperationException) { }

            if (insurance_company != "" && insurance_company != null)
            {
                КомпанияЛукапДляАктов = (from x in КомпанииЛукапДляАктов
                                         where x.Company == insurance_company
                                       select x).FirstOrDefault();
            }
            else
            {
                КомпанияЛукапДляАктов = (from x in КомпанииЛукапДляАктов
                                        where x.id == 1
                                       select x).FirstOrDefault();
            }

        }

        private DataGrid _itemsControl_commis = null;
        #region  Event Handlers

        private void Commissions_ControlAvailable(object send, ControlAvailableEventArgs e)
        {
            _itemsControl_commis = e.Control as DataGrid;

            if (_itemsControl_commis == null)
            {
                return;
            }

            _itemsControl_commis.SelectionMode = DataGridSelectionMode.Extended;
        }

        private void ActGridKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                ActEditID = QueryAct.SelectedItem.ID;
            }
        }
        #endregion

        partial void Status_Changed()
        {
            if (Status == null || Status == 0)
            {
                Status1b = 1; Status1e = 5;
                Status2b = 1; Status2e = 5;
            }
            else
            if (Except)
            {
                if (Status > 1)
                {
                    Status1b = 1; Status1e = Status.Value - 1;
                }
                else
                {
                    Status1b = 2; Status1e = 5;
                }
                if (Status < 5)
                {
                    Status2b = Status.Value + 1; Status2e = 5;
                }
                else
                {
                    Status2b = 1; Status2e = 4;
                }
            }
            else
            {
                Status1b = Status.Value; Status1e = Status.Value;
                Status2b = Status.Value; Status2e = Status.Value;
            }
            QueryAct.Load();
        }

        partial void QueryActDeleteSelected_CanExecute(ref bool result)
        {
            if (QueryAct.SelectedItem != null)
            {
                result = QueryAct.SelectedItem.Status != 2 &&
                         QueryAct.SelectedItem.Status != 4 &&
                         QueryAct.SelectedItem.Status != 5;
            }
            else
            {
                result = false;
            }
        }

        partial void QueryActDeleteSelected_Execute()
        {
            if (this.QueryAct.SelectedItem.Status != 3)
            {
                this.QueryAct.SelectedItem.Delete();
            }
            else
                if (this.ShowMessageBox("Удалить акт?", "Акт еще не оплачен", MessageBoxOption.YesNo) == System.Windows.MessageBoxResult.Yes)
                {
                    this.QueryAct.SelectedItem.Delete();
                }
        }
        
        partial void ActToExcel_Execute()
        {
            string sd = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\InsuranceCommissions";
            if (!System.IO.Directory.Exists(sd))
            {
                System.IO.Directory.CreateDirectory(sd);
            }
            sd = sd + "\\" + QueryAct.SelectedItem.Year;
            if (!System.IO.Directory.Exists(sd))
            {
                System.IO.Directory.CreateDirectory(sd);
            }
            sd = sd + "\\" + QueryAct.SelectedItem.Month;
            if (!System.IO.Directory.Exists(sd))
            {
                System.IO.Directory.CreateDirectory(sd);
            }
            sd = sd + "\\" + QueryAct.SelectedItem.Source2;
            if (!System.IO.Directory.Exists(sd))
            {
                System.IO.Directory.CreateDirectory(sd);
            }

            string ExcelFile = sd + "\\" + QueryAct.SelectedItem.Company.Replace("\"", "").Replace(" ", "_") + ".xls";
            string VBSFile = sd + "\\" + QueryAct.SelectedItem.Company.Replace("\"", "").Replace(" ", "_") + ".vbs";

                if (AutomationFactory.IsAvailable)
                {
                    dynamic shell = AutomationFactory.CreateObject("Shell.Application");
                    Company = QueryAct.SelectedItem.Company;
                    РеквизитыКомпаний.Load();
                    РеквизитыКомпаний.FirstOrDefault();

                    try
                    {
                        System.IO.File.Delete(ExcelFile);
                    }
                    catch (System.IO.IOException) {
                        this.ShowMessageBox("Excel файл уже открыт, необходимо закрыть файл и запустить отчет заново.");
                        return;
                    }

                    FileStream f = System.IO.File.OpenWrite(ExcelFile);
                    string DateExcel;
                    DateExcel = QueryAct.SelectedItem.c_Date.ToString().Substring(0, 10);
                    if (QueryAct.SelectedItem.OKFromCompDate != null)
                    {
                        DateExcel = QueryAct.SelectedItem.OKFromCompDate.ToString().Substring(0, 10);
                    }
                    this.Application.fwrite(f,
                              "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta http-equiv=Content-Type content=\"text/html; charset=windows-1251\"><meta name=ProgId content=Excel.Sheet><meta name=Generator content=\"Microsoft Excel 11\"></head><BODY bgcolor=\"#FFFFFF\"><br><div align=center>Акт об оказанных услугах №<b>" + QueryAct.SelectedItem.Number + "</b><br>за период с <b>"
                              + new DateTime(QueryAct.SelectedItem.Year, QueryAct.SelectedItem.Month, 1).ToString().Substring(0, 10) + " по "
                              + new DateTime(QueryAct.SelectedItem.Year, QueryAct.SelectedItem.Month, 1).AddMonths(1).AddDays(-1).ToString().Substring(0, 10) + "</b><br>между <b>ООО \"Банк ПСА Финанс РУС\"</b> и <b>"
                              + РеквизитыКомпаний.SelectedItem.Company + "</b><br>по агентскому договору №<b>"
                              + РеквизитыКомпаний.SelectedItem.AgentDeal + "</b><br></div><div align=right>_<u>"
                              + DateExcel + "</u>_<br><br><br></div>\r\n<TABLE border=1>\r\n<th>№ п/п</th>\r\n<th>Номер договора</th>\r\n<th>Дата договора</th>\r\n<th>Год финансирования</th>\r\n<th>Регион </th>\r\n<th>Страхователь (ФИО)</th>\r\n<th>№ Страхового полиса</th>\r\n<th>VIN</th>\r\n<th>Город</th>\r\n<th>Год страхования ТС</th>\r\n<th>Дата начала действия полиса</th>\r\n<th>Способ оплаты</th>\r\n<th>Страховая премия</th>\r\n<th>Номер взноса</th>\r\n<th>Сумма взноса</th>\r\n<th>Дата взноса</th>\r\n<th>% комиссии</th>\r\n<th>Сумма комиссии (включая НДС)</th>\r\n<th>Место покупки полиса</th>\r\n<th>Основание</th>\r\n<th>Примечание</th>\r\n\r\n");

                    int i = 0;
                    double SumQty = 0;
                    string s;

                    IOrderedEnumerable<Комиссия> КомиссияС = (from x in Комиссия orderby x.Comment, x.Договор.Number select x);

                    foreach (Комиссия OneCom in КомиссияС)
                    {
                        i++;
                        s = "<tr><td align=right>" + i.ToString() + "</td><td>";
                        s = s + OneCom.Договор.Number.TrimEnd() + "</td><td>";
                        if (OneCom.Договор.DealDate.ToString() != "" && OneCom.Договор.DealDate != null)
                            s = s + OneCom.Договор.DealDate.ToString().Substring(0, 10);
                        s = s + "</td><td>";
                        s = s + OneCom.Source2.ToString() + "</td><td>";
                        s = s + OneCom.Договор.Region + "</td><td>";
                        s = s + OneCom.Договор.FIO + "</td><td>";
                        s = s + OneCom.NumKASKO + "</td><td>";
                        s = s + OneCom.VIN + "</td><td>";                        
                        s = s + OneCom.DealerCity + "</td><td>";
                        s = s + OneCom.YearInsurance + "</td><td>";
                        if (OneCom.DateKASKO != null)
                        {
                            s = s + OneCom.DateKASKO.ToString().Substring(0, 10);
                        }
                        s = s + "</td><td>";
                        s = s + OneCom.TypePay + "</td><td align=right style='mso-number-format:Standard;'>";

                        if (OneCom.PremiumQty != null)
                        {
                            s = s + Math.Round(OneCom.PremiumQty.Value, 2);
                        };
                        s = s + "</td><td align=left style='mso-number-format:\"\\@\";'>";
                        s = s + OneCom.install_num + "</td><td>";
                        if (OneCom.install_qty != null)
                        {
                            s = s + Math.Round(OneCom.install_qty.Value, 2);
                        };
                        
                        s = s + "</td><td>";
                        if (OneCom.install_date != null && OneCom.install_date.ToString().Length >= 10 )
                        {
                            s = s + OneCom.install_date.ToString().Substring(0, 10);
                        }
                        s = s + "</td><td align=right style='mso-number-format:Percent;'>";
                        s = s + OneCom.Prc + "</td><td align=right style='mso-number-format:Standard;'>";
                        if (OneCom.ComQty != null)
                        {
                            s = s + Math.Round(OneCom.ComQty.Value, 2);
                            SumQty = SumQty + (double)OneCom.ComQty.Value;
                        }
                        s = s + "</td><td>";
                        s = s + OneCom.place + "</td><td>";
                        s = s + OneCom.Comment + "</td><td>";
                        s = s + OneCom.Comment2 + "</td></tr>";
                        s = s + "\r\n";
                        this.Application.fwrite(f, s);
                    }
                    this.Application.fwrite(f, "<tr><td colspan=7 align=right>ИТОГО:</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td align=right style='mso-number-format:Standard;'>" + SumQty + "</td><td></td><td></td></tr></TABLE><br>");
                    string[] months = {"Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"};
                    this.Application.fwrite(f, "<TABLE border=0><tr><td colspan=4></td><td colspan=8>Услуга Исполнителем оказана надлежащим образом. К оплате за " + months[new DateTime(QueryAct.SelectedItem.Year, QueryAct.SelectedItem.Month, 1).Month - 1] + " месяц " + new DateTime(QueryAct.SelectedItem.Year, QueryAct.SelectedItem.Month, 1).Year + " года:</td></tr>");

                    this.Application.fwrite(f, "<tr><td colspan=2></td><td colspan=2 align=center style='mso-number-format:Standard;'><b>" + SumQty + "</b></td><td colspan=8>" + this.Application.num2str(SumQty) + "</td></tr>");
                    this.Application.fwrite(f, "<tr><td></td><td colspan=3 align=center>Итого к оплате:</td><td colspan=8>" + SumQty.ToString("0.00").Replace(".", ",") + " " + this.Application.num2str(SumQty) + "</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4></td><td align=center><b>в т.ч. НДС</b></td><td colspan=2 align=center style='mso-number-format:Standard;'><b>" + Math.Round(SumQty * /*18*/ this.Application.NDS / (this.Application.NDS + 100) /*118*/, 2).ToString("0.00") + "</b></td><td colspan=8>" + this.Application.num2str(Math.Round(SumQty * this.Application.NDS / (this.Application.NDS + 100) /*18 / 118*/, 2)) + "</td></tr></TABLE><br>");

                    this.Application.fwrite(f, "<div align=center>ПОДПИСИ СТОРОН:</div><br><TABLE border=0><tr><td></td><td colspan=4 align=center>");
                    this.Application.fwrite(f, "<TABLE border=0>");
                    this.Application.fwrite(f, "<tr><td colspan=4><b>ООО \"Банк ПСА Финанс РУС\"</b></td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4></td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>_________________________________" + this.Application.ActSignatory + "</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>                              м.п.</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4></td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>Адрес: 105120, Россия, город Москва, 2-й Сыромятнический переулок, дом 1, 7 этаж</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>ИНН 7750004288, КПП 770901001</td></tr>");
                    this.Application.fwrite(f, "<tr></tr>");

                    this.Application.fwrite(f, "<tr><td colspan=4>БИК 044525339</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>Корр/cчет: № 30101810845250000339 в Отделении 1 Москва</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>Банк получателя: ООО «Банк ПСА Финанс РУС»</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>р/с		" + РеквизитыКомпаний.SelectedItem.Account + "</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>ОКПО 86555411, ОГРН 1087711000024</td></tr>");
                    
                    this.Application.fwrite(f, "<tr><td colspan=4></td></tr><tr><td colspan=4></td></tr>");
                    this.Application.fwrite(f, "</TABLE></td><td></td><td></td><td></td><td colspan=7 align=center><TABLE border=0>");
                    this.Application.fwrite(f, "<tr><td colspan=4><b>" + РеквизитыКомпаний.SelectedItem.Company + "</b></td></tr>");

                    this.Application.fwrite(f, "<tr><td colspan=4></td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>_________________________________                        "+РеквизитыКомпаний.SelectedItem.Signer+"</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>                              м.п.</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4></td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>Адрес: " + РеквизитыКомпаний.SelectedItem.Address+"</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>ИНН " + РеквизитыКомпаний.SelectedItem.INN+"</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>БИК " + РеквизитыКомпаний.SelectedItem.BIC+"</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>корр./счет " + РеквизитыКомпаний.SelectedItem.CorAcc + "</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>р/с " + РеквизитыКомпаний.SelectedItem.CurAcc+" в " + РеквизитыКомпаний.SelectedItem.Bank+ "</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>ОГРН " + РеквизитыКомпаний.SelectedItem.OGRN+"</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>ОКПО " + РеквизитыКомпаний.SelectedItem.OKPO + "</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>ОКОНХ " + РеквизитыКомпаний.SelectedItem.OKONH + "</td></tr>");
                    this.Application.fwrite(f, "<tr><td colspan=4>Тел. " + РеквизитыКомпаний.SelectedItem.Phone + "</td></tr>");
                    this.Application.fwrite(f, "</TABLE></td></tr></TABLE></BODY>\r\n</HTML>\r\n\r\n");

                    f.Close();
                    f = System.IO.File.OpenWrite(VBSFile);
                    this.Application.fwrite(f, "Const xlValidateList = 3\r\n");
                    this.Application.fwrite(f, "Const xlThin = 2\r\n");
                    this.Application.fwrite(f, "Const xlContinuous = 1\r\n");
                    this.Application.fwrite(f, "Const xlDown = -4121\r\n");
                    this.Application.fwrite(f, "Const xlOpenXMLWorkbook = 51\r\n");
                    this.Application.fwrite(f, "Const xlR1C1 = -4150\r\n");

                    this.Application.fwrite(f, "Set fso = CreateObject(\"Scripting.FileSystemObject\")\r\n");
                    this.Application.fwrite(f, "strFilePath=fso.GetParentFolderName(wscript.ScriptFullName) & \"\\\"\r\n");
                    this.Application.fwrite(f, "strExtension=fso.GetExtensionName(wscript.ScriptFullName)\r\n");
                    this.Application.fwrite(f, "strFileName=Replace(Wscript.ScriptName, \".\" & strExtension, \"\")\r\n");
                    this.Application.fwrite(f, "strFileFullname=strFilePath & strFileName & \".xls\"\r\n");
                    this.Application.fwrite(f, "strFileFullnameX=strFileFullname & \"x\"\r\n");
                    this.Application.fwrite(f, "strFileNameX=strFileName & \".xlsx\"\r\n");
                    this.Application.fwrite(f, "Set objExcel = CreateObject(\"Excel.Application\")\r\n");
                    this.Application.fwrite(f, "objExcel.Visible = False\r\n");
                    this.Application.fwrite(f, "Set wrbook = objExcel.Workbooks.Open(strFileFullname) \r\n");

                    this.Application.fwrite(f, "objExcel.Cells(1, 22).FormulaR1C1 = \"полис не сдан/ не найден\"\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(2, 22).FormulaR1C1 = \"полис не оплачен\"\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(3, 22).FormulaR1C1 = \"номер полиса не принадлежит нашей страховой компании\"\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(4, 22).FormulaR1C1 = \"не найден (просьба уточнить данные)\"\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(5, 22).FormulaR1C1 = \"нет пролонгации\"\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(6, 22).FormulaR1C1 = \"Выгодоприобреталь не Банк ПСА Финанс РУС\"\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(7, 22).FormulaR1C1 = \"Согласовано\"\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(1, 22).Font.ThemeColor = 1\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(2, 22).Font.ThemeColor = 1\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(3, 22).Font.ThemeColor = 1\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(4, 22).Font.ThemeColor = 1\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(5, 22).Font.ThemeColor = 1\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(6, 22).Font.ThemeColor = 1\r\n");
                    this.Application.fwrite(f, "objExcel.Cells(7, 22).Font.ThemeColor = 1\r\n");
                    this.Application.fwrite(f, "objExcel.Columns(20).ColumnWidth = 14\r\n");
                    this.Application.fwrite(f, "objExcel.Columns(21).ColumnWidth = 28\r\n");

                    this.Application.fwrite(f, "Set objRange = objExcel.Range(\"A9\")\r\n");
                    this.Application.fwrite(f, "imax = objRange.End(xlDown).Row\r\n");
                    this.Application.fwrite(f, "i = 9\r\n");

                    this.Application.fwrite(f, "RS = objExcel.ReferenceStyle\r\n");
                    this.Application.fwrite(f, "objExcel.ReferenceStyle = xlR1C1\r\n");

                    this.Application.fwrite(f, "While i < imax And objExcel.Cells(i, 1).Value <> \"ИТОГО:\"\r\n");
                    this.Application.fwrite(f, "  objExcel.Cells(i, 20).Validation.Add xlValidateList,,,\"Согласовано;Перенос;Отказ\"\r\n");
                    this.Application.fwrite(f, "  objExcel.Cells(i, 20).Validation.IgnoreBlank = False\r\n");
                    this.Application.fwrite(f, "  objExcel.Cells(i, 21).Validation.Add xlValidateList,,,\"=IF(RC[-1]=\"\"Перенос\"\";R1C22:R2C22;IF(RC[-1]=\"\"Отказ\"\";R3C22:R6C22;R7C22))\"\r\n");
                    this.Application.fwrite(f, "  objExcel.Cells(i, 21).Validation.IgnoreBlank = False\r\n");
                    this.Application.fwrite(f, "  objExcel.Cells(i, 21).BorderAround xlContinuous, xlThin\r\n");

                    this.Application.fwrite(f, "  i = i + 1\r\n");
                    this.Application.fwrite(f, "WEnd\r\n");
                    this.Application.fwrite(f, "objExcel.ReferenceStyle = RS\r\n");
                    this.Application.fwrite(f, "objExcel.Application.DisplayAlerts = False\r\n");
                    this.Application.fwrite(f, "wrbook.SaveAs strFileFullnameX, xlOpenXMLWorkbook\r\n");
                    this.Application.fwrite(f, "objExcel.Application.DisplayAlerts = True\r\n");
                    this.Application.fwrite(f, "objExcel.Visible = True\r\n");
                    this.Application.fwrite(f, "Dim oShell\r\n");
                    this.Application.fwrite(f, "Set oShell = CreateObject(\"WScript.Shell\")\r\n");
                    this.Application.fwrite(f, "oShell.AppActivate strFileNameX\r\n");
                    this.Application.fwrite(f, "Set wrbook = Nothing\r\n");
                    this.Application.fwrite(f, "Set objExcel = Nothing\r\n");
                    this.Application.fwrite(f, "Set oShell = Nothing\r\n");
                    this.Application.fwrite(f, "Set fso = Nothing\r\n");
                    this.Application.fwrite(f, "Set objRange = Nothing\r\n");
                    f.Close();
                    shell.ShellExecute(VBSFile);
                }
                else
                {
                    this.ShowMessageBox("Automation not available");
                }
        }

        partial void ImportFromExcel_CanExecute(ref bool result)
        {
            result = this.Application.User.HasPermission(Permissions.SecurityAdministration);
        }

        partial void Акты_InitializeDataWorkspace(List<IDataService> saveChangesTo)
        {
            if (insurance_company != "" && insurance_company != null)
            {
                КомпанияЛукапДляАктов = (from x in КомпанииЛукапДляАктов
                                         where x.Company == insurance_company
                                         select x).FirstOrDefault();
            }
        }

        private DataGrid _itemsControl_ActGrid = null;
        #region  Event Handlers
        private void Acts_ControlAvailable(object send, ControlAvailableEventArgs e)
        {
            _itemsControl_ActGrid = e.Control as DataGrid;

            if (_itemsControl_ActGrid == null)
            {
                return;
            }

            _itemsControl_ActGrid.SelectionMode = DataGridSelectionMode.Extended;
        }
        #endregion

        partial void QueryActAddAndEditNew_Execute()
        {
            this.Application.ShowДобавитьАкт(КомпанияЛукапДляАктов.Company);
        }

        partial void QueryActEditSelected_Execute()
        {
            this.Application.ShowАктРедактировать(QueryAct.SelectedItem.ID);
        }

        partial void ComMassUpdate_Execute()
        {
            if (_itemsControl_commis == null) { return; }
            
            if (_itemsControl_commis.SelectedItems.Count == 0)
            {
                this.ShowMessageBox("Не выбрано ни одной записи. Необходимо отметить одну или более записей!");
                return;
            }
            status_new = 2;
            this.OpenModalWindow("ModalWindowStatus");
        }

        partial void OK_Status_Update_Execute()
        {
            if (this.ShowMessageBox("Поменять статусы у отмеченных комиссий?", "(Всего комиссий отмечено: " + _itemsControl_commis.SelectedItems.Count.ToString() + ")", MessageBoxOption.YesNo) == System.Windows.MessageBoxResult.No)
            {
                this.CloseModalWindow("ModalWindowStatus");
                return;
            }
            int TotalComUpdated;
            TotalComUpdated = 0;
            foreach (Комиссия item in _itemsControl_commis.SelectedItems)
            {
                item.Status = status_new;
                TotalComUpdated += 1;
            }
            this.CloseModalWindow("ModalWindowStatus");
            this.DataWorkspace.insuranceData.SaveChanges();
            this.ShowMessageBox("Всего комиссий изменено: " + TotalComUpdated.ToString());
        }

        partial void Cancel_Status_Update_Execute()
        {
            this.CloseModalWindow("ModalWindowStatus");
        }

        partial void Except_Changed()
        {
            Status_Changed();
        }

        partial void CommissionsGridEditSelected_Execute()
        {
            this.Application.ShowКомиссияДетали(Комиссия.SelectedItem.ID);
        }

        partial void Акты_Activated()
        {
            this.FindControl("QueryActGrid").Focus();
            this.Application.ActSignatory = (from n in this.DataWorkspace.insuranceData.Настройки
                                            where n.Name == "ActSignatory"
                                           select n).FirstOrDefault().StrValue;
        }

        partial void Акты_Closing(ref bool cancel)
        {
            insurance_company = null;
        }

        partial void КомпанияЛукапДляАктов_Changed()
        {
            QueryAct.Load();
            if (КомпанияЛукапДляАктов != null)
            {
                insurance_company = КомпанияЛукапДляАктов.Company;
            };
        }

        partial void InAct_Execute()
        {
            if (_itemsControl_commis == null) { return; }

            if (_itemsControl_commis.SelectedItems.Count == 0)
            {
                this.ShowMessageBox("Не выбрано ни одной записи. Необходимо отметить одну или более записей!");
                return;
            }
            foreach (Комиссия item in _itemsControl_commis.SelectedItems)
            {
                item.InAct = false;
            }
        }

        partial void ActEditID_Changed()
        {
            this.Application.ShowАктРедактировать(ActEditID);
        }

        partial void QueryActEditSelected_CanExecute(ref bool result)
        {
            result = (QueryAct.SelectedItem != null);
        }

        partial void CommissionsGridDeleteSelected_Execute()
        {
            if (_itemsControl_commis == null) { return; }

            if (_itemsControl_commis.SelectedItems.Count > 0)
            {
                List<decimal> IDs_del = new List<decimal>();

                foreach (Комиссия item in _itemsControl_commis.SelectedItems)
                {
                    IDs_del.Add(item.ID);
                }
                foreach (decimal ID_del in IDs_del)
                {
                    Комиссия com_d;
                    com_d = (from x in this.DataWorkspace.insuranceData.Комиссии
                             where x.ID == ID_del
                             select x).FirstOrDefault();
                    if (com_d != null)
                    {
                        com_d.Delete();
                    }
                }
            }
        }

        partial void ActMassUpdate_Execute()
        {
            if (_itemsControl_ActGrid == null) { return; }

            if (_itemsControl_ActGrid.SelectedItems.Count == 0)
            {
                this.ShowMessageBox("Не выбрано ни одной записи. Необходимо отметить одну или более записей!");
                return;
            }
            status_new_act = 1;
            this.OpenModalWindow("ModalWindowStatusAct");
        }

        partial void OK_Status_Act_Update_Execute()
        {
            if (this.ShowMessageBox("Поменять статусы у отмеченных актов?", "(Всего актов отмечено: " + _itemsControl_ActGrid.SelectedItems.Count.ToString() + ")", MessageBoxOption.YesNo) == System.Windows.MessageBoxResult.No)
            {
                this.CloseModalWindow("ModalWindowStatusAct");
                return;
            }
            int TotalActUpdated;
            TotalActUpdated = 0;
            foreach (Акт item in _itemsControl_ActGrid.SelectedItems)
            {
                item.Status = status_new_act;
                TotalActUpdated += 1;
            }
            this.CloseModalWindow("ModalWindowStatusAct");
            this.DataWorkspace.insuranceData.SaveChanges();
            this.ShowMessageBox("Всего актов изменено: " + TotalActUpdated.ToString());
        }

        partial void Cancel_Status_Act_Update_Execute()
        {
            this.CloseModalWindow("ModalWindowStatusAct");
        }

        partial void ActToExcelOld_Execute()
        {

            string sd = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\InsuranceCommissions";
            if (!System.IO.Directory.Exists(sd))
            {
                System.IO.Directory.CreateDirectory(sd);
            }
            sd = sd + "\\" + QueryAct.SelectedItem.Year;
            if (!System.IO.Directory.Exists(sd))
            {
                System.IO.Directory.CreateDirectory(sd);
            }
            sd = sd + "\\" + QueryAct.SelectedItem.Month;
            if (!System.IO.Directory.Exists(sd))
            {
                System.IO.Directory.CreateDirectory(sd);
            }
            sd = sd + "\\" + QueryAct.SelectedItem.Source2;
            if (!System.IO.Directory.Exists(sd))
            {
                System.IO.Directory.CreateDirectory(sd);
            }

            string ExcelFile = sd + "\\" + QueryAct.SelectedItem.Company.Replace("\"", "").Replace(" ", "_") + ".xls";

            if (AutomationFactory.IsAvailable)
            {
                dynamic shell = AutomationFactory.CreateObject("Shell.Application");
                Company = QueryAct.SelectedItem.Company;
                РеквизитыКомпаний.Load();
                РеквизитыКомпаний.FirstOrDefault();

                try
                {
                    System.IO.File.Delete(ExcelFile);
                }
                catch (System.IO.IOException)
                {
                    this.ShowMessageBox("Excel файл уже открыт, необходимо закрыть файл и запустить отчет заново.");
                    return;
                }

                FileStream f = System.IO.File.OpenWrite(ExcelFile);
                string DateExcel;
                DateExcel = QueryAct.SelectedItem.c_Date.ToString().Substring(0, 10);
                if (QueryAct.SelectedItem.OKFromCompDate != null)
                {
                    DateExcel = QueryAct.SelectedItem.OKFromCompDate.ToString().Substring(0, 10);
                }
                this.Application.fwrite(f, "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta http-equiv=Content-Type content=\"text/html; charset=windows-1251\"><meta name=ProgId content=Excel.Sheet><meta name=Generator content=\"Microsoft Excel 11\"></head><BODY bgcolor=\"#FFFFFF\"><br><div align=center>Акт об оказанных услугах №<b>" + QueryAct.SelectedItem.Number + "</b><br>за период с <b>"
                          + new DateTime(QueryAct.SelectedItem.Year, QueryAct.SelectedItem.Month, 1).ToString().Substring(0, 10) + " по "
                          + new DateTime(QueryAct.SelectedItem.Year, QueryAct.SelectedItem.Month, 1).AddMonths(1).AddDays(-1).ToString().Substring(0, 10) + "</b><br>между <b>ООО \"Банк ПСА Финанс РУС\"</b> и <b>"
                          + РеквизитыКомпаний.SelectedItem.Company + "</b><br>по агентскому договору №<b>"
                          + РеквизитыКомпаний.SelectedItem.AgentDeal + "</b><br></div><div align=right>_<u>"
                          + DateExcel + "</u>_<br><br><br></div>\r\n<TABLE border=1>\r\n<th>№ п/п</th>\r\n<th>Номер договора</th>\r\n<th>Дата договора</th>\r\n<th>Регион </th>\r\n<th>Страхователь (ФИО)</th>\r\n<th>Вид страхования</th>\r\n<th>№ Страхового полиса</th>\r\n<th>Город</th>\r\n<th>Страховая премия</th>\r\n<th>Год страхования ТС</th>\r\n<th>Дата начала действия полиса</th>\r\n<th>Способ оплаты</th>\r\n<th>% комиссии</th>\r\n<th>Сумма комиссии (включая НДС)</th>\r\n<th>Основание</th>\r\n\r\n");

                int i = 0;
                double SumQty = 0;
                string s;

                IOrderedEnumerable<Комиссия> КомиссияС = (from x in Комиссия orderby x.Comment, x.Договор.Number select x); //x.Город, x.ComQty select x); 

                foreach (Комиссия OneCom in КомиссияС)
                {
                    i++;
                    s = "<tr><td align=right>" + i.ToString() + "</td><td>";
                    s = s + OneCom.Договор.Number.TrimEnd() + "</td><td>";
                    if (OneCom.Договор.DealDate.ToString() != "" && OneCom.Договор.DealDate != null)
                        s = s + OneCom.Договор.DealDate.ToString().Substring(0, 10);
                    s = s + "</td><td>";
                    s = s + OneCom.Договор.Region + "</td><td>";
                    s = s + OneCom.Договор.FIO + "</td><td>";
                    s = s + /*OneCom.TypePay*/ "КАСКО" + "</td><td align=left style='mso-number-format:\"\\@\";'>";
                    s = s + OneCom.NumKASKO + "</td><td>";
                    s = s + OneCom.DealerCity + "</td><td align=right style='mso-number-format:Standard;'>";

                    if (OneCom.install_num != null)
                    {
                        if (OneCom.install_num.ToString() != "")
                        {
                            s = s + Math.Round((OneCom.install_qty == null ? 0 : OneCom.install_qty.Value), 2);
                        }
                        else
                        {
                            s = s + Math.Round(OneCom.PremiumQty.Value, 2);
                        }
                    }
                    else
                    {
                        if (OneCom.PremiumQty != null)
                        {
                            s = s + Math.Round(OneCom.PremiumQty.Value, 2);
                        }
                    }
                    s = s + "</td><td>";

                    s = s + OneCom.YearInsurance + "</td><td>";
                    if (OneCom.DateKASKO != null)
                    {
                        s = s + OneCom.DateKASKO.ToString().Substring(0, 10);
                    }
                    s = s + "</td><td>";
                    s = s + OneCom.TypePay + "</td><td align=right style='mso-number-format:Percent;'>";
                    s = s + OneCom.Prc + "</td><td align=right style='mso-number-format:Standard;'>";
                    if (OneCom.ComQty != null)
                    {
                        s = s + Math.Round(OneCom.ComQty.Value, 2) + "</td><td>";
                        SumQty = SumQty + (double)OneCom.ComQty.Value;
                    }

                    s = s + OneCom.Comment + "</td></tr>"; //...
                    this.Application.fwrite(f, s);
                }
                this.Application.fwrite(f, "<tr><td colspan=7 align=right>ИТОГО:</td><td></td><td></td><td></td><td></td><td></td><td></td><td align=right style='mso-number-format:Standard;'>" + SumQty + "</td><td></td></tr></TABLE><br>");
                string[] months = { "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };
                this.Application.fwrite(f, "<TABLE border=0><tr><td colspan=4></td><td colspan=8>Услуга Исполнителем оказана надлежащим образом. К оплате за " + months[new DateTime(QueryAct.SelectedItem.Year, QueryAct.SelectedItem.Month, 1).Month - 1] + " месяц " + new DateTime(QueryAct.SelectedItem.Year, QueryAct.SelectedItem.Month, 1).Year + " года:</td></tr>");
                
                this.Application.fwrite(f, "<tr><td colspan=2></td><td colspan=2 align=center style='mso-number-format:Standard;'><b>" + SumQty + "</b></td><td colspan=8>" + this.Application.num2str(SumQty) + "</td></tr>");
                this.Application.fwrite(f, "<tr><td></td><td colspan=3 align=center>Итого к оплате:</td><td colspan=8>" + SumQty.ToString("0.00").Replace(".", ",") + " " + this.Application.num2str(SumQty) + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4></td><td align=center><b>в т.ч. НДС</b></td><td colspan=2 align=center style='mso-number-format:Standard;'><b>" + Math.Round(SumQty * this.Application.NDS / (this.Application.NDS + 100) /*18 / 118*/, 2).ToString("0.00") + "</b></td><td colspan=8>" + this.Application.num2str(Math.Round(SumQty * this.Application.NDS / (this.Application.NDS + 100) /*18 / 118*/, 2)) + "</td></tr></TABLE><br>");

                this.Application.fwrite(f, "<div align=center>ПОДПИСИ СТОРОН:</div><br><TABLE border=0><tr><td></td><td colspan=4 align=center>");
                this.Application.fwrite(f, "<TABLE border=0>");
                this.Application.fwrite(f, "<tr><td colspan=4><b>ООО \"Банк ПСА Финанс РУС\"</b></td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4></td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>_________________________________" + this.Application.ActSignatory + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>                              м.п.</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4></td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>Адрес: 105120, Россия, город Москва, 2-й Сыромятнический переулок, дом 1, 7 этаж</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>ИНН 7750004288, КПП 770901001</td></tr>");

                this.Application.fwrite(f, "<tr></tr>");

                this.Application.fwrite(f, "<tr><td colspan=4>БИК 044525339</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>Корр/cчет: № 30101810845250000339 в Отделении 1 Москва</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>Банк получателя: ООО «Банк ПСА Финанс РУС»</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>р/с		" + РеквизитыКомпаний.SelectedItem.Account + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>ОКПО 86555411, ОГРН 1087711000024</td></tr>");

                this.Application.fwrite(f, "<tr><td colspan=4></td></tr><tr><td colspan=4></td></tr>");
                this.Application.fwrite(f, "</TABLE></td><td></td><td></td><td></td><td colspan=7 align=center><TABLE border=0>");
                this.Application.fwrite(f, "<tr><td colspan=4><b>" + РеквизитыКомпаний.SelectedItem.Company + "</b></td></tr>");

                this.Application.fwrite(f, "<tr><td colspan=4></td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>_________________________________                        " + РеквизитыКомпаний.SelectedItem.Signer + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>                              м.п.</td></tr>");

                this.Application.fwrite(f, "<tr><td colspan=4></td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>Адрес: " + РеквизитыКомпаний.SelectedItem.Address + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>ИНН " + РеквизитыКомпаний.SelectedItem.INN + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>БИК " + РеквизитыКомпаний.SelectedItem.BIC + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>корр./счет " + РеквизитыКомпаний.SelectedItem.CorAcc + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>р/с " + РеквизитыКомпаний.SelectedItem.CurAcc + " в " + РеквизитыКомпаний.SelectedItem.Bank + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>ОГРН " + РеквизитыКомпаний.SelectedItem.OGRN + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>ОКПО " + РеквизитыКомпаний.SelectedItem.OKPO + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>ОКОНХ " + РеквизитыКомпаний.SelectedItem.OKONH + "</td></tr>");
                this.Application.fwrite(f, "<tr><td colspan=4>Тел. " + РеквизитыКомпаний.SelectedItem.Phone + "</td></tr>");
                this.Application.fwrite(f, "</TABLE></td></tr></TABLE></BODY>\r\n</HTML>\r\n\r\n");

                f.Close();
                shell.ShellExecute(ExcelFile);
            }
            else
            {
                this.ShowMessageBox("Automation not available");
            }
        }
    }
}
