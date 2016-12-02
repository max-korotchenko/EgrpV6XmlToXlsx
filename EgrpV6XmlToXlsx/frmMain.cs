using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;

namespace EgrpV6XmlToXlsx
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if (ofdInput.ShowDialog() == DialogResult.OK)
                tbFileName.Text = ofdInput.FileName;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(tbFileName.Text) && File.Exists(tbFileName.Text))
                ThreadPool.QueueUserWorkItem(ProcessFile, tbFileName.Text);
            else
                MessageBox.Show(@"Некорректный путь к XML-файлу!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void tbFileName_DragDrop(object sender, DragEventArgs e)
        {
            string[] strings = (string[])e.Data.GetData(DataFormats.FileDrop, true);
            tbFileName.Text = strings[0];
        }

        private void tbFileName_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.All : DragDropEffects.None;
        }

        private void UISetBusyMode()
        {
            if (btnBrowse.InvokeRequired)
                btnBrowse.BeginInvoke(new Action(() => btnBrowse.Enabled = false));
            if (statusStrip.InvokeRequired)
                statusStrip.BeginInvoke(new Action(() => tsProgressBar.Visible = true));
        }

        private void UISetFreeMode()
        {
            if (btnBrowse.InvokeRequired)
                btnBrowse.BeginInvoke(new Action(() => btnBrowse.Enabled = true));
            if (statusStrip.InvokeRequired)
                statusStrip.BeginInvoke(new Action(() => tsProgressBar.Visible = false));
        }

        private void ProcessFile(object fullFileName)
        {
            UISetBusyMode();
            string file = (string)fullFileName;

            try
            {
                XDocument doc = XDocument.Load(file);
                string rightOwner = "Лист1";
                XElement subjectEl = doc.Descendants("Subject").SingleOrDefault();
                XElement g = null;
                if (subjectEl != null && subjectEl.Element("Governance") != null && subjectEl.Element("Governance").Element("Content") != null)
                {
                    rightOwner = subjectEl.Element("Governance").Element("Content").Value;
                }
                List<XElement> objRights = doc.Descendants("ObjectRight").ToList();

                using (ExcelPackage pck = new ExcelPackage())
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add(rightOwner);
                    ws.Cells[1, 1].Value = "№"; ws.Column(35).Style.Numberformat.Format = "0";
                    #region Поля Object
                    ws.Cells[1, 2].Value = "Уникальный ID объекта";
                    ws.Cells[1, 3].Value = "Кадастровый или условный номер";
                    ws.Cells[1, 4].Value = "Код типа объекта недвижимости";
                    ws.Cells[1, 5].Value = "Текстовое описание типа объекта недвижимости";
                    ws.Cells[1, 6].Value = "Наименование объекта недвижимости";
                    ws.Cells[1, 7].Value = "Код назначения";
                    ws.Cells[1, 8].Value = "Текстовое описание назначения";
                    ws.Cells[1, 9].Value = "Целевое назначение (категория) земель";
                    ws.Cells[1, 10].Value = "Текстовое описание целевоего назначения(категории) земель";
                    ws.Cells[1, 11].Value = "Значение площади";
                    ws.Cells[1, 12].Value = "Значение площади текстом";
                    ws.Cells[1, 13].Value = "Единица измерений";
                    ws.Cells[1, 14].Value = "Инвентарный номер, литер";
                    ws.Cells[1, 15].Value = "Этажность (этаж)";
                    ws.Cells[1, 16].Value = "Номера на поэтажном плане";
                    ws.Cells[1, 17].Value = "Уникальный ID адреса";
                    ws.Cells[1, 18].Value = "Суммарное неформализованное описание";
                    ws.Cells[1, 19].Value = "Регион РФ или страна регистрации";
                    ws.Cells[1, 20].Value = "ОКАТО";
                    ws.Cells[1, 21].Value = "КЛАДР";
                    ws.Cells[1, 22].Value = "Почтовый индекс";
                    ws.Cells[1, 23].Value = "Район";
                    ws.Cells[1, 24].Value = "Муниципальное образование";
                    ws.Cells[1, 25].Value = "Городской район";
                    ws.Cells[1, 26].Value = "Сельсовет";
                    ws.Cells[1, 27].Value = "Населенный пункт";
                    ws.Cells[1, 28].Value = "Улица";
                    ws.Cells[1, 29].Value = "Дом";
                    ws.Cells[1, 30].Value = "Корпус";
                    ws.Cells[1, 31].Value = "Строение";
                    ws.Cells[1, 32].Value = "Квартира";
                    ws.Cells[1, 33].Value = "Иное";
                    ws.Cells[1, 34].Value = "Состав сложной вещи";
                    #endregion
                    #region Поля Registration
                    ws.Cells[1, 35].Value = "Уникальный ID записи о праве (ограничении)";
                    ws.Cells[1, 36].Value = "Номер государственной регистрации";
                    ws.Cells[1, 37].Value = "Код права";
                    ws.Cells[1, 38].Value = "Вид права";
                    ws.Cells[1, 39].Value = "Дата государственной регистрации";
                    ws.Cells[1, 40].Value = "Дата прекращения права";
                    ws.Cells[1, 41].Value = "Доля";
                    ws.Cells[1, 42].Value = "Значение доли текстом";
                    ws.Cells[1, 43].Value = "Документы - основания регистрации права";
                    #endregion
                    #region Поля Encumbrance
                    ws.Cells[1, 44].Value = "Ограничения права";
                    #endregion
                    var titleRng = ws.Cells[1, 1, 1, 47];
                    titleRng.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    titleRng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    titleRng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    titleRng.AutoFilter = true;
                    titleRng.Style.Font.Bold = true;

                    for (int i = 0; i < objRights.Count; i++)
                    {
                        ws.Cells[i + 2, 1].Value = Convert.ToInt32(objRights[i].Attribute("ObjectNumber").Value);
                        #region Object
                        ws.Cells[i + 2, 2].Value = objRights[i].Element("Object").Element("ID_Object").Value;
                        ws.Cells[i + 2, 3].Value = objRights[i].Element("Object").Element("CadastralNumber") != null
                            ? objRights[i].Element("Object").Element("CadastralNumber").Value 
                            : (objRights[i].Element("Object").Element("ConditionalNumber") != null 
                                ? objRights[i].Element("Object").Element("ConditionalNumber").Value
                                : "");
                        ws.Cells[i + 2, 4].Value = objRights[i].Element("Object").Element("ObjectType").Value;
                        ws.Cells[i + 2, 5].Value = objRights[i].Element("Object").Element("ObjectTypeText") != null ? objRights[i].Element("Object").Element("ObjectTypeText").Value : "";
                        ws.Cells[i + 2, 6].Value = objRights[i].Element("Object").Element("Name").Value;
                        ws.Cells[i + 2, 7].Value = objRights[i].Element("Object").Element("Assignation_Code") != null ? objRights[i].Element("Object").Element("Assignation_Code").Value : "";
                        ws.Cells[i + 2, 8].Value = objRights[i].Element("Object").Element("Assignation_Code_Text") != null ? objRights[i].Element("Object").Element("Assignation_Code_Text").Value : "";
                        ws.Cells[i + 2, 9].Value = objRights[i].Element("Object").Element("GroundCategory") != null ? objRights[i].Element("Object").Element("GroundCategory").Value : "";
                        ws.Cells[i + 2, 10].Value = objRights[i].Element("Object").Element("GroundCategoryText") != null ? objRights[i].Element("Object").Element("GroundCategoryText").Value : "";
                        ws.Cells[i + 2, 11].Value = objRights[i].Element("Object").Element("Area") != null && objRights[i].Element("Object").Element("Area").Element("Area") != null ? objRights[i].Element("Object").Element("Area").Element("Area").Value : "";
                        ws.Cells[i + 2, 12].Value = objRights[i].Element("Object").Element("Area") != null ? objRights[i].Element("Object").Element("Area").Element("AreaText").Value : "";
                        ws.Cells[i + 2, 13].Value = objRights[i].Element("Object").Element("Area") != null && objRights[i].Element("Object").Element("Area").Element("Unit") != null ? objRights[i].Element("Object").Element("Area").Element("Unit").Value : "";
                        ws.Cells[i + 2, 14].Value = objRights[i].Element("Object").Element("Inv_No") != null ? objRights[i].Element("Object").Element("Inv_No").Value : "";
                        ws.Cells[i + 2, 15].Value = objRights[i].Element("Object").Element("Floor") != null ? objRights[i].Element("Object").Element("Floor").Value : "";
                        ws.Cells[i + 2, 16].Value = objRights[i].Element("Object").Element("FloorPlan_No") != null ? string.Join("; ", objRights[i].Element("Object").Element("FloorPlan_No").Elements("Explication").Select(el => el.Value)) : "";
                        ws.Cells[i + 2, 17].Value = objRights[i].Element("Object").Element("Address").Element("ID_Address") != null ? objRights[i].Element("Object").Element("Address").Element("ID_Address").Value : "";
                        ws.Cells[i + 2, 18].Value = objRights[i].Element("Object").Element("Address").Element("Content") != null ? objRights[i].Element("Object").Element("Address").Element("Content").Value : "";
                        ws.Cells[i + 2, 19].Value = objRights[i].Element("Object").Element("Address").Element("Region") != null 
                            ? objRights[i].Element("Object").Element("Address").Element("Region").Attribute("Code").Value + " - " + objRights[i].Element("Object").Element("Address").Element("Region").Attribute("Name").Value 
                            : (objRights[i].Element("Object").Element("Address").Element("Region") != null 
                                ? objRights[i].Element("Object").Element("Address").Element("Region").Attribute("Code").Value + " - " + objRights[i].Element("Object").Element("Address").Element("Region").Attribute("Name").Value 
                                : "");
                        ws.Cells[i + 2, 20].Value = objRights[i].Element("Object").Element("Address").Element("Code_OKATO") != null ? objRights[i].Element("Object").Element("Address").Element("Code_OKATO").Value : "";
                        ws.Cells[i + 2, 21].Value = objRights[i].Element("Object").Element("Address").Element("Code_KLADR") != null ? objRights[i].Element("Object").Element("Address").Element("Code_KLADR").Value : "";
                        ws.Cells[i + 2, 22].Value = objRights[i].Element("Object").Element("Address").Element("Postal_Code") != null ? objRights[i].Element("Object").Element("Address").Element("Postal_Code").Value : "";
                        ws.Cells[i + 2, 23].Value = objRights[i].Element("Object").Element("Address").Element("District") != null 
                            ? objRights[i].Element("Object").Element("Address").Element("District").Attribute("Name").Value + 
                                (objRights[i].Element("Object").Element("Address").Element("District").Attribute("Type") != null ? " " + objRights[i].Element("Object").Element("Address").Element("District").Attribute("Type").Value : "")
                            : "";
                        ws.Cells[i + 2, 24].Value = objRights[i].Element("Object").Element("Address").Element("City") != null
                            ? objRights[i].Element("Object").Element("Address").Element("City").Attribute("Name").Value +
                                (objRights[i].Element("Object").Element("Address").Element("City").Attribute("Type") != null ? " " + objRights[i].Element("Object").Element("Address").Element("City").Attribute("Type").Value : "")
                            : "";
                        ws.Cells[i + 2, 25].Value = objRights[i].Element("Object").Element("Address").Element("Urban_District") != null
                            ? objRights[i].Element("Object").Element("Address").Element("Urban_District").Attribute("Name").Value +
                                (objRights[i].Element("Object").Element("Address").Element("Urban_District").Attribute("Type") != null ? " " + objRights[i].Element("Object").Element("Address").Element("Urban_District").Attribute("Type").Value : "")
                            : "";
                        ws.Cells[i + 2, 26].Value = objRights[i].Element("Object").Element("Address").Element("Soviet_Village") != null
                            ? objRights[i].Element("Object").Element("Address").Element("Soviet_Village").Attribute("Name").Value +
                                (objRights[i].Element("Object").Element("Address").Element("Soviet_Village").Attribute("Type") != null ? " " + objRights[i].Element("Object").Element("Address").Element("Soviet_Village").Attribute("Type").Value : "")
                            : "";
                        ws.Cells[i + 2, 27].Value = objRights[i].Element("Object").Element("Address").Element("Locality") != null
                            ? objRights[i].Element("Object").Element("Address").Element("Locality").Attribute("Name").Value +
                                (objRights[i].Element("Object").Element("Address").Element("Locality").Attribute("Type") != null ? " " + objRights[i].Element("Object").Element("Address").Element("Locality").Attribute("Type").Value : "")
                            : "";
                        ws.Cells[i + 2, 28].Value = objRights[i].Element("Object").Element("Address").Element("Street") != null
                            ? objRights[i].Element("Object").Element("Address").Element("Street").Attribute("Name").Value +
                                (objRights[i].Element("Object").Element("Address").Element("Street").Attribute("Type") != null ? " " + objRights[i].Element("Object").Element("Address").Element("Street").Attribute("Type").Value : "")
                            : "";
                        ws.Cells[i + 2, 29].Value = objRights[i].Element("Object").Element("Address").Element("Level1") != null
                            ? objRights[i].Element("Object").Element("Address").Element("Level1").Attribute("Name").Value +
                                (objRights[i].Element("Object").Element("Address").Element("Level1").Attribute("Type") != null ? " " + objRights[i].Element("Object").Element("Address").Element("Level1").Attribute("Type").Value : "")
                            : "";
                        ws.Cells[i + 2, 30].Value = objRights[i].Element("Object").Element("Address").Element("Level2") != null
                            ? objRights[i].Element("Object").Element("Address").Element("Level2").Attribute("Name").Value +
                                (objRights[i].Element("Object").Element("Address").Element("Level2").Attribute("Type") != null ? " " + objRights[i].Element("Object").Element("Address").Element("Level2").Attribute("Type").Value : "")
                            : "";
                        ws.Cells[i + 2, 31].Value = objRights[i].Element("Object").Element("Address").Element("Level3") != null
                            ? objRights[i].Element("Object").Element("Address").Element("Level3").Attribute("Name").Value +
                                (objRights[i].Element("Object").Element("Address").Element("Level3").Attribute("Type") != null ? " " + objRights[i].Element("Object").Element("Address").Element("Level3").Attribute("Type").Value : "")
                            : "";
                        ws.Cells[i + 2, 32].Value = objRights[i].Element("Object").Element("Address").Element("Apartment") != null
                            ? objRights[i].Element("Object").Element("Address").Element("Apartment").Attribute("Name").Value +
                                (objRights[i].Element("Object").Element("Address").Element("Apartment").Attribute("Type") != null ? " " + objRights[i].Element("Object").Element("Address").Element("Apartment").Attribute("Type").Value : "")
                            : "";
                        ws.Cells[i + 2, 33].Value = objRights[i].Element("Object").Element("Address").Element("Other") != null ? objRights[i].Element("Object").Element("Address").Element("Other").Value : "";
                        ws.Cells[i + 2, 34].Value = objRights[i].Element("Object").Element("Complex") != null ? string.Join("; ", objRights[i].Element("Object").Element("Complex").Elements("Explication").Select(el => el.Value)) : "";
                        #endregion
                        #region Registration
                        ws.Cells[i + 2, 35].Value = objRights[i].Element("Registration").Element("ID_Record").Value;
                        ws.Cells[i + 2, 36].Value = objRights[i].Element("Registration").Element("RegNumber").Value;
                        ws.Cells[i + 2, 37].Value = objRights[i].Element("Registration").Element("Type").Value;
                        ws.Cells[i + 2, 38].Value = objRights[i].Element("Registration").Element("Name").Value;
                        ws.Cells[i + 2, 39].Value = objRights[i].Element("Registration").Element("RegDate").Value;
                        ws.Cells[i + 2, 41].Value = objRights[i].Element("Registration").Element("Share") != null ? objRights[i].Element("Registration").Element("Share").Attribute("Numerator").Value + "/" + objRights[i].Element("Registration").Element("Share").Attribute("Denominator").Value : "";
                        ws.Cells[i + 2, 42].Value = objRights[i].Element("Registration").Element("ShareText") != null ? objRights[i].Element("Registration").Element("ShareText").Value : "";
                        if (objRights[i].Element("Registration").Elements("DocFound") != null && objRights[i].Element("Registration").Elements("DocFound").Count() > 0)
                        {
                            string documents = "";
                            List<XElement> docsFoundList = objRights[i].Element("Registration").Elements("DocFound").ToList();
                            foreach (XElement docF in docsFoundList)
                            {
                                documents += "Уникальный ID документа: " + docF.Element("ID_Document").Value;
                                documents += ", суммарное описание: " + docF.Element("Content").Value;
                                documents += ", тип документа: " + docF.Element("Type_Document").Value;
                                documents += ", наименование документа: " + docF.Element("Name").Value;
                                documents += docF.Element("Series") != null ? ", серия документа: " + docF.Element("Series").Value : "";
                                documents += docF.Element("Number") != null ? ", номер документа: " + docF.Element("Number").Value : "";
                                documents += docF.Element("Date") != null ? ", дата выдачи документа: " + docF.Element("Date").Value : "";
                                documents += docF.Element("IssueOrgan") != null ? ", организация, выдавшая документ: " + docF.Element("IssueOrgan").Value : "";
                                documents += ".\r\n";
                            }
                            ws.Cells[i + 2, 43].Value = documents;
                        }
                        #endregion
                        #region Encumbrances
                        ws.Cells[i + 2, 44].Value = objRights[i].Elements("Encumbrance") != null ? string.Join(".\r\n\r\n", objRights[i].Elements("Encumbrance").Select(el => GetEncumbranceString(el))) : "";
                        #endregion
                    }
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    pck.SaveAs(new FileInfo(file.Replace(".xml", ".xlsx")));
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            UISetFreeMode();
            MessageBox.Show(@"Преобразование завершено!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string GetEncumbranceString(XElement el)
        {
            string enc = "";
            enc = "Уникальный ID записи об ограничении: " + el.Element("ID_Record").Value;
            enc += ", номер государственной регистрации: " + el.Element("RegNumber").Value;
            enc += ", код ограничения: " + el.Element("Type").Value;
            enc += ", вид ограничения: " + el.Element("Name").Value;
            enc += el.Element("ShareText") != null ? ", предмет ограничения: " + el.Element("ShareText").Value : "";
            enc += ", дата гос. регистрации: " + el.Element("RegDate").Value;
            enc += el.Element("Duration") != null && el.Element("Duration").Element("Started") != null ? ", дата начала действия: " + el.Element("Duration").Element("Started").Value : "";
            enc += el.Element("Duration") != null && el.Element("Duration").Element("Stopped") != null ? ", дата прекращения действия: " + el.Element("Duration").Element("Stopped").Value : "";
            enc += el.Element("Duration") != null && el.Element("Duration").Element("Term") != null ? ", продолжительность: " + el.Element("Duration").Element("Term").Value : "";
            if (el.Elements("Owner") != null && el.Elements("Owner").Count() > 0)
            {
                string owners = "";
                List<XElement> ownersList = el.Elements("Owner").ToList();
                foreach (XElement own in ownersList)
                {
                    owners += "уникальный ID субъекта: " + own.Element("ID_Subject").Value;
                    if (own.Element("Person") != null)
                    {
                        string person = " (ФЛ), ";

                        owners += person;
                    }
                    else if (own.Element("Organization") != null)
                    {
                        string organization = " (ЮЛ), ";

                        owners += organization;
                    }
                    else if (own.Element("Governance") != null)
                    {
                        string governance = " (субъект публичного права), ";

                        owners += governance;
                    }
                    owners += ";\r\n";
                }
                enc += ";\r\nлица, в пользу которых ограничиваются права: " + owners;
            }
            enc += (el.Element("AllShareOwner") != null ? ";\r\nучастники долевого строительства по договорам участия в долевом строительстве: " + el.Element("AllShareOwner").Value : "");
            if (el.Elements("DocFound") != null && el.Elements("DocFound").Count() > 0)
            {
                string documents = "";
                List<XElement> docsFoundList = el.Elements("DocFound").ToList();
                foreach (XElement docF in docsFoundList)
                {
                    documents += "уникальный ID документа: " + docF.Element("ID_Document").Value;
                    documents += ", суммарное описание: " + docF.Element("Content").Value;
                    documents += ", тип документа: " + docF.Element("Type_Document").Value;
                    documents += ", наименование документа: " + docF.Element("Name").Value;
                    documents += docF.Element("Series") != null ? ", серия документа: " + docF.Element("Series").Value : "";
                    documents += docF.Element("Number") != null ? ", номер документа: " + docF.Element("Number").Value : "";
                    documents += docF.Element("Date") != null ? ", дата выдачи документа: " + docF.Element("Date").Value : "";
                    documents += docF.Element("IssueOrgan") != null ? ", организация, выдавшая документ: " + docF.Element("IssueOrgan").Value : "";
                    documents += ";\r\n";
                }
                if (documents.Length > 0)
                    documents = documents.Substring(0, documents.Length - 3);
                enc += ";\r\nдокументы-основания для регистрации ограничения: " + documents;
            }
            return enc;
        }
    }
}