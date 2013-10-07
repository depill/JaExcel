using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using JaAPI;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace JaExcel
{
    [ComVisible(true)]
    public class JaRibbon : Office.IRibbonExtensibility
    {
        private string APIKey = "";
        private Office.IRibbonUI ribbon;

        public void GetAPIKey(Office.IRibbonControl control, String Text)
        {
            APIKey = Text;
        }

        public void LeitJa(Office.IRibbonControl RibbonControl)
        {
            Excel.Worksheet ws = ((Excel.Worksheet)Globals.JaExcel.Application.ActiveSheet);
            Excel.Range rng = ws.UsedRange;
            int numberOfRows = rng.Rows.Count;
            int c = 1;
            int addpi = 12;
            int dalkuri = 1;

            string[] dalkar = new string[11] { "Nafn", "Starfsheiti", "Heimilisfang", "Póstnúmer", "Bær", "Bannmerktur", "Veffang", "Netfang", "Fax", "Aðalsímanúmer", "Aukasímanúmer hér eftir" };


            for (int r = 1; r <= numberOfRows; r++)
            {

                if (r == 1)
                {
                    if (ws.Cells[r, c].Value.ToLower() == "kennitala")
                    {
                        dalkuri = 1;
                        foreach (string dalkur in dalkar)
                        {
                            ws.Cells[r, dalkuri + c] = dalkur;
                            dalkuri++;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Fyrsti dálkur og fyrsta röð þarf að vera heitir Kennitala");
                        break;
                    }
                }
                else if (ws.Cells[r, c].Value != null)
                {
                    string kennitala = ws.Cells[r, c].Value.ToString();

                    // Athugum eftir hvernig kennitala gæti orðið
                    // Ef það er - í kennitölunni þurfum við að henda því úr
                    if (kennitala.Contains("-"))
                    {
                        kennitala = kennitala.Replace("-", "");
                    }

                    if (kennitala.Length == 9)
                    {
                        kennitala = "0" + kennitala;
                    }

                    // Köllum í Já API
                    JaStruct response = Ja_API.FetchPhone(kennitala, APIKey);
                    if (response.error == "Unauthorized")
                    {
                        MessageBox.Show("API lykill ekki gildur");
                        break;
                    }
                    else
                    {
                        if (response.items.Count > 0)
                        {
                            // Köllum bara í fyrsta itemið vegna þess að við höfum bara eitt row til að vinna með
                            ItemsStruct adili = response.items[0];
                            ws.Cells[r, 2] = adili.name;
                            ws.Cells[r, 3] = adili.occupation;
                            ws.Cells[r, 4] = adili.address;
                            ws.Cells[r, 5] = adili.postalCode;
                            ws.Cells[r, 6] = adili.postalStation;
                            ws.Cells[r, 7] = adili.solicitationProhibited.ToString();
                            ws.Cells[r, 8] = adili.url;
                            ws.Cells[r, 9] = adili.email;
                            ws.Cells[r, 10] = adili.faxNumber;
                            ws.Cells[r, 11] = adili.phoneNumber.phoneNumber;

                            addpi = 12;
                            if (adili.additionalPhonenumbers != default(List<PhoneNumberStruct>))
                            {
                                foreach (PhoneNumberStruct addp in adili.additionalPhonenumbers)
                                {
                                    ws.Cells[r, addpi] = addp.phoneNumber;
                                    addpi++;
                                }
                            }

                        }
                        else
                        {
                            for (int coli = 2; coli <= 20; coli++)
                            {
                                ws.Cells[r, coli] = null;
                            }
                        }

                    }

                }
                else
                {
                    MessageBox.Show("Las allar kennitölur og kom að auðum reit, leit lokið");
                    break;
                }
            }
            MessageBox.Show("Leit lokið");
            
        }

        public JaRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("JaExcel.JaRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
