using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace unire_fisiere_excel
{
    public partial class Form1 : Form
    {
        int ok;
        Excel.Application ex;
        Excel.Workbook doc, docUnite;
        Excel.Worksheet pag, pagUniteP1, pagUniteP2, pagUniteP3, pagUniteM;
        int pu1, pu2, pu3, pum;
        public Form1()
        {
            InitializeComponent();
            ok = 0;
            ex = new Excel.Application();
            docUnite = ex.Workbooks.Add();
            docUnite.Worksheets.Add();
            docUnite.Worksheets.Add();
            docUnite.Worksheets.Add();
            pagUniteP1 = (Excel.Worksheet)docUnite.Worksheets.get_Item(1);
            pagUniteP2 = (Excel.Worksheet)docUnite.Worksheets.get_Item(2);
            pagUniteP3 = (Excel.Worksheet)docUnite.Worksheets.get_Item(3);
            pagUniteM = (Excel.Worksheet)docUnite.Worksheets.get_Item(4);
            pagUniteP1.Name = "Premiul 1";
            pagUniteP2.Name = "Premiul 2";
            pagUniteP3.Name = "Premiul 3";
            pagUniteM.Name = "Mentiune";
            pagUniteP1.Cells[1, 1] = "Nume";
            pagUniteP1.Cells[1, 2] = "Clasa";
            pagUniteP1.Cells[1, 3] = "Media";
            pu1 = pu2 = pu3 = pum = 1;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            od1.InitialDirectory = Directory.GetCurrentDirectory();
            if(od1.ShowDialog() == DialogResult.OK)
            {
                numeDirector.Text = Path.GetDirectoryName( od1.FileName);
                lista.Items.Clear();
                string[] l = Directory.GetFiles(numeDirector.Text);
                for(int i = 0; i< l.Length; i++)
                {
                    if(Path.GetExtension( l[i])==".xlsx")
                    {
                        lista.Items.Add(Path.GetFileNameWithoutExtension(l[i]));
                        doc = ex.Workbooks.Open(l[i]);
                        pag = doc.Sheets[1];
                        for(int j = 2;pag.Cells[j,3].Value2!=null; j++)
                        {
                            if(pag.Cells[j, 3].Value2=="I")
                            {
                                pu1++;
                                pagUniteP1.Cells[pu1, 1] = pag.Cells[j, 1];
                                pagUniteP1.Cells[pu1, 2] = Path.GetFileNameWithoutExtension(l[i]);
                                pagUniteP1.Cells[pu1, 3] = pag.Cells[j, 2];
                            }
                        }
                    }
                }

                /// salvare finala
                if (File.Exists(numeDirector.Text + @"\Unite.xlsx"))
                    File.Delete(numeDirector.Text + @"\Unite.xlsx");
                
                docUnite.SaveAs(numeDirector.Text + @"\Unite.xlsx");
                docUnite.Close();
                ex.Quit();
            }
        }
    }
}
