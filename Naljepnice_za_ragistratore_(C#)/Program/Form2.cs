using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Naljepnice_za_ragistratore
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // varijablama dodjeljujemo vrijednosti iz polja '.Text'
            string ime_poduzeca = text1.Text; // Unaprijed definirano u Designu. . . upravlja samo administrator i tvorac aplikacije!
            string mjesto_poduzeca = text2.Text; // Unaprijed definirano u Designu. . . upravlja samo administrator i tvorac aplikacije!
            string godina = text3.Text;
            string naslov1 = text4.Text;
            string naslov2 = text5.Text;
            string naslov3 = text6.Text;
            string podnaslov1 = text7.Text;
            string podnaslov2 = text8.Text;
            string podnaslov3 = text9.Text;
            string broj1 = text10.Text;
            string broj2 = text11.Text;
            string datum1 = text12.Text;
            string datum2 = text13.Text;


            // ..........Putanje............
            // templejt
            string tmpPath = @"G:\PROJEKTI 2021\Naljepnice_za_ragistratore_(C#)\Templates\registrator_naljepnica.docx";

            // spremanje gotove naljepnice  // spremanje pod imenom 'text3.Text + text4.Text - naljepnica registratora.pdf'
            string outputName = @"G:\PROJEKTI 2021\Naljepnice_za_ragistratore_(C#)\Printane_naljepnice\" + text3.Text + " " + text4.Text + " - naljepnica_registratora.pdf";

            // nevidljiva datoteka 
            string shadowFile = @"G:\PROJEKTI 2021\Naljepnice_za_ragistratore_(C#)\temp.docx";

            // kreiraj nevidljivu datoteku 
            System.IO.File.Copy(tmpPath, shadowFile, true);


            // MS WORD
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(shadowFile);

            // ime poduzeca
            object oBookMark = "ime_poduzeca";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = ime_poduzeca;

            // mjesto poduzeca
            oBookMark = "mjesto_poduzeca";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = mjesto_poduzeca;

            // godina
            oBookMark = "godina";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = godina;

            // naslov1
            oBookMark = "naslov1";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = naslov1;

            // naslov2
            oBookMark = "naslov2";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = naslov2;

            // naslov3
            oBookMark = "naslov3";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = naslov3;

            // podnaslov1
            oBookMark = "podnaslov1";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = podnaslov1;

            // podnaslov2
            oBookMark = "podnaslov2";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = podnaslov2;

            // podnaslov3
            oBookMark = "podnaslov3";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = podnaslov3;

            // broj OD 
            oBookMark = "broj1";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = broj1;

            // broj DO
            oBookMark = "broj2";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = broj2;

            // datum OD
            oBookMark = "datum1";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = datum1;

            // datum DO
            oBookMark = "datum2";
            doc.Bookmarks.get_Item(ref oBookMark).Range.Text = datum2;

            // Pdf format
            doc.ExportAsFixedFormat(outputName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            // provjera printa
            doc.PrintPreview();
            // zatvori, ugasi, kraj aplikacije
            doc.Close();
            app.Quit();
            System.IO.File.Delete(shadowFile);

            // Poruka o printanju
            MessageBox.Show("Printanje je uspješno!");
        }
        
        // gumb 'Odustani'
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}

/*  
--------------------------------------------------------------------------------------------

                             Produced by:
                                      -- Severin Knežević --  
                                Email: knezevicseverin@gmail.com

--------------------------------------------------------------------------------------------
 */