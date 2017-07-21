using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Ecoview_V2._0
{
    public class PrintPreviewDialogSelectPrinter : PrintPreviewDialog
    {
        public PrintPreviewDialogSelectPrinter()
        {
            Shown += myPrintPreview_Shown;
        }


        public void myPrintPreview_Shown(object sender, EventArgs e)
        {
            //Get the toolstrip from the base control
            ToolStrip ts = (ToolStrip)this.Controls[1];
            //Get the print button from the toolstrip
            ToolStripItem printItem = ts.Items[0];//"printToolStripButton"

            ToolStripItem myPrintItem;
            myPrintItem = ts.Items.Add(printItem.Text, printItem.Image, new EventHandler(MybtnPrint_Click));
            myPrintItem.DisplayStyle = ToolStripItemDisplayStyle.Image;
            //Relocate the item to the beginning of the toolstrip
            ts.Items.Insert(0, myPrintItem);
            //Remove the orginal button
            ts.Items.Remove(printItem);
        }

        public void MybtnPrint_Click(object sender, EventArgs e)
        {
            PrintDialog dlgPrint = new PrintDialog();
            try
            {
                dlgPrint.AllowSelection = true;
                dlgPrint.ShowNetwork = true;
                dlgPrint.Document = this.Document;
                if (dlgPrint.ShowDialog() == DialogResult.OK)
                {
                    this.Document.Print();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Print Error: " + ex.Message);
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // PrintPreviewDialogSelectPrinter
            // 
            this.ClientSize = new System.Drawing.Size(400, 300);
            this.Name = "PrintPreviewDialogSelectPrinter";
            this.Load += new System.EventHandler(this.PrintPreviewDialogSelectPrinter_Load);
            this.ResumeLayout(false);

        }

        private void PrintPreviewDialogSelectPrinter_Load(object sender, EventArgs e)
        {

        }
    }
}
