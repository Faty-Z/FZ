using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;


namespace mytest
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(WorkWithDocument);
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument += new Word.ApplicationEvents4_NewDocumentEventHandler(WorkWithDocument);

            // Using InsertBefore method inserts text   
            this.Application.ActiveDocument.Content.InsertBefore("Text @ the          Start - "); 
            // Using InsertAfter method inserts text  
            this.Application.ActiveDocument.Content.InsertAfter(" - Text @          the End");

            // There is a second method called "selection method for inserting text after and before
            // Using Selection Object inserting text after the text   
            //this.Application.Selection.InsertAfter(" - Text @ the End"); 
            // Using Selection Object inserting text before the text   
            //this.Application.Selection.InsertBefore("Text @ the Start - ");

            
        }

        private void WorkWithDocument(Microsoft.Office.Interop.Word.Document Doc)
        {
            // A method to add text in range
            //Specify a range at the beginning of a document and insert the text New Text
            // Word.Range rng = this.Application.ActiveDocument.Range(0, 0);
            // rng.Text = "New Text";

            //Select the Range object, which has expanded from one character to the length of the inserted text.
            //rng.Select();

            // To replace text in a range
            //Word.Range rng = this.Application.ActiveDocument.Range(0, 12);

            //Replace those characters with the string New Text.
            //rng.Text = "New Text";

            // Select the range
            // rng.Select();



            // Initializing the Range object  
                Word.Range PacktRangeSelect;
            // Check the sentence count  
            if (this.Sentences.Count >= 1)
            {
             // Set the start and ent point has object       
                object pktStartFrom = this.Sentences[2].Start;
                object pktStopHere = this.Sentences[5].End;
             // Assign the selection range       
                PacktRangeSelect = this.Range(ref pktStartFrom, ref pktStopHere);
            // Select the sentence using Select() method      
                PacktRangeSelect.Select();
            }
            else
            {
                return;
            }

        }
        


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
           

            
        }

    }
}
