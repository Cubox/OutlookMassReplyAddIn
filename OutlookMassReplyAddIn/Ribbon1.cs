using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace OutlookMassReplyAddIn
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        public void MassReply(Office.IRibbonControl control)
        {
            Outlook.Selection selection = control.Context;
            OpenFileDialog templateSelect = new OpenFileDialog();
            templateSelect.InitialDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Templates");
            if (templateSelect.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            try
            {
                templateSelect.OpenFile().Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Invalid template");
                return;
            }

            Outlook.MailItem template = selection.Application.CreateItemFromTemplate(templateSelect.FileName) as Outlook.MailItem;
            template.Display(false);

            Word.Document templateEditor = template.GetInspector.WordEditor as Word.Document;
            if (templateEditor == null)
            {
                MessageBox.Show("Error getting Word Document");
                return;
            }
            string templateFile = Path.GetTempFileName();
            templateEditor.Content.ExportFragment(templateFile, Word.WdSaveFormat.wdFormatDocumentDefault); // I hate you outlook

            string[] tmpAttachment = new string[template.Attachments.Count]; // I really hate you.
            bool gc = template.Attachments.Count > 1;
            int i = 0;
            foreach (Outlook.Attachment attachement in template.Attachments)
            {
                tmpAttachment[i] = Path.Combine(Path.GetTempPath(), attachement.FileName);
                attachement.SaveAsFile(tmpAttachment[i]); // Holy fuck outlook
                i++;
            }

            template.Close(Outlook.OlInspectorClose.olDiscard);

            foreach (System.Object oMail in selection)
            {
                Outlook.MailItem mail = oMail as Outlook.MailItem;
                if (mail == null)
                {
                    MessageBox.Show("Not an email. Aborted");
                    return;
                }

                Outlook.MailItem reply = mail.Reply();
                reply.Display(false);
                Word.Document replyEditor = reply.GetInspector.WordEditor as Word.Document;
                if (replyEditor == null)
                {
                    MessageBox.Show("Error getting Word Document");
                    return;
                }

                replyEditor.Range(0, 0).ImportFragment(templateFile);

                i = 0;
                foreach (Outlook.Attachment attachment in template.Attachments)
                {
                    reply.Attachments.Add(tmpAttachment[i]);
                    i++;
                }
                reply.Send();
                if (gc) // Fuck you outlook
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            foreach (string file in tmpAttachment)
            {
                File.Delete(file);
            }
            File.Delete(templateFile);

            template.Delete();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookMassReplyAddIn.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

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
