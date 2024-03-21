using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Exception = System.Exception;

namespace OutlookCopyAttachmentsToClipboard
{
    public partial class RibbonCopyAllAttachments
    {
        private void RibbonCopyAllAttachments_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
                if (selection != null && selection.Count > 0)
                {
                    foreach (Microsoft.Office.Interop.Outlook.MailItem mailItem in selection)
                    {
                        if (mailItem.Attachments.Count > 0)
                        {
                            System.Collections.Specialized.StringCollection paths = new System.Collections.Specialized.StringCollection();

                            for (int i = 1; i <= mailItem.Attachments.Count; i++)
                            {
                                Microsoft.Office.Interop.Outlook.Attachment attachment = mailItem.Attachments[i];
                                // Temporary folder to save attachments
                                string tempPath = System.IO.Path.GetTempPath() + attachment.FileName;
                                attachment.SaveAsFile(tempPath);
                                paths.Add(tempPath);
                            }

                            // Copy the files' paths to the clipboard
                            Clipboard.SetFileDropList(paths);
                            MessageBox.Show("Selected attachments copied to clipboard.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //Microsoft.Office.Interop.Outlook.Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
            string uniqueFolderName = Path.GetRandomFileName();
            string saveFolderPath = Path.Combine(Path.GetTempPath(), uniqueFolderName);

            // Create the directory
            Directory.CreateDirectory(saveFolderPath);
            try
            {
                Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
                if (selection != null && selection.Count > 0)
                {
                    foreach (Microsoft.Office.Interop.Outlook.MailItem mailItem in selection)
                    {
                        if (mailItem.Attachments.Count > 0)
                        {
                            System.Collections.Specialized.StringCollection paths = new System.Collections.Specialized.StringCollection();

                            for (int i = 1; i <= mailItem.Attachments.Count; i++)
                            {
                                Microsoft.Office.Interop.Outlook.Attachment attachment = mailItem.Attachments[i];
                                // Temporary folder to save attachments
                                string filePath = Path.Combine(saveFolderPath, attachment.FileName);
                                attachment.SaveAsFile(filePath);
                            }
                        }
                    }
                    System.Diagnostics.Process.Start(saveFolderPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }
    }
}
