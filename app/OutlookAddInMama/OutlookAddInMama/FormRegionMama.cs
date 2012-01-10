using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace OutlookAddInMama
{
    partial class FormRegionMama
    {
        /**
         * This small project doesn't make use of MVVM as a normal one would do.
         * And of course UI and program logic code could be better seperated.
         */

        // the template mail. everything is copied from here and attached with the respective file, references are replaced
        private MailItem template;

        // the running outlook instance
        private Outlook.Application application;

        // threshold when to warn user about sending a huge amount of mails
        private const int threshold = 10;

        #region Formularbereichsfactory

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("OutlookAddInMama.FormRegionMama")]
        public partial class FormRegionMamaFactory
        {
           private void FormRegionMamaFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
            }
        }

        #endregion

        #region event functions

        private void FormRegionMama_FormRegionShowing(object sender, System.EventArgs e)
        {
            this.reloadSettings();
        }

        private void FormRegionMama_FormRegionClosed(object sender, System.EventArgs e)
        {
            //release COM-Objects
            Marshal.ReleaseComObject(template);
            Marshal.ReleaseComObject(application);
        }

        private void FormRegionMama_Load(object sender, EventArgs e)
        {
            this.template = (Outlook.MailItem)Marshal.GetUniqueObjectForIUnknown(Marshal.GetIUnknownForObject(this.OutlookItem));
            this.application = (Outlook.Application)Marshal.GetUniqueObjectForIUnknown(Marshal.GetIUnknownForObject(this.OutlookFormRegion.Application));
        }

        private void buttonSend_Click(object sender, EventArgs e)
        {
            this.storeSettings();

            List<MailItem> mailList = createMails(this.textBoxPattern.Text, this.textBoxLocation.Text);
            this.sendMails(mailList);
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            this.chooseFolder();
        }

        #endregion

        #region functions

        private void reloadSettings()
        {
            // reload last settings
            this.textBoxPattern.Text = Properties.Settings.Default.Pattern;
            this.textBoxLocation.Text = Properties.Settings.Default.Directory;
            this.checkBoxPreview.Checked = Properties.Settings.Default.Preview;
        }

        private void storeSettings()
        {
            // store current settings
            Properties.Settings.Default.Pattern = this.textBoxPattern.Text;
            Properties.Settings.Default.Directory = this.textBoxLocation.Text;
            Properties.Settings.Default.Preview = this.checkBoxPreview.Checked;
            Properties.Settings.Default.Save();
        }

        private void sendMails(List<MailItem> mailList)
        {
            bool proceed = true;

            if (mailList.Count > threshold) proceed = (MessageBox.Show("This would create " + mailList.Count + " messages! Proceed?", "Mama says: A helluva lot of messages", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes);

            int successCount = 0;
            foreach (MailItem singleMail in mailList)
            {
                //cleanup
                if (!proceed)
                {
                    singleMail.Move(this.application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems));
                    singleMail.Delete();
                }
                else if (this.checkBoxPreview.Checked)
                {
                    try
                    {
                        singleMail.Move(this.application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts));
                        singleMail.Display();
                        successCount++;
                    }
                    catch (System.Exception err)
                    {
                        MessageBox.Show(err.ToString(), "Mama says: Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    try
                    {
                        singleMail.Send();
                        successCount++;
                    }
                    catch (System.Exception err)
                    {
                        MessageBox.Show(err.ToString(), "Mama says: Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            if (proceed)
            {
                DialogResult result = MessageBox.Show(successCount + " of " + mailList.Count + " Messages have been successfully created for send. Close template mail?", "Mama says: Successful", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes) this.OutlookFormRegion.Inspector.Close(OlInspectorClose.olPromptForSave);
            }
        }

        /// <summary>
        /// Returns a collection of new MailItems. One for every single file found in directory path
        /// and matching pattern pattern.
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        private List<MailItem> createMails(string pattern, string path)
        {
            //what will be returned
            List<MailItem> ret = new List<MailItem>();

            //the regular expression to which against is matched, ignore case (so does Windows with filenames...)
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

            //first get all files
            String[] allFiles = Directory.GetFiles(path);

            //and then look which matches
            foreach (String fileFullPath in allFiles)
            {
                //the fileName is used for match, the fileFullPath later on to attach the file
                String fileName = fileFullPath;
                if (fileFullPath.StartsWith(path)) fileName = fileFullPath.Substring(path.Length+1);

                Match match = regex.Match(fileName);

                if (match.Success)
                {
                    //Beware: This places a new item in the outbox...
                    MailItem mail = template.Copy();
                    //...but later on it's moved ;-)

                    /* This would create a brand new MailItem in memory without any object in
                     * the outbox, but you would have to copy all properties by yourself:
                     * 
                     * MailItem mail = this.OutlookFormRegion.Application.CreateItem(Outlook.OlItemType.olMailItem);
                    */

                    Dictionary<string, string> replacementDictionary = new Dictionary<string, string>();
                    //if the user used groups in his regex, now references ($1...n) are searched in To, CC, BCC and Body and replaced 
                    GroupCollection coll = match.Groups;
                    for (int i = 0; i < coll.Count; i++)
                    {
                        replacementDictionary.Add("" + i, coll[i].Value);
                    }

                    //since neither simple replace nor regex replace seemed to be appropriate, i quickly wrote a replacement parser
                    ReplacementParser rp = new ReplacementParser(replacementDictionary);

                    mail.To = rp.replaceAll(mail.To);
                    mail.CC = rp.replaceAll(mail.CC);
                    mail.BCC = rp.replaceAll(mail.BCC);
                    mail.Subject = rp.replaceAll(mail.Subject);
                    mail.Body = rp.replaceAll(mail.Body);


                    //finally attach the file
                    mail.Attachments.Add(fileFullPath);

                    //and add the new item to the collection
                    ret.Add(mail);
                }
            }

            return ret;
        }

        private void chooseFolder()
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxLocation.Text = folderBrowserDialog.SelectedPath;
            }
        }

#endregion

        #region helping functions

        #endregion
    }
}
