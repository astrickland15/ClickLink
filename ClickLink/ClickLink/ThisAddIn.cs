using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Outlook.Application;




namespace ClickLink
{
    public partial class ThisAddIn
    {
        

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           
        }

                              
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

		#region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }

        #endregion

        public void Delist()
        {
            string folderName = "Office365 Blacklist";

            List<string> delistedIPs = new List<string>();
            List<MailItem> unreadEmails = new List<MailItem>();

            //set inbox as the main folder
            Folder inBox = (Folder)
                Application.ActiveExplorer().Session.GetDefaultFolder
                (OlDefaultFolders.olFolderInbox);

            Application.ActiveExplorer().CurrentFolder = inBox.Folders[folderName];
            Application.ActiveExplorer().CurrentFolder.Display();

            //search for unread emails in folder
            Items items = inBox.Folders[folderName].Items;
            MailItem mail = items.Find("[Unread] =true");

            if (mail == null)
            {
                MessageBox.Show("Sorry, there are no unread items in this folder", "No Unread Emails");
                return;
            }
            while (mail != null)
            {
                unreadEmails.Add(mail);
                mail = items.FindNext();
            }
            foreach (MailItem unread in unreadEmails)
            {
                string html = unread.HTMLBody;
                unread.UnRead = false;

                //extract the link
                MatchCollection urls = Regex.Matches(html, @"<a\shref=""(?<url>.*?)"">(?<text>.*?)</a>");
                foreach (Match url in urls)
                {
                    string link = url.Groups["url"].Value + " -- Text = " + url.Groups["text"].Value;
                    link = ProperLink(link);
                                                    
                    int firstPos = link.IndexOf("ip=");
                    int secPos = link.IndexOf("&ttl");
                    string result = link.Substring(firstPos, secPos - firstPos);
                    string ip = result.Substring(result.IndexOf("=") + 1);

                    //click Delist button on website

                    WebBrowser webBrowser1 = new WebBrowser();
                    webBrowser1.Navigate(link);
                    webBrowser1.DocumentCompleted += webBrowser1_DocumentCompleted;

                    void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
                    {
                        HtmlElementCollection elc = ((WebBrowser)sender).Document.GetElementsByTagName("input");

                        foreach (HtmlElement el in elc)
                        {
                            if (el.GetAttribute("className").Equals("btn btn-default"))
                            {
                                el.InvokeMember("Click");
                                delistedIPs.Add(ip);
			    	if (unreadEmails.Count == delistedIPs.Count)
                                {
                                    var message = String.Join(Environment.NewLine, delistedIPs);
                                    MessageBox.Show(("The following IPs have been delisted: \n\n" + message), "Delistings Complete");
                                }
                            }
			}
                    }
                }
            }
        }

        public static string ProperLink(string input)
        {
            input = input.Replace("amp;", "");
            input = input.Remove(input.IndexOf(" ") + 1);
            return input;
        }
    }
}
          

      








   


        
    





