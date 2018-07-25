using Outlook = Microsoft.Office.Interop.Outlook;
using System.Configuration;

namespace OutlookMessageAllert
{
    public class EmailSubjectFiltering
    {



        public void outLookApp_NewMailEx(string entryIdCollection)
        {
            try
            {
                Outlook.Application application = new Outlook.Application();
                Outlook._NameSpace nameSpace = application.GetNamespace("MAPI");

                var configuration = ConfigurationManager.AppSettings;
                var stopWordsText = configuration["stopList"];
                string[] stopWordsArray = stopWordsText.Split(',');
                Outlook.MailItem item = nameSpace.GetItemFromID(entryIdCollection);
                if (item != null)
                {
                    if (item.MessageClass == "IPM.Note")
                    {
                        foreach (var filter in stopWordsArray)
                        {
                            if (!string.IsNullOrEmpty(item.Subject))
                            {
                                if (item.Subject.ToUpper().Contains(filter.ToUpper()))
                                {
                                    ShowMessage(@"Urgent mail recived");
                                    break;
                                }
                            }
                            if (item.Importance == Outlook.OlImportance.olImportanceHigh)
                            {
                                ShowMessage(@"Urgent mail recived");
                                break;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ShowMessage(ex.Message);
            }
        }

        private void ShowMessage(string text)
        {
            MessageAlert messageAlert = new MessageAlert();
            messageAlert.label1.Text = text;
            messageAlert.TopMost = true;
            messageAlert.ShowDialog();
        }
    }
}
