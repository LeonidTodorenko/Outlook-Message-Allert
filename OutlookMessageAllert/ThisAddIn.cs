namespace OutlookMessageAllert
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            EmailSubjectFiltering emailSubjectFiltering = new EmailSubjectFiltering();
            Application.NewMailEx += emailSubjectFiltering.outLookApp_NewMailEx;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }
    }
}
