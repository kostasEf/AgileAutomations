using AgileAutomations.Interface;
using AgileAutomations.Model;
using SimpleBrowser;

namespace AgileAutomations.Helper
{
    public class HtmlHelper : IHtmlHelper
    {
        public string SubmitContactForm(ContactFormData contactFormData)
        {
            var browser = new Browser();
            var referenceNo = string.Empty;

            browser.UserAgent = "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.10 (KHTML, like Gecko) Chrome/8.0.552.224 Safari/534.10";

            browser.Navigate("https://agileautomations.co.uk/home/inputform");
            if (LastRequestFailed(browser)) return referenceNo;

            var name = browser.Find("ContactName");
            var mail = browser.Find("ContactEmail");
            var sub = browser.Find("ContactSubject");
            var message = browser.Find("Message");
            var sendButton = browser.Find(ElementType.Button, FindBy.Value, "Send message");

                
            if (!sendButton.Exists || !name.Exists || !mail.Exists || !sub.Exists || !message.Exists)
            {
                return referenceNo;
            }

            name.Value = contactFormData.Name;
            mail.Value = contactFormData.Email;
            sub.Value = contactFormData.Subject;
            message.Value = contactFormData.Message;

            sendButton.Click();
            if (LastRequestFailed(browser)) return referenceNo;

            referenceNo = browser.Find("ReferenceNo").Value.Split('-')[1].Trim();


            return referenceNo;
        }

        private bool LastRequestFailed(Browser browser)
        {
            return browser.LastWebException != null;
        }

    }
}