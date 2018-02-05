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
            string referenceNo = string.Empty;

            browser.UserAgent = "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.10 (KHTML, like Gecko) Chrome/8.0.552.224 Safari/534.10";

            browser.Navigate("https://agileautomations.co.uk/home/inputform");
            if (LastRequestFailed(browser))
            {
                return referenceNo;
            }

            HtmlResult name = browser.Find("ContactName");
            HtmlResult mail = browser.Find("ContactEmail");
            HtmlResult sub = browser.Find("ContactSubject");
            HtmlResult message = browser.Find("Message");
            HtmlResult sendButton = browser.Find(ElementType.Button, FindBy.Value, "Send message");

                
            if (!sendButton.Exists || !name.Exists || !mail.Exists || !sub.Exists || !message.Exists)
            {
                return referenceNo;
            }

            name.Value = contactFormData.Name;
            mail.Value = contactFormData.Email;
            sub.Value = contactFormData.Subject;
            message.Value = contactFormData.Message;

            sendButton.Click();
            if (LastRequestFailed(browser))
            {
                return referenceNo;
            }

            referenceNo = browser.Find("ReferenceNo").Value.Split('-')[1].Trim();


            return referenceNo;
        }

        private bool LastRequestFailed(Browser browser)
        {
            return browser.LastWebException != null;
        }

    }
}