using AgileAutomations.Model;

namespace AgileAutomations.Interface
{
    public interface IHtmlHelper
    {
        string SubmitContactForm(ContactFormData contactFormData);
    }
}