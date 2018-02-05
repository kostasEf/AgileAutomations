using AgileAutomations.Helper;
using AgileAutomations.Interface;
using AgileAutomations.Model;
using NUnit.Framework;

namespace Test
{
    [TestFixture]
    public class HtmlHelperTests
    {
        private IHtmlHelper htmlHelper;
        private ContactFormData contactFormData;

        [OneTimeSetUp]
        public void Initialize()
        {
            htmlHelper = new HtmlHelper();

            contactFormData = new ContactFormData
            {
                Name = "Name",
                Email = "Email",
                Subject = "Subject",
                Message = "Message",
                Reference = "123"
            };
        }

        [Test]
        public void SubmitContactForm_WhenCalled_ReturnsReferenceNo()
        {
            Assert.IsFalse(string.IsNullOrWhiteSpace(htmlHelper.SubmitContactForm(contactFormData)));
        }
    }
}