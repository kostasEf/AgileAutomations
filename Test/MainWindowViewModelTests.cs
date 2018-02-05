using System.Collections.ObjectModel;
using System.Linq;
using AgileAutomations.Interface;
using AgileAutomations.Model;
using AgileAutomations.ViewModel;
using Moq;
using NUnit.Framework;

namespace Test
{
    [TestFixture]
    public class MainWindowViewModelTests
    {
        private IMainWindowViewModel mainWindowViewModel;
        private Mock<IExcelHelper> excelHelperMock;
        private Mock<IFileBrowser> fileBrowserMock;
        private Mock<IHtmlHelper> htmlHelperMock;
        private const string EXCEL_PATH = @"legit path";
        private ObservableCollection<ContactFormData> contactFormDataCollection;

        [SetUp]
        public void Initialize()
        {
            excelHelperMock = new Mock<IExcelHelper>();
            fileBrowserMock = new Mock<IFileBrowser>();
            htmlHelperMock = new Mock<IHtmlHelper>();

            mainWindowViewModel = new MainWindowViewModel(excelHelperMock.Object, fileBrowserMock.Object, htmlHelperMock.Object);
            contactFormDataCollection = new ObservableCollection<ContactFormData>
            {
                new ContactFormData {Name = "Name", Email = "Email", Subject = "Subject", Message = "Message", Reference = "123"},
                new ContactFormData {Name = "Name", Email = "Email", Subject = "Subject", Message = "Message", Reference = "321"}
            };
        }

        private static object[] dataSource =
        {
            new object[]
            {
                string.Empty
            },
            new object[]
            {
                " "
            },
            new object[]
            {
                null
            }
        };

        [Test]
        public void ExtractRows_WhenCalled_ContactFormDataCollectionGetsPopulated()
        {
            //Setup
            fileBrowserMock.Setup(fbm => fbm.GetFullName()).Returns(EXCEL_PATH);
            excelHelperMock.Setup(epm => epm.ExtractRows(It.IsAny<string>())).Returns(contactFormDataCollection);

            //Act
            mainWindowViewModel.ExtractRowsCommand.Execute(null);

            //Assert
            Assert.IsTrue(mainWindowViewModel.ContactFormDataCollection.Any());
        }

        [Test]
        public void ExtractRows_WhenCalled_FeedbackGetsSet()
        {
            //Setup
            fileBrowserMock.Setup(fbm => fbm.GetFullName()).Returns(EXCEL_PATH);
            excelHelperMock.Setup(epm => epm.ExtractRows(It.IsAny<string>())).Returns(contactFormDataCollection);

            //Act
            mainWindowViewModel.ExtractRowsCommand.Execute(null);

            //Assert
            Assert.IsTrue(string.Equals(mainWindowViewModel.Feedback, $"Rows found: {mainWindowViewModel.ContactFormDataCollection.Count}"));
        }

        [Test]
        public void ExtractRows_WhenCalled_FileSelectedBecomesTrue()
        {
            //Setup
            fileBrowserMock.Setup(fbm => fbm.GetFullName()).Returns(EXCEL_PATH);
            excelHelperMock.Setup(epm => epm.ExtractRows(It.IsAny<string>())).Returns(contactFormDataCollection);

            //Act
            mainWindowViewModel.ExtractRowsCommand.Execute(null);

            //Assert
            Assert.IsTrue(mainWindowViewModel.FileSelected);
        }

        [Test,TestCaseSource(nameof(dataSource))]
        public void ExtractRows_WhenPathIsInvalid_DoesNotExecuteExcelParser(string path)
        {
            //Setup
            fileBrowserMock.Setup(fbm => fbm.GetFullName()).Returns(path);

            //Act
            mainWindowViewModel.ExtractRowsCommand.Execute(null);

            //Assert
            excelHelperMock.Verify(epm => epm.ExtractRows(It.IsAny<string>()), Times.Never);
        }

        [Test,TestCaseSource(nameof(dataSource))]
        public void ExtractRows_WhenPathIsInvalid_DoesNotSetFeedBack(string path)
        {
            //Setup
            fileBrowserMock.Setup(fbm => fbm.GetFullName()).Returns(path);

            //Act
            mainWindowViewModel.ExtractRowsCommand.Execute(null);

            //Assert
            Assert.IsTrue(string.IsNullOrWhiteSpace(mainWindowViewModel.Feedback));
        }

        [Test, TestCaseSource(nameof(dataSource))]
        public void ExtractRows_WhenPathIsInvalid_DoesNotSetFileSelected(string path)
        {
            //Setup
            fileBrowserMock.Setup(fbm => fbm.GetFullName()).Returns(path);

            //Act
            mainWindowViewModel.ExtractRowsCommand.Execute(null);

            //Assert
            Assert.IsFalse(mainWindowViewModel.FileSelected);
        }

        [Test]
        public void SubmitForm_WhenCalled_FeedbackGetsSet()
        {
            //Act
            mainWindowViewModel.SubmitFormCommand.Execute(null);

            //Assert
            Assert.IsTrue(string.Equals(mainWindowViewModel.Feedback, "Success"));
        }

        [Test]
        public void SubmitForm_WhenCalled_FileSelectedBecomesFalse()
        {
            //Act
            mainWindowViewModel.SubmitFormCommand.Execute(null);

            //Assert
            Assert.IsFalse(mainWindowViewModel.FileSelected);
        }
    }
}
