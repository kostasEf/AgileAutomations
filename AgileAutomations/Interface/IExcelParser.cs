using System.Collections.Generic;
using AgileAutomations.Model;

namespace AgileAutomations.Interface
{
    public interface IExcelParser
    {
        void AddReferenceNumbers(string path, IList<ContactFormData> contactFormDataList);
        IList<ContactFormData> ExtractRows(string path);
    }
}