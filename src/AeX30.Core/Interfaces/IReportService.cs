using AeX30.Core.Entities;

namespace AeX30.Core.Interfaces
{
    public interface IReportService
    {
        bool SaveReport(string templatePath, string saveAsPath, Report report);
    }
}
