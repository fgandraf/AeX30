using AeX30.Core.Entities;

namespace AeX30.Core.Interfaces
{
    public interface IRequestService
    {
        Request LoadFromFile(string filePath);
    }
}
