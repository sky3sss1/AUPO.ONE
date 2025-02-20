using Domain;

namespace Application.Interfaces;

public interface IParser
{
    public List<Vulnerability> MapFromByte(byte[] byteArray);
}
