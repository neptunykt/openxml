using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlServer.WinWordRenderer;

public abstract class OpenXmlElementService
{
    protected Body _body;
    protected OpenXmlElementService(Body body)
    {
        _body = body;
    }
    /// <summary>
    /// Get List of OpenXmlElement from Body
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    public IEnumerable<T> GetOpenXmlElement<T>() where T : OpenXmlElement =>
        _body.Descendants<T>().ToList();

    /// <summary>
    /// Remove OpenXmlElement
    /// </summary>
    /// <param name="openXmlElement"></param>
    /// <typeparam name="T"></typeparam>
    public void RemoveOpenXmlElement<T>(T openXmlElement) where T : OpenXmlElement? =>
        _body.RemoveChild(openXmlElement);
}