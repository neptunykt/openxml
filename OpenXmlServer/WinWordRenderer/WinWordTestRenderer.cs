using DocumentFormat.OpenXml.Packaging;
using OpenXmlClient.Classes.Models;
using OpenXmlServer.WinWordRenderer.HeaderFooterRenderer;
using OpenXmlServer.WinWordRenderer.TableRenderer;
using OpenXmlServer.WinWordRenderer.TextRenderer;

namespace OpenXmlServer.WinWordRenderer;

 public static class WinWordTestRenderer
    {
        /// <summary>
        /// For test purposes only
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="winWordRenderModel"></param>
        /// <param name="folderName"></param>
        /// <param name="saveFileName"></param>
        public static void Render(string fileName, WinWordRenderModel winWordRenderModel, string folderName, string saveFileName)
        {
            var directory = GetDirectoryPath();
            var byteArray = File.ReadAllBytes(Path.Combine($"{directory}/{folderName}", fileName));
            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                using (var wordDocument = WordprocessingDocument.Open(stream, true))
                {
                    var headerFooterFormatter = new WinWordHeaderFooterRenderer(wordDocument);
                    // Replace in header and footer
                    headerFooterFormatter.Render(winWordRenderModel.HeaderFooterStorage);
                    var body = wordDocument.MainDocumentPart?.Document.Body;
                    var winWordTextService = new WinWordTextService(body);
                    var winWordTableService = new WinWordTableService(winWordTextService, body);
                    var textRenderer = new WinWordGlobalTextRenderer(winWordTextService);
                    textRenderer.Render(winWordRenderModel);
                    var numberingTextRenderer = new WinWordNumberingTextRenderer(winWordTextService);
                    numberingTextRenderer.Render(winWordRenderModel);
                    var textGeneratorRenderer = new WinWordTextGeneratorRenderer(winWordTextService);
                    textGeneratorRenderer.Render(winWordRenderModel);
                    var tableRenderer = new WinWordTableRenderer(winWordTableService, winWordTextService);
                    tableRenderer.Clone(winWordRenderModel);
                    tableRenderer.Render(winWordRenderModel);
                }
                File.WriteAllBytes(Path.Combine($"{directory}/{folderName}", saveFileName), stream.ToArray()); 
            }
        }

        public static string GetDirectoryPath(string localPath = null)
        {
            var currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var rootDirectory = currentDirectory?.Replace(@"bin/Debug/netcoreapp3.1/", "");
            return localPath == null ? rootDirectory : $@"{rootDirectory}/{localPath}";
        }
        
        
    }