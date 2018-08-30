using CommentAddInWeb.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace CommentAddInWeb
{
    public class OOXml
    {
        /// <summary>
        /// Convert Stream to CommentRange
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static ICollection<CommentRange> ConvertStreamToCommentRange(Stream stream)
        {
            using (var wordDoc = WordprocessingDocument.Open(stream, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                var document = mainPart.Document;
                var comments = mainPart.WordprocessingCommentsPart.Comments.ChildElements;
                List<CommentRange> commentRanges = new List<CommentRange>();

                foreach (Comment comment in comments)
                {
                    CommentRange commentRange = new CommentRange();
                    commentRange.Id = comment.Id;
                    commentRange.Initials = comment.Initials;
                    commentRange.Date = comment.Date;
                    commentRange.Author = comment.Author;
                    commentRange.Text = comment.InnerText;
                    OpenXmlElement rangeStart = document.Descendants<CommentRangeStart>().Where(c => c.Id == commentRange.Id).FirstOrDefault();
                    rangeStart = rangeStart.NextSibling();

                    List<OpenXmlElement> referenced = new List<OpenXmlElement>();
                    //append range text until the commentRangeEnd with the same comment id
                    while (!(rangeStart is CommentRangeEnd && ((CommentRangeEnd)rangeStart).Id.Value == comment.Id.Value))
                    {
                        referenced.Add(rangeStart);
                        rangeStart = rangeStart.NextSibling();
                    }

                    foreach (var ele in referenced)
                    {
                        if (!string.IsNullOrWhiteSpace(ele.InnerText))
                        {
                            commentRange.CommentedText += ele.InnerText;
                        }
                    }
                    commentRanges.Add(commentRange);
                }
                return commentRanges;
            }

        }
        /// <summary>
        /// Convert File Stream to Range Object
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static ICollection<Range> ConvertStreamToRange(Stream stream)
        {
            using (var wordDoc = WordprocessingDocument.Open(stream, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                var document = mainPart.Document;
                var comments = mainPart.WordprocessingCommentsPart.Comments.ChildElements;
                foreach (Comment comment in comments)
                {
                    CommentRange commentRange = new CommentRange();
                    commentRange.Id = comment.Id;
                    commentRange.Initials = comment.Initials;
                    commentRange.Date = comment.Date;
                    commentRange.Author = comment.Author;
                    commentRange.Text = comment.InnerText;
                    OpenXmlElement rangeStart = document.Descendants<CommentRangeStart>().Where(c => c.Id == commentRange.Id).FirstOrDefault();
                    List<OpenXmlElement> referenced = new List<OpenXmlElement>();
                    List<Range> ranges = new List<Range>();
                    rangeStart = rangeStart.NextSibling();

                    while (!(rangeStart is CommentRangeEnd))
                    {

                        referenced.Add(rangeStart);
                        rangeStart = rangeStart.NextSibling();
                        ranges.Add(new Range { Text = rangeStart.InnerText  });
                    }
                }
                return null;
            }
        }
        /// <summary>
        /// Returns a System.IO.Packaging.Package stream for the given word open XML.
        /// </summary>
        /// <param name="wordOpenXML">The word open XML.</param>
        /// <returns></returns>
        public static MemoryStream GetPackageStreamFromWordOpenXML(string wordOpenXML)
        {
            XDocument doc = XDocument.Parse(wordOpenXML);
            XNamespace pkg =
               "http://schemas.microsoft.com/office/2006/xmlPackage";
            XNamespace rel =
                "http://schemas.openxmlformats.org/package/2006/relationships";
            Package InmemoryPackage = null;
            MemoryStream memStream = new MemoryStream();
            using (InmemoryPackage = Package.Open(memStream, FileMode.Create))
            {
                // add all parts (but not relationships)
                foreach (var xmlPart in doc.Root
                    .Elements()
                    .Where(p =>
                        (string)p.Attribute(pkg + "contentType") !=
                        "application/vnd.openxmlformats-package.relationships+xml"))
                {
                    string name = (string)xmlPart.Attribute(pkg + "name");
                    string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                    if (contentType.EndsWith("xml"))
                    {
                        Uri u = new Uri(name, UriKind.Relative);
                        PackagePart part = InmemoryPackage.CreatePart(u, contentType,
                            CompressionOption.SuperFast);
                        using (Stream str = part.GetStream(FileMode.Create))
                        using (XmlWriter xmlWriter = XmlWriter.Create(str))
                            xmlPart.Element(pkg + "xmlData")
                                .Elements()
                                .First()
                                .WriteTo(xmlWriter);
                    }
                    else
                    {
                        Uri u = new Uri(name, UriKind.Relative);
                        PackagePart part = InmemoryPackage.CreatePart(u, contentType,
                            CompressionOption.SuperFast);
                        using (Stream str = part.GetStream(FileMode.Create))
                        using (BinaryWriter binaryWriter = new BinaryWriter(str))
                        {
                            string base64StringInChunks =
                           (string)xmlPart.Element(pkg + "binaryData");
                            char[] base64CharArray = base64StringInChunks
                                .Where(c => c != '\r' && c != '\n').ToArray();
                            byte[] byteArray =
                                System.Convert.FromBase64CharArray(base64CharArray,
                                0, base64CharArray.Length);
                            binaryWriter.Write(byteArray);
                        }
                    }
                }
                foreach (var xmlPart in doc.Root.Elements())
                {
                    string name = (string)xmlPart.Attribute(pkg + "name");
                    string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                    if (contentType ==
                        "application/vnd.openxmlformats-package.relationships+xml")
                    {
                        // add the package level relationships
                        if (name == "/_rels/.rels")
                        {
                            foreach (XElement xmlRel in
                                xmlPart.Descendants(rel + "Relationship"))
                            {
                                string id = (string)xmlRel.Attribute("Id");
                                string type = (string)xmlRel.Attribute("Type");
                                string target = (string)xmlRel.Attribute("Target");
                                string targetMode =
                                    (string)xmlRel.Attribute("TargetMode");
                                if (targetMode == "External")
                                    InmemoryPackage.CreateRelationship(
                                        new Uri(target, UriKind.Absolute),
                                        TargetMode.External, type, id);
                                else
                                    InmemoryPackage.CreateRelationship(
                                        new Uri(target, UriKind.Relative),
                                        TargetMode.Internal, type, id);
                            }
                        }
                        else
                        // add part level relationships
                        {
                            string directory = name.Substring(0, name.IndexOf("/_rels"));
                            string relsFilename = name.Substring(name.LastIndexOf('/'));
                            string filename =
                                relsFilename.Substring(0, relsFilename.IndexOf(".rels"));
                            PackagePart fromPart = InmemoryPackage.GetPart(
                                new Uri(directory + filename, UriKind.Relative));
                            foreach (XElement xmlRel in
                                xmlPart.Descendants(rel + "Relationship"))
                            {
                                string id = (string)xmlRel.Attribute("Id");
                                string type = (string)xmlRel.Attribute("Type");
                                string target = (string)xmlRel.Attribute("Target");
                                string targetMode =
                                    (string)xmlRel.Attribute("TargetMode");
                                if (targetMode == "External")
                                    fromPart.CreateRelationship(
                                        new Uri(target, UriKind.Absolute),
                                        TargetMode.External, type, id);
                                else
                                    fromPart.CreateRelationship(
                                        new Uri(target, UriKind.Relative),
                                        TargetMode.Internal, type, id);
                            }
                        }
                    }
                }
                InmemoryPackage.Flush();
            }
            return memStream;
        }

        /// <summary>
        /// create word file
        /// </summary>
        /// <param name="wordOpenXML"></param>
        /// <param name="filePath"></param>
        public static void CreatePackageFromWordOpenXML(string wordOpenXML, string filePath)
        {
            string packageXmlns = "http://schemas.microsoft.com/office/2006/xmlPackage";
            Package newPkg = Package.Open(filePath, FileMode.Create);

            try
            {
                XPathDocument xpDocument = new XPathDocument(new StringReader(wordOpenXML));
                XPathNavigator xpNavigator = xpDocument.CreateNavigator();

                XmlNamespaceManager nsManager = new XmlNamespaceManager(xpNavigator.NameTable);
                nsManager.AddNamespace("pkg", packageXmlns);
                XPathNodeIterator xpIterator = xpNavigator.Select("//pkg:part", nsManager);

                while (xpIterator.MoveNext())
                {
                    Uri partUri = new Uri(xpIterator.Current.GetAttribute("name", packageXmlns), UriKind.Relative);

                    PackagePart pkgPart = newPkg.CreatePart(partUri, xpIterator.Current.GetAttribute("contentType", packageXmlns));

                    // Set this package part's contents to this XML node's inner XML, sans its surrounding xmlData element.
                    string strInnerXml = xpIterator.Current.InnerXml
                        .Replace("<pkg:xmlData xmlns:pkg=\"" + packageXmlns + "\">", "")
                        .Replace("</pkg:xmlData>", "");
                    byte[] buffer = Encoding.UTF8.GetBytes(strInnerXml);
                    pkgPart.GetStream().Write(buffer, 0, buffer.Length);
                }

                newPkg.Flush();
            }
            finally
            {
                newPkg.Close();
            }
        }

    }
}