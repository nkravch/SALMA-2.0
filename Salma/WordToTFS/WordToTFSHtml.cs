using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WordToTFS
{
    internal class WordToTFSHtml
    {
        /// <summary>
        /// Word html
        /// </summary>
        public string Html { get; private set; }

        /// <summary>
        /// List type enum
        /// </summary>
        private enum ListType
        {
            bulleted,
            numbered
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public WordToTFSHtml(string html)
        {
            Html = html;
        }

        /// <summary>
        /// Return TFS html
        /// </summary>
        public string GetTFSHtml()
        {
            HtmlDocument wordDoc = new HtmlDocument();
            HtmlDocument tfsDoc = new HtmlDocument();
            wordDoc.LoadHtml(Html);
            DoHtml(wordDoc.DocumentNode, tfsDoc.DocumentNode);
            return tfsDoc.DocumentNode.InnerHtml;
        }

        /// <summary>
        /// Create TFS html
        /// </summary>
        /// <param name="inputNode">Html node</param>
        /// <param name="outputNode">Html node</param>
        private void DoHtml(HtmlNode inputNode, HtmlNode outputNode)
        {
            HtmlNode appendNode = null;

            switch (inputNode.NodeType)
            {
                case HtmlNodeType.Document:
                    appendNode = outputNode;
                    break;
                case HtmlNodeType.Element:
                    if (IsList(inputNode))
                    {
                        DoList(inputNode, outputNode);
                        return;
                    }
                    else
                        appendNode = outputNode.AppendChild(inputNode.CloneNode(false));
                    break;
                case HtmlNodeType.Text:
                case HtmlNodeType.Comment:
                default:
                    if (inputNode.InnerHtml.Trim().Length == 0)
                        return;
                    appendNode = outputNode.AppendChild(inputNode.CloneNode(false));
                    break;
            }

            if (inputNode.HasChildNodes)
                foreach (var child in inputNode.ChildNodes)
                    DoHtml(child, appendNode);
        }

        /// <summary>
        /// Is list
        /// </summary>
        /// <param name="currentNode">Html node</param>
        private bool IsList(HtmlNode currentNode)
        {
            if (currentNode.NodeType == HtmlNodeType.Element 
                && currentNode.OriginalName.ToLower().Equals("p")
                && currentNode.HasAttributes)
                return Regex.IsMatch(currentNode.GetAttributeValue("style", "").ToLower(), "mso-list");
            /*
            if (currentNode.NodeType == HtmlNodeType.Element && currentNode.HasAttributes)
                return Regex.IsMatch(currentNode.GetAttributeValue("class", "").ToLower(), "msolistparagraph");*/

            return false;
        }

        /// <summary>
        /// Create new list and list items
        /// </summary>
        /// <param name="inputNode">Html node</param>
        /// <param name="outputNode">Html node</param>
        private void DoList(HtmlNode inputNode, HtmlNode outputNode)
        {
            HtmlNode appendNode = null;

            ListType type = GetListType(inputNode);

            if (IsNewList(inputNode, outputNode, type))
            {
                //appendNode = outputNode.AppendChild(HtmlNode.CreateNode("<div></div>"));

                if (type == ListType.bulleted)
                    //appendNode = appendNode.AppendChild(HtmlNode.CreateNode("<ul></l>"));
                    appendNode = outputNode.AppendChild(HtmlNode.CreateNode("<ul></l>"));
                else
                    //appendNode = appendNode.AppendChild(HtmlNode.CreateNode("<ol></ol>"));
                    appendNode = outputNode.AppendChild(HtmlNode.CreateNode("<ol></ol>"));
            }
            else
            {
                //appendNode = outputNode.ChildNodes.Where(n => n.OriginalName.ToLower().Equals("div")).LastOrDefault();
                appendNode = outputNode.LastChild; //appendNode.LastChild;
            }

            string value = GetListItemValue(inputNode);

            HtmlNode listItem = HtmlNode.CreateNode("<li></li>");
            HtmlNode textItem = HtmlNode.CreateNode("<span>" + value + "<br></span>");
            //textItem.SetAttributeValue("style", "font-family: Symbol;");
            listItem.AppendChild(textItem);
            appendNode.AppendChild(listItem);
        }


        /// <summary>
        /// Return list type - bulleted or numbered.
        /// </summary>
        /// <param name="currentNode">Html node</param>
        private ListType GetListType(HtmlNode currentNode)
        {
            int count = currentNode.Descendants().Where(n => n.OriginalName.ToLower().Equals("span") &&
                Regex.IsMatch(n.GetAttributeValue("style", "").ToLower(), @"symbol|courier\s*new|wingdings")).Count();

            if (count > 0)
                return ListType.bulleted;

            return ListType.numbered;
        }

        /// <summary>
        /// Return value from list item.
        /// </summary>
        /// <param name="currentNode">Html node</param>
        private string GetListItemValue(HtmlNode currentNode)
        {
            HtmlNode commentNode = currentNode.Descendants().Where(n => n.NodeType == HtmlNodeType.Comment && 
                n.InnerHtml.ToLower().Equals("<![if !supportlists]>")).FirstOrDefault();

            if (commentNode != null)
            {
                HtmlNode textNode = commentNode.NextSibling;

                if (textNode == null)
                    return "";

                textNode = textNode.NextSibling;

                if (textNode != null)
                    return textNode.InnerHtml.Trim().Replace("<o:p></o:p>", "");
            }

            return "";
           
        }

        /// <summary>
        /// Is new list in selection
        /// </summary>
        /// <param name="inputNode">Html node</param>
        /// <param name="outputNode">Html node</param>
        /// <param name="type">List type</param>
        private bool IsNewList(HtmlNode inputNode, HtmlNode outputNode, ListType type)
        {

            HtmlNode divNode = outputNode; //.ChildNodes.Where(n => n.OriginalName.ToLower().Equals("div")).LastOrDefault();

            if (divNode.LastChild == null)//if (divNode == null)
                return true;

            if (type == ListType.bulleted && !divNode.LastChild.OriginalName.ToLower().Equals("ul"))
                return true;

            if (type == ListType.numbered && !divNode.LastChild.OriginalName.ToLower().Equals("ol"))
                return true;

            return false;
        }
    }
}
