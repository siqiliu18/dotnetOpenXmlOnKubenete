using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace SiqiSecondPptMerge.Properties
{
    public class OpenXmlHelper
    {
        private readonly string baseDatasetsRelativePath;

        //default constructer
        public OpenXmlHelper(string inputBasePath)
        {
            baseDatasetsRelativePath = inputBasePath;
        }

        public Drawing.TableCell CreateTextCell(string text)
        {
            Drawing.TableCell tc = new Drawing.TableCell(
                                new Drawing.TextBody(
                                    new Drawing.BodyProperties(),
                                new Drawing.Paragraph(
                                    new Drawing.Run(
                                        new Drawing.Text(text)))),
                                new Drawing.TableCellProperties());

            return tc;
        }

        public void SwapPhoto(SlidePart slidePart, string imgId)
        {
            Drawing.Blip blip = slidePart.Slide.Descendants<Drawing.Blip>().First();
            blip.Embed = imgId;
            slidePart.Slide.Save();
        }

        public void SwapPlaceholderText(SlidePart slidePart, string placeholder, string value)
        {
            //Find and get all the placeholder text locations 
            List<Drawing.Text> textList = slidePart.Slide.Descendants<Drawing.Text>()
                                        .Where(t => t.Text.Equals(placeholder)).ToList();

            //Swap the placeholder text with other text
            foreach (Drawing.Text text in textList)
            {
                text.Text = value;
            }
        }

        public string GetAbsolutePath(string relativePath)
        {
            FileInfo _dataRoot = new FileInfo(typeof(MainClass).Assembly.Location);
            string assemblyFolderPath = _dataRoot.Directory.FullName;

            string fullPath = Path.Combine(assemblyFolderPath, baseDatasetsRelativePath + relativePath);

            return fullPath;
        }
    }
}
