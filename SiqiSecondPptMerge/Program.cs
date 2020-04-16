using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SiqiSecondPptMerge.Properties;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace SiqiSecondPptMerge
{
    class MainClass
    {
        //private static readonly string BaseDatasetsRelativePath = @"../../../Data";
        //private static readonly string TemplatePath = GetAbsolutePath(BaseDatasetsRelativePath + "/template.pptx");
        //private static readonly string OutputPath = GetAbsolutePath(BaseDatasetsRelativePath + "/output.pptx");
        //private static readonly string ImagePath = GetAbsolutePath(BaseDatasetsRelativePath + "/testingImage.png");

        public static void Main(string[] args)
        {
            OpenXmlHelper openXmlProcessor = new OpenXmlHelper(@"../../../Data");

            string TemplatePath = openXmlProcessor.GetAbsolutePath("/template.pptx");
            string OutputPath = openXmlProcessor.GetAbsolutePath("/output.pptx");
            string ImagePath = openXmlProcessor.GetAbsolutePath("/testingImage.png");

            File.Copy(TemplatePath, OutputPath, true);

            using (PresentationDocument presentationDoc = PresentationDocument.Open(OutputPath, true))
            {
                PresentationPart presPart = presentationDoc.PresentationPart;

                // Step 1: Identify the proper Part Id
                SlidePart contentSlidePart = (SlidePart)presPart.GetPartById("rId2");

                // Step 2: Replace one image with external file
                string imageRel = "rIdImg";
                int imageRelId = 1;
                var imgId = imageRel + imageRelId;

                ImagePart imagePart = contentSlidePart.AddImagePart(ImagePartType.Jpeg, imgId);

                using (Image image = Image.FromFile(ImagePath))
                {
                    using (MemoryStream m = new MemoryStream())
                    {
                        image.Save(m, image.RawFormat);
                        m.Position = 0;
                        imagePart.FeedData(m);

                        openXmlProcessor.SwapPhoto(contentSlidePart, imgId);
                    }
                }

                // Step 3: Replace text matched by the key
                openXmlProcessor.SwapPlaceholderText(contentSlidePart, "{{Title}}", "Testing Title for IBM");

                // Step 4: Fill out multi value fields by table
                var tupleList = new List<(string category, string model, string price)>
                {
                    ("Automobile", "Ford", "$25K"),
                    ("Automobile", "Toyota", "$30K"),
                    ("Computer", "IBM PC", "$2.5K"),
                    ("Laptop", "Dell", "$1K"),
                    ("Laptop", "Microsoft", "$2K")
                };

                Drawing.Table tbl = contentSlidePart.Slide.Descendants<Drawing.Table>().First();

                foreach (var row in tupleList)
                {
                    Drawing.TableRow tr = new Drawing.TableRow();
                    tr.Height = 100;
                    tr.Append(openXmlProcessor.CreateTextCell(row.category));
                    tr.Append(openXmlProcessor.CreateTextCell(row.model));
                    tr.Append(openXmlProcessor.CreateTextCell(row.price));
                    tbl.Append(tr);
                }

                // Step 5: Save the presentation
                presPart.Presentation.Save();
            }
        }

        //static Drawing.TableCell CreateTextCell(string text)
        //{
        //    Drawing.TableCell tc = new Drawing.TableCell(
        //                        new Drawing.TextBody(
        //                            new Drawing.BodyProperties(),
        //                        new Drawing.Paragraph(
        //                            new Drawing.Run(
        //                                new Drawing.Text(text)))),
        //                        new Drawing.TableCellProperties());

        //    return tc;
        //}

        //static void SwapPhoto(SlidePart slidePart, string imgId)
        //{
        //    Drawing.Blip blip = slidePart.Slide.Descendants<Drawing.Blip>().First();
        //    blip.Embed = imgId;
        //    slidePart.Slide.Save();
        //}

        //static void SwapPlaceholderText(SlidePart slidePart, string placeholder, string value)
        //{
        //    //Find and get all the placeholder text locations 
        //    List<Drawing.Text> textList = slidePart.Slide.Descendants<Drawing.Text>()
        //                                .Where(t => t.Text.Equals(placeholder)).ToList();

        //    //Swap the placeholder text with other text
        //    foreach (Drawing.Text text in textList)
        //    {
        //        text.Text = value;
        //    }
        //}

        //public static string GetAbsolutePath(string relativePath)
        //{
        //    FileInfo _dataRoot = new FileInfo(typeof(MainClass).Assembly.Location);
        //    string assemblyFolderPath = _dataRoot.Directory.FullName;

        //    string fullPath = Path.Combine(assemblyFolderPath, relativePath);

        //    return fullPath;
        //}
    }
}
