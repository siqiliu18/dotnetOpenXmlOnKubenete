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
        public static void Main(string[] args)
        {
            string dataPath = @"../../../Data";
            OpenXmlHelper openXmlProcessor = new OpenXmlHelper(dataPath);

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

            // set watcher to watch output ppt changes
            Watcher watcher = new Watcher(dataPath);
            watcher.Run();
        }
    }
}
