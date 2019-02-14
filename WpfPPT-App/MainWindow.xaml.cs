using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using Syncfusion.Presentation;
using HtmlAgilityPack;
using System.Diagnostics;
using System.Windows.Documents;

namespace WpfPPT_App
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string filepath = AppDomain.CurrentDomain.BaseDirectory;
        string staticFilePath= Environment.CurrentDirectory + "\\data\\Img\\";

        public MainWindow()
        {
            InitializeComponent();         

        }
        private void CreatePPT_Btn_Click(object sender, RoutedEventArgs e)
        {
            CreateInitialPPt();
            // Repalce_Pic();
        }

        private void Suggest_Btn_Click(object sender, RoutedEventArgs e)
        {
               //getKeyWords(Txt_Blk.Text.Trim());
                getPictresFromWeb();
        }

        private void CreateInitialPPt()
        {
            String Title, TextDesc;
            Title = Title_txtBx.Text;
            TextDesc = Txt_Blk.Text;
            string today = DateTime.Now.ToShortDateString();

            //Creates a new ppt doc
            IPresentation ppt_doc = Presentation.Create();

            //Adding a initial slide to the ppt
            ISlide slide = ppt_doc.Slides.Add(SlideLayoutType.PictureWithCaption);

            //Specify the fill type and fill color for the slide background 
            slide.Background.Fill.FillType = FillType.Solid;
            slide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(232, 241, 229);

            //Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide
            IShape titleShape = slide.Shapes[0] as IShape;
            titleShape.TextBody.AddParagraph(Title).HorizontalAlignment = HorizontalAlignmentType.Center;

            //Adding a TextBox to the slide
            IShape shape = slide.AddTextBox(80, 200, 500, 100);
            shape.TextBody.AddParagraph(TextDesc);

            String imgName = GetImageName();
            imgName = staticFilePath + imgName;
            //Gets a picture as stream.
            Stream pictureStream = File.Open(imgName, FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Save the ppt
            ppt_doc.Save(Title + ".pptx");

            //Dispose the image stream
            pictureStream.Dispose();

            //closing the ppt
            ppt_doc.Close();
        }
        private void getPictresFromWeb()
        {
            String query,q;
            q = getBlockTextKeyWords();          

            if (Title_txtBx.Text.Length > 1)
                query = Title_txtBx.Text.Trim();              
            else//default text instead of error message
                query = "ApplePie";
            //adding textblocks bold
            if (!q.Equals(null))
                query = query +"+"+ q;


                int startPosition = 1;
            Boolean filterSimilarResults = true;
            String SafeSearchFiltering = "Moderate";


            string requestUrl = string.Format("http://images.google.com/images?" +
                                "q={0}&start={1}&filter={2}&safe={3}",
                                query, startPosition.ToString(),
                                (filterSimilarResults) ? 1.ToString() : 0.ToString(), SafeSearchFiltering);

            //WebClient client = new WebClient();
            //String url = "https://www.google.co.uk/images/srpr/logo11w.png";
            // client.DownloadFile(url, ".\\data\\img\\xyz.html");           


            WebClient client = new WebClient();
            string html = client.DownloadString(requestUrl);

            // Load the Html into the agility pack
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);

            // Now, using LINQ to get all Images
            List<HtmlNode> imageNodes = null;
            imageNodes = (from HtmlNode node in doc.DocumentNode.SelectNodes("//img")
                          where node.Name == "img"
                          select node).ToList();
            int i = 0;
            foreach (HtmlNode node in imageNodes)
            {
                client.DownloadFile(new Uri(node.Attributes["src"].Value), staticFilePath +"Photo"+ i + ".png");
                //Console.WriteLine(node.Attributes["src"].Value);//for testing
                if (i == 20)
                    break;
                i++;
            }
            i = 0;
            for (int rowIndex = 0; rowIndex < ScrabbleBoard.RowDefinitions.Count; rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < ScrabbleBoard.ColumnDefinitions.Count; columnIndex++)
                {
                    string uri = staticFilePath+"Photo" + (i + columnIndex) + ".png";
                    var imageSource = new BitmapImage(new Uri(uri));
                    var image = new Image { Source = imageSource };
                    Grid.SetRow(image, rowIndex);
                    Grid.SetColumn(image, columnIndex);
                    ScrabbleBoard.Children.Add(image);
                }
                i += ScrabbleBoard.ColumnDefinitions.Count;
            }
            client.Dispose();

        }

        private void Repalce_Pic()
        {

            //Opens an existing Presentation.
            IPresentation pptxDoc = Presentation.Open("Sample.pptx");

            //Retrieves the first slide from the Presentation.
            ISlide slide = pptxDoc.Slides[0];

            //Retrieves the first picture from the slide.
            IPicture picture = slide.Pictures[0];

            WebClient client = new WebClient();

            //Gets the new picture as stream.
            byte[] data = client.DownloadData(filepath);

            //Creates instance for memory stream
            MemoryStream memoryStream = new MemoryStream(data);

            //Replaces the existing image with new image.
            picture.ImageData = memoryStream.ToArray();

            //Saves the Presentation to the file system.
            pptxDoc.Save("Output.pptx");

            //Closes the Presentation
            pptxDoc.Close();
        }

        private String GetImageName()
        {
            for (int rowIndex = 0; rowIndex < ScrabbleBoard.RowDefinitions.Count; rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < ScrabbleBoard.ColumnDefinitions.Count; columnIndex++)
                {
                    CheckBox chkbx =(CheckBox) ScrabbleBoard.FindName("checkBox" + rowIndex.ToString() + columnIndex.ToString());
                    if (chkbx.IsChecked.Value)
                    {
                        string suffix;
                        if (rowIndex < 1)
                            suffix = columnIndex.ToString();
                        else
                            suffix = rowIndex.ToString() + columnIndex.ToString();

                        string imgName = "Photo" + suffix + ".png";
                        Txt_Blk.Text = imgName;
                        return imgName;

                    }
                }
            }
            return "Image.jpg";
        }
        private string getBlockTextKeyWords()
        {
            string result = null;            
            List<Inline> inlines = textBlock.Inlines.ToList();
            foreach (Run line in inlines)
            {
                if (line.FontWeight.Equals(FontWeights.Bold))
                    result += line.Text + "+";

            }
            return result;
        }
    }
}
