using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Excel4Engineers.WPFControls
{
    /// <summary>
    /// Interaction logic for MultiImageDisplay.xaml
    /// </summary>
    public partial class MultiImageDisplay : Window
    {

        public MultiImageDisplay(string title, System.Drawing.Image[] images, double scale = 1)
        {
            InitializeComponent();

            this.Title = title;

            Images = images.Select(i => ImageSourceHelper.CreateImagesource(i)).ToArray();

            Image.Width = Images.Max(i => i.Width) * scale;
            Image.Height = Images.Max(i => i.Height) * scale;

            SetIndex(0);
        }


        private int _Index;
        public int Index
        {
            get => _Index; 
        }

        private void SetIndex(int i)
        {
            _Index = i;
            if(i <= 0)
            {
                _Index = 0; //so index never goes below 0
                PrevButton.IsEnabled = false;
                NextButton.IsEnabled = true;
            }
            else if (i >= Images.Length - 1)
            {
                _Index = Images.Length - 1; //so index never goes beyond the last
                PrevButton.IsEnabled = true;
                NextButton.IsEnabled = false;
            }
            else
            {
                PrevButton.IsEnabled = true;
                NextButton.IsEnabled = true;
            }

            Image.Source = Images[i];
        }

        public ImageSource[] Images { get; private set; }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Close the dialog
            this.DialogResult = true;
        }

        public static void ShowImageDialog(string title, System.Drawing.Image[] images, double scale = 1)
        {
            MultiImageDisplay dialog = new MultiImageDisplay(title, images, scale);
            bool? result = dialog.ShowDialog();
        }

        private ImageSource CreateImagesource(System.Drawing.Image image)
        {
            using (var ms = new MemoryStream())
            {
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                ms.Seek(0, SeekOrigin.Begin);

                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.StreamSource = ms;
                bitmapImage.EndInit();

                return bitmapImage;
            }
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            SetIndex(Index + 1);
        }

        private void PrevButton_Click(object sender, RoutedEventArgs e)
        {
            SetIndex(Index - 1);
        }
    }
}
