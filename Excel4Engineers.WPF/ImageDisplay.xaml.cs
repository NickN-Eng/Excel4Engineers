using System;
using System.Collections.Generic;
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

namespace Excel4Engineers.WPF
{
    /// <summary>
    /// Interaction logic for ImageDisplay.xaml
    /// </summary>
    public partial class ImageDisplay : Window
    {
        public ImageDisplay(string title, System.Drawing.Image image)
        {
            InitializeComponent();

            Image.Source = image;
        }


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Close the dialog
            this.DialogResult = true;
        }

        public static void ShowImageDialog(string title, Image image)
        {
            ImageDisplay dialog = new ImageDisplay(title, image);
            bool? result = dialog.ShowDialog();

            if (result == true)
            {
                // The user clicked OK
            }
        }

        private ImageSource GenerateBarcode(Image image)
        {
            using (var ms = new MemoryStream())
            {
                image.Save(ms, ImageFormat.Bmp);
                ms.Seek(0, SeekOrigin.Begin);

                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.StreamSource = ms;
                bitmapImage.EndInit();

                return bitmapImage;
            }
        }
    }
}
