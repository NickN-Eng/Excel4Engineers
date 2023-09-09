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

namespace Excel4Engineers.WPFControls
{
    /// <summary>
    /// Interaction logic for ImageDisplay.xaml
    /// </summary>
    public partial class ImageDisplay : Window
    {
        public ImageDisplay(string title, System.Drawing.Image image, double scale = 1)
        {
            InitializeComponent();

            this.Title = title;
            Image.Source = ImageSourceHelper.CreateImagesource(image);
            Image.Width = Image.Source.Width * scale;
            Image.Height = Image.Source.Height * scale;
        }


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Close the dialog
            this.DialogResult = true;
        }

        public static void ShowImageDialog(string title, System.Drawing.Image image, double scale = 1)
        {
            ImageDisplay dialog = new ImageDisplay(title, image,scale);
            bool? result = dialog.ShowDialog();
        }


    }
}
