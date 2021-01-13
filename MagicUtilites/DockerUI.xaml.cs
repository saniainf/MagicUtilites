using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using corel = Corel.Interop.VGCore;

namespace MagicUtilites
{
    public partial class DockerUI : UserControl
    {
        private corel.Application corelApp;
        private Styles.StylesController stylesController;
        public DockerUI(object app)
        {
            InitializeComponent();
            try
            {
                this.corelApp = app as corel.Application;
                stylesController = new Styles.StylesController(this.Resources, this.corelApp);
            }
            catch
            {
                global::System.Windows.MessageBox.Show("VGCore Erro");
            }

        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();
        }

        private void ConvertToCurves(object sender, RoutedEventArgs e)
        {
            MakeToAllPages((s) =>
            {
                if (s.Type == corel.cdrShapeType.cdrTextShape) // если текст
                    s.ConvertToCurves(); // перевести в кривые
            });
        }

        private void BitmapToCMYK(object sender, RoutedEventArgs e)
        {
            MakeToAllPages((s) =>
            {
                if (s.Type == corel.cdrShapeType.cdrBitmapShape) // если картинка
                    if (s.Bitmap.Mode != corel.cdrImageType.cdrCMYKColorImage) // цветовая модель не CMYK
                        s.Bitmap.ConvertTo(corel.cdrImageType.cdrCMYKColorImage); // конвертировать в CMYK
            });
        }

        private void UniformFillToCMYK(object sender, RoutedEventArgs e)
        {
            MakeToAllPages((s) =>
            {
                if (s.CanHaveFill) // у объекта может быть заливка
                    if (s.Fill.Type == corel.cdrFillType.cdrUniformFill) // заливка сплошная
                        if (s.Fill.UniformColor.Type != corel.cdrColorType.cdrColorCMYK) // цветовая модель не CMYK
                            s.Fill.UniformColor.ConvertToCMYK(); // конвертировать в CMYK
            });
        }

        private void OutlineFillToCMYK(object sender, RoutedEventArgs e)
        {
            MakeToAllPages((s) =>
            {
                if (s.CanHaveOutline) // у объекта может быть обводка
                    if (s.Outline.Type == corel.cdrOutlineType.cdrOutline) // обводка есть
                        if (s.Outline.Color.Type != corel.cdrColorType.cdrColorCMYK) // цветовая модель не CMYK
                            s.Outline.Color.ConvertToCMYK(); // конвертировать в CMYK
            });
        }

        private void FountainFillToCMYK(object sender, RoutedEventArgs e)
        {
            MakeToAllPages((s) =>
            {
                if (s.CanHaveFill) // у объекта может быть заливка
                    if (s.Fill.Type == corel.cdrFillType.cdrFountainFill) // заливка градиент
                    {
                        foreach (corel.FountainColor c in s.Fill.Fountain.Colors) // перебор всех ключей в градиенте
                        {
                            if (c.Color.Type != corel.cdrColorType.cdrColorCMYK) // цветовая модель не CMYK
                                c.Color.ConvertToCMYK(); // конвертировать в CMYK
                        }
                    }
            });
        }

        private void ResampleBitmap(object sender, RoutedEventArgs e)
        {
            MakeToAllPages((s) =>
            {
                int resolution = 300;
                if (s.Type == corel.cdrShapeType.cdrBitmapShape) // если картинка
                    if (s.Bitmap.ResolutionX != resolution || s.Bitmap.ResolutionY != resolution) // разрешение не совпадает с заданным
                        s.Bitmap.Resample(0, 0, true, resolution, resolution); // изменяем разрешение на заданное
            });
        }

        private void MakeToAllPages(Action<corel.Shape> action)
        {
            if (corelApp.ActiveDocument == null)
                return;
            corelApp.BeginDraw();
            foreach (corel.Page page in corelApp.ActiveDocument.Pages)
            {
                MakeToShapeRange(page.Shapes.All(), action);
            }
            corelApp.EndDraw();
        }

        private void MakeToShapeRange(corel.ShapeRange sr, Action<corel.Shape> action)
        {
            foreach (corel.Shape shape in sr)
            {
                if (shape.Type == corel.cdrShapeType.cdrGroupShape)
                    MakeToShapeRange(shape.Shapes.All(), action);

                if (shape.PowerClip != null)
                    MakeToShapeRange(shape.PowerClip.Shapes.All(), action);

                action(shape);
            }
        }
    }
}
