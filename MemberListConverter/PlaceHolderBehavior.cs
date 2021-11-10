using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PlaceHolderSample
{
    // プレースホルダーを表示する添付ビヘイビア
    public static class PlaceHolderBehavior
    {
        // プレースホルダーとして表示するテキスト
        public static readonly DependencyProperty PlaceHolderTextProperty = DependencyProperty.RegisterAttached(
            "PlaceHolderText",
            typeof(string),
            typeof(PlaceHolderBehavior),
            new PropertyMetadata(null, OnPlaceHolderChanged));

        private static void OnPlaceHolderChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox == null)
            {
                return;
            }

            var placeHolder = e.NewValue as string;
            var handler = CreateEventHandler(placeHolder);
            if (string.IsNullOrEmpty(placeHolder))
            {
                textBox.TextChanged -= handler;
            }
            else
            {
                textBox.TextChanged += handler;
                if (string.IsNullOrEmpty(textBox.Text))
                {
                    textBox.Background = CreateVisualBrush(placeHolder);
                    // textBox.Background = new SolidColorBrush(Colors.White);
                    // textBox.Foreground = CreateVisualBrush(placeHolder);
                }
            }
        }

        private static TextChangedEventHandler CreateEventHandler(string placeHolder)
        {
            // TextChanged イベントをハンドルし、TextBox が未入力のときだけ
            // プレースホルダーを表示するようにする。
            return (sender, e) =>
            {
                // 背景に TextBlock を描画する VisualBrush を使って
                // プレースホルダーを実現
                var textBox = (TextBox)sender;
                if (string.IsNullOrEmpty(textBox.Text))
                {
                    // textBox.Background = new SolidColorBrush(Colors.White);
                    textBox.Background = CreateVisualBrush(placeHolder);
                    // textBox.Foreground = CreateVisualBrush(placeHolder);
                }
                else
                {
                    textBox.Background = new SolidColorBrush(Colors.Transparent);
                    // textBox.Foreground = new SolidColorBrush(Colors.Black);
                }
            };
        }

        private static VisualBrush CreateVisualBrush(string placeHolder)
        {
            var visual = new Label()
            {
                Content = placeHolder,
                Padding = new Thickness(5, 1, 1, 1),
                Foreground = new SolidColorBrush(Colors.DarkGray),
                FontWeight = FontWeights.Light,
                FontStyle = FontStyles.Italic,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Center,
            };
            return new VisualBrush(visual)
            {
                Stretch = Stretch.None,
                TileMode = TileMode.None,
                AlignmentX = AlignmentX.Left,
                AlignmentY = AlignmentY.Center,
            };
        }

        public static void SetPlaceHolderText(TextBox textBox, string placeHolder)
        {
            textBox.SetValue(PlaceHolderTextProperty, placeHolder);
        }

        public static string GetPlaceHolderText(TextBox textBox)
        {
            return textBox.GetValue(PlaceHolderTextProperty) as string;
        }
    }
}
