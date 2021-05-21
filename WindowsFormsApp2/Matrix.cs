using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Matrix : Form
    {
        public Matrix()
        {
            InitializeComponent();

            byte secondNumber = 1;
            byte thirdNumber = 1;
            for (byte column = 0; column < 20; column++)
            {
                for (byte row = 0; row < 8; row++)
                {
                    mainTable.Controls.Add(new Label
                    {
                        Text = row.ToString() + secondNumber.ToString() + thirdNumber.ToString(),
                        TextAlign = ContentAlignment.MiddleCenter
                    }, column, row + 3);
                }
                thirdNumber++;
                if ((column + 1) % 4 == 0)
                {
                    secondNumber++;
                    thirdNumber = 1;
                }
            }

        }

        #region Методы Paint некоторых ячеек матрицы для вертикального отображения текста

        private void Base_Paint(object sender, PaintEventArgs e)
        {
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Center;
            string Text = "Base";

            SizeF txt = e.Graphics.MeasureString(Text, Font);
            SizeF sz = e.Graphics.VisibleClipBounds.Size;

            e.Graphics.TranslateTransform(sz.Width, 0);
            e.Graphics.RotateTransform(90);
            e.Graphics.DrawString(Text, Font, Brushes.Black, new RectangleF(0, 12, sz.Height, sz.Width), format);
            e.Graphics.ResetTransform();
        }

        private void Structure_Paint(object sender, PaintEventArgs e)
        {
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Center;
            string Text = "Struct";

            SizeF txt = e.Graphics.MeasureString(Text, Font);
            SizeF sz = e.Graphics.VisibleClipBounds.Size;

            e.Graphics.TranslateTransform(sz.Width, 0);
            e.Graphics.RotateTransform(90);
            e.Graphics.DrawString(Text, Font, Brushes.Black, new RectangleF(0, 12, sz.Height, sz.Width), format);
            e.Graphics.ResetTransform();

        }

        private void Measures_Paint(object sender, PaintEventArgs e)
        {
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Center;
            string Text = "Measures";

            SizeF txt = e.Graphics.MeasureString(Text, Font);
            SizeF sz = e.Graphics.VisibleClipBounds.Size;

            e.Graphics.TranslateTransform(sz.Width, 0);
            e.Graphics.RotateTransform(90);
            e.Graphics.DrawString(Text, Font, Brushes.Black, new RectangleF(0, 12, sz.Height, sz.Width), format);
            e.Graphics.ResetTransform();
        }

        private void Facilities_Paint(object sender, PaintEventArgs e)
        {
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Center;
            string Text = "Funds";

            SizeF txt = e.Graphics.MeasureString(Text, Font);
            SizeF sz = e.Graphics.VisibleClipBounds.Size;

            e.Graphics.TranslateTransform(sz.Width, 0);
            e.Graphics.RotateTransform(90);
            e.Graphics.DrawString(Text, Font, Brushes.Black, new RectangleF(0, 12, sz.Height, sz.Width), format);
            e.Graphics.ResetTransform();
        }

        private void Stages_Paint(object sender, PaintEventArgs e)
        {
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Center;
            string Text = "<<< Stages";
            Font font = new Font("Microsoft Sans Serif", 12, FontStyle.Bold);

            SizeF txt = e.Graphics.MeasureString(Text, Font);
            SizeF sz = e.Graphics.VisibleClipBounds.Size;

            e.Graphics.TranslateTransform(0, sz.Height);
            e.Graphics.RotateTransform(270);
            e.Graphics.DrawString(Text, font, Brushes.Black, new RectangleF(0, 15, sz.Height, sz.Width), format);
            e.Graphics.ResetTransform();
        }

        #endregion

        private void mainTable_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label12_Click(object sender, System.EventArgs e)
        {

        }

        private void label22_Click(object sender, System.EventArgs e)
        {

        }

        private void label14_Click(object sender, System.EventArgs e)
        {

        }

        private void label8_Click(object sender, System.EventArgs e)
        {

        }

        private void Structure1_Click(object sender, System.EventArgs e)
        {

        }

        private void label20_Click(object sender, System.EventArgs e)
        {

        }

        private void Base1_Click(object sender, System.EventArgs e)
        {

        }
    }
}
