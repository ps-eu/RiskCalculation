using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace WindowsFormsApp2
{
    public partial class Graphic : Form
    {
        public Graphic(DataGridView datagrid)
        {
            InitializeComponent();
            chart.Series.Clear();

            // зелёная линия требуемого профиля безопасности
            chart.Series.Add("Required Cybersecurity Profile");
            chart.Series[0].ChartType = SeriesChartType.Line;
            chart.Series[0].BorderWidth = 2;
            chart.Series[0].XValueType = ChartValueType.UInt32;
            chart.Series[0].YValueType = ChartValueType.Single;
            chart.Series[0].Color = System.Drawing.Color.Green;

            // красная линия достигнутого профиля безопасности
            chart.Series.Add("Achieved Cybersecurity Profile");
            chart.Series[1].ChartType = SeriesChartType.Line;
            chart.Series[1].BorderWidth = 4;
            chart.Series[1].XValueType = ChartValueType.UInt32;
            chart.Series[1].YValueType= ChartValueType.Single;
            chart.Series[1].Color = System.Drawing.Color.Red;

            // интервалы осей Х и У
            chart.ChartAreas[0].AxisX.Interval = 1;
            chart.ChartAreas[0].AxisY.Interval = 0.1;

            // цикл для добавления в график каждой точки с проверкой на null-значение
            for (ushort i = 1; i < datagrid.Rows.Count - 1; i++)
            {
                // координата Х для точки-номера из столбца "Перечень показателей", общая координата для всех линий графика
                var X = datagrid.Rows[i].Cells[1].Value != null ? System.Convert.ToInt32(datagrid.Rows[i].Cells[1].Value.ToString().Replace('.',',')) : 0; 

                // координата У достигнутого профиля безопасности 
                var Y_a = datagrid.Rows[i].Cells[5].Value != null ? System.Convert.ToSingle(datagrid.Rows[i].Cells[5].Value.ToString().Replace('.', ',')) : 0;

                // координата У требуемого профиля безопасности
                var Y_r = datagrid.Rows[i].Cells[4].Value != null ? System.Convert.ToSingle(datagrid.Rows[i].Cells[4].Value.ToString().Replace('.', ',')) : 0;

                // ToString().Replace('.',',') необходимо для избежания ошибки, если пользователь ввёл в дробном числе точку, а не запятую (0.6, а не 0,6)

                // добавление точек на график
                chart.Series[0].Points.AddXY(X, Y_r);
                chart.Series[1].Points.AddXY(X, Y_a);
            }
        }

        private void chart_Click(object sender, System.EventArgs e)
        {

        }
    }
}
