using System;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class MainForm : Form
    {
        float[] groups = new float[28]; // массив для удобного подсчёта ячеек в столбце "Степень выполнения групп"
        public MainForm()
        {
            InitializeComponent();
            Grid.AutoGenerateColumns = false; 
            Grid.AllowUserToAddRows = false; 
            Grid.Rows.Add("", "m", "№", "ai", "Omp", "Od", "Odai", "S пр", "O groups", "O", "S");
            Grid.Rows[0].ReadOnly = true;
            int m = 1; // переменная перечней показателей (второй столбец таблицы)
            int n_matrix = 131; // переменная номера элемента матрицы (третий столбец таблицы)
            for (int i = 1; i <= 7; i++) // і также обозначает номер этапа
            {
                if (i == 7) n_matrix += 10; // корректировка для соответствия исходной таблице
                for (int j = 0; j < 4; j++)
                {
                    var row = Grid.Rows[Grid.Rows.Add(i, m, n_matrix)];
                    row.Cells["Column10"].Value = ""; // пустые последние два столбца, 
                                                      // Grid_CellPainting сделал визуальное слияние ячеек
                    row.Cells["Column11"].Value = "";
                    m++;
                    n_matrix++;
                }
                n_matrix += 96;
            }
            Grid.Rows.Add("", "", "", "", "", "", "", 0); // последний ряд со счётчиком для столбца "Сравнение профилей"
            Grid.Rows[Grid.Rows.Count - 1].ReadOnly = true;
            Grid_CellValidated(new object(), new DataGridViewCellEventArgs(0, 0)); // инициализация ячеек-счётчиков нулями при старте программы, 
                                                                                   // когда пользователь ещё не нажал на какой-либо элемент формы
        }

        // проверка на то, что в смежных ячейках одного столбца значения совпадают
        // необходимо для визуального слияния ячеек в столбцах "Номер этапа", "Степень выполнения группы", 
        //      "Качественная оценка" и "Количественная оценка"
        bool IsTheSameCellValue(int column, int row)
        {
            DataGridViewCell cell1 = Grid[column, row];
            DataGridViewCell cell2 = Grid[column, row - 1];
            if (cell1.Value == null || cell2.Value == null)
                return false;
            return cell1.Value.ToString() == cell2.Value.ToString();
        }

        // убирает нижние границы и верхние границы, если значения в смежных ячейках одного столбца совпадают
        // необходимо для визуального слияния ячеек в столбцах "Номер этапа", "Степень выполнения группы", 
        //      "Качественная оценка" и "Количественная оценка"
        private void Grid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex == Grid.Rows.Count - 1)
            {
                e.AdvancedBorderStyle.Bottom = Grid.AdvancedCellBorderStyle.Bottom;
            }
            if (e.RowIndex < 1 || e.ColumnIndex < 0 ) 
                return;
            if (e.ColumnIndex >= 3 && e.ColumnIndex <= 7)
            {
                e.AdvancedBorderStyle.Top = Grid.AdvancedCellBorderStyle.Top;
                return;
            }
            if (IsTheSameCellValue(e.ColumnIndex, e.RowIndex))
                e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
            else
                e.AdvancedBorderStyle.Top = Grid.AdvancedCellBorderStyle.Top;
        }

        // обнуляет значение ячейки, если оно совпадает со значением смежной ячейки того же столбца
        // необходимо для визуального слияния ячеек в столбцах "Номер этапа", "Степень выполнения группы", 
        //      "Качественная оценка" и "Количественная оценка"
        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == 0 || (e.ColumnIndex >=3 && e.ColumnIndex <= 7))
                return;
            if (IsTheSameCellValue(e.ColumnIndex, e.RowIndex))
            {
                e.Value = "";
                e.FormattingApplied = true;
            }
        }

        // при каждом вводе значения в ячейку автоматически делает необходимые подсчёты
        private void Grid_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = Grid.Rows[e.RowIndex];
                if (row.Cells[3].Value != null && row.Cells[4].Value != null && row.Cells[5].Value != null)
                {
                    string valueA = row.Cells[3].Value.ToString().Replace('.', ','); // значение коэффициента важности
                    string valueB = row.Cells[4].Value.ToString().Replace('.', ','); // значение требуемого профиля безопасности
                    string valueC = row.Cells[5].Value.ToString().Replace('.', ','); // значение достигнутого профиля безопасности
                    float result; // нужно для метода float.TryParse()
                    if (float.TryParse(valueA, out result) && float.TryParse(valueB, out result) && float.TryParse(valueC, out result))
                    {
                        row.Cells[6].Value = Convert.ToSingle(valueA) * Convert.ToSingle(valueC); // коэффициент важности * профиль безопасности достигнутый
                        groups[e.RowIndex - 1] = Convert.ToSingle(row.Cells[6].Value);
                        if (Convert.ToSingle(valueB) == Convert.ToSingle(valueC)) row.Cells[7].Value = 1; // ячейка сравнение профилей
                        else row.Cells[7].Value = 0;                                                      // принимает значения - или 1
                    }

                    float quality_mark = 0; // столбец "Качественная оценка", среднее арифметическое значений "Степень выполнения групп"
                    for (ushort i = 0; i < 28; i += 4)
                    {
                        float acc = 0; // аккумулятор
                        for (ushort j = 0; j < 4; j++)
                        {
                            acc += groups[j + i];
                        }
                        quality_mark += acc;
                        Grid.Rows[i + 2].Cells[8].Value = acc; // визуальное слияние двух ячеек столбца "Степень выполнения групп"
                        Grid.Rows[i + 3].Cells[8].Value = acc;
                    }
                    Grid.Rows[13].Cells[9].Value = quality_mark / (float) 7; // подсчёт значения для столбца "Качественная оценка"

                    ushort counter = 0; // сумма всех значений столбца "Сравнение профилей"
                    ushort result_1 = 0; // нужно для метода UInt16.TryParse()
                    for (ushort i = 0; i < Grid.Rows.Count - 1; i++)
                    {
                        if (Grid.Rows[i].Cells[7].Value != null && UInt16.TryParse(Grid.Rows[i].Cells[7].Value.ToString(), out result_1))
                            counter += Convert.ToUInt16(Grid.Rows[i].Cells[7].Value);
                    }
                    Grid.Rows[Grid.Rows.Count - 1].Cells[7].Value = counter;
                    Grid.Rows[13].Cells[10].Value = (float) counter / (float) 28; // значение стоблца "Количественная оценка", дробное число от 0 до 1
                }

            }
        }

        // метод для предотвращения ввода нечисловых символов в ячейки
        private void Grid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (Grid.CurrentCell.ColumnIndex == 3 || Grid.CurrentCell.ColumnIndex == 4 || Grid.CurrentCell.ColumnIndex == 5)
            {
                e.Control.KeyPress += new KeyPressEventHandler(Grid_KeyPress);
            }
        }
        // метод для предотвращения ввода нечисловых символов в ячейки
        private void Grid_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != ',')
            {
                e.Handled = true;
            }
        }

        private void AboutprogramToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This program was created by Vladimir Pastukhov and Ivan Voitenko, students of the BIKS-21 " +
                "Faculty of Information Technologies of Kiev National University.", "About the program");
        }

        private void ReferenceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This program calculates information security risks based on the " +
                "entered data. User data is entered in cells with a green background, the rest" +
                "the values ​​are calculated automatically. Information about the number of the matrix element can be found in the" +
                "menu \" ISMS Matrix ");
        }

        private void matrixToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Matrix matrix = new Matrix();
            matrix.Show();
        }

        // код для экспорта данных в Excel, источник вдохновения: https://code.msdn.microsoft.com/office/How-to-Export-DataGridView-62f1f8ff
        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // создание объекта Excel  
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {
                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Risk calculation";

                ushort cellRowIndex = 1;
                ushort cellColumnIndex = 1;

                // считывание значений каждой ячейки
                for (ushort i = 0; i < Grid.Rows.Count - 1; i++)
                {
                    for (ushort j = 0; j < Grid.Columns.Count; j++)
                    {
                        // индексы Excel начитаются с 1,1. В первом ряду будут заголовки столбцов 
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = Grid.Columns[j].HeaderText;
                        }
                        else
                        {
                           worksheet.Cells[cellRowIndex, cellColumnIndex] = (Grid.Rows[i].Cells[j].Value != null) ? Grid.Rows[i].Cells[j].Value.ToString() : "" ;
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                // диалоговое окно для сохранение в файл 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Files Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Data exported successfully");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error!");
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }

        }

        // код для импрота данных из таблицы Excel. Проверено только с форматом xlsx (версия Excel 2010)
        private void importToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Files Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            dialog.FilterIndex = 1;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
                // понятия не имею, что обозначают аргументы в вызове метода excel.Workbooks.Open
                Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Open(dialog.FileName, 0, true, 5, "", "", true,
                                                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
                workbook.ReadOnlyRecommended = true;
                try
                {
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets.get_Item(1);

                    for (byte i = 2; i < 30; i ++) // 2 - 29 - номера ячеек в стоблцах таблицы (сверху вниз), значения которых нужно импортировать
                    {
                        for (byte j = 4; j < 7; j++) // 4 - 6 - номера ячеек в рядках таблицы (слева направо), значения которых нужно импортировать
                        {
                            // проверка на null-значение вместо null импортируется пустая строка
                            string cellValue = (worksheet.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value != null ? 
                                (worksheet.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value.ToString() : "";

                            Grid.Rows[i - 1].Cells[j - 1].Value = cellValue;
                            // i - 1, j - 1  - поправка на смещение индексов
                        }
                    }

                    for (ushort i = 0; i < Grid.Rows.Count; i++)
                        Grid_CellValidated(new object(), new DataGridViewCellEventArgs(0, i)); // "триггер" для того, чтобы значения ячеек,
                                                                                               // в которые не вводятся данные, 
                                                                                               // начали рассчитываться. Небольшой костыль
                    MessageBox.Show("Data imported successfully");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error!");
                }
                finally
                {
                    excel.Quit();
                    workbook = null;
                    excel = null;
                }

            }
        }

        // построение графика
        private void graphykToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Graphic gr = new Graphic(Grid);
            gr.Show();
        }

        private void Grid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    } 
}
