using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace стпо1
{
    public partial class Form1 : Form
    {
        bool errorExeCalled = false;// контроль однократного вызова
        double currentEps;
        string rawA;
        string rawB;
        string rawH;

        public Form1()
        {
            InitializeComponent();
            InitTable();
        }

        void InitTable()
        {
            dgv.RowCount = 20;
            dgv.ColumnCount = 3;

            dgv.Columns[0].HeaderText = "Парабола";
            dgv.Columns[1].HeaderText = "Трапеции";
            dgv.Columns[2].HeaderText = "Монте-Карло";

            for (int i = 0; i < 20; i++)
                dgv.Rows[i].HeaderCell.Value = $"{i + 1}";
        }

        // ======= ЭТАЛОН (Ньютон–Лейбниц) ==========
        double ExactIntegral(double a, double b, int n)
        {
            double Fa = 0;
            double Fb = 0;

            for (int k = 0; k <= n; k++)
            {
                // первообразная: x^(k+1)
                Fa += Math.Pow(a, k + 1);
                Fb += Math.Pow(b, k + 1);
            }

            return Fb - Fa;
        }

        // ============ ВЫЗОВ Integral3x.exe ============
        double CallIntegralExe(double a, double b, double h, int method, int degree)
        {
            string args = $"{a} {b} {h} {method}";
            for (int i = 1; i <= degree + 1; i++)
                args += $" {i}";

            string output = CallIntegralExeRaw(args);

            Match m = Regex.Match(output, @"S\s*=\s*([-+]?\d+(?:[.,]\d+)?)");
            if (!m.Success)
                throw new Exception("Integral3x.exe не вернул результат");

            string number = m.Groups[1].Value.Replace(',', '.');
            return double.Parse(number, CultureInfo.InvariantCulture);
        }

        string CallIntegralExeRaw(string args)
        {
            string exePath = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Integral3x.exe"
            );

            if (!System.IO.File.Exists(exePath))
                throw new Exception("Integral3x.exe не найден");

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = args,

                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardInput = true,
                CreateNoWindow = true
            };

            using (Process p = Process.Start(psi))
            {
                // закрываем system("pause")
                p.StandardInput.WriteLine();
                p.StandardInput.Flush();

                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();

                return output;
            }
        }

        // ================= КНОПКА =================
        private void buttonCalc_Click(object sender, EventArgs e)
        {
            textBoxLog.Clear();
            errorExeCalled = false;
            rawA = textBoxA.Text.Trim();
            rawB = textBoxB.Text.Trim();
            rawH = textBoxH.Text.Trim();

            if (!double.TryParse(textBoxEPSt.Text, out double eps) || eps < 0)
            {
                MessageBox.Show(
                    "Невозможно сохранить данные.\nПогрешность EPSt не задана или задана некорректно.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            currentEps = eps;


            if (!TryReadInput(out double a, out double b, out double h, out string expectedError))
            {
                HandleInputError(a, b, h, 1, 1, expectedError);
                return;
            }

            for (int n = 1; n <= 20; n++)
                {
                    double exact = ExactIntegral(a, b, n);

                    for (int m = 1; m <= 3; m++)
                    {
                        try
                        {
                            double num = CallIntegralExe(a, b, h, m, n);
                            double EPSf = Math.Abs(num - exact);

                            var cell = dgv.Rows[n - 1].Cells[m - 1];
                            cell.Value = EPSf.ToString();

                            cell.Style.BackColor =
                                EPSf <= eps ? Color.LightGreen : Color.LightCoral;
                        }
                        catch (Exception ex)
                        {
                            HandleInputError(a, b, h, m, n, ex.Message);
                            return;
                        }
                    }
                }
        }

        private void buttonSaveExcel_Click(object sender, EventArgs e)
        {
            if (currentEps <= 0)
            {
                MessageBox.Show(
                    "Невозможно сохранить данные.\nПогрешность EPSt не задана или задана некорректно.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            if (!TryReadInput(out double a, out double b, out double h, out string expectedError))
            {
                MessageBox.Show(
                    "Невозможно сохранить данные.\nПроверьте корректность исходных данных.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            SaveResultsToExcel(a, b, h, currentEps);
        }

        void HandleInputError(double a, double b, double h, int method, int degree, string expected)
        {
            if (errorExeCalled) return;
            errorExeCalled = true;

            List<string> argsList = new List<string>();

            if (!string.IsNullOrWhiteSpace(rawA)) argsList.Add(rawA);
            if (!string.IsNullOrWhiteSpace(rawB)) argsList.Add(rawB);
            if (!string.IsNullOrWhiteSpace(rawH)) argsList.Add(rawH);
            argsList.Add(method.ToString());
            argsList.Add(degree.ToString());

            string args = string.Join(" ", argsList);

            string rawOutput = CallIntegralExeRaw(args);

            string actual = rawOutput.TrimEnd();

            // удаляем последнюю строку (pause)
            int lastNewLine = actual.LastIndexOf('\n');
            if (lastNewLine >= 0)
                actual = actual.Substring(0, lastNewLine).TrimEnd();

            textBoxLog.Text += "Негативный тест-кейс (ошибка ввода)\r\n\r\n" + "Ожидаемый отклик:\r\n" + expected + "\r\n\r\n" + "Фактический отклик:\r\n"
                + actual;
        }

        void SaveResultsToExcel(double a, double b, double h, double eps)
        {
            Excel.Application excel = new Excel.Application();
            excel.Workbooks.Add();

            Excel.Worksheet sheet = excel.ActiveSheet;
            sheet.Name = "Результаты интегрирования";

            int row = 1;

            // ===== СЛУЖЕБНАЯ ИНФОРМАЦИЯ =====
            sheet.Cells[row++, 1] = "Исходные данные";
            sheet.Cells[row++, 1] = "Левая граница a:";
            sheet.Cells[row - 1, 2] = a;

            sheet.Cells[row++, 1] = "Правая граница b:";
            sheet.Cells[row - 1, 2] = b;

            sheet.Cells[row++, 1] = "Шаг h:";
            sheet.Cells[row - 1, 2] = h;

            sheet.Cells[row++, 1] = "Допустимая погрешность EPSt:";
            sheet.Cells[row - 1, 2] = eps;

            sheet.Cells[row++, 1] = "Эталонный метод:";
            sheet.Cells[row - 1, 2] = "Ньютон–Лейбниц";

            row += 2;

            // ===== ЗАГОЛОВКИ ТАБЛИЦЫ =====
            sheet.Cells[row, 1] = "Степень полинома";
            sheet.Cells[row, 2] = "Парабола";
            sheet.Cells[row, 3] = "Трапеции";
            sheet.Cells[row, 4] = "Монте-Карло";

            Excel.Range header = sheet.Range[
                sheet.Cells[row, 1],
                sheet.Cells[row, 4]
            ];
            header.Font.Bold = true;
            header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            row++;

            // ===== ДАННЫЕ ИЗ DataGridView =====
            for (int i = 0; i < dgv.RowCount; i++)
            {
                sheet.Cells[row, 1] = i + 1;

                for (int j = 0; j < dgv.ColumnCount; j++)
                {
                    var cell = dgv.Rows[i].Cells[j];
                    sheet.Cells[row, j + 2] = cell.Value;

                    // цвет ячейки
                    if (cell.Style.BackColor == Color.LightGreen)
                        sheet.Cells[row, j + 2].Interior.Color =
                            ColorTranslator.ToOle(Color.LightGreen);

                    if (cell.Style.BackColor == Color.LightCoral)
                        sheet.Cells[row, j + 2].Interior.Color =
                            ColorTranslator.ToOle(Color.LightCoral);
                }
                row++;
            }

            // автоширина
            sheet.Columns.AutoFit();

            excel.Visible = true;
        }

        // ============ ПРОВЕРКА ВВОДА ============
        bool TryReadInput(out double a, out double b, out double h, out string expectedErrors)
        {
            expectedErrors = "";
            bool hasError = false;
            bool aFlag = false;
            bool bFlag = false;

            a = 0;
            b = 0;
            h = 0;

            // ===== a =====
            if (string.IsNullOrWhiteSpace(rawA))
            {

                expectedErrors = "Число параметров не соответствует ожидаемому и должно быть, как минимум 5!";
                hasError = true;
                return !hasError;

            }
            else if (!double.TryParse(rawA, out a))
            {
                expectedErrors += "Левая граница диапазона не является числом!\r\n";
                hasError = true;
            }

            // ===== b =====
            if (string.IsNullOrWhiteSpace(rawB))
            {

                expectedErrors = "Число параметров не соответствует ожидаемому и должно быть, как минимум 5!";
                hasError = true;
                return !hasError;

            }
            else if (!double.TryParse(rawB, out b))
            {
                expectedErrors += "Правая граница диапазона не является числом!\r\n";
                hasError = true;
            }

            // ===== a < b =====
            if (a >= b && !aFlag && !bFlag)
            {
                expectedErrors += "Левая граница диапазона должна быть < правой границы диапазона!\r\n";
                hasError = true;
            }

            // ===== h =====
            if (string.IsNullOrWhiteSpace(rawH))
            {

                expectedErrors = "Число параметров не соответствует ожидаемому и должно быть, как минимум 5!";
                hasError = true;
                return !hasError;

            }
            else if (!double.TryParse(rawH, out h))
            {
                h = 0;

                if (h < 0.000001 || h > 0.5)
                {
                    expectedErrors += "Шаг интегрирования должен быть в пределах [0.000001;0.5]\r\n";
                    hasError = true;
                }

            }

            return !hasError;
        }

        private void textBoxA_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
