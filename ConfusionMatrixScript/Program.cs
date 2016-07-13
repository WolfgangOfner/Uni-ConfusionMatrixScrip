using RDotNet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Threading.Tasks;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;


namespace ConfusionMatrixScript
{
    class Program
    {
        public static REngine engine = REngine.GetInstance();

        private static int countRows = 0;

        static void Main(string[] args)
        {
            // path to parameter file
            string parameterFile = args[0];
            string[] readText = File.ReadAllLines(parameterFile);

            // confusion matrix array starts at index 8
            int index = 8;

            // path to .R script
            string scriptPath = readText[0];
            // path to .xls file (after calculation)
            string downloadPath = readText[1];
            // size of confusion matrix array
            int dataAmount = Convert.ToInt32(readText[2]);

            bool[] checkBoxes = new bool[] { Convert.ToBoolean(readText[3]), Convert.ToBoolean(readText[4]), Convert.ToBoolean(readText[5]), Convert.ToBoolean(readText[6]), Convert.ToBoolean(readText[7]) };

            // confusion matrix
            var array = new string[Convert.ToInt32(Math.Sqrt(dataAmount)), Convert.ToInt32(Math.Sqrt(dataAmount))];

            for (int i = 0; i < Math.Sqrt(dataAmount); i++)
            {
                for (int j = 0; j < Math.Sqrt(dataAmount); j++)
                {
                    array[i, j] = readText[index++];
                }
            }

            FileStream stream = File.OpenRead(scriptPath);

            engine.Evaluate(stream);
            stream.Dispose();
            stream.Close();

            SetCountRows(checkBoxes);

            var doubleArray = ConvertStringArrayToDoubleArray(array);
            //    List<IEnumerable> resultCollection = new List<IEnumerable>();

            NumericMatrix matrix = ConvertArrayToMatrix(doubleArray);
            var resultArray = FillResultArray(array);

            //  resultCollection.Add(new int[] { resultArray.GetLength(0), resultArray.GetLength(1) });                              ///Adding the size of the array to the collection to know the iteration number of the for-loops in javascript

            var userAccuracy = CalculateUserAccuracy(matrix);
            resultArray = InsertUserAccuracyInResultArray(userAccuracy, resultArray);

            var producerAccuracy = CalculateProducerAccuracy(matrix);
            resultArray = InsertProducerAccuracyInResultArray(producerAccuracy, resultArray);

            var overallAccuracy = CalculateOverallAccuracy(matrix);
            resultArray = InsertOverallAccuracyInResultArray(overallAccuracy, resultArray);

            if (checkBoxes[0])
            {
                var portmanteauAccuracy = CalculatePortmanteauAccuracy(matrix);
                resultArray = InsertPortmanteauAccuracyInResultArray(portmanteauAccuracy, resultArray);
            }

            if (checkBoxes[1])
            {
                var kappa = CalculateKappa(matrix);
                resultArray = InsertKappaInResultArray(kappa, resultArray);
            }

            if (checkBoxes[2])
            {
                var ami = CalculateAMI(matrix);
                resultArray = InsertAMIInResultArray(ami, resultArray);
            }

            if (checkBoxes[3])
            {
                var quantityDisaggreement = CalculateQuantityDisagreement(matrix);
                resultArray = InsertQuantityDisagreementInResultArray(quantityDisaggreement, resultArray);
            }

            if (checkBoxes[4])
            {
                var allocationDisagreement = CalculateAllocationDisagreement(matrix);
                resultArray = InsertAllocationDisagreementInResultArray(allocationDisagreement, resultArray);
            }
         
            CreateExcel(resultArray, downloadPath);

            Process p = Process.GetCurrentProcess();

            Environment.ExitCode = 0;
            p.CloseMainWindow();
            p.Close();

        }

        private static void CreateExcel(string[,] array, string path)
        {
            FileInfo newFile = new FileInfo(path);

            using (var package = new ExcelPackage(newFile))
            { 
                ExcelWorksheet excelSheet = package.Workbook.Worksheets.Add("ConfusionMatrix");
          
                for (int i = 0; i < array.GetLength(0); i++)
                {
                    for (int j = 0; j < array.GetLength(1); j++)
                    {
                        excelSheet.Cells[i + 1, j + 1].Value = array[i, j];
                        excelSheet.Column(j + 1).AutoFit();

                        if (j == array.GetLength(1) - 1 && i != 0 && i < array.GetLength(1) - 3)
                        {
                            excelSheet.Cells[i + 1, j + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            excelSheet.Cells[i + 1, j + 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                            excelSheet.Cells[i + 1, j + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.LightGray);
                        }

                        if (i >= array.GetLength(1) - 1 && j != 0 && excelSheet.Cells[i + 1, j + 1].Value != null)
                        {
                            excelSheet.Cells[i + 1, j + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            excelSheet.Cells[i + 1, j + 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                            excelSheet.Cells[i + 1, j + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.LightGray);
                        }

                        if (i == array.GetLength(1) - 2 && j == array.GetLength(1) - 1)
                        {
                            excelSheet.Cells[i + 1, j + 1].Style.Font.Bold = true;
                        }
                    }
                }

                excelSheet.Row(1).Style.Font.Bold = true;
                excelSheet.Column(1).Style.Font.Bold = true;

                package.Save();
            }
        }

        /// <summary>
        /// Counts which values are set true, to know the row amount of the result table.
        /// </summary>
        /// <param name="checkBoxes"></param>
        private static void SetCountRows(bool[] checkBoxes)
        {
            countRows = 0;

            foreach (var boolean in checkBoxes)
            {
                if (boolean)
                {
                    countRows++;
                }
            }
        }


        /// <summary>
        /// Creating result array.
        /// </summary>
        /// <param name="array">The array of the excel table on the website</param>
        /// <returns></returns>
        private static string[,] FillResultArray(string[,] array)
        {
            string[,] resultArray;

            if (countRows == 0)
            {
                resultArray = new string[array.GetLength(1) + countRows + 3, array.GetLength(1) + 3];
            }
            else
            {
                resultArray = new string[array.GetLength(1) + countRows + 4, array.GetLength(1) + 3];
            }

            double rowTotalAmount;

            resultArray[array.GetLength(1), array.GetLength(1)] = "0";
            resultArray[array.GetLength(1), 0] = "Col total";
            resultArray[0, array.GetLength(1)] = "Row total";


            for (int i = 0; i < array.GetLength(1); i++)
            {
                rowTotalAmount = 0;

                for (int j = 0; j < array.GetLength(1); j++)
                {
                    if (i != 0 && j != 0)
                    {
                        rowTotalAmount += Convert.ToDouble(array[i, j]);
                        resultArray[array.GetLength(1), j] = (Convert.ToDouble(resultArray[array.GetLength(1), j]) + Convert.ToDouble(array[i, j])).ToString();
                    }

                    if (j == array.GetLength(1) - 1 && i != 0)
                    {
                        resultArray[i, j] = array[i, j];
                        resultArray[i, array.GetLength(0)] = rowTotalAmount.ToString();
                        resultArray[array.GetLength(1), array.GetLength(1)] = (Convert.ToDouble(resultArray[array.GetLength(1), array.GetLength(1)]) + rowTotalAmount).ToString();
                    }
                    else
                    {
                        resultArray[i, j] = array[i, j];
                    }
                }

            }

            return resultArray;
        }

        private static string[,] InsertUserAccuracyInResultArray(NumericVector userAccuracy, string[,] resultArray)
        {
            int n = resultArray.GetLength(1) - 1;

            resultArray[n, 0] = "User's accuracy";

            for (int m = 1; m <= userAccuracy.Length; m++)
            {
                resultArray[n, m] = Math.Round((userAccuracy[m - 1] * 100), 2).ToString() + "%";
            }

            return resultArray;
        }

        private static string[,] InsertProducerAccuracyInResultArray(NumericVector producerAccuracy, string[,] resultArray)
        {
            int n = resultArray.GetLength(1) - 1;

            resultArray[0, n] = "Producer's accuracy";

            for (int m = 1; m <= producerAccuracy.Length; m++)
            {
                resultArray[m, n] = Math.Round((producerAccuracy[m - 1] * 100), 2).ToString() + "%";
            }

            return resultArray;
        }

        private static string[,] InsertOverallAccuracyInResultArray(NumericVector overallAccuracy, string[,] resultArray)
        {
            resultArray[resultArray.GetLength(1) - 2, resultArray.GetLength(1) - 1] = "Overall accuracy";

            resultArray[resultArray.GetLength(1) - 1, resultArray.GetLength(1) - 1] = Math.Round((overallAccuracy[0] * 100), 2).ToString() + "%";

            return resultArray;
        }

        private static string[,] InsertPortmanteauAccuracyInResultArray(NumericVector portmanteauAccuracy, string[,] resultArray)
        {
            int n = resultArray.GetLength(0) - countRows;
            countRows--;

            resultArray[n, 0] = "Portmanteau accuracy";

            for (int m = 1; m <= portmanteauAccuracy.Length; m++)
            {
                resultArray[n, m] = Math.Round((portmanteauAccuracy[m - 1] * 100), 2).ToString() + "%";
            }

            return resultArray;
        }

        private static string[,] InsertKappaInResultArray(NumericVector kappa, string[,] resultArray)
        {
            int n = resultArray.GetLength(0) - countRows;
            countRows--;

            resultArray[n, 0] = "Kappa";
            resultArray[n, 1] = Math.Round(kappa[0], 3).ToString();

            return resultArray;
        }

        private static string[,] InsertAMIInResultArray(NumericVector ami, string[,] resultArray)
        {
            int n = resultArray.GetLength(0) - countRows;
            countRows--;

            resultArray[n, 0] = "AMI";
            resultArray[n, 1] = Math.Round(ami[0], 3).ToString();

            return resultArray;
        }

        private static string[,] InsertQuantityDisagreementInResultArray(NumericVector quantityDisaggreement, string[,] resultArray)
        {
            int n = resultArray.GetLength(0) - countRows;
            countRows--;

            resultArray[n, 0] = "Quantity disagreement";
            resultArray[n, 1] = Math.Round(quantityDisaggreement[0], 3).ToString();

            return resultArray;
        }

        private static string[,] InsertAllocationDisagreementInResultArray(NumericVector allocationDisagreement, string[,] resultArray)
        {
            int n = resultArray.GetLength(0) - countRows;

            resultArray[n, 0] = "Allocation disagreement";
            resultArray[n, 1] = Math.Round(allocationDisagreement[0], 3).ToString();

            return resultArray;
        }

        private static NumericVector CalculateUserAccuracy(NumericMatrix matrix)
        {

            return engine.Evaluate("user(matrix)").AsNumeric();
        }

        private static NumericVector CalculateProducerAccuracy(NumericMatrix matrix)
        {
            return engine.Evaluate("producer(matrix)").AsNumeric();
        }

        private static NumericVector CalculateOverallAccuracy(NumericMatrix matrix)
        {
            return engine.Evaluate("overacc(matrix)").AsNumeric();
        }

        private static NumericVector CalculatePortmanteauAccuracy(NumericMatrix matrix)
        {
            return engine.Evaluate("port(matrix)").AsNumeric();
        }

        private static NumericVector CalculateKappa(NumericMatrix matrix)
        {
            return engine.Evaluate("kappa(matrix)").AsNumeric();
        }

        private static NumericVector CalculateAMI(NumericMatrix matrix)
        {
            return engine.Evaluate("ami(matrix)").AsNumeric();
        }

        private static NumericVector CalculateQuantityDisagreement(NumericMatrix matrix)
        {
            return engine.Evaluate("quant.dis(matrix)").AsNumeric();
        }

        private static NumericVector CalculateAllocationDisagreement(NumericMatrix matrix)
        {
            return engine.Evaluate("alloc.dis(matrix)").AsNumeric();
        }

        private static double[,] ConvertStringArrayToDoubleArray(string[,] array)
        {
            double[,] doubleArray = new double[array.GetLength(0) - 1, array.GetLength(0) - 1];
            for (int i = 1; i <= doubleArray.GetLength(0); i++)
            {
                for (int j = 1; j <= doubleArray.GetLength(1); j++)
                {
                    doubleArray[i - 1, j - 1] = Convert.ToDouble(array[i, j]);
                }
            }

            return doubleArray;
        }

        private static NumericMatrix ConvertArrayToMatrix(double[,] array)
        {
            var dimension = array.GetLength(0);

            var dimensions = engine.CreateNumericVector(new double[] { dimension });
            dimensions = engine.CreateNumericVector(new double[] { dimension });
            engine.SetSymbol("dimension", dimensions);

            NumericVector arrayVector = engine.CreateNumericVector(dimension * dimension);

            int i = 0;
            foreach (var item in array)
            {
                arrayVector[i] = item;
                i++;
            }
            engine.SetSymbol("array", arrayVector);

            return engine.Evaluate("matrix <- matrix(array,nrow=dimension,ncol=dimension)").AsNumericMatrix();
        }


    }
}
