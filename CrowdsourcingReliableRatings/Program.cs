using MathNet.Numerics;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CrowdsourcingReliableRatings
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string debugPath = Directory.GetCurrentDirectory();
            var excelDirectory = Directory.GetParent(debugPath).Parent.Parent;
            Console.WriteLine("Load the text file...");
            var excelFilePath = new DirectoryInfo(String.Concat(excelDirectory, "/", ExcelConstants.NameOfExcelFile));
            FileInfo existingFile = new FileInfo(excelFilePath.ToString());
            Console.WriteLine("File Loaded.");

           

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //delete previous worksheets
                DeletePreviousWorksheets(package);

                //Get the first worksheet in the workbook
                ExcelWorksheet evaluations = package.Workbook.Worksheets[0];

                //foreach worker create worksheet and make calculations
                //for (int column = 1; column <= ExcelConstants.CountOfWorkers; column++)
                for (int column = 1; column <= ExcelConstants.FakeCountOfWorkers; column++)
                {

                    package.Workbook.Worksheets.Add(ExcelConstants.NameOfEachWorkerSheet + column.ToString());
                    var workSheet = package.Workbook.Worksheets[column];

                    // create headers
                    CreateHeaders(column, workSheet);

                    //Calculate average
                    CalculateAverage(evaluations, workSheet);

                    //Map worker's evaluations
                    MapWorkerEvaluations(evaluations, column, workSheet);

                    //calculate difference from average
                    DifferenceFromAverage(evaluations, workSheet);

                    //Calculate linear model
                    double slope = 0;
                    double interval = 0;
                    FindLinearModel(workSheet, ref slope, ref interval);

                    //Calculate debiased evaluation
                    DebiasedEvaluation(workSheet, slope, interval);


                    //Number of times voted higher - lower and Score OverUnder 
                    OverUnderScore(workSheet);
                }
                //Calculate average debiased evaluations and store that info foreach individual worker
                CalculateAverageDebiasedEvaluations(package);

                //Calculate Distance Score
                CalculateDistanceScore(package);

                //AutofitCells
                AutofitCells(package);

                //save workbook
                package.Save();
            } // the using statement automatically calls Dispose() which closes the package.

            Console.WriteLine();
            Console.WriteLine("Algorithm excecuted successfully");
            Console.WriteLine();
            Console.ReadLine();
        }

        private static void CalculateDistanceScore(ExcelPackage package)
        {
            for (int workerSheet = 1; workerSheet <= ExcelConstants.FakeCountOfWorkers; workerSheet++)
            {
                var workSheet = package.Workbook.Worksheets[workerSheet];
                double sumOfDebiasedSub = 0;
                for (int task = 2; task < ExcelConstants.CountOfTasks + 2; task++)
                {
                    sumOfDebiasedSub = sumOfDebiasedSub + Math.Pow(((double)workSheet.Cells[task, 8].Value - (double)workSheet.Cells[task, 12].Value), 2);
                }
                workSheet.Cells[2, 13].Value = 1 / (1 + (Math.Sqrt(sumOfDebiasedSub)));
            }
        }

        private static void CalculateAverageDebiasedEvaluations(ExcelPackage package)
        {
            for (int column = 1; column <= ExcelConstants.FakeCountOfWorkers; column++)
            {
                for (int taskStep = 2; taskStep < ExcelConstants.CountOfTasks + 2; taskStep++)
                {
                    double sumDebiasedEvaluation = 0;
                    for (int debiasedEvaluation = 1; debiasedEvaluation <= ExcelConstants.FakeCountOfWorkers; debiasedEvaluation++)
                    {
                        var workSheet = package.Workbook.Worksheets[debiasedEvaluation];

                        sumDebiasedEvaluation = sumDebiasedEvaluation + (double)workSheet.Cells[taskStep, 8].Value;
                    }
                    for (int debiasedEvaluation = 1; debiasedEvaluation <= ExcelConstants.FakeCountOfWorkers; debiasedEvaluation++)
                    {
                        var workSheet = package.Workbook.Worksheets[debiasedEvaluation];
                        workSheet.Cells[taskStep, 12].Value = sumDebiasedEvaluation / ExcelConstants.FakeCountOfWorkers;
                    }
                }
            }
        }

        private static void DeletePreviousWorksheets(ExcelPackage package)
        {
            int workSheetsCount = package.Workbook.Worksheets.Count;
            for (int k = 1; k < workSheetsCount; k++)
            {
                package.Workbook.Worksheets.Delete(1);
            }
        }

        private static void AutofitCells(ExcelPackage package)
        {
            for (int sheet = 1; sheet <= ExcelConstants.FakeCountOfWorkers; sheet++)
            {
                var workSheet = package.Workbook.Worksheets[sheet];
                workSheet.Cells.AutoFitColumns();
            }
        }

        private static void DebiasedEvaluation(ExcelWorksheet workSheet, double slope, double interval)
        {
            for (int row = 2; row < ExcelConstants.CountOfTasks + 2; row++)
            {
                double debiasedEvaluation = 0;
                var workerTaskEvaluation = (double)workSheet.Cells[row, 3].Value;
                debiasedEvaluation = workerTaskEvaluation - (slope * workerTaskEvaluation + interval);
                workSheet.Cells[row, 8].Value = debiasedEvaluation;
            }
        }

        private static void FindLinearModel(ExcelWorksheet workSheet, ref double slope, ref double interval)
        {
            var xData = new List<double>();
            var yData = new List<double>();
            for (int row = 2; row < ExcelConstants.CountOfTasks + 2; row++)
            {
                xData.Add((double)workSheet.Cells[row, 3].Value);
                yData.Add((double)workSheet.Cells[row, 4].Value);
            }

            Tuple<double, double> linearModel = Fit.Line(xData.ToArray(), yData.ToArray());
            workSheet.Cells["E2"].Value = $"y={linearModel.Item2}x+{linearModel.Item1}";
            slope = linearModel.Item2;
            interval = linearModel.Item1;
            workSheet.Cells["F2"].Value = slope;
            workSheet.Cells["G2"].Value = interval;
        }

        private static void OverUnderScore(ExcelWorksheet workSheet)
        {
            string lastCell = (ExcelConstants.CountOfTasks + 1).ToString();
            workSheet.Cells["I2"].Value = workSheet.Cells["d2:d" + lastCell].Where(w => (double)w.Value >= 0).Count();
            workSheet.Cells["J2"].Value = workSheet.Cells["d2:d" + lastCell].Where(w => (double)w.Value < 0).Count();
            var diff = Math.Abs((int)workSheet.Cells["I2"].Value - (int)workSheet.Cells["J2"].Value);
            var score = (double)diff / ExcelConstants.CountOfTasks;
            workSheet.Cells["K2"].Value = score;
        }

        private static void DifferenceFromAverage(ExcelWorksheet evaluations, ExcelWorksheet workSheet)
        {
            for (int row = 2; row < ExcelConstants.CountOfTasks + 2; row++)
            {
                double diffFromAverage = 0;
                for (int col = 2; col < ExcelConstants.CountOfWorkers + 2; col++)
                {
                    diffFromAverage = diffFromAverage + (double)evaluations.Cells[row, col].Value;
                }
                workSheet.Cells[row, 4].Value = (double)workSheet.Cells[row, 3].Value - (double)workSheet.Cells[row, 2].Value;
            }
        }

        private static void MapWorkerEvaluations(ExcelWorksheet evaluations, int column, ExcelWorksheet workSheet)
        {
            for (int row = 2; row < ExcelConstants.CountOfTasks + 2; row++)
            {
                workSheet.Cells[row, 3].Value = evaluations.Cells[row, column + 1].Value;
            }
        }

        private static void CalculateAverage(ExcelWorksheet evaluations, ExcelWorksheet workSheet)
        {
            for (int row = 2; row < ExcelConstants.CountOfTasks + 2; row++)
            {
                double average = 0;
                for (int col = 2; col < ExcelConstants.CountOfWorkers + 2; col++)
                {
                    average = average + (double)evaluations.Cells[row, col].Value;
                }
                workSheet.Cells[row, 2].Value = average / ExcelConstants.CountOfWorkers;
            }
        }

        private static void CreateHeaders(int i, ExcelWorksheet workSheet)
        {
            workSheet.Cells["B1"].Value = ExcelConstants.TaskAverage;
            workSheet.Cells["C1"].Value = ExcelConstants.NameOfEachWorker + i.ToString();
            workSheet.Cells["D1"].Value = ExcelConstants.DifferenceFromAverage;
            workSheet.Cells["E1"].Value = ExcelConstants.LinearRegressionModel;
            workSheet.Cells["F1"].Value = ExcelConstants.Slope;
            workSheet.Cells["G1"].Value = ExcelConstants.Interval;
            workSheet.Cells["H1"].Value = ExcelConstants.DebiasedEvaluation;
            workSheet.Cells["I1"].Value = ExcelConstants.VotedHigherCount;
            workSheet.Cells["J1"].Value = ExcelConstants.VotedLowerCount;
            workSheet.Cells["K1"].Value = ExcelConstants.OverUnderScore;
            workSheet.Cells["L1"].Value = ExcelConstants.AverageDebiasedEvaluation;
            workSheet.Cells["M1"].Value = ExcelConstants.DistanceScore;
            workSheet.Cells["N1"].Value = ExcelConstants.FuzzyLogicWeight;
            for (int j = 1; j <= ExcelConstants.CountOfTasks; j++)
            {
                workSheet.Cells[j + 1, 1].Value = ExcelConstants.NameOfTask + " " + j.ToString();
            }
        }
    }
}
