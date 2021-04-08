using AI.Fuzzy.Library;
using CrowdsourcingReliableRatings.Models;
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
                for (int column = 1; column <= ExcelConstants.CountOfWorkers; column++)
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

                //Calculate fuzzy logic weight
                CalculateFuzzyLogicWeight(package);

                //Calculate worker weight
                CalculateWorkerWeight(package);

                //Create last worksheet with the final evaluations
                FillEvaluationsWorksheet(package);

                //Create ChartAmountOfAnswersPerRating
                ChartAmountOfAnswersPerRating(package);

                //Create ChartAverageRatingPerUser
                ChartAverageRatingPerUser(package);

                //RMSE
                CalculateRMSE(package);

                //Rmse varying population - Use only when altering CountOfWorkers and comment out CalculateRMSE
                //CalculateRMSEVaryingPopulation(package);

                //gather fuzzylogic weight foreach worker
                GatherFuzzyLogicWeightsForeachWorker(package);


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

        private static void GatherFuzzyLogicWeightsForeachWorker(ExcelPackage package)
        {
            int chartWeightPerUserWorkSheetPosition = ExcelConstants.CountOfWorkers + 5;
            package.Workbook.Worksheets.Add(ExcelConstants.ChartWeightPerUser);
            var chartWeightPerUser = package.Workbook.Worksheets[chartWeightPerUserWorkSheetPosition];
            chartWeightPerUser.Cells[1, 1].Value = "Worker";
            chartWeightPerUser.Cells[1, 2].Value = "Fuzzy logic Weight";
            //count ratings
            var userWeightList = new List<double>();
            for (int workerStep = 1; workerStep <= ExcelConstants.CountOfWorkers; workerStep++)
            {
                var worksheet = package.Workbook.Worksheets[workerStep];
                userWeightList.Add((double)worksheet.Cells[2, 14].Value);
            }
            userWeightList.Sort();
            for (int workerStep = 2; workerStep <= ExcelConstants.CountOfWorkers + 1; workerStep++)
            {
                chartWeightPerUser.Cells[workerStep, 1].Value = workerStep - 1;
                chartWeightPerUser.Cells[workerStep, 2].Value = userWeightList[workerStep - 2];
            }
        }

        private static void CalculateRMSEVaryingPopulation(ExcelPackage package)
        {
            int chartRMSEVaryingSampleSizeWorkSheetPosition = ExcelConstants.CountOfWorkers + 4;
            package.Workbook.Worksheets.Add(ExcelConstants.ChartRMSESampleSize);
            var chartRMSEVaryingSampleSize = package.Workbook.Worksheets[chartRMSEVaryingSampleSizeWorkSheetPosition];
            chartRMSEVaryingSampleSize.Cells[1, 1].Value = "User Population";
            chartRMSEVaryingSampleSize.Cells[1, 2].Value = "RMSE Random";
            chartRMSEVaryingSampleSize.Cells[1, 3].Value = "RMSE Average";
            chartRMSEVaryingSampleSize.Cells[1, 4].Value = "RMSE Recommended";
            var workerInfoList = new List<WorkerInfo>();
            int sampleSize = 15;
            for (int workerStep = 1; workerStep <= ExcelConstants.CountOfWorkers; workerStep++)
            {
                var workerSheet = package.Workbook.Worksheets[workerStep];
                var workerInfo = new WorkerInfo((double)workerSheet.Cells[2, 14].Value);
                for (int taskStep = 2; taskStep < ExcelConstants.CountOfTasks + 2; taskStep++)
                {
                    workerInfo.evaluations.Add(
                        new Evaluation(taskStep - 1, (double)workerSheet.Cells[taskStep, 2].Value, (double)workerSheet.Cells[taskStep, 3].Value, (double)workerSheet.Cells[taskStep, 8].Value)
                        );
                }
                workerInfoList.Add(workerInfo);
            }
            for (int population = 30; population < 71; population += 20)
            {
                var topKWorkers = workerInfoList
                    .OrderByDescending(o => o.FuzzyLogicWeight)
                    .Take(sampleSize)
                    .ToList();

                //i want 3 values: Random, Average from users with high score, our approach
                var listOfRandomInt = new List<int>();
                for (int i = 0; i < sampleSize; i++)
                {
                    int randomNumber = 0;
                    do
                    {
                        var ran = new Random();
                        randomNumber = ran.Next(0, ExcelConstants.CountOfWorkers);
                    } while (listOfRandomInt.Any(a => a == randomNumber));
                    listOfRandomInt.Add(randomNumber);
                }

                double randomRMSE = 0, averageRMSE = 0, recommendedRMSE = 0;

                List<WorkerInfo> randomWorkers = workerInfoList.Where((worker, index) => listOfRandomInt.Contains(index)).ToList();
                GetRMSE(topKWorkers, randomWorkers, ref randomRMSE, ref averageRMSE, ref recommendedRMSE);

                listOfRandomInt.Clear();
                int step = (population - 10) / 20 + 1;
                chartRMSEVaryingSampleSize.Cells[step, 1].Value = population;
                chartRMSEVaryingSampleSize.Cells[step, 2].Value = randomRMSE;
                chartRMSEVaryingSampleSize.Cells[step, 3].Value = averageRMSE;
                chartRMSEVaryingSampleSize.Cells[step, 4].Value = recommendedRMSE;
            }
        }

        private static void CalculateRMSE(ExcelPackage package)
        {
            var workerInfoList = new List<WorkerInfo>();
            //var evaluationsSheet = package.Workbook.Worksheets[0];
            int chartRMSESampleSizeWorkSheetPosition = ExcelConstants.CountOfWorkers + 4;
            package.Workbook.Worksheets.Add(ExcelConstants.ChartRMSESampleSize);
            var chartRMSESampleSize = package.Workbook.Worksheets[chartRMSESampleSizeWorkSheetPosition];

            chartRMSESampleSize.Cells[1, 1].Value = "Sample Size";
            chartRMSESampleSize.Cells[1, 2].Value = "RMSE Random";
            chartRMSESampleSize.Cells[1, 3].Value = "RMSE Average";
            chartRMSESampleSize.Cells[1, 4].Value = "RMSE Recommended";
            for (int workerStep = 1; workerStep <= ExcelConstants.CountOfWorkers; workerStep++)
            {
                var workerSheet = package.Workbook.Worksheets[workerStep];
                var workerInfo = new WorkerInfo((double)workerSheet.Cells[2, 14].Value);
                for (int taskStep = 2; taskStep < ExcelConstants.CountOfTasks + 2; taskStep++)
                {
                    workerInfo.evaluations.Add(
                        new Evaluation(taskStep - 1, (double)workerSheet.Cells[taskStep, 2].Value, (double)workerSheet.Cells[taskStep, 3].Value, (double)workerSheet.Cells[taskStep, 8].Value)
                        );
                }
                workerInfoList.Add(workerInfo);
            }
            for (int topK = 5; topK < 71; topK += 5)
            {
                var topKWorkers = workerInfoList
                    .OrderByDescending(o => o.FuzzyLogicWeight)
                    .Take(topK)
                    .ToList();

                //i want 3 values: Random, Average from users with high score, our approach
                var listOfRandomInt = new List<int>();
                for (int i = 0; i < topK; i++)
                {
                    int randomNumber = 0;
                    do
                    {
                        var ran = new Random();
                        randomNumber = ran.Next(0, ExcelConstants.CountOfWorkers);
                    } while (listOfRandomInt.Any(a => a == randomNumber));
                    listOfRandomInt.Add(randomNumber);
                }

                double randomRMSE = 0, averageRMSE = 0, recommendedRMSE = 0;

                List<WorkerInfo> randomWorkers = workerInfoList.Where((worker, index) => listOfRandomInt.Contains(index)).ToList();
                GetRMSE(topKWorkers, randomWorkers, ref randomRMSE, ref averageRMSE, ref recommendedRMSE);

                listOfRandomInt.Clear();
                int step = topK / 5 + 1;
                chartRMSESampleSize.Cells[step, 1].Value = topK;
                chartRMSESampleSize.Cells[step, 2].Value = randomRMSE;
                chartRMSESampleSize.Cells[step, 3].Value = averageRMSE;
                chartRMSESampleSize.Cells[step, 4].Value = recommendedRMSE;
            }
        }

        private static void GetRMSE(List<WorkerInfo> topKWorkers, List<WorkerInfo> randomWorkers, ref double randomRMSE, ref double averageRMSE, ref double recommendedRMSE)
        {
            //sqrt((a-b)/c)
            //b is workerEvaluation.TaskAverage
            double innerPow = 0;
            for (int taskStep = 1; taskStep < ExcelConstants.CountOfTasks; taskStep++)
            {
                //foreach taskstep find average
                double workerEvaluationsSum = 0;
                foreach (var workerInfo in randomWorkers)
                {
                    workerEvaluationsSum += workerInfo.evaluations[taskStep - 1].WorkerEvaluation;
                }
                double workerEvaluationsAverage = workerEvaluationsSum / randomWorkers.Count;
                innerPow += Math.Pow(workerEvaluationsAverage - randomWorkers.First().evaluations[taskStep - 1].TaskAverage, 2);
            }
            randomRMSE = Math.Sqrt(innerPow / ExcelConstants.CountOfTasks);


            innerPow = 0;
            for (int taskStep = 1; taskStep < ExcelConstants.CountOfTasks; taskStep++)
            {
                //foreach taskstep find average
                double workerEvaluationsSum = 0;
                foreach (var workerInfo in topKWorkers)
                {
                    workerEvaluationsSum += workerInfo.evaluations[taskStep - 1].WorkerEvaluation;
                }
                double workerEvaluationsAverage = workerEvaluationsSum / topKWorkers.Count;
                innerPow += Math.Pow(workerEvaluationsAverage - topKWorkers.First().evaluations[taskStep - 1].TaskAverage, 2);
            }
            averageRMSE = Math.Sqrt(innerPow / ExcelConstants.CountOfTasks);

            innerPow = 0;
            //first find weight
            double totalWeight = 0;
            foreach (var workerInfo in topKWorkers)
            {
                totalWeight += workerInfo.FuzzyLogicWeight;
            }
            foreach (var workerInfo in topKWorkers)
            {
                workerInfo.TotalFuzzyLogicWeight = totalWeight;
            }

            for (int taskStep = 1; taskStep < ExcelConstants.CountOfTasks; taskStep++)
            {
                //foreach taskstep find average
                double workerEvaluationsAverage = 0;
                foreach (var workerInfo in topKWorkers)
                {
                    workerEvaluationsAverage += workerInfo.evaluations[taskStep - 1].WorkerDebiasedEvaluation * workerInfo.AssignedWeight;
                }
                //double workerEvaluationsAverage = workerEvaluationsSum / topKWorkers.Count;
                innerPow += Math.Pow(workerEvaluationsAverage - topKWorkers.First().evaluations[taskStep - 1].TaskAverage, 2);
            }
            recommendedRMSE = Math.Sqrt(innerPow / ExcelConstants.CountOfTasks);
        }

        private static void ChartAverageRatingPerUser(ExcelPackage package)
        {
            var evaluationsSheet = package.Workbook.Worksheets[0];
            int chartAverageRatingPerUserWorkSheetPosition = ExcelConstants.CountOfWorkers + 3;
            package.Workbook.Worksheets.Add(ExcelConstants.ChartAverageRatingPerUser);
            var chartAverageRatingPerUser = package.Workbook.Worksheets[chartAverageRatingPerUserWorkSheetPosition];
            chartAverageRatingPerUser.Cells[1, 1].Value = "Average Rating per User";
            chartAverageRatingPerUser.Cells[1, 2].Value = "Average Rating";
            //count ratings
            var averageUserRatingList = new List<double>();
            double averageUserRating = 0;
            for (int workerStep = 2; workerStep <= ExcelConstants.CountOfWorkers + 1; workerStep++)
            {
                for (int taskStep = 2; taskStep < ExcelConstants.CountOfTasks + 2; taskStep++)
                {
                    averageUserRating += (double)evaluationsSheet.Cells[taskStep, workerStep].Value;
                }
                averageUserRatingList.Add(averageUserRating / ExcelConstants.CountOfTasks);
                averageUserRating = 0;
            }
            averageUserRatingList.Sort();
            for (int workerStep = 2; workerStep <= ExcelConstants.CountOfWorkers + 1; workerStep++)
            {
                chartAverageRatingPerUser.Cells[workerStep, 1].Value = workerStep - 1;
                chartAverageRatingPerUser.Cells[workerStep, 2].Value = averageUserRatingList[workerStep - 2];
            }
        }

        private static void ChartAmountOfAnswersPerRating(ExcelPackage package)
        {
            var evaluationsSheet = package.Workbook.Worksheets[0];
            int chartOneWorkSheetPosition = ExcelConstants.CountOfWorkers + 2;
            package.Workbook.Worksheets.Add(ExcelConstants.ChartAmountOfAnswersPerRating);
            var ChartAmountOfAnswersPerRating = package.Workbook.Worksheets[chartOneWorkSheetPosition];
            ChartAmountOfAnswersPerRating.Cells[1, 1].Value = "Ratings";
            ChartAmountOfAnswersPerRating.Cells[1, 2].Value = "Amount of answers per lane change request";
            //count ratings
            var listOfRatings = new List<double>();
            for (int column = 2; column <= ExcelConstants.CountOfWorkers + 1; column++)
            {
                for (int taskStep = 2; taskStep < ExcelConstants.CountOfTasks + 2; taskStep++)
                {
                    listOfRatings.Add((double)evaluationsSheet.Cells[taskStep, column].Value);
                }
            }
            for (int i = 1; i < 6; i++)
            {
                ChartAmountOfAnswersPerRating.Cells[i + 1, 1].Value = i;
                ChartAmountOfAnswersPerRating.Cells[i + 1, 2].Value = listOfRatings.Where(w => w == i).Count();
            }
        }

        private static void FillEvaluationsWorksheet(ExcelPackage package)
        {
            int lastWorkSheet = ExcelConstants.CountOfWorkers + 1;
            package.Workbook.Worksheets.Add(ExcelConstants.NameOfOutputEvaluations);
            var evaluationsOutputWorkSheet = package.Workbook.Worksheets[lastWorkSheet];
            for (int j = 1; j <= ExcelConstants.CountOfTasks; j++)
            {
                evaluationsOutputWorkSheet.Cells[j + 1, 1].Value = ExcelConstants.NameOfTask + " " + j.ToString();
            }
            evaluationsOutputWorkSheet.Cells["B1"].Value = ExcelConstants.TaskAverage;
            evaluationsOutputWorkSheet.Cells["C1"].Value = ExcelConstants.AverageDebiasedEvaluation;
            evaluationsOutputWorkSheet.Cells["D1"].Value = ExcelConstants.DeMaliciousAndDebiasedEvaluation;

            for (int task = 1; task <= ExcelConstants.CountOfTasks; task++)
            {
                evaluationsOutputWorkSheet.Cells[task + 1, 1].Value = ExcelConstants.NameOfTask + " " + task.ToString();
                double finalTaskEvaluation = 0;
                for (int workerSheet = 1; workerSheet <= ExcelConstants.CountOfWorkers; workerSheet++)
                {
                    var workerWorkSheet = package.Workbook.Worksheets[workerSheet];
                    var workerDebiasedEvaluation = (double)workerWorkSheet.Cells[task + 1, 8].Value;
                    var workerWeight = (double)workerWorkSheet.Cells[2, 15].Value;
                    finalTaskEvaluation += workerDebiasedEvaluation * workerWeight;
                }
                evaluationsOutputWorkSheet.Cells[task + 1, 4].Value = finalTaskEvaluation;
            }
            for (int task = 1; task <= ExcelConstants.CountOfTasks; task++)
            {
                var workerOneWorkSheet = package.Workbook.Worksheets[1];
                evaluationsOutputWorkSheet.Cells[task + 1, 1].Value = ExcelConstants.NameOfTask + " " + task.ToString();
                evaluationsOutputWorkSheet.Cells[task + 1, 2].Value = workerOneWorkSheet.Cells[task + 1, 2].Value;
                evaluationsOutputWorkSheet.Cells[task + 1, 3].Value = workerOneWorkSheet.Cells[task + 1, 12].Value;
            }
        }

        private static void CalculateWorkerWeight(ExcelPackage package)
        {
            double totalFuzzyLogicWeight = 0.0;
            for (int workerSheet = 1; workerSheet <= ExcelConstants.CountOfWorkers; workerSheet++)
            {
                var workSheet = package.Workbook.Worksheets[workerSheet];
                if (Double.IsNaN((double)workSheet.Cells[2, 14].Value))
                {
                    var k = 3;
                }
                totalFuzzyLogicWeight += (double)workSheet.Cells[2, 14].Value;
            }
            for (int workerSheet = 1; workerSheet <= ExcelConstants.CountOfWorkers; workerSheet++)
            {
                var workSheet = package.Workbook.Worksheets[workerSheet];
                workSheet.Cells[2, 15].Value = (double)workSheet.Cells[2, 14].Value / totalFuzzyLogicWeight;
            }
        }

        private static void CalculateFuzzyLogicWeight(ExcelPackage package)
        {
            for (int workerSheet = 1; workerSheet <= ExcelConstants.CountOfWorkers; workerSheet++)
            {
                var workSheet = package.Workbook.Worksheets[workerSheet];

                var overUnderScore = workSheet.Cells[2, 11].Value;
                var distanceScore = workSheet.Cells[2, 13].Value;
                var workerWeightFuzzySystem = FuzzyController.GetMamdaniFuzzySystem();
                FuzzyVariable fvDistanceScore = workerWeightFuzzySystem.InputByName("distanceScore");
                FuzzyVariable fvOverUnderScore = workerWeightFuzzySystem.InputByName("overUnderScore");
                FuzzyVariable fvWeight = workerWeightFuzzySystem.OutputByName("workerWeight");

                // Associate input values with input variables
                Dictionary<FuzzyVariable, double> inputValues = new Dictionary<FuzzyVariable, double>();
                inputValues.Add(fvDistanceScore, (double)distanceScore);
                inputValues.Add(fvOverUnderScore, (double)overUnderScore);

                // Calculate result: one output value for each output variable
                Dictionary<FuzzyVariable, double> result = workerWeightFuzzySystem.Calculate(inputValues);

                workSheet.Cells[2, 14].Value = Double.IsNaN((double)result[fvWeight]) ? 0.001 : (object)(double)result[fvWeight];
            }
        }

        private static void CalculateDistanceScore(ExcelPackage package)
        {
            var listOfWorkerWeights = new List<double>();
            for (int worker = 1; worker <= ExcelConstants.CountOfWorkers; worker++)
            {
                var workerSheet = package.Workbook.Worksheets[worker];
                double sumOfInnerPower = 0;
                for (int task = 2; task < ExcelConstants.CountOfTasks + 2; task++)
                {
                    double sumOfDebiasedEval = 0;
                    for (int workerSheetStep = 1; workerSheetStep <= ExcelConstants.CountOfWorkers; workerSheetStep++)
                    {
                        var workSheet = package.Workbook.Worksheets[workerSheetStep];
                        sumOfDebiasedEval += (double)workSheet.Cells[task, 8].Value;
                    }

                    var innerPower = Math.Pow(((double)workerSheet.Cells[task, 8].Value - sumOfDebiasedEval / ExcelConstants.CountOfWorkers), 2);
                    sumOfInnerPower += innerPower;
                }
                //workerSheet.Cells[2, 13].Value = 1 / (1 + (Math.Sqrt(sumOfInnerPower)));

                listOfWorkerWeights.Add(1 / (1 + (Math.Sqrt(sumOfInnerPower))));
            }

            //normalize findings
            var ratio = 1.0 / listOfWorkerWeights.Max();
            for (int workerSheet = 1; workerSheet <= ExcelConstants.CountOfWorkers; workerSheet++)
            {
                var workSheet = package.Workbook.Worksheets[workerSheet];

                workSheet.Cells[2, 13].Value = listOfWorkerWeights[workerSheet - 1] * ratio;  // (double)result[fvWeight];
            }

        }

        private static void CalculateAverageDebiasedEvaluations(ExcelPackage package)
        {
            for (int column = 1; column <= ExcelConstants.CountOfWorkers; column++)
            {
                for (int taskStep = 2; taskStep < ExcelConstants.CountOfTasks + 2; taskStep++)
                {
                    double sumDebiasedEvaluation = 0;
                    for (int debiasedEvaluation = 1; debiasedEvaluation <= ExcelConstants.CountOfWorkers; debiasedEvaluation++)
                    {
                        var workSheet = package.Workbook.Worksheets[debiasedEvaluation];

                        sumDebiasedEvaluation = sumDebiasedEvaluation + (double)workSheet.Cells[taskStep, 8].Value;
                    }
                    for (int debiasedEvaluation = 1; debiasedEvaluation <= ExcelConstants.CountOfWorkers; debiasedEvaluation++)
                    {
                        var workSheet = package.Workbook.Worksheets[debiasedEvaluation];
                        workSheet.Cells[taskStep, 12].Value = sumDebiasedEvaluation / ExcelConstants.CountOfWorkers;
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
            for (int sheet = 1; sheet <= ExcelConstants.CountOfWorkers + 1; sheet++)
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
            workSheet.Cells["O1"].Value = ExcelConstants.WorkerWeight;
            for (int j = 1; j <= ExcelConstants.CountOfTasks; j++)
            {
                workSheet.Cells[j + 1, 1].Value = ExcelConstants.NameOfTask + " " + j.ToString();
            }
        }
    }
}
