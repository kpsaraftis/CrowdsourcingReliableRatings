using System;

namespace CrowdsourcingReliableRatings
{
    public static class ExcelConstants
    {
        public static string NameOfExcelFile = "TaskEvaluations.xlsx";
        public static string NameOfMainSheet = "Evaluations";
        public static string NameOfTask = "Task";
        public static int CountOfTasks = 100;
        public static int CountOfWorkers = 50;
        public static int FakeCountOfWorkers = 3;
        public static string NameOfEachWorkerSheet = "WorkerInfo";
        public static string NameOfEachWorker = "Worker Evaluations ";
        public static string TaskAverage = "TaskAverage";
        public static string DifferenceFromAverage = "Difference From Average";
        public static string VotedHigherCount = "Voted Higher Count";
        public static string VotedLowerCount = "Voted Lower Count";
        public static string OverUnderScore = "Over-Under Score";
        public static string LinearRegressionModel = "Linear Regression Model";
        public static string DebiasedEvaluation = "Debiased Evaluation";
        public static string Slope = "Slope";
        public static string Interval = "Interval";
        public static string AverageDebiasedEvaluation = "Average Debiased Evaluation";
        public static string DistanceScore = "Distance Score";
    }
}
