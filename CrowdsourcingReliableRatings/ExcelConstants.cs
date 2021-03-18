using System;

namespace CrowdsourcingReliableRatings
{
    public static class ExcelConstants
    {
        public static string NameOfExcelFile = "ReliableTaskEvaluations.xlsx";
        public static string NameOfMainSheet = "Evaluations";
        public static string NameOfTask = "Task";
        public static int CountOfTasks = 132;
        public static int CountOfWorkers = 70;
        public static string NameOfEachWorkerSheet = "WorkerInfo";
        public static string NameOfEachWorker = "Worker Evaluations ";
        public static string NameOfOutputEvaluations = "Evaluations Output";
        public static string TaskAverage = "TaskAverage";
        public static string DifferenceFromAverage = "Difference From Average";
        public static string VotedHigherCount = "Voted Higher Count";
        public static string VotedLowerCount = "Voted Lower Count";
        public static string OverUnderScore = "Over-Under Score";
        public static string LinearRegressionModel = "Linear Regression Model";
        public static string DebiasedEvaluation = "Debiased Evaluation";
        public static string DeMaliciousAndDebiasedEvaluation = "De-Malicious and Debiased Evaluation";
        public static string Slope = "Slope";
        public static string Interval = "Interval";
        public static string AverageDebiasedEvaluation = "Average Debiased Evaluation";
        public static string DistanceScore = "Distance Score";
        public static string FuzzyLogicWeight = "Fuzzy Logic Weight";
        public static string WorkerWeight = "Worker Weight";
    }
}
