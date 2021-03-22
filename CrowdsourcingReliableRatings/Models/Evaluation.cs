using System;
using System.Collections.Generic;
using System.Text;

namespace CrowdsourcingReliableRatings.Models
{
    public class Evaluation
    {
        public int TaskID { get; set; }
        //col 2
        public double TaskAverage { get; set; }
        //col 3
        public double WorkerEvaluation { get; set; }
        //col 8
        public double WorkerDebiasedEvaluation { get; set; }
        
        public Evaluation(int taskID, double taskAverage, double workerEvaluation, double workerDebiasedEvaluation)
        {
            TaskID = taskID;
            WorkerEvaluation = workerEvaluation;
            TaskAverage = taskAverage;
            WorkerDebiasedEvaluation = workerDebiasedEvaluation;
        }
    }
}
