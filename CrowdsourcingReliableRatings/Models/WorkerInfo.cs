using System;
using System.Collections.Generic;
using System.Text;

namespace CrowdsourcingReliableRatings.Models
{
    public class WorkerInfo
    {
        //col 14
        public readonly double FuzzyLogicWeight;

        public double TotalFuzzyLogicWeight { get; set; }

        public double AssignedWeight { get { return this.FuzzyLogicWeight / this.TotalFuzzyLogicWeight; } }

        public List<Evaluation> evaluations { get; set; }

        public WorkerInfo(double fuzzyLogicWeight)
        {
            this.FuzzyLogicWeight = fuzzyLogicWeight;
            this.evaluations = new List<Evaluation>();
        }

    }
}
