using AI.Fuzzy.Library;
using System;
using System.Collections.Generic;
using System.Text;

namespace CrowdsourcingReliableRatings
{
    internal static class FuzzyController
    {
        internal static MamdaniFuzzySystem GetMamdaniFuzzySystem()
        {
            MamdaniFuzzySystem workerWeightFuzzySystem = new MamdaniFuzzySystem();

            // Create input variables for the system
            FuzzyVariable distanceScore = new FuzzyVariable("distanceScore", 0.0, 1.0);
            distanceScore.Terms.Add(new FuzzyTerm("low", new TriangularMembershipFunction(0.0, 0.5, 0.7)));
            distanceScore.Terms.Add(new FuzzyTerm("average", new TriangularMembershipFunction(0.7, 0.8, 0.9)));
            distanceScore.Terms.Add(new FuzzyTerm("excellent", new TriangularMembershipFunction(0.8, 0.9, 1.1)));
            workerWeightFuzzySystem.Input.Add(distanceScore);

            FuzzyVariable overUnderScore = new FuzzyVariable("overUnderScore", 0.0, 1.0);
            overUnderScore.Terms.Add(new FuzzyTerm("low", new TriangularMembershipFunction(0.0, 0.5, 0.7)));
            overUnderScore.Terms.Add(new FuzzyTerm("average", new TriangularMembershipFunction(0.4, 0.7, 0.9)));
            overUnderScore.Terms.Add(new FuzzyTerm("excellent", new TriangularMembershipFunction(0.8, 0.9, 1.1)));
            workerWeightFuzzySystem.Input.Add(overUnderScore);

            // Create output variables for the system
            FuzzyVariable fvWeight = new FuzzyVariable("workerWeight", 0.0, 1.0);
            fvWeight.Terms.Add(new FuzzyTerm("low", new TriangularMembershipFunction(0.0, 0.15, 0.5)));
            fvWeight.Terms.Add(new FuzzyTerm("belowAverage", new TriangularMembershipFunction(0.1, 0.2, 0.55)));
            fvWeight.Terms.Add(new FuzzyTerm("average", new TriangularMembershipFunction(0.55, 0.6, 0.85)));
            fvWeight.Terms.Add(new FuzzyTerm("aboveAverage", new TriangularMembershipFunction(0.75, 0.85, 0.95)));
            fvWeight.Terms.Add(new FuzzyTerm("excellent", new TriangularMembershipFunction(0.9, 0.95, 1.0)));
            workerWeightFuzzySystem.Output.Add(fvWeight);

            // Create fuzzy rules
            MamdaniFuzzyRule rule1 = workerWeightFuzzySystem.ParseRule("if (distanceScore is low ) and (overUnderScore is low) then workerWeight is low");
            MamdaniFuzzyRule rule2 = workerWeightFuzzySystem.ParseRule("if (distanceScore is low ) and (overUnderScore is average) then workerWeight is belowAverage");
            MamdaniFuzzyRule rule3 = workerWeightFuzzySystem.ParseRule("if (distanceScore is low ) and (overUnderScore is excellent) then workerWeight is average");
            MamdaniFuzzyRule rule4 = workerWeightFuzzySystem.ParseRule("if (distanceScore is average ) and (overUnderScore is low) then workerWeight is belowAverage");
            MamdaniFuzzyRule rule5 = workerWeightFuzzySystem.ParseRule("if (distanceScore is average ) and (overUnderScore is average) then workerWeight is average");
            MamdaniFuzzyRule rule6 = workerWeightFuzzySystem.ParseRule("if (distanceScore is average ) and (overUnderScore is excellent) then workerWeight is aboveAverage");
            MamdaniFuzzyRule rule7 = workerWeightFuzzySystem.ParseRule("if (distanceScore is excellent ) and (overUnderScore is low) then workerWeight is average");
            MamdaniFuzzyRule rule8 = workerWeightFuzzySystem.ParseRule("if (distanceScore is excellent ) and (overUnderScore is average) then workerWeight is aboveAverage");
            MamdaniFuzzyRule rule9 = workerWeightFuzzySystem.ParseRule("if (distanceScore is excellent ) and (overUnderScore is excellent) then workerWeight is excellent");

            //Add fuzzy rules
            workerWeightFuzzySystem.Rules.Add(rule1);
            workerWeightFuzzySystem.Rules.Add(rule2);
            workerWeightFuzzySystem.Rules.Add(rule3);
            workerWeightFuzzySystem.Rules.Add(rule4);
            workerWeightFuzzySystem.Rules.Add(rule5);
            workerWeightFuzzySystem.Rules.Add(rule6);
            workerWeightFuzzySystem.Rules.Add(rule7);
            workerWeightFuzzySystem.Rules.Add(rule8);
            workerWeightFuzzySystem.Rules.Add(rule9);

            return workerWeightFuzzySystem;
        }
    }
}
