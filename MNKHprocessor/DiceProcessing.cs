using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Text;
using System.Diagnostics;
using System.Numerics;

namespace MNKHprocessor
{
    class DiceProcessing
    {
        static bool COLORING = true;
        static int crit_max = 2;

        static string plan_in = "plan.txt";
        static string dice_in = "dice.txt";
        //https://en.wikipedia.org/wiki/Longest_common_subsequence_problem
        static int LcsLength(string a, string b) {
            int[,] C = new int[a.Length + 1, b.Length + 1]; // (a, b).Length + 1
            for (int i = 0; i < a.Length; i++)
                C[i, 0] = 0;
            for (int j = 0; j < b.Length; j++)
                C[0, j] = 0;
            for (int i = 1; i <= a.Length; i++)
                for (int j = 1; j <= b.Length; j++) {
                    if (a[i - 1] == b[j - 1])//i-1,j-1
                        C[i, j] = C[i - 1, j - 1] + 1;
                    else
                        C[i, j] = Math.Max(C[i, j - 1], C[i - 1, j]);
                }
            return C[a.Length, b.Length];
        }
        static ActionData FindAction(string action_name, List<SectionData> sections) {
            string action_name_lower = action_name.ToLower();
            ActionData action = null;
            int best_closeness = 0;
            bool have_conflict = true;
            foreach (SectionData section in sections) {
                foreach (ActionData evalAction in section.actions) {
                    string name2 = evalAction.name;
                    int closeness = LcsLength(name2.ToLower(), action_name_lower);
                    if (closeness * 2 > Math.Min(name2.Length, action_name.Length)) {
                        closeness = closeness * 10 - Math.Abs(name2.Length - action_name.Length);
                        if (closeness == best_closeness) {
                            have_conflict = true;
                        } else if (closeness > best_closeness) {
                            have_conflict = false;
                            action = evalAction;
                            best_closeness = closeness;
                        }
                    }
                }
            }
            Debug.Assert(!have_conflict);
            return action;
        }
        static string Colorize(string str, string col) {
            if (!COLORING || col == "") return str;
            return string.Concat("[COLOR=rgb(", col, ")]", str, "[/COLOR]");
        }
        public static void ProcessDice(TurnData turn) {
            //Process dice rolls into queue
            Queue<int> dice_raw = new Queue<int>();
            using (StreamReader sr = File.OpenText(dice_in)) {
                string s;
                double dice_total = 0;
                Regex isDice = new Regex(@"\G(\d+)\s?", RegexOptions.IgnoreCase);
                while ((s = sr.ReadLine()) != null) {
                    foreach (Match m in isDice.Matches(s)) {
                        int dice_parsed = Int32.Parse(m.Groups[1].Value);
                        dice_total += dice_parsed;
                        dice_raw.Enqueue(dice_parsed);
                    }
                }
                Console.WriteLine("Dice average: " + dice_total / dice_raw.Count);
                Console.WriteLine();
            }
            //read plan and format the dice
            List<string> planList = new List<string>();
            string last_section = "";
            using (StreamReader sr = File.OpenText(plan_in)) {
                Regex isFocusAction = new Regex(@"Dedicate Focus[^(0-9]*\((.*?)\)", RegexOptions.IgnoreCase);
                //Regex isResourceCounter = new Regex(@"-*\[X?\]\d+/\d+ Resources", RegexOptions.IgnoreCase);
                string s;
                while ((s = sr.ReadLine()) != null) {
                    if (s.StartsWith("--[")) {
                        Match m_focus = isFocusAction.Match(s);
                        if (m_focus.Success) {
                            string target_name = m_focus.Groups[1].Value;
                            ActionData action = FindAction(target_name, turn.sections);
                            if (action.section_name.Contains("Service")) {
                                action.bonus += 15;
                            } else {
                                action.bonus += 5;
                            }
                            continue;
                        }
                        planList.Add(s);
                    }
                }
            }
            Regex isDiceAction = new Regex(@"--\[X?\]\s*(.*?)(?:,|:|\s)+(\d) Di", RegexOptions.IgnoreCase);
            Regex isAnyAction = new Regex(@"--\[X?\]\s*(.*)", RegexOptions.IgnoreCase);
            //Regex skip = new Regex(@"\d+/\d+ (di|Resource)", RegexOptions.IgnoreCase);
            int total_cost = 0;
            Dictionary<string, int> indicators = new();
            foreach (string s in planList) {
                int last_position = -1;
                Match m = isDiceAction.Match(s);
                bool anyMatch = false;
                if (!m.Success) {
                    m = isAnyAction.Match(s);
                    anyMatch = true;
                }
                if (m.Success) {
                    string action_plan_name = m.Groups[1].Value;
                    string dice_formatted = "";
                    string dice_final = "";
                    string color = "";
                    string crits = "";
                    ActionData action = FindAction(action_plan_name, turn.sections);
                    string action_name = action.name;//Use cannonical name instead of the plan's
                    if (last_section != action.section_name) {
                        last_section = action.section_name;
                        Console.WriteLine();
                    }
                    if (action.position <= last_position) {
                        Console.WriteLine(">>WARNING: ACTION ORDER MISMATCH!");
                        Console.WriteLine(">>" + action_name + " at " + action.position + " but last position was " + last_position);
                    }
                    last_position = action.position;
                    int dice_count = 1;
                    if (!anyMatch) {
                        dice_count = Int32.Parse(m.Groups[2].Value);
                    }
                    total_cost += dice_count * action.rpd;
                    switch (action.type) {
                        case ACTION_TYPES.NORMAL:
                        case ACTION_TYPES.DC:
                            StringBuilder str = new StringBuilder();
                            int total_sum = action.prog_curr + action.bonus * dice_count;
                            double prob = CalcChance(dice_count, action.prog_max - total_sum);
                            double prob2 = CalcChance(dice_count, action.prog_max - total_sum - 10);
                            double prob3 = CalcChance(dice_count, action.prog_max - total_sum - 15);
                            if (action.type == ACTION_TYPES.NORMAL) {
                                str.Append(action.prog_curr).Append('+');
                            }
                            str.Append(string.Concat("(", dice_count, "*", action.bonus, ")"));
                            for (int i = 1; i <= dice_count; i++) {
                                int dice = dice_raw.Dequeue();
                                total_sum += dice;
                                str.Append('+').Append(dice.ToString());
                                if (dice == 1) { crits += Colorize(" (Nat 1)", "255,0,0"); }
                                if (dice == 2) { crits += Colorize(" (Nat 2)", "255,0,0"); }
                                if (dice == 100) { crits += Colorize(" (Nat 100)", "0,255,0"); }
                            }
                            if (total_sum < action.prog_max && prob2 >= 0.1) {
                                color = "255,0,0";
                            }
                            if (total_sum + 15 >= action.prog_max && total_sum < action.prog_max) {
                                crits += Colorize(" (omake?)", "50,200,50");
                            }
                            if (total_sum >= action.prog_max) {
                                color = "0,255,0";
                            }

                            if (total_sum >= action.prog_max * 4) {
                                crits += Colorize(" (4x target)", "150,100,100");
                            } else if (total_sum >= action.prog_max * 2) {
                                crits += Colorize(" (2x target)", "200,50,50");
                            }
                            if (prob3 > 0) {
                                crits += Colorize(string.Concat("(",
                                    prob.ToString("P"), "/",
                                    prob3.ToString("P"), ")"), "150,150,150");
                            }

                            str.Append('=');
                            if (action.type == ACTION_TYPES.DC) {
                                dice_final = string.Concat(total_sum, "/DC ", action.prog_max);
                            } else {
                                dice_final = string.Concat(total_sum, "/", action.prog_max);
                            }

                            dice_formatted = str.ToString();
                            if (total_sum + 15 >= action.prog_max) {
                                foreach (KeyValuePair<string, int> ind in action.indicators) {
                                    indicators[ind.Key] = indicators.GetValueOrDefault(ind.Key, 0) + ind.Value;
                                }
                            }
                            break;
                        case ACTION_TYPES.REFORM:
                            int tmp_sum = action.bonus * dice_count;
                            dice_formatted = string.Concat("(", dice_count, "*", action.bonus, ")");
                            for (int i = 1; i <= dice_count; i++) {
                                int dice = dice_raw.Dequeue();
                                tmp_sum += dice;
                                dice_formatted += "+" + dice.ToString();
                                if (dice <= crit_max) { crits += Colorize(" (Nat " + dice + ")", "255,0,0"); }
                                if (dice == 100) { crits += Colorize(" (Nat 100)", "0,255,0"); }
                            }
                            dice_formatted += "=";
                            dice_final = tmp_sum.ToString();
                            break;
                    }
                    Console.Write(Colorize(action_name, color));
                    Console.Write(" ");
                    Console.Write(dice_formatted);
                    Console.Write(Colorize(dice_final, color));
                    Console.Write(crits);
                    Console.WriteLine();
                }
            }
            //check that the plan used all dice
            Debug.Assert(dice_raw.Count == 0);
            Console.WriteLine();
            Console.WriteLine("Total cost: " + total_cost + " R");
            Console.WriteLine("Indicator adjustments from actions (assuming no additional stages or other changes)");
            foreach (string ind_name in TurnProcessor.indicator_names) {
                Console.WriteLine(ind_name + ": " + indicators.GetValueOrDefault(ind_name, 0));
            }
        }

        static double CalcChance(int dice, int target) { //returns the probability that Nd100 dice is at least target
            BigInteger POSSIBILITIES = 0;
            if (target <= dice) return 1;
            if (target > dice * 100) return 0;
            for (int index = dice; index < target; index++) {
                for (int k = 0; k <= (index - dice) / 100; k++) {
                    BigInteger temp = nCr(dice, k);
                    temp *= nCr(index - 100 * k - 1, dice - 1);
                    if ((k / 2) * 2 == k) {
                        POSSIBILITIES += temp;
                    } else {
                        POSSIBILITIES -= temp;
                    }
                }
            }
            return 1 - (Double)((POSSIBILITIES * 1_000_000) / BigInteger.Pow(100, dice)) / 1_000_000;
        }
        //Based on https://stackoverflow.com/questions/26311984/permutation-and-combination-in-c-sharp
        static BigInteger nCr(int n, int r) {
            // naive: return Factorial(n) / (Factorial(r) * Factorial(n - r));
            return nPr(n, r) / Factorial(r);
        }

        static BigInteger nPr(int n, int r) {
            // naive: return Factorial(n) / Factorial(n - r);
            return FactorialDivision(n, n - r);
        }

        private static BigInteger FactorialDivision(int topFactorial, int divisorFactorial) {
            BigInteger result = 1;
            for (int i = topFactorial; i > divisorFactorial; i--)
                result *= i;
            return result;
        }

        private static BigInteger Factorial(int n) {
            if (n <= 1)
                return 1;
            BigInteger result = 1;
            for (int i = 1; i <= n; i++)
                result *= i;
            return result;
        }
    }
}