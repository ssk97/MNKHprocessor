using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace MNKHprocessor
{
    enum ACTION_TYPES { NORMAL, DC, REFORM, NOROLL };
    [DebuggerDisplay("{name}")]
    class ActionData
    {
        public string name;
        public int rpd;
        public int prog_curr, prog_max;
        public int bonus;
        public ACTION_TYPES type;
        public int forced_dice;
        public Dictionary<string, int> indicators;
        public int position;
        public string section_name;
    }
    [DebuggerDisplay("{name}")]
    class SectionData
    {
        public string name;
        public List<ActionData> actions;
        public string dice;
        public int section_modifier;
        public int forced_count;
    }
    [DebuggerDisplay("{name}")]
    class TurnData
    {
        public List<SectionData> sections;
        public int turn;
        public int free_dice;
        public int max_resources;
        public static int crit_max = 2;
    }
    class TurnProcessor
    {
        static string turn_in = "text/turnPost.txt";
        public static string[] section_names = new string[] { "Infrastructure", "Heavy Industry", "Rocketry", "Light Industry", "Chemical Industry", "Agriculture", "Services", "Bureaucracy", "Ministry Actions" };
        public static string[] indicator_names = new string[] { "General Labor", "Educated Labor", "Electricity", "Steel", "Coal", "Non-Ferrous", "Petroleum Fuels", "Petrochemicals" };
        public static string[] indicator_short_names = new string[] { "GL", "EL", "E", "S", "C", "NF", "PF", "P" };
        static int get_global_bonus() { //Affects reform rolls too
            return 0; //Management XP
        }
        static int get_bonus(string name, SectionData section) {
            int bonus = 0;
            bonus += 10; //MNKh education
            bonus += 9; //Economics Education
            bonus += 5; //Stat Planning
            bonus += 5; //Telecomms

            bonus += get_global_bonus();
            if (section.name == "Infrastructure") {
                bonus += 5;
            }
            if (section.name == "Services") {
                bonus += 10;
            }
            bonus += section.section_modifier;
            return bonus;
        }

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
        public static ActionData FindAction(string action_name, List<SectionData> sections) {
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

        public static TurnData ProcessTurn() {
            List<SectionData> sections = new();
            //Process the turn post to collect progress info
            int section_curr = 0;
            int turn = 0;
            int max_resources = 0;
            int free_dice = 0;
            using (StreamReader sr = File.OpenText(turn_in)) {
                Regex turnCheck = new Regex(@"Turn (\d+)", RegexOptions.IgnoreCase);
                Regex resourceCounter = new Regex(@"(\+|-|Base |with )\s*(\d+)", RegexOptions.IgnoreCase);
                string s;
                while ((s = sr.ReadLine()) != null) {
                    Match turn_m = turnCheck.Match(s);
                    if (turn_m.Success) {
                        turn = Int32.Parse(turn_m.Groups[1].Value);
                        Console.WriteLine("turn " + turn);
                    } else if (s.StartsWith("Resources per Turn")) {
                        MatchCollection count_matches = resourceCounter.Matches(s);
                        foreach (Match count_m in count_matches) {
                            Console.WriteLine(count_m.Value);
                            if (count_m.Value.StartsWith('-')) {
                                max_resources -= Int32.Parse(count_m.Groups[2].Value);
                            } else {
                                max_resources += Int32.Parse(count_m.Groups[2].Value);
                            }
                        }
                        Console.WriteLine("total " + max_resources);
                    } else if (s.StartsWith("Free dice")) {
                        Regex diceCheck = new Regex(@"(\d+) Dice", RegexOptions.IgnoreCase);
                        Match dice_m = diceCheck.Match(s);
                        Debug.Assert(dice_m.Success);
                        free_dice = Int32.Parse(dice_m.Groups[1].Value);
                        break;
                    }
                }
                Regex isActionProg = new Regex(@"\[\](.*?):.*?\((\d+) Resources per Dice (\d+)/(\d+)\)", RegexOptions.IgnoreCase);
                Regex isActionDC = new Regex(@"\[\](.*?):.*DC (\d+)", RegexOptions.IgnoreCase);
                Regex isActionAny = new Regex(@"\[\](.*?):", RegexOptions.IgnoreCase);
                Regex isActionForced = new Regex(@"(.*?):.*?(\d+) Dice", RegexOptions.IgnoreCase);
                Regex checkMultiAction = new Regex(@"(\p{L}+)(/\p{L}+)+", RegexOptions.IgnoreCase);
                Regex diceCount = new Regex(@"(\d+) Dice", RegexOptions.IgnoreCase);
                Regex sectionModProgress = new Regex(@"(-?\d+)/Dice Malus", RegexOptions.IgnoreCase);
                //Regex rocketryModifier = new Regex(@"Cancel Project.*?-(\d+) Dice", RegexOptions.IgnoreCase);

                string tmp_indicator_list = string.Concat("(", String.Join("|", indicator_names), ")");
                bool inPriceSection = false;
                bool initialBureaucracy = false;
                Regex priceModifier = new Regex(@"(-?\d+) RpD ([a-zA-Z ]+)", RegexOptions.IgnoreCase);
                //language=regex
                Regex indicatorCheckDirect = new Regex(@"\W(-?\d+)(\s?CI\d+)?\s" + tmp_indicator_list + @"(?! per)", RegexOptions.IgnoreCase);
                int current_action_position = 0;
                while ((s = sr.ReadLine()) != null) {
                    if (inPriceSection) {
                        Match rpdMod = priceModifier.Match(s);
                        if (rpdMod.Success) {
                            int modifier = Int32.Parse(rpdMod.Groups[1].Value);
                            string category = rpdMod.Groups[2].Value;
                            foreach (var sect in sections) {
                                if (sect.name == category || category.Contains("Universal")) {
                                    foreach (var act in sect.actions) {
                                        if (act.type == ACTION_TYPES.NORMAL) {
                                            act.rpd += modifier;
                                            if (act.rpd < 0) { act.rpd = 0; }
                                        }
                                    }
                                }
                            }
                        }
                        continue;
                    }
                    Match m = isActionProg.Match(s);
                    ACTION_TYPES type = ACTION_TYPES.NORMAL;
                    int forced_dice = 0;
                    if (!m.Success) {//Check for other types of actions than ones with "Resources per Dice"
                        m = isActionDC.Match(s);
                        if (m.Success) {
                            type = ACTION_TYPES.DC;
                        } else {
                            m = isActionAny.Match(s);
                            if (!m.Success && initialBureaucracy) {
                                m = isActionForced.Match(s);
                                if (m.Success) {
                                    forced_dice = Int32.Parse(m.Groups[2].Value);
                                }
                            }
                            if (s.Contains("Unrolled)")) {
                                type = ACTION_TYPES.NOROLL;
                            } else {
                                type = ACTION_TYPES.REFORM;
                            }
                        }
                    }

                    if (!m.Success) {//If not an action, check if it's changing to a new section
                        foreach (string section_name in section_names) {
                            if (s.StartsWith(section_name)) {
                                Match sect_m = diceCount.Match(s);
                                string tmp_dice = "Unlimited";
                                if (sect_m.Success) {
                                    tmp_dice = sect_m.Groups[1].Value;
                                }
                                Match sect_m2 = sectionModProgress.Match(s);
                                int tmp_mod = 0;
                                if (sect_m2.Success) {
                                    tmp_mod = Int32.Parse(sect_m2.Groups[1].Value);
                                }
                                SectionData tmp = new SectionData {
                                    actions = new List<ActionData>(),
                                    name = section_name,
                                    dice = tmp_dice,
                                    section_modifier = tmp_mod,
                                };
                                sections.Add(tmp);
                                section_curr = sections.Count - 1;
                                current_action_position = 0;
                                if (section_name == "Bureaucracy") {
                                    initialBureaucracy = true;
                                }
                            }
                        }
                    } else {//Add action info into list
                            //setup the indicators
                        if (forced_dice == 0) {
                            initialBureaucracy = false;
                        } else {
                            sections[section_curr].forced_count += 1;
                        }
                        Dictionary<string, int> indicators = new();
                        MatchCollection indicatorsDirect = indicatorCheckDirect.Matches(s);
                        foreach (Match ind in indicatorsDirect) { //Indicators without half-values
                            int amount = Int32.Parse(ind.Groups[1].Value);
                            //CI = group 2
                            string name = ind.Groups[3].Value;
                            if (name == "per") break;
                            indicators.Add(name, amount);
                        }

                        Match multi_check = checkMultiAction.Match(m.Groups[1].Value);//check only in the name section
                        if (!multi_check.Success) { //Not a multi-action
                            ActionData new_action = new ActionData() {
                                name = m.Groups[1].Value,
                                type = type,
                                indicators = indicators,
                                position = current_action_position++,
                                section_name = sections[section_curr].name,
                                forced_dice = forced_dice,
                            };
                            switch (type) {
                                case ACTION_TYPES.REFORM:
                                    new_action.rpd = 0;
                                    new_action.prog_curr = 0;
                                    new_action.prog_max = 0;
                                    new_action.bonus = get_global_bonus();
                                    break;
                                case ACTION_TYPES.DC:
                                    new_action.rpd = 0;
                                    new_action.prog_curr = 0;
                                    new_action.prog_max = Int32.Parse(m.Groups[2].Value);
                                    new_action.bonus = get_bonus(new_action.name, sections[section_curr]);
                                    break;
                                case ACTION_TYPES.NORMAL:
                                    new_action.rpd = Int32.Parse(m.Groups[2].Value);
                                    new_action.prog_curr = Int32.Parse(m.Groups[3].Value);
                                    new_action.prog_max = Int32.Parse(m.Groups[4].Value);
                                    new_action.bonus = get_bonus(new_action.name, sections[section_curr]);
                                    break;
                            }
                            sections[section_curr].actions.Add(new_action);
                        } else {
                            //break down a multi-action into separate action entries
                            Debug.Assert(forced_dice == 0, "Forced dice incompatible with multi-actions");
                            List<string> subaction_names = new List<string>();
                            subaction_names.Add(multi_check.Groups[1].Value);
                            foreach (Capture cap in multi_check.Groups[2].Captures) {
                                subaction_names.Add(cap.Value.Substring(1));//remove the "/"
                            }
                            foreach (string subaction_name in subaction_names) {
                                string replaced_name = m.Groups[1].Value.Replace(multi_check.Value, subaction_name);
                                ActionData new_action = new ActionData() {
                                    name = replaced_name,
                                    type = type,
                                    indicators = indicators,
                                    position = current_action_position++,
                                    section_name = sections[section_curr].name,
                                };
                                if (type == ACTION_TYPES.NORMAL) {
                                    new_action.rpd = Int32.Parse(m.Groups[2].Value);
                                } else {
                                    new_action.rpd = 0;
                                }
                                //language=regex
                                Regex find_subaction_progress = new Regex(subaction_name + @"\W+(\d+)/(\d+)", RegexOptions.IgnoreCase);
                                Match subaction_progress = find_subaction_progress.Match(s);
                                if (subaction_progress.Success) {
                                    new_action.prog_curr = Int32.Parse(subaction_progress.Groups[1].Value);
                                    new_action.prog_max = Int32.Parse(subaction_progress.Groups[2].Value);
                                } else {
                                    if (type == ACTION_TYPES.NORMAL) {
                                        new_action.prog_curr = Int32.Parse(m.Groups[3].Value);
                                        new_action.prog_max = Int32.Parse(m.Groups[4].Value);
                                    } else { Debug.Assert(type == ACTION_TYPES.NOROLL); }
                                }
                                new_action.bonus = get_bonus(new_action.name, sections[section_curr]);
                                sections[section_curr].actions.Add(new_action);
                            }
                        }
                    }
                    if (s.StartsWith("Current Economic Prices")) {
                        inPriceSection = true;
                    }
                    if (s.StartsWith("Plan Effects: ")) {
                        break; //TODO: Add resource changes?
                    }
                }
            }
            return new TurnData { sections = sections, turn = turn, free_dice = free_dice, max_resources = max_resources };
        }
    }
}