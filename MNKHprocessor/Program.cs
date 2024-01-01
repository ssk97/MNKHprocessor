using System;

namespace MNKHprocessor
{
    class MainProgram
    {
        static void Main(string[] args) {
            TurnData turn = TurnProcessor.ProcessTurn();
            Console.WriteLine();
            Console.WriteLine("Press \"S\" to create a spreadsheet, or \"P\" to process dice roll results.");
            bool repeat = true;
            while (repeat) {
                ConsoleKeyInfo key = Console.ReadKey();
                Console.WriteLine();
                if (key.KeyChar == 'p' || key.KeyChar == 'P') {
                    repeat = false;
                    Console.WriteLine();
                    Console.WriteLine();
                    DiceProcessing.ProcessDice(turn);
                } else if (key.KeyChar == 's' || key.KeyChar == 'S') {
                    repeat = false;
                    //https://stackoverflow.com/questions/25134024/clean-up-excel-interop-objects-with-idisposable/25135685#25135685
                    //https://stackoverflow.com/questions/17130382/understanding-garbage-collection-in-net/17131389#17131389
                    try {
                        Spreadsheets.RenderSpreadsheet(turn);
                    } finally {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                } else if (key.KeyChar == 'q' || key.KeyChar == 'Q') {
                    repeat = false;
                } else {
                    Console.WriteLine("Key " + key + " pressed, not valid. Try again (or Q to quit)!");
                }
            }
        }
    }
}
