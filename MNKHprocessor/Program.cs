using System;

namespace MNKHprocessor
{
    class MainProgram
    {
        static void Main(string[] args) {
            TurnData turn = TurnProcessor.ProcessTurn();
            Console.WriteLine("Turn "+turn.turn);
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
                    Spreadsheets.RenderSpreadsheet(turn);
                } else if (key.KeyChar == 'q' || key.KeyChar == 'Q') {
                    repeat = false;
                } else {
                    Console.WriteLine("Key " + key + " pressed, not valid. Try again (or Q to quit)!");
                }
            }
        }
    }
}
