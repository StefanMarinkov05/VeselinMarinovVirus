using System;
using System.ComponentModel.Design;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;

namespace VeselinMarinovVirus
{
    [Command(PackageIds.MyCommand)]
    internal sealed class MyCommand : BaseCommand<MyCommand>
    {
        private static bool isRunning = false;
        private static CancellationTokenSource? cts;

        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            if (isRunning)
            {
                Stop();
                await VS.MessageBox.ShowAsync("VeselinMarinovVirus", "Error 420: No more cannabis!");
                return;
            }

            var docView = await VS.Documents.GetActiveDocumentViewAsync();
            if (docView?.TextView == null)
            {
                await VS.MessageBox.ShowWarningAsync("VeselinMarinovVirus", "Отвори файл първо!");
                return;
            }

            var textBuffer = docView.TextBuffer;
            string originalText = textBuffer.CurrentSnapshot.GetText();
            var lines = originalText.Split('\n');
            int maxLength = 0;
            foreach (var line in lines)
                maxLength = Math.Max(maxLength, line.Length);

            char[,] grid = new char[lines.Length, maxLength];
            for (int i = 0; i < lines.Length; i++)
            {
                var padded = lines[i].PadRight(maxLength, ' ');
                for (int j = 0; j < padded.Length; j++)
                {
                    grid[i, j] = padded[j];
                }
            }

            cts = new CancellationTokenSource();
            isRunning = true;

            _ = Task.Run(async () =>
            {
                try
                {
                    while (!cts.IsCancellationRequested)
                    {
                        bool moved = StepWithGravity(grid);

                        string updatedText = GridToString(grid);
                        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                        textBuffer.Replace(
                            new Span(0, textBuffer.CurrentSnapshot.Length),
                            updatedText
                        );

                        if (!moved)
                        {
                            // Stop if nothing moved = fully settled
                            Stop();
                            break;
                        }

                        await Task.Delay(50, cts.Token);
                    }
                }
                catch (OperationCanceledException)
                {
                    // expected
                }
            });
        }

        private static bool StepWithGravity(char[,] grid)
        {
            int rows = grid.GetLength(0);
            int cols = grid.GetLength(1);
            bool moved = false;
            var random = new Random();

            // Gravity pass: from bottom-1 to top
            for (int i = rows - 2; i >= 0; i--)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (grid[i, j] != ' ' && grid[i + 1, j] == ' ')
                    {
                        grid[i + 1, j] = grid[i, j];
                        grid[i, j] = ' ';
                        moved = true;
                    }
                    // Add small chance to slide left/right
                    else if (grid[i, j] != ' ')
                    {
                        int dir = random.Next(2) == 0 ? -1 : 1;
                        int newJ = j + dir;

                        if (newJ >= 0 && newJ < cols && grid[i + 1, newJ] == ' ')
                        {
                            grid[i + 1, newJ] = grid[i, j];
                            grid[i, j] = ' ';
                            moved = true;
                        }
                    }
                }
            }

            return moved;
        }

        private static string GridToString(char[,] grid)
        {
            var sb = new StringBuilder();
            int rows = grid.GetLength(0);
            int cols = grid.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    sb.Append(grid[i, j]);
                }
                if (i != rows - 1)
                    sb.Append('\n');
            }

            return sb.ToString();
        }

        public static void Stop()
        {
            if (isRunning)
            {
                cts?.Cancel();
                isRunning = false;
            }
        }
    }

    [Command(PackageIds.StopCommand)]
    internal sealed class StopCommand : BaseCommand<StopCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            MyCommand.Stop();
            //await VS.MessageBox.ShowAsync("VeselinMarinovVirus", "Error 421: Like 420 but with 1 more joint!");
        }
    }
}
