using Maomi.TTS.Windows;
using System.CommandLine;

namespace ttscommannd;

internal class Program
{
    static async Task<int> Main(string[] args)
    {
        Console.WriteLine(string.Join(" ", args));

#if DEBUG
        Console.WriteLine("输入命令");
        var cs = Console.ReadLine();
        args = cs.Split(" ").Where(x => !string.IsNullOrEmpty(x)).ToArray();
#endif

        var printInstalledVoices = new Argument<string?>(
            name: "v",
            description: "print installed voices.")
        {
            Arity = ArgumentArity.ZeroOrOne
        };

        // 构建命令行参数
        var rootCommand = new RootCommand("请输入要执行的命令.");
        rootCommand.AddArgument(printInstalledVoices);

        // 解析参数调用
        rootCommand.SetHandler(async (printInstalledVoices) =>
        {
            try
            {
                if (!string.IsNullOrEmpty(printInstalledVoices))
                {
                    PrintVoicesList();
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }

        }, printInstalledVoices);

        return await rootCommand.InvokeAsync(args);
    }

    static void PrintVoicesList()
    {
        var installedVoices = SpeechHelper.GetInstalledVoices();
        Console.WriteLine("Available voices:");

        foreach (var voice in installedVoices)
        {
            var info = voice.VoiceInfo;
            Console.WriteLine($"Name: {info.Name}");
            Console.WriteLine($"Enable: {voice.Enabled}");
            Console.WriteLine($"Culture: {info.Culture}");
            Console.WriteLine($"Age: {info.Age}");
            Console.WriteLine($"Gender: {info.Gender}");
            Console.WriteLine($"Description: {info.Description}");
            Console.WriteLine($"ID: {info.Id}");
            Console.WriteLine(new string('-', 40));
        }
    }
}
