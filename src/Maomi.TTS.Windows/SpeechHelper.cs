using System.Globalization;
using System.Speech.AudioFormat;
using System.Speech.Synthesis;

namespace Maomi.TTS.Windows;

/// <summary>
/// Speech synthesis service.<br />
/// 语音合成服务.
/// </summary>
public static class SpeechHelper
{
    /// <summary>
    /// 播放注音.
    /// </summary>
    private const string SSML =
"""
<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='zh-CN'>
{0}
</speak>
""";

    /// <summary>
    /// 播放注音，带时间间隔.
    /// </summary>
    private const string SSMLTime =
"""
<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='zh-CN'>
    <voice name="{0}">
    {1}
    </voice>
</speak>
""";

    private static readonly SpeechSynthesizer _synth = new SpeechSynthesizer();

    /// <summary>
    /// 取消正在朗读的语音.
    /// </summary>
    public static void SpeakAsyncCancelAll()
    {
        _synth.SpeakAsyncCancelAll();
    }

    /// <summary>
    /// 获取系统中已经安装的语音风格.
    /// </summary>
    /// <returns>语音风格.</returns>
    public static IReadOnlyCollection<InstalledVoice> GetInstalledVoices()
    {
        return _synth.GetInstalledVoices();
    }

    /// <summary>
    /// 设置当前语音风格.
    /// </summary>
    /// <param name="name">语音风格名称.</param>
    public static void SelectVoice(string name)
    {
        _synth.SelectVoice(name);
    }

    /// <summary>
    /// 实时朗读文字.
    /// </summary>
    /// <param name="text">要播放的文本.</param>
    /// <param name="rate">语速.</param>
    /// <param name="volume">音量.</param>
    /// <param name="language">语言.</param>
    public static void Speak(string text, int rate = 0, int volume = 100, string language = "zh-CN")
    {
        _synth.SpeakAsyncCancelAll();

        using SpeechSynthesizer synth = new SpeechSynthesizer()
        {
            Rate = rate,
            Volume = volume >= 100 ? 100 : volume,
        };

        synth.SetOutputToDefaultAudioDevice();

        PromptBuilder promptBuilder = new()
        {
            Culture = new CultureInfo(language)
        };

        promptBuilder.AppendText(text);
        synth.Speak(promptBuilder);
    }

    /// <summary>
    /// 合成 .wav 语音文件.
    /// </summary>
    /// <param name="text">要播放的文本.</param>
    /// <param name="wavFilePath">wav 语音文件存储位置.</param>
    /// <param name="rate">语速.</param>
    /// <param name="volume">音量.</param>
    /// <param name="language">语言.</param>
    /// <param name="speechAudioFormatInfo">自定义频率、声道等.<see cref="SpeechAudioFormatInfo"/>.</param>
    public static void SpeakToWav(string text, string wavFilePath, int rate = 0, int volume = 100, string language = "zh-CN", SpeechAudioFormatInfo? speechAudioFormatInfo = null)
    {
        if (speechAudioFormatInfo == null)
        {
            speechAudioFormatInfo = new SpeechAudioFormatInfo(16000, AudioBitsPerSample.Sixteen, AudioChannel.Mono);
        }

        using SpeechSynthesizer synth = new SpeechSynthesizer()
        {
            Rate = rate,
            Volume = volume >= 100 ? 100 : volume
        };

        synth.SetOutputToWaveFile(wavFilePath, speechAudioFormatInfo);

        PromptBuilder promptBuilder = new()
        {
            Culture = new CultureInfo(language)
        };

        promptBuilder.AppendText(text);
        synth.Speak(promptBuilder);
    }

    /// <summary>
    /// 异步朗读语音.
    /// </summary>
    /// <param name="text">要播放的文本.</param>
    /// <param name="rate">语速.</param>
    /// <param name="volume">音量.</param>
    /// <param name="language">语言.</param>
    /// <returns>语音播放状态 <see cref="Prompt"/>.</returns>
    public static Prompt SpeakAsync(string text, int rate = 0, int volume = 100, string language = "zh-CN")
    {
        _synth.SpeakAsyncCancelAll();

        using SpeechSynthesizer synth = new SpeechSynthesizer()
        {
            Rate = rate,
            Volume = volume >= 100 ? 100 : volume
        };

        PromptBuilder promptBuilder = new()
        {
            Culture = new CultureInfo(language)
        };

        promptBuilder.AppendText(text);
        return synth.SpeakAsync(promptBuilder);
    }

    /// <summary>
    /// 异步朗读语音.
    /// </summary>
    /// <param name="text">要播放的文本.</param>
    /// <param name="rate">语速.</param>
    /// <param name="volume">音量.</param>
    /// <param name="language">语言.</param>
    /// <param name="token">任务状态.</param>
    /// <returns><see cref="Task"/>.</returns>
    public static async Task TaskSpeakAsync(string text, int rate = 0, int volume = 100, string language = "zh-CN", CancellationToken token = default)
    {
        _synth.SpeakAsyncCancelAll();

        using SpeechSynthesizer synth = new SpeechSynthesizer()
        {
            Rate = rate,
            Volume = volume >= 100 ? 100 : volume
        };

        synth.SetOutputToDefaultAudioDevice();

        PromptBuilder promptBuilder = new()
        {
            Culture = new CultureInfo(language)
        };
        promptBuilder.AppendText(text);

        var prompt = synth.SpeakAsync(promptBuilder);
        if (prompt.IsCompleted)
        {
            await Task.CompletedTask;
        }

        TaskCompletionSource completionSource = new TaskCompletionSource();
        _ = Task.Factory.StartNew(async () =>
        {
            try
            {
                while (!prompt.IsCompleted && !token.IsCancellationRequested)
                {
                    await Task.Delay(10).ConfigureAwait(false);
                }
            }
            finally
            {
                completionSource.TrySetResult();
            }
        });

        await completionSource.Task;
    }

    /// <summary>
    /// 异步朗读拼音(注音).
    /// </summary>
    /// <param name="pinyin">拼音、注音.</param>
    /// <param name="rate">语速.</param>
    /// <param name="volume">音量.</param>
    /// <param name="voice">声音风格.</param>
    /// <returns>语音播放状态 <see cref="Prompt"/>.</returns>
    public static Prompt SpeakSsmlAsync(string pinyin, int rate = 0, int volume = 100, string voice = "zh-CN-YunxiNeural")
    {
        const string PinyinSSML = "<phoneme  alphabet='ipa' ph='{0}'></phoneme>";
        var phoneme = string.Format(PinyinSSML, pinyin);
        string ssml = string.Format(SSML, voice, phoneme);

        _synth.SpeakAsyncCancelAll();

        using SpeechSynthesizer synth = new SpeechSynthesizer()
        {
            Rate = rate,
            Volume = volume >= 100 ? 100 : volume
        };

        synth.SetOutputToDefaultAudioDevice();

        return synth.SpeakSsmlAsync(ssml);
    }

    /// <summary>
    /// 异步朗读拼音(注音).
    /// </summary>
    /// <param name="pinyin">拼音、注音.</param>
    /// <param name="rate">语速.</param>
    /// <param name="volume">音量.</param>
    /// <param name="voice">声音风格.</param>
    /// <param name="token">状态.</param>
    /// <returns><see cref="Task"/>.</returns>
    public static async Task TaskSpeakSsmlAsync(string pinyin, int rate = 0, int volume = 100, string voice = "zh-CN-YunxiNeural", CancellationToken token = default)
    {
        const string PinyinSSML = "<phoneme  alphabet='ipa' ph='{0}'></phoneme>";
        var phoneme = string.Format(PinyinSSML, pinyin);
        string ssml = string.Format(SSML, voice, phoneme);

        _synth.SpeakAsyncCancelAll();

        using SpeechSynthesizer synth = new SpeechSynthesizer()
        {
            Rate = rate,
            Volume = volume >= 100 ? 100 : volume
        };

        synth.SetOutputToDefaultAudioDevice();

        var prompt = synth.SpeakSsmlAsync(ssml);
        if (prompt.IsCompleted)
        {
            await Task.CompletedTask;
        }

        TaskCompletionSource completionSource = new TaskCompletionSource();
        _ = Task.Factory.StartNew(async () =>
        {
            try
            {
                while (!prompt.IsCompleted && !token.IsCancellationRequested)
                {
                    await Task.Delay(1).ConfigureAwait(false);
                }
            }
            finally
            {
                completionSource.TrySetResult();
            }
        });

        await completionSource.Task;
    }

    /// <summary>
    /// 同步朗读拼音(注音).
    /// </summary>
    /// <param name="pinyin">拼音、注音.</param>
    /// <param name="rate">语速.</param>
    /// <param name="volume">音量.</param>
    /// <param name="voice">声音风格.</param>
    public static void SpeakSsml(string pinyin, int rate = 0, int volume = 100, string voice = "zh-CN-YunxiNeural")
    {
        const string PinyinSSML = "<phoneme  alphabet='ipa' ph='{0}'></phoneme>";
        var phoneme = string.Format(PinyinSSML, pinyin);
        string ssml = string.Format(SSML, voice, phoneme);

        _synth.SpeakAsyncCancelAll();

        using SpeechSynthesizer synth = new SpeechSynthesizer()
        {
            Rate = rate,
            Volume = volume >= 100 ? 100 : volume
        };

        synth.SetOutputToDefaultAudioDevice();
        synth.Speak(ssml);
    }

    /// <summary>
    /// 同步朗读语音，可定义停顿间隔.
    /// </summary>
    /// <param name="texts">多个语句.</param>
    /// <param name="pauseTime">停顿时间.</param>
    /// <param name="rate">语速.</param>
    /// <param name="volume">音量.</param>
    /// <param name="voice">声音风格.</param>
    public static void SpeakSsml(string[] texts, int pauseTime = 500, int rate = 0, int volume = 100, string voice = "zh-CN-YunxiNeural")
    {
        List<string> xml = new List<string>();
        foreach (var item in texts)
        {
            xml.Add(string.Format("<phoneme  alphabet='sapi' ph='{0} {1}'>{0}</phoneme>", item, pauseTime));
        }

        string ssml = string.Format(SSMLTime, voice, string.Join(Environment.NewLine, xml));

        _synth.SpeakAsyncCancelAll();

        using SpeechSynthesizer synth = new SpeechSynthesizer()
        {
            Rate = rate,
            Volume = volume >= 100 ? 100 : volume
        };

        synth.SetOutputToDefaultAudioDevice();
        synth.SpeakSsml(ssml);
    }

    /// <summary>
    /// 异步朗读语音，可定义停顿间隔.
    /// </summary>
    /// <param name="texts">多个语句.</param>
    /// <param name="pauseTime">停顿时间.</param>
    /// <param name="rate">语速.</param>
    /// <param name="volume">音量.</param>
    /// <param name="voice">声音风格.</param>
    /// <returns>语音播放状态 <see cref="Prompt"/>.</returns>
    public static Prompt SpeakSsmlAsync(string[] texts, int pauseTime = 500, int rate = 0, int volume = 100, string voice = "zh-CN-YunxiNeural")
    {
        List<string> xml = new List<string>();
        foreach (var item in texts)
        {
            xml.Add(string.Format("<phoneme  alphabet='sapi' ph='{0} {1}'>{0}</phoneme>", item, pauseTime));
        }

        string ssml = string.Format(SSMLTime, voice, string.Join(Environment.NewLine, xml));

        using SpeechSynthesizer synth = new SpeechSynthesizer()
        {
            Rate = rate,
            Volume = volume >= 100 ? 100 : volume
        };

        synth.SetOutputToDefaultAudioDevice();

        return synth.SpeakSsmlAsync(ssml);
    }

    /// <summary>
    /// 异步朗读语音，可定义停顿间隔.
    /// </summary>
    /// <param name="texts">多个语句.</param>
    /// <param name="pauseTime">停顿时间.</param>
    /// <param name="rate">语速.</param>
    /// <param name="volume">音量.</param>
    /// <param name="voice">声音风格.</param>
    /// <param name="token">状态.</param>
    /// <returns><see cref="Task"/>.</returns>
    public static async Task TaskSpeakSsmlAsync(string[] texts, int pauseTime = 500, int rate = 0, int volume = 100, string voice = "zh-CN-YunxiNeural", CancellationToken token = default)
    {
        List<string> xml = new List<string>();
        foreach (var item in texts)
        {
            xml.Add(string.Format("<phoneme  alphabet='sapi' ph='{0} {1}'>{0}</phoneme>", item, pauseTime));
        }

        string ssml = string.Format(SSMLTime, voice, string.Join(Environment.NewLine, xml));

        using SpeechSynthesizer synth = new SpeechSynthesizer()
        {
            Rate = rate,
            Volume = volume >= 100 ? 100 : volume
        };

        synth.SetOutputToDefaultAudioDevice();

        var prompt = synth.SpeakSsmlAsync(ssml);
        if (prompt.IsCompleted)
        {
            await Task.CompletedTask;
        }

        TaskCompletionSource completionSource = new TaskCompletionSource();
        _ = Task.Factory.StartNew(async () =>
        {
            try
            {
                while (!prompt.IsCompleted && !token.IsCancellationRequested)
                {
                    await Task.Delay(1).ConfigureAwait(false);
                }
            }
            finally
            {
                completionSource.TrySetResult();
            }
        });

        await completionSource.Task;
    }
}
