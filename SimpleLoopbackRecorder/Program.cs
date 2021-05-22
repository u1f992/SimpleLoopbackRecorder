using System;
using System.IO;
using System.Diagnostics;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace SimpleLoopbackRecorder
{
    public class Program
    {
        private const string SLR_VERSION = "1.0";
        private const string SLR_USEAGE = "Usage:\n\tslr {-o | --output} <filename> [Options]\nOptions:\n\t-o, --output <filename>\tspecify output\n\t-d, --device <id>\tspecify device\n\t-l, --list\t\tprint device list\n\t-v, --version\t\tprint version\n\t-h, --help\t\tprint help";

        static void Main(string[] args)
        {
            string outputFilePath = "";
            MMDevice _MMDevice = null;
            if (!AnalyzeArgs(args, ref outputFilePath, ref _MMDevice))
            {
                if (outputFilePath == "")
                {
                    // output指定がないor不正
                    // デバイス指定が不正
                    Console.WriteLine("\n{0}\n", SLR_USEAGE);
                    Environment.Exit(1);
                }
                else
                {
                    // version,help
                    Environment.Exit(0);
                }
            }

            var capture = new WasapiLoopbackCapture(_MMDevice);
            var writer = new WaveFileWriter(outputFilePath, capture.WaveFormat);
            capture.DataAvailable += (s, a) =>
            {
                writer.Write(a.Buffer, 0, a.BytesRecorded);
            };
            capture.RecordingStopped += (s, a) =>
            {
                writer.Dispose();
                writer = null;
                capture.Dispose();
            };

            // 何も再生していない時間も録音する
            var silence = new WasapiOut();
            silence.Init(new SilenceProvider(capture.WaveFormat).ToSampleProvider());
            silence.Play();

            var stopwatch = new Stopwatch();

            stopwatch.Start();
            capture.StartRecording();
            while (capture.CaptureState != NAudio.CoreAudioApi.CaptureState.Stopped)
            {
                Console.Write("\r{0:hh\\:mm\\:ss\\.fff}", stopwatch.Elapsed);

                // キー入力で録音終了
                if (Console.KeyAvailable) capture.StopRecording();
            }
            stopwatch.Stop();
            silence.Stop();

            // キー入力を破棄
            Console.ReadKey(true);
            Console.WriteLine("\r{0:hh\\:mm\\:ss\\.fff}\n\"{1}\" has been saved.", stopwatch.Elapsed, outputFilePath);
        }

        /// <summary>
        /// 引数を処理する
        /// </summary>
        /// <param name="args">コマンドライン引数</param>
        /// <param name="outputFilePath"></param>
        /// <returns>
        /// 正常 outptFilePath != "",
        /// 中断 ret == false
        /// </returns>
        private static bool AnalyzeArgs(string[] args, ref string outputFilePath, ref MMDevice _MMDevice)
        {
            string ret = "";
            _MMDevice = null;

            // -o, -dは一度のみ検出する
            bool flag_output, flag_device;
            flag_output = false;
            flag_device = false;

            for (int i = 0; i < args.Length; i++)
            {
                var tmp = args[i];
                if ((tmp == "-o" || tmp == "--output") && i != args.Length - 1 && !flag_output)
                {
                    // 出力ファイルを指定
                    i++;
                    ret = args[i];
                    if (!ValidatePath(ref ret) || args[i].Substring(0, 1) == "-")
                    {
                        Console.Error.WriteLine("Error: Invalid filename.");
                        ret = "";
                        return false;
                    }
                    flag_output = true;
                }
                else if ((tmp == "-d" || tmp == "--device") && i != args.Length - 1 && !flag_device)
                {
                    // 使用するデバイスを指定
                    i++;
                    var deviceID = args[i];
                    _MMDevice = GetDeviceFromID(deviceID);
                    if (_MMDevice == null)
                    {
                        Console.Error.WriteLine("Error: Invalid device ID.");
                        ret = "";
                        return false;
                    }
                    flag_device = true;
                }
                else if (((tmp == "-o" || tmp == "--output" || tmp == "-d" || tmp == "--device") && i == args.Length - 1) ||
                            (ret == "" && args.Length > 2 && i == args.Length - 1) ||
                                (!(tmp == "-l" || tmp == "--list" || tmp == "-v" || tmp == "--version" || tmp == "-h" || tmp == "--help") && args.Length == 1))
                {
                    // 直後に値が必要な引数が最後にある
                    // 2つ以上の引数がある場合に、最後の引数なのにoutputが指定されていない
                    // -l, -v, -h以外の引数が単独で渡されている
                    Console.Error.WriteLine("Error: Invalid arguments.");
                    ret = "";
                    return false;
                }
                else if (tmp == "-l" || tmp == "--list")
                {
                    // デバイス一覧を表示
                    Console.WriteLine(GenerateDeviceList());
                    outputFilePath = "?list";
                    return false;
                }
                else if (tmp == "-v" || tmp == "--version")
                {
                    // バージョン情報
                    Console.WriteLine("SimpleLoopbackRecorder version {0}", SLR_VERSION);
                    outputFilePath = "?version";
                    return false;
                }
                else if (tmp == "-h" || tmp == "--help")
                {
                    // ヘルプ
                    Console.WriteLine(SLR_USEAGE);
                    Console.WriteLine(GenerateDeviceList());
                    outputFilePath = "?help";
                    return false;
                }
            }

            if (!flag_device)
            {
                var enumerator = new MMDeviceEnumerator();
                _MMDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            }

            outputFilePath = ret;
            return true;
        }

        /// <summary>
        /// <a href="https://stackoverflow.com/questions/3067479/determine-via-c-sharp-whether-a-string-is-a-valid-file-path">
        /// パス文字列の検証
        /// </a>
        /// </summary>
        /// <param name="path">検証するパス</param>
        /// <returns></returns>
        private static bool ValidatePath(ref string path)
        {
            string CurDir = Environment.CurrentDirectory;

            // 無効な文字が含まれていないか
            if (path.IndexOfAny(Path.GetInvalidPathChars()) == -1)
            {
                try
                {
                    // 相対パスであればカレントディレクトリを結合する
                    if (!Path.IsPathRooted(path))
                    {
                        path = Path.Combine(CurDir, path);
                    }
                    var stream = File.Create(path);
                    stream.Close();

                    // Exceptions from FileInfo Constructor:
                    //   System.ArgumentNullException:
                    //     fileName is null.
                    //
                    //   System.Security.SecurityException:
                    //     The caller does not have the required permission.
                    //
                    //   System.ArgumentException:
                    //     The file name is empty, contains only white spaces, or contains invalid
                    //     characters.
                    //
                    //   System.IO.PathTooLongException:
                    //     The specified path, file name, or both exceed the system-defined maximum
                    //     length. For example, on Windows-based platforms, paths must be less than
                    //     248 characters, and file names must be less than 260 characters.
                    //
                    //   System.NotSupportedException:
                    //     fileName contains a colon (:) in the middle of the string.
                    FileInfo fileInfo = new FileInfo(path);

                    // Exceptions using FileInfo.Length:
                    //   System.IO.IOException:
                    //     System.IO.FileSystemInfo.Refresh() cannot update the state of the file or
                    //     directory.
                    //
                    //   System.IO.FileNotFoundException:
                    //     The file does not exist.-or- The Length property is called for a directory.
                    bool throwEx = fileInfo.Length == -1;

                    // Exceptions using FileInfo.IsReadOnly:
                    //   System.UnauthorizedAccessException:
                    //     Access to fileName is denied.
                    //     The file described by the current System.IO.FileInfo object is read-only.
                    //     -or- This operation is not supported on the current platform.
                    //     -or- The caller does not have the required permission.
                    throwEx = fileInfo.IsReadOnly;

                    return true;
                }
                catch (ArgumentNullException)
                {
                    //   System.ArgumentNullException:
                    //     fileName is null.
                }
                catch (System.Security.SecurityException)
                {
                    //   System.Security.SecurityException:
                    //     The caller does not have the required permission.
                }
                catch (ArgumentException)
                {
                    //   System.ArgumentException:
                    //     The file name is empty, contains only white spaces, or contains invalid
                    //     characters.
                }
                catch (UnauthorizedAccessException)
                {
                    //   System.UnauthorizedAccessException:
                    //     Access to fileName is denied.
                }
                catch (PathTooLongException)
                {
                    //   System.IO.PathTooLongException:
                    //     The specified path, file name, or both exceed the system-defined maximum
                    //     length. For example, on Windows-based platforms, paths must be less than
                    //     248 characters, and file names must be less than 260 characters.
                }
                catch (NotSupportedException)
                {
                    //   System.NotSupportedException:
                    //     fileName contains a colon (:) in the middle of the string.
                }
                catch (FileNotFoundException)
                {
                    // System.FileNotFoundException
                    //  The exception that is thrown when an attempt to access a file that does not
                    //  exist on disk fails.
                }
                catch (IOException)
                {
                    //   System.IO.IOException:
                    //     An I/O error occurred while opening the file.
                }
                catch (Exception)
                {
                    // Unknown Exception. Might be due to wrong case or nulll checks.
                }
            }
            else
            {
                // Path contains invalid characters
            }
            return false;
        }

        private static string GenerateDeviceList()
        {
            string ret = "Device List:\n\tFriendlyName :\tID\n";

            var enumerator = new MMDeviceEnumerator();
            foreach (var wasapi in enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                ret += $"\t\"{wasapi.FriendlyName}\" : {wasapi.ID}\n";
            }

            return ret;
        }

        private static MMDevice GetDeviceFromID(string deviceID)
        {
            var enumerator = new MMDeviceEnumerator();
            foreach (var wasapi in enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                if (wasapi.ID == deviceID) return wasapi;
            }
            return null;
        }
    }
}
