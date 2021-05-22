using System;
using System.IO;
using System.Diagnostics;
using System.Threading;
using NAudio.Wave;

namespace SimpleLoopbackRecorder
{
    public class Program
    {
        private const string SLR_VERSION = "0.1";
        private const string SLR_USEAGE = "Usage:\n\tslr {-o | --output} <filename> [Options]\nOptions:\n\t-o, --output\tspecify output\n\t-v, --version\tprint version\n\t-h, --help\tprint help";

        static void Main(string[] args)
        {
            string outputFilePath = "";
            if (!AnalyzeArgs(args, ref outputFilePath))
            {
                if (outputFilePath == "")
                {
                    // output指定がないor不正
                    Console.Error.WriteLine("\n{0}", SLR_USEAGE);
                    Environment.Exit(1);
                }
                else
                {
                    // version,help
                    Environment.Exit(0);
                }
            }

            var capture = new WasapiLoopbackCapture();
            var writer = new WaveFileWriter(outputFilePath, capture.WaveFormat);

            // 何も再生していない時間も録音するため
            var silence = new WaveOutEvent();
            silence.Init(new SilenceProvider(capture.WaveFormat).ToSampleProvider());
            silence.Play();

            var stopwatch = new Stopwatch();

            capture.DataAvailable += (s, a) =>
            {
                writer.Write(a.Buffer, 0, a.BytesRecorded);
                if (writer.Position > capture.WaveFormat.AverageBytesPerSecond * 5)
                {
                    capture.StopRecording();
                }
            };
            capture.RecordingStopped += (s, a) =>
            {
                writer.Dispose();
                writer = null;
                capture.Dispose();
            };

            stopwatch.Start();
            capture.StartRecording();
            while (capture.CaptureState != NAudio.CoreAudioApi.CaptureState.Stopped)
            {
                Thread.Sleep(1);

                Console.Write("\r\t{0}", stopwatch.Elapsed);
                
            }
            stopwatch.Stop();
            silence.Stop();

            Console.WriteLine("\r\t{0}\n\"{1}\" has been saved.", stopwatch.Elapsed, outputFilePath);
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
        private static bool AnalyzeArgs(string[] args, ref string outputFilePath)
        {
            string ret = "";

            for (int i = 0; i < args.Length; i++)
            {
                var tmp = args[i];
                if ((tmp == "-o" || tmp == "--output") && i != args.Length - 1)
                {
                    i++;
                    ret = args[i];
                    if (!ValidatePath(ref ret))
                    {
                        ret = "";
                    }
                }
                else if (tmp == "-h" || tmp == "--help")
                {
                    Console.WriteLine(SLR_USEAGE);
                    outputFilePath = "?help";
                    return false;
                }
                else if (tmp == "-v" || tmp == "--version")
                {
                    Console.WriteLine("SimpleLoopbackRecorder version {0}", SLR_VERSION);
                    outputFilePath = "?version";
                    return false;
                }
            }

            if (ret == "")
            {
                Console.Error.WriteLine("Error: Invalid designation.");
                return false;
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
    }
}
