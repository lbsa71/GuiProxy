using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace lbsa71.GuiProxy
{
    internal static class WinFormExtensions
    {
        /// <summary>
        ///  Helper extension method to make calling lambdas on UI thread and wait for completion simpler.
        /// </summary>
        /// <param name="control">What control should govern the invoke</param>
        /// <param name="action">What to do</param>
        public static Task Invoke(this Control control, Action action)
        {
            TaskCompletionSource<bool> taskCompletionSource = new TaskCompletionSource<bool>();

            Action innerAction = () =>
            {
                action();

                taskCompletionSource.SetResult(true);
            };

            control.Invoke(innerAction);

            return taskCompletionSource.Task;
        }
    }

    public class GuiProxy
    {
        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        private const uint BM_CLICK = 0x00F5;
        private const uint WM_SETTEXT = 0x000C;
        private const uint WM_GETTEXT = 0xD;
        private const uint WM_GETTEXTLENGTH = 0xE;

        private static readonly string processName = "DummyApp";
        private Form marshallingForm;
        private IntPtr targetApplicationWindowHandle;
        private List<IntPtr> children;
        private IEnumerable<string> childClassnames;
        private IEnumerable<string> childTexts;

        private void SetText(int index, string str)
        {
            WithControl(index, () =>
            {
                var control = children[index];

                SendMessage(control, WM_SETTEXT, IntPtr.Zero, str);
            });
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern bool SendMessage(IntPtr hWnd, uint msg, int wParam, StringBuilder lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint msg,
            IntPtr wParam, string lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, IntPtr processId);

        [DllImport("user32.dll")]
        private static extern IntPtr AttachThreadInput(IntPtr idAttach,
            IntPtr idAttachTo, bool fAttach);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        private static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            var gch = GCHandle.FromIntPtr(pointer);
            var list = gch.Target as List<IntPtr>;

            if (list == null)
            {
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
            }

            list.Add(handle);
            //  You can modify this to check to see if you want to cancel the operation, then return a null here
            return true;
        }

        public static List<IntPtr> GetChildWindows(IntPtr parent)
        {
            var result = new List<IntPtr>();
            var listHandle = GCHandle.Alloc(result);
            try
            {
                var childProc = new EnumWindowsProc(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                {
                    listHandle.Free();
                }
            }
            return result;
        }

        private string GetText(int index)
        {
            return GetText(children[index]);
        }

        private string GetText(IntPtr control)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder(255); // or length from call with GETTEXTLENGTH

            marshallingForm.Invoke(() =>
            {
                var result = SendMessage(
                    control,
                    WM_GETTEXT,
                    sb.Capacity,
                    sb);
            }).Wait();

            return sb.ToString();
        }

        private void WithControl(int index, Action action)
        {
            var control = children[index];

            marshallingForm.Invoke(() =>
            {
                SetForegroundWindow(targetApplicationWindowHandle);

                var activeWindowThread = GetWindowThreadProcessId(targetApplicationWindowHandle, IntPtr.Zero);

                var thisWindowThread = GetWindowThreadProcessId(marshallingForm.Handle, IntPtr.Zero);

                AttachThreadInput(activeWindowThread, thisWindowThread, true);

                SetFocus(control);

                AttachThreadInput(activeWindowThread, thisWindowThread, false);

                action();

            }).Wait();
        }

        private void PressButton(int i)
        {
            var buttonToClick = children[i];

            SendMessage(buttonToClick, BM_CLICK, (IntPtr)0, (IntPtr)0);
        }

        private static void InvokeSendKeys(string str)
        {
            SendKeys.SendWait(str);
        }

        public bool Start()
        {
            KillAllTargetProcesses();

            var process = Process.Start(processName + ".exe");

            SleepUntil(TimeSpan.FromSeconds(5), () =>
            {
                targetApplicationWindowHandle = process.MainWindowHandle;

                return targetApplicationWindowHandle != (IntPtr) 0;
            });

            TaskCompletionSource<bool> taskCompletionSource = new TaskCompletionSource<bool>();

            Task.Run(() =>
            {
                marshallingForm = new Form();
                marshallingForm.Show();

                taskCompletionSource.SetResult(true);

                Application.Run(marshallingForm);
            });

            taskCompletionSource.Task.Wait();

            marshallingForm.Invoke(() =>
            {
                children = GetChildWindows(targetApplicationWindowHandle);

                childClassnames = children.Select(l =>
                {
                    var sb = new StringBuilder(256);

                    GetClassName(l, sb, 256);

                    return sb.ToString();
                });

                childTexts = children.Select(GetText);

            }).Wait();

            return true;
        }

        private bool SleepUntil(TimeSpan timeout, Func<bool> func)
        {
            var cancelBy = DateTime.Now + timeout;

            do
            {
                var completed = func();

                if (completed)
                {
                    return true;
                }
                else
                {
                    Thread.Sleep(100);
                }

            } while (DateTime.Now < cancelBy);

            return false;
        }

        public override bool OnStop()
        {
            KillAllTargetProcesses();


            marshallingForm.Invoke(() => { marshallingForm.Hide(); });
            marshallingForm = null;

            return true;
        }

        private static void KillAllTargetProcesses()
        {
            var processes = Process.GetProcessesByName(processName);
            if (processes.Any())
            {
                foreach (var process in processes)
                {
                    process.Kill();
                }
            }
        }
    }
}
